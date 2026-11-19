# -*- coding: utf-8 -*-
"""preview.py — 把 .pptx 近似渲染成 PNG，用于"没有 PowerPoint/LibreOffice 时用眼睛验收排版"。

为什么需要这个文件
------------------
目标环境可能没有 LibreOffice / WPS 命令行、没有 PIL/numpy/ffmpeg，
PPTX 无法转图片。2026-10 的课程 PPT 任务里因此手写过一个 530 行渲染器，
在 43/合并稿上逐页目视检查，定位了"行距 0.1 倍导致叠字""版式与页面双图标"
等只有看图才能发现的问题。本模块是那次实现的产品化版本：

* 数据一律来自 `read.py` 的只读 IR（Deck / Slide / Shape / Para / Run），不重复解析结构；
* 单位一律按 deck 自身尺寸换算，字号只经 `units.pt_of_sz`，不在本模块留第二套 EMU 常量；
* 纯标准库 + fontTools：PNG 用 zlib 解码，文字取字形轮廓后自己扫描线光栅化；
* JPEG 不解码（不内置解码器），按框比例画浅灰占位 + 左下角标注 "JPEG WxH"——
  这足以判断"这里有没有图、图占多大、和图注是否对齐"，而不会假装画出来。

能力边界（有意为之，写在这里防止误用）
--------------------------------------
已支持：纯色背景、图片背景（PNG）、矩形/圆角矩形/椭圆/线/四向箭头、图片（PNG 解码，
含透明通道与 srcRect 裁切）、文本框（真实字形、粗体、行距 spc_ratio、对齐 l/ctr/r、
垂直锚点 t/ctr/b、按框宽换行）、版式/母版继承的图片（校徽就属于这一类）。
绘制所需的非几何属性全部来自 `read.py` 的 IR 字段（anchor / insets / line_w /
flip_h / flip_v / crop / Slide.bg_image），本模块不重复解析 XML 结构。
不支持（近似或跳过，结果仍可用于排版检查）：渐变填充、阴影/发光、形状旋转、
半透明填充（solidFill 的 a:alpha 未进 IR）、wrap="none"、艺术字、
表格（本语料 graphicFrame 为 0）、JPEG 解码、EMF/WMF 矢量图；
run 未写 sz 且形状 lstStyle 也没有时按 18pt 兜底。
继承来的**占位符**（"单击此处编辑母版标题样式"）一律不画——它们只是提示框。

用法
----
    import preview
    pngs = preview.render("deck-b.pptx", "out/preview", pages=[1, 3], width=1400)
    preview.contact_sheet(pngs, "out/preview/overview.png", cols=4)
    # CLI
    python preview.py deck-b.pptx out/preview
    python preview.py deck-b.pptx out/preview --pages 1-4,7 --width 1400 --sheet overview.png
    python preview.py --self-test            # 渲染 deck-b.pptx 前 2 页并自检
"""
import math
import os
import struct
import sys
import time
import zipfile
import zlib

_ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
if _ENGINE_DIR not in sys.path:
    sys.path.insert(0, _ENGINE_DIR)

try:                       # 支持 `import preview` 与 `from engine import preview` 两种用法
    from . import units, read
except ImportError:        # pragma: no cover
    import units
    import read

try:
    from fontTools.ttLib import TTFont
    from fontTools.pens.recordingPen import RecordingPen
except ImportError as _exc:  # pragma: no cover - 只在没装 fontTools 的机器上触发
    TTFont = None
    RecordingPen = None
    _FONTTOOLS_ERROR = str(_exc)
else:
    _FONTTOOLS_ERROR = ""


# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------
#: 默认字体：鸿蒙系统的方正黑体（与 2026-10 参考实现一致）。MAX_PPT_FONT 可覆盖，
#: 想分别指定常规/粗体时用 os.pathsep（":"）分隔两个路径。
DEFAULT_FONT_REGULAR = "/system/fonts/FZHeiT-SC-Regular.ttf"
DEFAULT_FONT_BOLD = "/system/fonts/FZHeiT-SC-Bold.ttf"
FONT_ENV = "MAX_PPT_FONT"

#: 画布宽度（像素）；高度按 deck 纵横比推出来，不拉伸。
DEFAULT_WIDTH = 1400

#: 主题色的固定近似表（Office 默认主题）。PPTX 的真值在 theme1.xml + clrMap 里，
#: 这里不解析主题，按约定用固定表近似：accent1..6 / dk1 dk2 lt1 lt2 / tx1 tx2 bg1 bg2。
SCHEME_COLORS = {
    "accent1": (0x44, 0x72, 0xC4),
    "accent2": (0xED, 0x7D, 0x31),
    "accent3": (0xA5, 0xA5, 0xA5),
    "accent4": (0xFF, 0xC0, 0x00),
    "accent5": (0x5B, 0x9B, 0xD5),
    "accent6": (0x70, 0xAD, 0x47),
    "dk1": (0x00, 0x00, 0x00),
    "lt1": (0xFF, 0xFF, 0xFF),
    "dk2": (0x44, 0x54, 0x6A),
    "lt2": (0xE7, 0xE6, 0xE6),
    "tx1": (0x00, 0x00, 0x00),
    "tx2": (0x44, 0x54, 0x6A),
    "bg1": (0xFF, 0xFF, 0xFF),
    "bg2": (0xE7, 0xE6, 0xE6),
    "hlink": (0x05, 0x63, 0xC1),
    "folHlink": (0x95, 0x4F, 0x72),
}
DEFAULT_TEXT_COLOR = (0x1A, 0x1A, 0x1A)      # 无显式颜色时的正文色

#: 渲染用行高（em 倍数）。units.LINE_HEIGHT=1.5 是"会不会溢出"的保守估算，
#: 这里要的是"看起来像不像 WPS"：WPS 微软雅黑实测单倍行高约 1.32em。
LINE_EM = 1.32

#: 图片占位符配色
PLACEHOLDER_BG = (0xE6, 0xE6, 0xE6)
PLACEHOLDER_LINE = (0xBF, 0xBF, 0xBF)
PLACEHOLDER_TEXT = (0x80, 0x80, 0x80)

_PNG_SIG = b"\x89PNG\r\n\x1a\n"
_SOF_MARKERS = frozenset((0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                          0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF))

#: 超过这个像素数就不解 PNG（纯 Python 逐字节反滤波太慢），退回占位符。
#: 5,000x1,000 = 5M 的横幅图可以解；再大就只在图上标尺寸。
MAX_PNG_PIXELS = int(os.environ.get("MAX_PPT_PNG_MAX_PIXELS", "8000000"))

#: 字形轮廓离散精度（每段贝塞尔取几个点）
_CURVE_STEPS = 8
#: 字形超采样倍数：2 表示 2x2 采样后平均成 alpha，得到抗锯齿边缘。
_GLYPH_SS = 2


# --------------------------------------------------------------------------
# 颜色
# --------------------------------------------------------------------------
def parse_color(value, default=None):
    """把 IR 里的颜色字符串转成 (r, g, b)。

    '191B1F' → (0x19, 0x1B, 0x1F)；'scheme:accent1' → 固定近似表；
    空值 / 无法解析 → default。
    """
    if not value:
        return default
    if value.startswith("scheme:"):
        return SCHEME_COLORS.get(value[7:], default)
    v = value.strip().lstrip("#")
    if len(v) == 6:
        try:
            return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
        except ValueError:
            return default
    if len(v) == 3:
        try:
            return tuple(int(c * 2, 16) for c in v)  # type: ignore[return-value]
        except ValueError:
            return default
    return default


# --------------------------------------------------------------------------
# 字体与字形光栅化
# --------------------------------------------------------------------------
_FONT_CACHE = {}      # bold(bool) -> dict
_GLYPH_CACHE = {}     # (ch, round(px*4), bold) -> (w, h, x0, y0, mask, adv) | None
_GLYPH_CACHE_LIMIT = 8000


def _font_paths():
    """解析常规/粗体字体路径；MAX_PPT_FONT 可覆盖（os.pathsep 分隔两个路径）。"""
    regular, bold = DEFAULT_FONT_REGULAR, DEFAULT_FONT_BOLD
    env = os.environ.get(FONT_ENV) or ""
    if env:
        parts = [p for p in env.split(os.pathsep) if p]
        if parts:
            regular = parts[0]
            if len(parts) > 1:
                bold = parts[1]
            elif "Bold" in regular or "bold" in regular:
                bold = regular
            else:
                # 只给一个字体时粗体仍尝试默认粗体；不存在则退回同一字体
                bold = DEFAULT_FONT_BOLD
    return regular, bold


def _font(bold):
    """加载并缓存字体信息（fontTools TTFont、cmap、字形集、hmtx、上伸部比例）。"""
    key = 1 if bold else 0
    info = _FONT_CACHE.get(key)
    if info is not None:
        return info
    if TTFont is None:
        raise RuntimeError(
            "preview.py 需要 fontTools 做字形光栅化，但 import 失败（%s）。"
            "请先 `pip install fonttools`。" % _FONTTOOLS_ERROR)
    regular, bold_path = _font_paths()
    path = bold_path if bold else regular
    if not os.path.exists(path):
        fallback = regular if os.path.exists(regular) else DEFAULT_FONT_REGULAR
        if not os.path.exists(fallback):
            raise RuntimeError(
                "找不到预览字体：%s。可用环境变量 %s 指定一个中文 TTF 字体路径。"
                % (path, FONT_ENV))
        path = fallback
    font = TTFont(path)
    upm = font["head"].unitsPerEm
    hhea = font["hhea"]
    info = {
        "path": path,
        "font": font,
        "glyphset": font.getGlyphSet(),
        "cmap": font.getBestCmap(),
        "hmtx": font["hmtx"],
        "upm": float(upm),
        # 基线：从行顶往下 ascent 处画基线；CJK 字体 ascent/upm ≈ 0.9
        "ascent": float(hhea.ascent) / float(upm),
        "descent": -float(hhea.descent) / float(upm),
    }
    _FONT_CACHE[key] = info
    return info


def _flatten_pen(pen_value):
    """RecordingPen 的操作序列 → 多边形列表（三次/二次贝塞尔按 _CURVE_STEPS 离散）。

    与 2026-10 参考实现同算法：二次曲线的控制点序列按 TrueType 约定取中点拆段。
    """
    polys, cur, pos = [], [], None
    for op, args in pen_value:
        if op == "moveTo":
            if cur:
                polys.append(cur)
            pos = args[0]
            cur = [pos]
        elif op == "lineTo":
            pos = args[0]
            cur.append(pos)
        elif op == "qCurveTo":
            pts = list(args)
            if pts and pts[-1] is None:      # TrueType 闭合轮廓
                pts = pts[:-1]
            if not pts:
                continue
            endp = pts[-1]
            ctrls = pts[:-1]
            prev = pos if pos is not None else ctrls[0]
            if not ctrls:
                cur.append(endp)
            else:
                for i, cp in enumerate(ctrls):
                    if i < len(ctrls) - 1:
                        nxt = ((cp[0] + ctrls[i + 1][0]) / 2.0,
                               (cp[1] + ctrls[i + 1][1]) / 2.0)
                    else:
                        nxt = endp
                    for t in range(1, _CURVE_STEPS + 1):
                        tt = t / float(_CURVE_STEPS)
                        mt = 1.0 - tt
                        cur.append((mt * mt * prev[0] + 2 * mt * tt * cp[0] + tt * tt * nxt[0],
                                    mt * mt * prev[1] + 2 * mt * tt * cp[1] + tt * tt * nxt[1]))
                    prev = nxt
                if prev != endp:
                    cur.append(endp)
            pos = endp
        elif op == "curveTo":
            c1, c2, endp = args[-3], args[-2], args[-1]
            prev = pos if pos is not None else c1
            for t in range(1, _CURVE_STEPS + 1):
                tt = t / float(_CURVE_STEPS)
                mt = 1.0 - tt
                cur.append((mt ** 3 * prev[0] + 3 * mt * mt * tt * c1[0]
                            + 3 * mt * tt * tt * c2[0] + tt ** 3 * endp[0],
                            mt ** 3 * prev[1] + 3 * mt * mt * tt * c1[1]
                            + 3 * mt * tt * tt * c2[1] + tt ** 3 * endp[1]))
            pos = endp
        elif op == "closePath":
            if cur:
                polys.append(cur)
                cur = []
    if cur:
        polys.append(cur)
    return polys


def _raster_mask(polys, ss):
    """把（已缩放到像素、y 已翻转的）多边形光栅化成 ss 倍超采样掩码。

    返回 (w, h, x0, y0, mask)；mask 为 bytearray，0/1 两值，尺寸是输出像素的 ss 倍。
    奇偶规则填充：字形里的孔洞（如"回"字内框）靠这个挖出来。
    """
    xs = [p[0] for c in polys for p in c]
    ys = [p[1] for c in polys for p in c]
    if not xs:
        return None
    x0 = int(math.floor(min(xs))) - 1
    y0 = int(math.floor(min(ys))) - 1
    w = max(1, int(math.ceil(max(xs))) + 1 - x0)
    h = max(1, int(math.ceil(max(ys))) + 1 - y0)
    sw, sh = w * ss, h * ss
    edges = []
    for contour in polys:
        for i in range(len(contour)):
            ax, ay = contour[i - 1]
            bx, by = contour[i]
            ax = (ax - x0) * ss
            ay = (ay - y0) * ss
            bx = (bx - x0) * ss
            by = (by - y0) * ss
            if ay == by:
                continue
            edges.append((ax, ay, bx, by))
    if not edges:
        return None
    mask = bytearray(sw * sh)
    ymin = max(0, int(math.floor(min(min(e[1], e[3]) for e in edges))))
    ymax = min(sh - 1, int(math.ceil(max(max(e[1], e[3]) for e in edges))))
    for y in range(ymin, ymax + 1):
        yc = y + 0.5
        crossings = []
        for ex0, ey0, ex1, ey1 in edges:
            if (ey0 <= yc < ey1) or (ey1 <= yc < ey0):
                crossings.append(ex0 + (yc - ey0) * (ex1 - ex0) / (ey1 - ey0))
        if len(crossings) < 2:
            continue
        crossings.sort()
        base = y * sw
        for i in range(0, len(crossings) - 1, 2):
            xa = max(0, int(math.ceil(crossings[i])))
            xb = min(sw - 1, int(math.floor(crossings[i + 1])))
            if xb >= xa:
                mask[base + xa:base + xb + 1] = b"\x01" * (xb - xa + 1)
    return (w, h, x0, y0, mask)


def _glyph(ch, px, bold):
    """取一个字符在指定像素字号下的字形位图（带缓存）。

    返回 (w, h, x0, y0, alpha_mask, advance_px)：
      * (x0, y0) 是字形包围盒相对基线的整数偏移；
      * alpha_mask 已是 1x 分辨率（_GLYPH_SS 超采样后平均），用于抗锯齿混色。
    无该字形时返回 None（调用方按 advance 留空）。
    """
    px = max(1.0, float(px))
    key = (ch, int(round(px * 4)), 1 if bold else 0)
    if key in _GLYPH_CACHE:
        return _GLYPH_CACHE[key]
    if len(_GLYPH_CACHE) > _GLYPH_CACHE_LIMIT:      # 防跨文档长期占用内存
        _GLYPH_CACHE.clear()
    info = _font(bold)
    gname = info["cmap"].get(ord(ch))
    if gname is None:
        _GLYPH_CACHE[key] = None
        return None
    try:
        glyph = info["glyphset"][gname]
        pen = RecordingPen()
        glyph.draw(pen)
        polys = _flatten_pen(pen.value)
    except Exception:
        _GLYPH_CACHE[key] = None
        return None
    if not polys:
        _GLYPH_CACHE[key] = None
        return None
    scale = px / info["upm"]
    scaled = [[(x * scale, -y * scale) for (x, y) in contour] for contour in polys]
    raw = _raster_mask(scaled, _GLYPH_SS)
    if raw is None:
        _GLYPH_CACHE[key] = None
        return None
    w, h, x0, y0, sub = raw
    ss = _GLYPH_SS
    sw = w * ss
    mask = bytearray(w * h)
    if ss == 2:
        for y in range(h):
            row = (y * 2) * sw          # 超采样图里第 2y 行的起点
            out = y * w
            for x in range(w):
                i = row + x * 2
                mask[out + x] = (sub[i] + sub[i + 1]
                                 + sub[i + sw] + sub[i + sw + 1]) * (255 // 4)
    else:  # pragma: no cover - 只支持 1/2 两种采样率
        for y in range(h):
            for x in range(w):
                mask[y * w + x] = 255 if sub[(y * sw) + x] else 0
    try:
        adv = info["hmtx"][gname][0] * scale
    except Exception:
        adv = px * 0.5
    result = (w, h, x0, y0, mask, adv)
    _GLYPH_CACHE[key] = result
    return result


def _char_advance(ch, px, bold):
    """字符前进宽度（像素）。空格按字体表；缺字按 0.5em 兜底。"""
    g = _glyph(ch, px, bold)
    if g is not None:
        return g[5]
    if ch == "\u3000":
        return px * 0.5
    if ch in (" ", "\t"):
        return px * 0.30
    return px * 0.5


# --------------------------------------------------------------------------
# 画布（R/G/B 三个平面，便于切片级像素搬运）
# --------------------------------------------------------------------------
class Canvas:
    """RGB 画布。三个通道分开存成 bytearray：矩形填充/图片缩放都能用切片赋值走 C 循环。"""

    def __init__(self, w, h, bg=(255, 255, 255)):
        self.w = int(w)
        self.h = int(h)
        self.p = [bytearray([bg[c]]) * (self.w * self.h) for c in range(3)]

    # ---------- 填充 ----------
    def fill_rect(self, x, y, w, h, color):
        """填充不透明矩形（切片赋值，速度只和行数有关）。"""
        if w <= 0 or h <= 0:
            return
        x0, x1 = _int_span(x, x + w, self.w)
        y0, y1 = _int_span(y, y + h, self.h)
        if x1 <= x0 or y1 <= y0:
            return
        n = x1 - x0
        spans = [bytes([color[c]]) * n for c in range(3)]
        for yy in range(y0, y1):
            base = yy * self.w + x0
            self.p[0][base:base + n] = spans[0]
            self.p[1][base:base + n] = spans[1]
            self.p[2][base:base + n] = spans[2]

    def fill_row(self, y, x0, x1, color):
        """填一行（x0/x1 已取整）；椭圆/圆角矩形按行调用。"""
        if y < 0 or y >= self.h:
            return
        x0 = max(0, int(x0))
        x1 = min(self.w, int(x1))
        n = x1 - x0
        if n <= 0:
            return
        base = y * self.w + x0
        self.p[0][base:base + n] = bytes([color[0]]) * n
        self.p[1][base:base + n] = bytes([color[1]]) * n
        self.p[2][base:base + n] = bytes([color[2]]) * n

    def fill_ellipse(self, x, y, w, h, color):
        """椭圆填充（逐行算左右边界，避免整体超采样）。"""
        if w <= 0 or h <= 0:
            return
        cx, cy = x + w / 2.0, y + h / 2.0
        rx, ry = w / 2.0, h / 2.0
        y0, y1 = _int_span(y, y + h, self.h)
        for yy in range(y0, y1):
            dy = (yy + 0.5 - cy) / ry
            if abs(dy) > 1.0:
                continue
            dx = rx * math.sqrt(max(0.0, 1.0 - dy * dy))
            self.fill_row(yy, cx - dx, cx + dx, color)

    def fill_rounded_rect(self, x, y, w, h, color, radius=None):
        """圆角矩形填充；radius 缺省用 roundRect 的 adj=16667（半径 = 1/6 短边）。"""
        if w <= 0 or h <= 0:
            return
        if radius is None:
            radius = min(w, h) / 6.0
        r = max(0.0, min(float(radius), w / 2.0, h / 2.0))
        if r <= 0.5:
            self.fill_rect(x, y, w, h, color)
            return
        y0, y1 = _int_span(y, y + h, self.h)
        top = y + r
        bot = y + h - r
        for yy in range(y0, y1):
            yc = yy + 0.5
            if yc < top:
                dy = top - yc
            elif yc > bot:
                dy = yc - bot
            else:
                dy = 0.0
            if dy <= 0.0:
                self.fill_row(yy, x, x + w, color)
            elif dy <= r:
                inset = r - math.sqrt(max(0.0, r * r - dy * dy))
                self.fill_row(yy, x + inset, x + w - inset, color)

    def fill_polygon(self, pts, color):
        """扫描线（奇偶规则）填充多边形，用于箭头等预设形状。"""
        if len(pts) < 3:
            return
        edges = []
        for i in range(len(pts)):
            ax, ay = pts[i - 1]
            bx, by = pts[i]
            if ay == by:
                continue
            edges.append((ax, ay, bx, by))
        if not edges:
            return
        ymin = max(0, int(math.floor(min(min(e[1], e[3]) for e in edges))))
        ymax = min(self.h - 1, int(math.ceil(max(max(e[1], e[3]) for e in edges))))
        for yy in range(ymin, ymax + 1):
            yc = yy + 0.5
            xs = []
            for ex0, ey0, ex1, ey1 in edges:
                if (ey0 <= yc < ey1) or (ey1 <= yc < ey0):
                    xs.append(ex0 + (yc - ey0) * (ex1 - ex0) / (ey1 - ey0))
            if len(xs) < 2:
                continue
            xs.sort()
            for i in range(0, len(xs) - 1, 2):
                self.fill_row(yy, xs[i], xs[i + 1], color)

    # ---------- 线 ----------
    def line(self, x0, y0, x1, y1, color, width=1.0):
        """画一条有宽度的线段（用四点多边形）。"""
        dx, dy = x1 - x0, y1 - y0
        length = math.hypot(dx, dy)
        w = max(0.8, float(width))
        if length < 1e-6:
            self.fill_rect(x0 - w / 2, y0 - w / 2, w, w, color)
            return
        nx, ny = -dy / length * w / 2.0, dx / length * w / 2.0
        self.fill_polygon([(x0 + nx, y0 + ny), (x1 + nx, y1 + ny),
                           (x1 - nx, y1 - ny), (x0 - nx, y0 - ny)], color)

    def rect_outline(self, x, y, w, h, color, width=1.0):
        """矩形描边（四条粗线，比逐像素描边省事）。"""
        if w <= 0 or h <= 0:
            return
        t = max(0.8, float(width))
        self.fill_rect(x, y, w, t, color)
        self.fill_rect(x, y + h - t, w, t, color)
        self.fill_rect(x, y, t, h, color)
        self.fill_rect(x + w - t, y, t, h, color)

    def polyline(self, pts, color, width=1.0, closed=False):
        n = len(pts)
        if n < 2:
            return
        for i in range(1, n):
            self.line(pts[i - 1][0], pts[i - 1][1], pts[i][0], pts[i][1], color, width)
        if closed:
            self.line(pts[-1][0], pts[-1][1], pts[0][0], pts[0][1], color, width)

    # ---------- 图片 ----------
    def blit(self, x, y, w, h, info, crop=None):
        """把解码好的图片贴进 (x, y, w, h) 框；最近邻缩放，支持 alpha 与裁剪。

        info 是 _media_info 的返回值：planes=(R, G, B[, A])，每平面 len(w*h)。
        """
        iw, ih = info["w"], info["h"]
        planes = info.get("planes")
        if not planes or w <= 0 or h <= 0:
            return
        sx0, sy0, sx1, sy1 = 0, 0, iw, ih
        if crop:
            l, t, r, b = crop
            sx0 = max(0, int(iw * l / 100000.0))
            sy0 = max(0, int(ih * t / 100000.0))
            sx1 = min(iw, iw - int(iw * r / 100000.0))
            sy1 = min(ih, ih - int(ih * b / 100000.0))
        cw = max(1, sx1 - sx0)
        ch = max(1, sy1 - sy0)
        x0i = int(round(x))
        y0i = int(round(y))
        wpx = max(1, int(round(w)))
        hpx = max(1, int(round(h)))
        # x 方向源列索引预计算（每次贴图只算一次）
        xs = None if cw == wpx else [sx0 + (dx * cw) // wpx for dx in range(wpx)]
        alpha = planes[3] if len(planes) > 3 else None
        pr, pg, pb = planes[0], planes[1], planes[2]
        xa = max(0, -x0i)
        xe = min(wpx, self.w - x0i)
        if xe <= xa:
            return
        if alpha is None:
            for dy in range(hpx):
                cy = y0i + dy
                if cy < 0 or cy >= self.h:
                    continue
                sy = sy0 + (dy * ch) // hpx
                base = sy * iw
                n = xe - xa
                dst = cy * self.w + x0i + xa
                if xs is None:
                    src = base + sx0 + xa
                    self.p[0][dst:dst + n] = pr[src:src + n]
                    self.p[1][dst:dst + n] = pg[src:src + n]
                    self.p[2][dst:dst + n] = pb[src:src + n]
                else:
                    idx = [base + i for i in xs[xa:xe]]
                    self.p[0][dst:dst + n] = bytes([pr[i] for i in idx])
                    self.p[1][dst:dst + n] = bytes([pg[i] for i in idx])
                    self.p[2][dst:dst + n] = bytes([pb[i] for i in idx])
        else:
            for dy in range(hpx):
                cy = y0i + dy
                if cy < 0 or cy >= self.h:
                    continue
                sy = sy0 + (dy * ch) // hpx
                base = sy * iw
                row = cy * self.w
                for dx in range(xa, xe):
                    si = base + (xs[dx] if xs is not None else sx0 + dx)
                    a = alpha[si]
                    if not a:
                        continue
                    di = row + x0i + dx
                    if a == 255:
                        self.p[0][di] = pr[si]
                        self.p[1][di] = pg[si]
                        self.p[2][di] = pb[si]
                    else:
                        inv = 255 - a
                        self.p[0][di] = (self.p[0][di] * inv + pr[si] * a) // 255
                        self.p[1][di] = (self.p[1][di] * inv + pg[si] * a) // 255
                        self.p[2][di] = (self.p[2][di] * inv + pb[si] * a) // 255

    # ---------- 文字 ----------
    def glyph(self, g, ox, oy, color):
        """把字形位图（_glyph 的返回值）画到 (ox, oy)：oy 是基线 y，ox 是笔位 x。"""
        gw, gh, gx0, gy0, mask = g[0], g[1], g[2], g[3], g[4]
        px0 = int(round(ox + gx0))
        py0 = int(round(oy + gy0))
        if py0 >= self.h or py0 + gh <= 0 or px0 >= self.w or px0 + gw <= 0:
            return
        cr, cg, cb = color
        xa = max(0, -px0)
        xe = min(gw, self.w - px0)
        if xe <= xa:
            return
        for yy in range(gh):
            cy = py0 + yy
            if cy < 0 or cy >= self.h:
                continue
            row = cy * self.w + px0
            mi = yy * gw
            for xx in range(xa, xe):
                a = mask[mi + xx]
                if not a:
                    continue
                di = row + xx
                if a == 255:
                    self.p[0][di] = cr
                    self.p[1][di] = cg
                    self.p[2][di] = cb
                else:
                    inv = 255 - a
                    self.p[0][di] = (self.p[0][di] * inv + cr * a) // 255
                    self.p[1][di] = (self.p[1][di] * inv + cg * a) // 255
                    self.p[2][di] = (self.p[2][di] * inv + cb * a) // 255

    # ---------- 输出 ----------
    def png(self, path):
        """编码为 8bit RGB PNG（filter=0；纯标准库 zlib）。"""
        chunks = []
        for y in range(self.h):
            row = bytearray(self.w * 3)
            base = y * self.w
            row[0::3] = self.p[0][base:base + self.w]
            row[1::3] = self.p[1][base:base + self.w]
            row[2::3] = self.p[2][base:base + self.w]
            chunks.append(b"\x00" + bytes(row))
        raw = b"".join(chunks)
        del chunks

        def chunk(tag, data):
            body = tag + data
            return struct.pack(">I", len(data)) + body + struct.pack(
                ">I", zlib.crc32(body) & 0xFFFFFFFF)

        blob = _PNG_SIG
        blob += chunk(b"IHDR", struct.pack(">IIBBBBB", self.w, self.h, 8, 2, 0, 0, 0))
        blob += chunk(b"IDAT", zlib.compress(raw, 6))
        blob += chunk(b"IEND", b"")
        with open(path, "wb") as fh:
            fh.write(blob)
        return path

    def nonwhite_ratio(self, tolerance=12):
        """非背景像素占比；自测用来判断"不是空白页"。"""
        total = self.w * self.h
        if not total:
            return 0.0
        count = 0
        pr, pg, pb = self.p
        for i in range(total):
            if (abs(pr[i] - 255) > tolerance or abs(pg[i] - 255) > tolerance
                    or abs(pb[i] - 255) > tolerance):
                count += 1
        return count / float(total)


def _int_span(a, b, limit):
    """浮点区间 → [max(0, round(a)), min(limit, round(b))) 的整数区间。"""
    return max(0, int(round(a))), min(limit, int(round(b)))


# --------------------------------------------------------------------------
# PNG 解码 / JPEG 尺寸 / 媒体信息
# --------------------------------------------------------------------------
_MASKS = {}


def _add_mod256(a, b):
    """两个等长 bytes 逐字节 (a+b) mod 256 —— PNG filter 2(Up) 用。

    用大整数 SWAR 做，避免逐字节 Python 循环：
      lo = (x & 0x7f..) + (y & 0x7f..)        # 每字节低 7 位相加，最多 0xFE，不跨字节进位
      z  = (lo & 0x7f..) | ((lo ^ (x&0x80..) ^ (y&0x80..)) & 0x80..)
    """
    n = len(a)
    if n == 0:
        return b""
    masks = _MASKS.get(n)
    if masks is None:
        masks = (int.from_bytes(b"\x7f" * n, "big"),
                 int.from_bytes(b"\x80" * n, "big"))
        _MASKS[n] = masks
    mask_low, mask_high = masks
    x = int.from_bytes(a, "big")
    y = int.from_bytes(b, "big")
    lo = (x & mask_low) + (y & mask_low)
    z = (lo & mask_low) | ((lo ^ (x & mask_high) ^ (y & mask_high)) & mask_high)
    return z.to_bytes(n, "big")


def _png_chunks(data):
    """解析 PNG 头与数据块；返回 dict（不反滤波）。"""
    pos = 8
    out = {"w": 0, "h": 0, "bd": 8, "ct": 0, "interlace": 0,
           "idat": bytearray(), "plte": None, "trns": None, "error": ""}
    n = len(data)
    while pos + 8 <= n:
        length = struct.unpack(">I", data[pos:pos + 4])[0]
        tag = data[pos + 4:pos + 8]
        body = data[pos + 8:pos + 8 + length]
        pos += 12 + length
        if tag == b"IHDR":
            if len(body) < 13:
                out["error"] = "IHDR 长度异常"
                return out
            out["w"], out["h"], out["bd"], out["ct"], _comp, _filt, out["interlace"] = \
                struct.unpack(">IIBBBBB", body[:13])
        elif tag == b"IDAT":
            out["idat"] += body
        elif tag == b"PLTE":
            out["plte"] = body
        elif tag == b"tRNS":
            out["trns"] = body
        elif tag == b"IEND":
            break
    return out


def _png_dims(data):
    """只读 IHDR 里的像素尺寸（不解码）。"""
    if len(data) >= 24 and data[12:16] == b"IHDR":
        return struct.unpack(">II", data[16:24])
    c = _png_chunks(data)
    return (c["w"], c["h"]) if c["w"] else None


def _jpeg_dims(data):
    """从 JPEG 的 SOF 段读像素尺寸；读不出返回 None。"""
    i, n = 2, len(data)
    while i + 4 <= n:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker == 0xFF:
            i += 1
            continue
        if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
            i += 2
            continue
        seg = struct.unpack(">H", data[i + 2:i + 4])[0]
        if marker in _SOF_MARKERS and i + 9 <= n:
            h, w = struct.unpack(">HH", data[i + 5:i + 9])
            return w, h
        i += 2 + seg
    return None


def _decode_png(data):
    """PNG → planes。返回 dict(kind, w, h, planes, error)。

    planes = (R, G, B[, A])，每平面 len(w*h) 的 bytes；带 alpha 时 4 个平面。
    只支持 8bit、非隔行；其余情况把尺寸带上、error 写清楚，由调用方画占位符。
    """
    c = _png_chunks(data)
    w, h, bd, ct = c["w"], c["h"], c["bd"], c["ct"]
    info = {"kind": "png", "w": w, "h": h, "planes": None, "error": ""}
    if not w or not h:
        info["error"] = "IHDR 缺失"
        return info
    if w * h > MAX_PNG_PIXELS:
        info["error"] = "像素数 %d 超过解码上限 %d" % (w * h, MAX_PNG_PIXELS)
        return info
    if bd != 8:
        info["error"] = "位深 %d（仅支持 8bit）" % bd
        return info
    if c["interlace"]:
        info["error"] = "隔行扫描（Adam7）"
        return info
    if ct not in (0, 2, 3, 4, 6):
        info["error"] = "颜色类型 %d 不支持" % ct
        return info
    bpp = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ct]
    try:
        raw = zlib.decompress(bytes(c["idat"]))
    except Exception as exc:
        info["error"] = "zlib 解压失败：%s" % exc
        return info
    stride = w * bpp
    need = h * (stride + 1)
    if len(raw) < need:
        info["error"] = "IDAT 数据不足（%d < %d）" % (len(raw), need)
        return info
    out = bytearray(h * stride)
    prev = bytes(stride)
    pos = 0
    try:
        for y in range(h):
            ft = raw[pos]
            pos += 1
            seg = raw[pos:pos + stride]
            pos += stride
            if ft == 0:
                line = seg
            elif ft == 2:
                line = _add_mod256(seg, prev) if y else seg
            elif ft == 1:
                line = bytearray(seg)
                for i in range(bpp, stride):
                    line[i] = (line[i] + line[i - bpp]) & 0xFF
            elif ft == 3:
                line = bytearray(seg)
                for i in range(bpp):
                    line[i] = (line[i] + (prev[i] >> 1)) & 0xFF
                for i in range(bpp, stride):
                    line[i] = (line[i] + ((line[i - bpp] + prev[i]) >> 1)) & 0xFF
            elif ft == 4:
                line = bytearray(seg)
                for i in range(bpp):
                    line[i] = (line[i] + prev[i]) & 0xFF
                for i in range(bpp, stride):
                    a = line[i - bpp]
                    b = prev[i]
                    cc = prev[i - bpp]
                    p = a + b - cc
                    pa = abs(p - a)
                    pb = abs(p - b)
                    pc = abs(p - cc)
                    if pa <= pb and pa <= pc:
                        pr = a
                    elif pb <= pc:
                        pr = b
                    else:
                        pr = cc
                    line[i] = (line[i] + pr) & 0xFF
            else:
                info["error"] = "filter 类型 %d 非法" % ft
                return info
            out[y * stride:(y + 1) * stride] = line
            prev = bytes(line)
    except Exception as exc:  # pragma: no cover - 防坏文件把渲染整体打断
        info["error"] = "反滤波失败：%s" % exc
        return info

    if ct == 2:                      # RGB
        info["planes"] = (bytes(out[0::3]), bytes(out[1::3]), bytes(out[2::3]))
    elif ct == 6:                    # RGBA
        info["planes"] = (bytes(out[0::4]), bytes(out[1::4]),
                          bytes(out[2::4]), bytes(out[3::4]))
    elif ct == 0:                    # 灰度
        g = bytes(out)
        info["planes"] = (g, g, g)
    elif ct == 4:                    # 灰度 + alpha
        g = bytes(out[0::2])
        info["planes"] = (g, g, g, bytes(out[1::2]))
    else:                            # 调色板（ct=3）
        plte = c["plte"]
        if not plte or len(plte) < 3:
            info["error"] = "调色板缺失"
            return info
        pal_r = plte[0::3] + bytes(256 - len(plte[0::3]))
        pal_g = plte[1::3] + bytes(256 - len(plte[1::3]))
        pal_b = plte[2::3] + bytes(256 - len(plte[2::3]))
        data_idx = bytes(out)
        planes = [data_idx.translate(pal_r),
                  data_idx.translate(pal_g),
                  data_idx.translate(pal_b)]
        if c["trns"]:
            alpha_tbl = c["trns"] + bytes(256 - len(c["trns"]))
            planes.append(data_idx.translate(alpha_tbl))
        info["planes"] = tuple(planes)
    if info["planes"]:
        n = w * h
        if len(info["planes"][0]) != n:
            info["planes"] = None
            info["error"] = "解码后像素数不符"
    return info


def _media_info(z, part, cache):
    """读并缓存一个媒体部件：PNG 解码、JPEG 只读尺寸、其他类型只记类型。"""
    if part in cache:
        return cache[part]
    try:
        data = z.read(part)
    except KeyError:
        info = {"kind": "missing", "w": None, "h": None, "planes": None,
                "error": "媒体部件不存在"}
        cache[part] = info
        return info
    if data[:8] == _PNG_SIG:
        info = _decode_png(data)
    elif data[:3] == b"\xff\xd8\xff":
        dims = _jpeg_dims(data)
        info = {"kind": "jpeg", "w": dims[0] if dims else None,
                "h": dims[1] if dims else None, "planes": None, "error": ""}
    else:
        info = {"kind": "other", "w": None, "h": None, "planes": None, "error": ""}
    info["part"] = part
    info["bytes"] = len(data)
    cache[part] = info
    return info


# --------------------------------------------------------------------------
# 文字排版与绘制
# --------------------------------------------------------------------------
#: 默认正文内边距（EMU）：PowerPoint 默认左右 0.1in、上下 0.05in。
#: read.Shape.insets 在 bodyPr 缺失时才落到 (0,0,0,0)，这里只做兜底。
_DEF_LR_INSET = units.inch(0.1)
_DEF_TB_INSET = units.inch(0.05)


def _shape_insets(shape):
    ins = getattr(shape, "insets", None) or (0, 0, 0, 0)
    l, t, r, b = ins
    if l == 0 and r == 0 and t == 0 and b == 0 and shape.paras:
        # bodyPr 没写内边距时 read 已经给了 PowerPoint 默认值；全 0 只可能是没有 txBody
        return (_DEF_LR_INSET, _DEF_TB_INSET, _DEF_LR_INSET, _DEF_TB_INSET)
    return (l, t, r, b)


def _run_pt(run):
    """run 的字号（pt）：显式 sz，否则 units 的默认 18pt。"""
    return units.pt_of_sz(run.sz if run.sz else None)


def _ppi_of(ppt_scale):
    """像素/EMU → 像素/英寸（1 英寸 = units.EMU_PER_INCH）。"""
    return ppt_scale * units.EMU_PER_INCH


def _text_lines(shape, ppt_scale):
    """把 shape 的段落排成行（每行是字符列表 [(ch, px, color, bold, extra)]）。

    返回 blocks = [{"lines": [...], "align": "l|ctr|r", "default_px": px, "spc_ratio": r}]
    """
    li, _ti, ri, _bi = _shape_insets(shape)
    box_w = max(1.0, shape.w * ppt_scale - (li + ri) * ppt_scale)
    ppi = _ppi_of(ppt_scale)
    blocks = []
    for para in shape.paras:
        chars = []
        for run in para.runs:
            text = run.text or ""
            if not text:
                continue
            px = _run_pt(run) * ppi / 72.0
            color = parse_color(run.color, DEFAULT_TEXT_COLOR)
            bold = bool(run.bold)
            if para.level:
                chars.append((" ", px, color, bold, para.level * px * 0.5))
            if para.has_bullet and text.strip():
                chars.append(("\u2022", px, color, bold, px * 0.62))
                chars.append((" ", px, color, bold, 0.0))
            for ch in text:
                chars.append((ch, px, color, bold, 0.0))
        if not chars:
            blocks.append({"lines": [[]], "align": para.align or "l",
                           "default_px": 18.0 * ppi / 72.0,
                           "spc_ratio": max(0.1, para.spc_ratio or 1.0)})
            continue
        lines = []
        cur = []
        cur_w = 0.0
        for item in chars:
            ch, px, _color, bold, extra = item
            adv = _char_advance(ch, px, bold) + extra
            if cur and cur_w + adv > box_w:
                lines.append(cur)
                cur, cur_w = [], 0.0
                if ch == " ":
                    continue
                adv = _char_advance(ch, px, bold)
            cur.append(item)
            cur_w += adv
        if cur:
            lines.append(cur)
        if not lines:
            lines = [[]]
        max_px = max((it[1] for ln in lines for it in ln), default=18.0 * ppi / 72.0)
        blocks.append({"lines": lines, "align": para.align or "l",
                       "default_px": max_px,
                       "spc_ratio": max(0.1, para.spc_ratio or 1.0)})
    return blocks


def draw_text(canvas, shape, ppt_scale):
    """在 shape 的框里排版并绘制文字（真实字形光栅化）。"""
    if not shape.paras:
        return
    blocks = _text_lines(shape, ppt_scale)
    if not blocks:
        return
    li, ti, ri, bi = _shape_insets(shape)
    li *= ppt_scale
    ri *= ppt_scale
    ti *= ppt_scale
    bi *= ppt_scale
    box_w = max(1.0, shape.w * ppt_scale - (li + ri))
    x_base = shape.x * ppt_scale + li
    y_base = shape.y * ppt_scale

    # 先算总高（每行行高 = max 字号 * LINE_EM * spc_ratio）
    line_infos = []      # (chars, line_h_px, max_px, align)
    total_h = 0.0
    for blk in blocks:
        ratio = blk["spc_ratio"]
        for ln in blk["lines"]:
            max_px = max((it[1] for it in ln), default=blk["default_px"])
            line_h = max_px * LINE_EM * ratio
            line_infos.append((ln, line_h, max_px, blk["align"]))
            total_h += line_h
    anchor = getattr(shape, "anchor", "") or "t"
    if anchor == "ctr":
        y = y_base + (shape.h * ppt_scale - total_h) / 2.0
    elif anchor == "b":
        y = y_base + shape.h * ppt_scale - bi - total_h
    else:
        y = y_base + ti
    ascent_ratio = _font(False)["ascent"]
    for chars, line_h, max_px, align in line_infos:
        width = 0.0
        for ch, px, _color, bold, extra in chars:
            width += _char_advance(ch, px, bold) + extra
        if align == "ctr":
            x = x_base + (box_w - width) / 2.0
        elif align == "r":
            x = x_base + box_w - width
        else:
            x = x_base
        baseline = y + ascent_ratio * max_px
        for ch, px, color, bold, extra in chars:
            x += extra
            g = _glyph(ch, px, bold)
            if g is None:
                x += _char_advance(ch, px, bold)
                continue
            canvas.glyph(g, x, baseline, color)
            x += g[5]
        y += line_h


def draw_label(canvas, text, x, y, px, color):
    """画一行不换行的标签（占位符标注、拼图页码用）；y 是行顶。"""
    ascent = _font(False)["ascent"]
    baseline = y + ascent * px
    cx = float(x)
    for ch in text:
        g = _glyph(ch, px, False)
        if g is None:
            cx += _char_advance(ch, px, False)
            continue
        canvas.glyph(g, cx, baseline, color)
        cx += g[5]
    return cx


# --------------------------------------------------------------------------
# 形状绘制
# --------------------------------------------------------------------------
_ARROWS = frozenset(("rightArrow", "leftArrow", "upArrow", "downArrow"))


def _arrow_points(geom, x, y, w, h, flip_h, flip_v):
    """四向箭头的多边形点列（OOXML 预设几何默认 adj1=adj2=50000）。"""
    ss = min(w, h)
    head = min(ss * 0.5, min(w, h) * 0.5)
    shaft = min(ss * 0.5, ss * 0.9)
    if geom == "rightArrow":
        x1 = x + w - head
        pts = [(x, y + h / 2 - shaft / 2), (x1, y + h / 2 - shaft / 2), (x1, y),
               (x + w, y + h / 2), (x1, y + h), (x1, y + h / 2 + shaft / 2),
               (x, y + h / 2 + shaft / 2)]
    elif geom == "leftArrow":
        x1 = x + head
        pts = [(x + w, y + h / 2 - shaft / 2), (x1, y + h / 2 - shaft / 2), (x1, y),
               (x, y + h / 2), (x1, y + h), (x1, y + h / 2 + shaft / 2),
               (x + w, y + h / 2 + shaft / 2)]
    elif geom == "downArrow":
        y1 = y + h - head
        pts = [(x + w / 2 - shaft / 2, y), (x + w / 2 - shaft / 2, y1),
               (x, y1), (x + w / 2, y + h), (x + w, y1),
               (x + w / 2 + shaft / 2, y1), (x + w / 2 + shaft / 2, y)]
    else:  # upArrow
        y1 = y + head
        pts = [(x + w / 2 - shaft / 2, y + h), (x + w / 2 - shaft / 2, y1),
               (x, y1), (x + w / 2, y), (x + w, y1),
               (x + w / 2 + shaft / 2, y1), (x + w / 2 + shaft / 2, y + h)]
    cx, cy = x + w / 2.0, y + h / 2.0
    if flip_h:
        pts = [(2 * cx - px, py) for (px, py) in pts]
    if flip_v:
        pts = [(px, 2 * cy - py) for (px, py) in pts]
    return pts


def _ellipse_outline_points(x, y, w, h, steps=48):
    return [(x + w / 2.0 + (w / 2.0) * math.cos(2 * math.pi * i / steps),
             y + h / 2.0 + (h / 2.0) * math.sin(2 * math.pi * i / steps))
            for i in range(steps)]


def _line_width_px(shape, ppt_scale):
    """线宽像素：a:ln/@w（EMU）→ 像素；未写时按 1pt。"""
    w = getattr(shape, "line_w", 0) or 0
    if not w:
        w = units.pt_emu(1)
    return max(0.9, w * ppt_scale)


def _placeholder(canvas, x, y, w, h, label):
    """浅灰图片占位：框 + 边框 + 左下角标注（JPEG 不解码时用）。"""
    if w < 2 or h < 2:
        return
    canvas.fill_rect(x, y, w, h, PLACEHOLDER_BG)
    canvas.rect_outline(x, y, w, h, PLACEHOLDER_LINE, 1.0)
    if not label or h < 16 or w < 24:
        return
    px = max(10.0, min(13.0, h * 0.12))
    text_w = sum(_char_advance(c, px, False) for c in label)
    if text_w > w - 8:
        px = max(8.0, px * (w - 8) / max(1.0, text_w))
    draw_label(canvas, label, x + 5, y + h - px * 1.55, px, PLACEHOLDER_TEXT)


def _media_label(info):
    """占位符文案：'JPEG 1260x1072' / 'PNG 5000x1000' / 'IMG'。"""
    iw, ih = info.get("w"), info.get("h")
    dims = (" %dx%d" % (iw, ih)) if iw and ih else ""
    kind = info.get("kind")
    if kind == "jpeg":
        return "JPEG" + dims
    if kind == "png":
        return "PNG" + dims
    if kind == "missing":
        return "无图片"
    return "IMG" + dims


def _draw_picture(canvas, ctx, shape, x, y, w, h):
    """画图片：PNG 解码上屏；JPEG/超大/解码失败按占位 + 尺寸标注。"""
    if not shape.media:
        _placeholder(canvas, x, y, w, h, "IMG?")
        return
    info = _media_info(ctx.z, shape.media, ctx.media)
    if info.get("planes"):
        canvas.blit(x, y, w, h, info, getattr(shape, "crop", ()) or None)
        return
    _placeholder(canvas, x, y, w, h, _media_label(info))


def draw_shape(canvas, ctx, shape, ppt_scale):
    """绘制页面上的一个形状（矩形/圆角矩形/椭圆/线/箭头/图片 + 文字）。"""
    if shape.x is None or shape.y is None or shape.w is None or shape.h is None:
        return
    x = shape.x * ppt_scale
    y = shape.y * ppt_scale
    w = shape.w * ppt_scale
    h = shape.h * ppt_scale
    if w <= 0 or h <= 0:
        return
    geom = (shape.geom or "rect").strip()

    if shape.kind == "pic":
        _draw_picture(canvas, ctx, shape, x, y, w, h)
        return

    # 填充/描边：spPr 显式值优先；没有显式值时用 p:style 的主题引用兜底
    # （read.Style 的两个字段只在"没有显式填充/线条"时才非空，正是 OOXML 的 fallback 语义）。
    fill = parse_color(shape.fill) or parse_color(shape.style_fill)
    line_color = parse_color(shape.line) or parse_color(shape.style_line)
    lw = _line_width_px(shape, ppt_scale) if line_color is not None else 0.0

    if shape.kind == "cxnSp" or geom == "line":
        color = line_color or fill or (0xA6, 0xA6, 0xA6)
        x0, y0, x1, y1 = x, y, x + w, y + h
        if shape.flip_h:
            x0, x1 = x1, x0
        if shape.flip_v:
            y0, y1 = y1, y0
        canvas.line(x0, y0, x1, y1, color, lw if lw else _line_width_px(shape, ppt_scale))
        draw_text(canvas, shape, ppt_scale)
        return

    if geom in _ARROWS:
        pts = _arrow_points(geom, x, y, w, h, shape.flip_h, shape.flip_v)
        if fill is not None:
            canvas.fill_polygon(pts, fill)
        if line_color is not None:
            canvas.polyline(pts, line_color, lw, closed=True)
    elif geom == "ellipse":
        if fill is not None:
            canvas.fill_ellipse(x, y, w, h, fill)
        if line_color is not None:
            canvas.polyline(_ellipse_outline_points(x + lw / 2, y + lw / 2,
                                                    max(0.0, w - lw), max(0.0, h - lw)),
                            line_color, lw, closed=True)
    elif geom in ("roundRect", "round2SameRect", "round1Rect"):
        radius = min(w, h) / 6.0
        if fill is not None:
            if line_color is not None and lw > 1.2:
                canvas.fill_rounded_rect(x, y, w, h, line_color, radius + lw / 2)
                canvas.fill_rounded_rect(x + lw, y + lw, max(1.0, w - 2 * lw),
                                         max(1.0, h - 2 * lw), fill,
                                         max(0.0, radius - lw / 2))
            else:
                canvas.fill_rounded_rect(x, y, w, h, fill, radius)
        elif line_color is not None:
            pts = []
            steps = 6
            for cx, cy, a0 in ((x + w - radius, y + radius, -math.pi / 2),
                               (x + w - radius, y + h - radius, 0.0),
                               (x + radius, y + h - radius, math.pi / 2),
                               (x + radius, y + radius, math.pi)):
                for i in range(steps + 1):
                    a = a0 + (math.pi / 2) * i / steps
                    pts.append((cx + radius * math.cos(a), cy + radius * math.sin(a)))
            canvas.polyline(pts, line_color, lw, closed=True)
    else:  # rect 及未知预设几何一律按矩形近似
        if fill is not None:
            canvas.fill_rect(x, y, w, h, fill)
            if line_color is not None:
                canvas.rect_outline(x, y, w, h, line_color, lw)
        elif line_color is not None:
            canvas.rect_outline(x, y, w, h, line_color, lw)

    draw_text(canvas, shape, ppt_scale)


# --------------------------------------------------------------------------
# 背景
# --------------------------------------------------------------------------
def _part_bg(deck, z, part, cache):
    """版式/母版的背景（read.py 只暴露页面级 bg/bg_image，这里兜底）。

    返回 ('color', (r,g,b)) / (None)；图片背景不在版式/母版上处理（常见语料为 bgRef 纯色）。
    """
    if not part:
        return None
    if part in cache:
        return cache[part]
    spec = None
    try:
        root = read.parse_xml(z.read(part))
        bg = root.find(read.q("p:cSld") + "/" + read.q("p:bg"))
        if bg is not None:
            bg_pr = bg.find(read.q("p:bgPr"))
            if bg_pr is not None:
                solid = bg_pr.find(read.q("a:solidFill"))
                if solid is not None:
                    node = solid.find(".//" + read.q("a:srgbClr"))
                    if node is not None and node.get("val"):
                        spec = ("color", parse_color(node.get("val")))
                    else:
                        node = solid.find(".//" + read.q("a:schemeClr"))
                        if node is not None and node.get("val"):
                            spec = ("color", parse_color("scheme:" + node.get("val")))
            if spec is None:
                bg_ref = bg.find(read.q("p:bgRef"))
                if bg_ref is not None:
                    node = bg_ref.find(".//" + read.q("a:schemeClr"))
                    if node is not None and node.get("val"):
                        spec = ("color", parse_color("scheme:" + node.get("val")))
    except (KeyError, Exception):
        spec = None
    cache[part] = spec
    return spec


def _paint_background(deck, z, canvas, slide, cache):
    """背景：页面纯色 → 页面图片 → 版式/母版纯色 → 白底。"""
    if slide.bg:
        color = parse_color(slide.bg)
        if color is not None:
            canvas.fill_rect(0, 0, canvas.w, canvas.h, color)
            return
    if slide.bg_image:
        info = _media_info(z, slide.bg_image, _BG_MEDIA[0])
        if info.get("planes"):
            canvas.blit(0, 0, canvas.w, canvas.h, info)
        else:
            _placeholder(canvas, 0, 0, canvas.w, canvas.h, _media_label(info))
        return
    for part in (slide.layout, slide.master):
        spec = _part_bg(deck, z, part, cache)
        if spec and spec[0] == "color" and spec[1] is not None:
            canvas.fill_rect(0, 0, canvas.w, canvas.h, spec[1])
            return


#: 为了让 draw_shape 不层层传 zip/缓存，背景图片解码期间把当前上下文放在这里。
#: 只在 render() 内部使用（单线程、同步），属于内部实现细节。
_BG_MEDIA = [None]


# --------------------------------------------------------------------------
# 渲染主流程
# --------------------------------------------------------------------------
class _Ctx:
    """一次 render() 调用的上下文：zip、画布尺寸、媒体缓存。"""

    def __init__(self, deck, pptx, width):
        self.deck = deck
        self.z = zipfile.ZipFile(pptx)
        self.W = int(width)
        self.H = max(1, int(round(self.W * deck.h / float(deck.w))))
        self.scale = self.W / float(deck.w)
        self.media = {}
        self.bg_cache = {}

    def close(self):
        self.z.close()


def render(pptx, outdir, pages=None, width=DEFAULT_WIDTH,
           with_layout_images=True, verbose=False):
    """把 pptx 渲染成 PNG，返回生成的文件路径列表（按 pages 顺序）。

    pages: None 表示全部；否则是 1-based 页号列表，如 [1, 3]。
    width: 单页输出宽度（像素），高度按画布纵横比推导。
    with_layout_images: 是否画版式/母版继承的图片与装饰（2026-10 的校徽在 layout 里）。
    """
    t_start = time.time()
    deck = read.Deck.load(pptx)
    os.makedirs(outdir, exist_ok=True)
    ctx = _Ctx(deck, pptx, width)
    try:
        if pages is None:
            selected = list(range(1, deck.n + 1))
        else:
            selected = []
            for p in pages:
                p = int(p)
                if not (1 <= p <= deck.n):
                    raise ValueError("页号 %d 超出 1—%d（%s）" % (p, deck.n, pptx))
                selected.append(p)
        _BG_MEDIA[0] = ctx.media
        paths = []
        for page_no in selected:
            t_page = time.time()
            slide = deck.slide(page_no)
            canvas = Canvas(ctx.W, ctx.H, (255, 255, 255))
            _paint_background(deck, ctx.z, canvas, slide, ctx.bg_cache)
            if with_layout_images:
                for sh in slide.inherited:
                    if sh.kind == "sp" and sh.placeholder:
                        continue        # 占位符只是提示框，不画
                    draw_shape(canvas, ctx, sh, ctx.scale)
            for sh in slide.shapes:
                draw_shape(canvas, ctx, sh, ctx.scale)
            path = os.path.join(outdir, "p%02d.png" % page_no)
            canvas.png(path)
            paths.append(path)
            if verbose:
                print("  p%02d → %s  (%.2fs)" % (page_no, path, time.time() - t_page))
        if verbose:
            elapsed = time.time() - t_start
            print("  共 %d 页，%.2fs，单页均值 %.2fs（宽 %d，画布 %dx%d）" % (
                len(paths), elapsed, elapsed / max(1, len(paths)), ctx.W, ctx.W, ctx.H))
        return paths
    finally:
        _BG_MEDIA[0] = None
        ctx.close()


# --------------------------------------------------------------------------
# 拼图总览
# --------------------------------------------------------------------------
def _read_png_planes(path):
    """读回本模块生成的 PNG（也可读其他 8bit PNG），返回 info dict 或 None。"""
    try:
        with open(path, "rb") as fh:
            data = fh.read()
    except OSError:
        return None
    if data[:8] != _PNG_SIG:
        return None
    return _decode_png(data)


def _resize_plane(plane, sw, sh, dw, dh):
    """最近邻缩放一个通道平面，返回 bytes(dw*dh)。整除时走切片（C 速度）。"""
    out = bytearray(dw * dh)
    if sw >= dw and sw % dw == 0:
        step = sw // dw
        for dy in range(dh):
            sy = (dy * sh) // dh
            out[dy * dw:(dy + 1) * dw] = plane[sy * sw:(sy + 1) * sw:step][:dw]
        return bytes(out)
    xs = [x * sw // dw for x in range(dw)]
    for dy in range(dh):
        sy = (dy * sh) // dh
        base = sy * sw
        out[dy * dw:(dy + 1) * dw] = bytes([plane[base + x] for x in xs])
    return bytes(out)


def contact_sheet(pngs, out_png, cols=2, cell_width=None):
    """把多页 PNG 拼成一张总览图（默认 2 列），返回输出路径。

    cols: 列数；cell_width: 单格宽度（默认 = 单页宽 // cols，保证整图宽度接近单页）。
    """
    pngs = [p for p in pngs if p]
    if not pngs:
        raise ValueError("contact_sheet 需要至少一张 PNG")
    infos = []
    for p in pngs:
        info = _read_png_planes(p)
        if info is None:
            raise ValueError("无法解码拼图输入：%s" % p)
        infos.append(info)
    cols = max(1, int(cols))
    src_w, src_h = infos[0]["w"], infos[0]["h"]
    if cell_width is None:
        cell_width = max(200, src_w // cols)
    cell_w = int(cell_width)
    cell_h = max(1, int(round(cell_w * src_h / float(src_w))))
    margin, gap, label_h = 16, 14, 18
    rows = int(math.ceil(len(infos) / float(cols)))
    sheet_w = margin * 2 + cols * cell_w + (cols - 1) * gap
    sheet_h = margin * 2 + rows * (label_h + cell_h) + (rows - 1) * gap
    canvas = Canvas(sheet_w, sheet_h, (246, 246, 246))
    for i, (path, info) in enumerate(zip(pngs, infos)):
        row, col = divmod(i, cols)
        cx = margin + col * (cell_w + gap)
        cy = margin + row * (cell_h + label_h + gap)
        label = os.path.splitext(os.path.basename(path))[0]
        draw_label(canvas, label, cx + 2, cy + 2, 13, (90, 90, 90))
        iy = cy + label_h
        canvas.fill_rect(cx, iy, cell_w, cell_h, (255, 255, 255))
        planes = info.get("planes")
        if not planes:
            draw_label(canvas, "解码失败", cx + 6, iy + 8, 13, (150, 60, 60))
            canvas.rect_outline(cx, iy, cell_w, cell_h, (200, 200, 200), 1)
            continue
        dw = min(cell_w, max(1, int(round(cell_w * info["w"] / float(src_w)))))
        dh = min(cell_h, max(1, int(round(cell_h * info["h"] / float(src_h)))))
        spans = [_resize_plane(planes[c], info["w"], info["h"], dw, dh)
                 for c in range(3)]
        for dy in range(dh):
            dst = (iy + dy) * sheet_w + cx
            src = dy * dw
            canvas.p[0][dst:dst + dw] = spans[0][src:src + dw]
            canvas.p[1][dst:dst + dw] = spans[1][src:src + dw]
            canvas.p[2][dst:dst + dw] = spans[2][src:src + dw]
        canvas.rect_outline(cx, iy, cell_w, cell_h, (200, 200, 200), 1)
    canvas.png(out_png)
    return out_png


# --------------------------------------------------------------------------
# 自测 / CLI
# --------------------------------------------------------------------------
DEFAULT_CORPUS = os.environ.get("MAKEPPT_CORPUS") or None
if DEFAULT_CORPUS and os.path.isfile(DEFAULT_CORPUS):
    SELFTEST_PPTX = DEFAULT_CORPUS
else:
    SELFTEST_PPTX = (os.path.join(DEFAULT_CORPUS, "deck-b.pptx")
                     if DEFAULT_CORPUS else None)
SELFTEST_OUT = os.path.join(_ENGINE_DIR, os.pardir, "out", "preview", "selftest")


def _nonwhite_ratio(path, tolerance=12):
    """解码 PNG 并数非白像素占比；自测用（>1% 才认为画出了东西）。"""
    info = _read_png_planes(path)
    if info is None or not info.get("planes"):
        return None
    total = info["w"] * info["h"]
    r, g, b = info["planes"][:3]
    count = 0
    for i in range(total):
        if (abs(r[i] - 255) > tolerance or abs(g[i] - 255) > tolerance
                or abs(b[i] - 255) > tolerance):
            count += 1
    return count / float(total)


def _self_test(pptx=None, outdir=None, width=DEFAULT_WIDTH):
    """渲染示例语料前 2 页并检查：路径存在、单页耗时、非空像素占比 > 1%。"""
    pptx = pptx or SELFTEST_PPTX
    if not pptx:
        print("SKIP preview.py 自测：未设置 MAKEPPT_CORPUS（示例语料目录）")
        return 0
    outdir = outdir or SELFTEST_OUT
    print("=== preview.py self-test ===")
    print("pptx:", pptx)
    if not os.path.exists(pptx):
        print("SKIP preview.py 自测：找不到示例语料 %s" % pptx)
        return 0
    t0 = time.time()
    paths = render(pptx, outdir, pages=[1, 2], width=width)
    elapsed = time.time() - t0
    ok = True
    for path in paths:
        ratio = _nonwhite_ratio(path)
        if ratio is None:
            print("  %s：PNG 读回失败" % path)
            ok = False
            continue
        status = "OK" if ratio > 0.01 else "空白页!"
        if ratio <= 0.01:
            ok = False
        print("  %s  非空像素 %.2f%%（>1%% 才算画出了东西）→ %s"
              % (path, ratio * 100, status))
    print("  耗时 %.2fs / %d 页，单页均值 %.2fs"
          % (elapsed, len(paths), elapsed / max(1, len(paths))))
    if ok:
        print("SELF-TEST OK — 预览渲染可用")
        return 0
    print("SELF-TEST FAILED")
    return 1


def parse_pages(spec):
    """把 '1-4,7,9-10' 解析成 [1,2,3,4,7,9,10]（1-based，可乱序）。"""
    if not spec:
        return None
    out = []
    for token in str(spec).split(","):
        token = token.strip()
        if not token:
            continue
        if "-" in token:
            a, _, b = token.partition("-")
            try:
                a, b = int(a), int(b)
            except ValueError:
                raise ValueError("页号区间不合法：%r" % token)
            if a > b:
                a, b = b, a
            out.extend(range(a, b + 1))
        else:
            out.append(int(token))
    if not out:
        raise ValueError("--pages 没有解析出任何页号：%r" % spec)
    return out


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    self_test = "--self-test" in argv
    argv = [a for a in argv if a != "--self-test"]
    pages = None
    width = DEFAULT_WIDTH
    sheet = None
    sheet_cols = 2
    json_out = False
    positional = []
    i = 0
    while i < len(argv):
        arg = argv[i]
        if arg == "--pages" and i + 1 < len(argv):
            pages = parse_pages(argv[i + 1])
            i += 2
        elif arg == "--width" and i + 1 < len(argv):
            width = int(argv[i + 1])
            i += 2
        elif arg == "--sheet" and i + 1 < len(argv):
            sheet = argv[i + 1]
            i += 2
        elif arg == "--cols" and i + 1 < len(argv):
            sheet_cols = int(argv[i + 1])
            i += 2
        elif arg == "--json":
            json_out = True
            i += 1
        elif arg.startswith("-"):
            print("未知参数：%s" % arg)
            return 2
        else:
            positional.append(arg)
            i += 1
    if self_test:
        return _self_test(positional[0] if positional else None,
                          positional[1] if len(positional) > 1 else None,
                          width)
    if not positional:
        print("用法：python preview.py <pptx> <outdir> [--pages 1-4,7] [--width 1400] "
              "[--sheet overview.png] [--cols 2] [--json]")
        return 2
    if len(positional) < 2:
        print("缺少 outdir 参数")
        return 2
    pptx, outdir = positional[0], positional[1]
    deck = read.Deck.load(pptx)
    print("渲染 %s（%d 页）× 宽 %d → %s" % (pptx, deck.n, width, outdir))
    t0 = time.time()
    paths = render(pptx, outdir, pages=pages, width=width, verbose=not json_out)
    elapsed = time.time() - t0
    selected = pages if pages else list(range(1, deck.n + 1))
    if sheet:
        sheet_path = sheet if os.path.isabs(sheet) or os.path.dirname(sheet) \
            else os.path.join(outdir, sheet)
        contact_sheet(paths, sheet_path, cols=sheet_cols)
        print("拼图总览：%s" % sheet_path)
    stats = {
        "pptx": pptx,
        "outdir": outdir,
        "width": width,
        "pages": selected,
        "pngs": paths,
        "seconds": round(elapsed, 3),
        "per_page_seconds": round(elapsed / max(1, len(paths)), 3),
    }
    if json_out:
        import json
        print(json.dumps(stats, ensure_ascii=False, indent=1))
    else:
        print("完成：%d 页，总耗时 %.2fs，单页均值 %.2fs"
              % (len(paths), elapsed, elapsed / max(1, len(paths))))
    return 0


if __name__ == "__main__":
    sys.exit(main())
