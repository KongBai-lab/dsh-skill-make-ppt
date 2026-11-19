# -*- coding: utf-8 -*-
"""edit.py — PPTX 原位改造（make-ppt engine）。

为什么要有这个模块
------------------
2026-10 做课程 PPT 时，90% 的工作不是"从零生成"，而是
**改已有 PPTX**：移页、删页、复制页、改一行字、统一校徽、写备注。当时没有
模块化工具，每个需求都是一份 300—700 行的一次性脚本，于是出过这些事故：

* 重跑生成器把手工修复覆盖掉（没有"原子写 + 先校验后替换"）；
* 同一个校徽在 layout 与页面里各放一份，成了"双图标"，修了三次；
* 复制页面时只复制了 slide XML，没同步 presentation.xml.rels 与
  [Content_Types].xml，WPS 打开报"文件损坏"；
* 单位换算散落在脚本里，`sz`（1/100 pt）被当成 pt、`spcPct` 写成 100。

本模块把上述经验固化为可复用 API。设计原则：

1. **只改局部字节**。所有 slide / layout / notesSlide 的修改都在字符串层面
   完成，未触碰的部分逐字节保留（不整体重排 XML、不重写命名空间），因此
   不会破坏 `mc:Ignorable`、a14 扩展、主题引用等 PowerPoint 敏感结构。
2. **单位只走 units**。任何写进 XML 的字号、坐标、行距都经 `units` 生成或
   校验；本模块不定义第二套常量。
3. **跨页操作四件套同步**：`sldIdLst`、`presentation.xml.rels`、
   `[Content_Types].xml`、`notesSlide` 关系。保存前做基础结构校验，任何
   断链/未覆盖部件/数值越界都会拒绝落盘。
4. **原子写**：先写 `<dst>.tmp`，校验通过后 `os.replace`；失败保留旧文件。
5. **幂等**：`delete_slide` / `duplicate_slide` / `unify_element` 重复执行
   不会产生重复部件、重复关系或重复图片（unify 连跑两次，第二次 0 改动）。

API（对外页号一律 1-based；内部 IR 列表 0-based）
------------------------------------------------
    ed = edit.Editor.open(src_pptx)          # 读全部部件到内存
    ed.deck                                  # read.Deck（只读 IR，含继承元素）
    ed.set_text(2, name="标题 2", text="新标题")
    ed.set_text(2, idx=3, text="新文字")      # idx 与 read.Shape.idx 一致
    ed.set_text(2, contains="旧片段", text="新文字")
    ed.replace_text(2, "旧", "新")            # 返回替换次数（支持跨 run）
    ed.delete_slide(3) / ed.move_slide(5, 1) / ed.duplicate_slide(2)
    ed.set_notes(2, "备注文字")               # 更新或新建 notesSlide
    ed.unify_element(picture_bytes=png, x=..., y=..., w=..., h=...,
                     target_layouts=["ppt/slideLayouts/slideLayout2.xml"])
    ed.renumber_superscripts({"¹": "⁴", "²": "⁵"})
    ed.save(out_pptx, verify=True)            # 先 tmp → 校验 → os.replace
    ed.diff(other_pptx)                       # 页数/文本/图片数差异

`save(dst, verify=...)` 的 verify 值域：

* `True`（默认）——跑 check；structure/units 类 ERROR 阻止落盘，
  inherit/layout 类 ERROR（如源文件自带的"双图标"）只打印提示不阻断，
  内容问题用 `unify_element` 修，不该让人连中间稿都存不下来；
* `"strict"` —— check 的所有 ERROR 都阻止落盘（交付前用）；
* `False` —— 不跑 check，只做基础结构校验（仍保证包没坏）。
  check 尚未就绪时自动降级并打印"check 未就绪，跳过校验"，不会崩。

命令行
------
    python edit.py <src.pptx> --self-test   # 一键回归（本节验收 5 条全跑）
    python edit.py <src.pptx> --summary     # 打印 read.Deck.summary()
    python edit.py <src.pptx> --diff <other.pptx>

自测用语料：优先 `tests/corpus/deck-a.pptx`，其次环境变量 `MAKEPPT_CORPUS`；
两者都没有时打印 SKIP，不报错。

依赖：仅标准库 + 同目录 units.py / read.py。
"""
import hashlib
import json
import os
import posixpath
import re
import sys
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, unescape

def _locate_engine_dir():
    """定位 engine 目录（兼容 `python edit.py`、runpy、exec 三种执行方式）。"""
    cands = []
    try:
        cands.append(os.path.dirname(os.path.abspath(__file__)))
    except NameError:                       # exec 方式执行时没有 __file__
        pass
    env = os.environ.get("MAKE_PPT_ENGINE")
    if env:
        cands.append(env)
    if sys.argv and sys.argv[0].endswith(".py") and os.path.isfile(sys.argv[0]):
        cands.append(os.path.dirname(os.path.abspath(sys.argv[0])))
    cands.append(os.getcwd())
    for c in cands:
        if c and os.path.isfile(os.path.join(c, "units.py")):
            return os.path.abspath(c)
    return os.path.abspath(cands[0] if cands else os.getcwd())


_ENGINE_DIR = _locate_engine_dir()
if _ENGINE_DIR not in sys.path:
    sys.path.insert(0, _ENGINE_DIR)

try:                       # 支持 `import edit` 与 `from engine import edit` 两种用法
    from . import units, read
except ImportError:        # pragma: no cover
    import units, read


# ==========================================================================
# 常量：命名空间 / 关系类型 / Content-Type
# ==========================================================================
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
PKG_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

REL_SLIDE = NS_R + "/slide"
REL_LAYOUT = NS_R + "/slideLayout"
REL_MASTER = NS_R + "/slideMaster"
REL_NOTES_SLIDE = NS_R + "/notesSlide"
REL_NOTES_MASTER = NS_R + "/notesMaster"
REL_IMAGE = NS_R + "/image"
REL_THEME = NS_R + "/theme"
REL_HYPERLINK = NS_R + "/hyperlink"

CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
CT_NOTES_SLIDE = ("application/vnd.openxmlformats-officedocument."
                  "presentationml.notesSlide+xml")
CT_LAYOUT = ("application/vnd.openxmlformats-officedocument."
             "presentationml.slideLayout+xml")
CT_MASTER = ("application/vnd.openxmlformats-officedocument."
             "presentationml.slideMaster+xml")
CT_NOTES_MASTER = ("application/vnd.openxmlformats-officedocument."
                   "presentationml.notesMaster+xml")
CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
CT_PRES = ("application/vnd.openxmlformats-officedocument."
           "presentationml.presentation.main+xml")
CT_PRES_PROPS = ("application/vnd.openxmlformats-officedocument."
                 "presentationml.presProps+xml")
CT_VIEW_PROPS = ("application/vnd.openxmlformats-officedocument."
                 "presentationml.viewProps+xml")
CT_TABLE_STYLES = ("application/vnd.openxmlformats-officedocument."
                   "presentationml.tableStyles+xml")
CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
CT_APP = ("application/vnd.openxmlformats-officedocument."
          "extended-properties+xml")
CT_CUSTOM = ("application/vnd.openxmlformats-officedocument."
             "custom-properties+xml")

# 顶层形状标签（read.py 的 _walk 认这几种；顺序只用于扫描，不用于输出）
BLOCK_TAGS = ("p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp",
              "p:contentPart")
_BLOCK_RE = re.compile(r"<(%s)(?=[\s/>])" % "|".join(re.escape(t)
                                                    for t in BLOCK_TAGS))

# Unicode 上标数字 ↔ 普通数字（renumber_superscripts 用）
SUP_TO_DIGIT = {"⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
                "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9"}
DIGIT_TO_SUP = dict((v, k) for k, v in SUP_TO_DIGIT.items())

# 内置最小 notesSlide 骨架（包里一个 notesSlide 都没有时用它兜底）。
# 结构对齐 PowerPoint 生成的 notesSlide：body 占位符 + sldNum 页码。
MIN_NOTES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:notes xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
    '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
    '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="备注占位符 1"/>'
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
    '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
    '<p:spPr><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>'
    '<a:p/></p:txBody></p:sp>'
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
    '</p:notes>'
) % (NS_A, NS_R, NS_P)


class EditError(RuntimeError):
    """改造操作本身不合法（页号/定位参数/XML 结构异常等）。"""


class VerifyError(EditError):
    """保存前校验失败（产出会拒绝落盘，旧文件保留）。"""


# ==========================================================================
# 基础工具：rels / 部件路径 / XML 块扫描
# ==========================================================================
def _rels_path(part):
    """ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels"""
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def _abs_target(part, target):
    """把关系里的相对 Target 换算成包内绝对部件名。"""
    t = target or ""
    if (not t) or t.startswith("/") or t.startswith(("http://", "https://",
                                                     "mailto:", "file:")):
        return t
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), t))


def _short_type(rel_type):
    return (rel_type or "").rsplit("/", 1)[-1]


def _parse_rels(raw):
    """rels XML 字节/字符串 → [{"id","type","target","mode"}]（保持原顺序）。"""
    if not raw:
        return []
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as e:      # 损坏的关系文件：明确报错，不静默
        raise EditError("关系文件无法解析：%s" % e)
    out = []
    for el in root:
        if el.tag.rsplit("}", 1)[-1] != "Relationship":
            continue
        out.append({
            "id": el.get("Id") or "",
            "type": el.get("Type") or "",
            "target": el.get("Target") or "",
            "mode": el.get("TargetMode") or "Internal",
        })
    return out


def _rels_xml(rows):
    """[{"id","type","target","mode"}] → rels XML 字符串。"""
    buf = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Relationships xmlns="%s">' % PKG_REL]
    for r in rows:
        target = escape(r.get("target") or "")
        extra = ' TargetMode="External"' if r.get("mode") == "External" else ""
        buf.append('<Relationship Id="%s" Type="%s" Target="%s"%s/>'
                   % (r.get("id") or "", escape(r.get("type") or ""),
                      target, extra))
    buf.append("</Relationships>")
    return "".join(buf)


def _next_rid(rows):
    """分配一个未占用的 rId。"""
    used = set(r.get("id") for r in rows)
    i = 1
    while ("rId%d" % i) in used:
        i += 1
    return "rId%d" % i


def _tag_open_end(xml, start):
    """从 start（'<' 位置）找开标签的 '>' 下标，处理属性值里的 '>'（转义过，安全）。"""
    gt = xml.find(">", start)
    if gt < 0:
        raise EditError("XML 标签未闭合（位置 %d）" % start)
    return gt


def _match_close(xml, start, tag, end=None):
    """从 start（`<tag` 开始处）找配对闭合标签之后的下标。自闭合直接返回。"""
    end = len(xml) if end is None else end
    pat = re.compile(r"</?%s(?=[\s/>])" % re.escape(tag))
    pos = start
    depth = 0
    while pos < end:
        m = pat.search(xml, pos, end)
        if not m:
            return None
        is_close = xml[m.start() + 1] == "/"
        gt = xml.find(">", m.end())
        if gt < 0 or gt >= end:
            return None
        self_closing = xml[gt - 1] == "/"
        if is_close:
            depth -= 1
            if depth == 0:
                return gt + 1
        elif not self_closing:
            depth += 1
        elif depth == 0:
            return gt + 1
        pos = gt + 1
    return None


def _scan_blocks(xml, start=0, end=None):
    """扫描区间内**顶层**形状元素块，返回 [(tag, s, e, block)]（s/e 绝对下标）。"""
    end = len(xml) if end is None else end
    out = []
    i = start
    while i < end:
        m = _BLOCK_RE.search(xml, i, end)
        if not m:
            break
        tag = m.group(1)
        s = m.start()
        e = _match_close(xml, s, tag, end)
        if e is None:                       # 损坏结构：跳过，避免死循环
            i = m.end()
            continue
        out.append((tag, s, e, xml[s:e]))
        i = e
    return out


def _sp_tree_span(xml):
    """返回 spTree 内容区间 (inner_start, inner_end)；空 spTree 返回 None。"""
    m = re.search(r"<p:spTree(?=[\s/>])", xml)
    if not m:
        return None
    gt = _tag_open_end(xml, m.start())
    if xml[gt - 1] == "/":
        return None
    close = xml.find("</p:spTree>", gt + 1)
    if close < 0:
        return None
    return (gt + 1, close)


def _walk_elements(xml):
    """按 read._walk 的顺序（grpSp 递归展开、grpSp 自身不计数）返回
    [(idx, tag, s, e, block)]，idx 从 1 开始，与 read.Shape.idx 对齐。"""
    span = _sp_tree_span(xml)
    out = []
    if not span:
        return out
    counter = [0]

    def rec(s, e):
        for tag, bs, be, block in _scan_blocks(xml, s, e):
            if tag == "p:grpSp":
                gt = _tag_open_end(xml, bs)
                close = xml.rfind("</p:grpSp>", gt + 1, be)
                if close > gt:
                    rec(gt + 1, close)
                continue
            counter[0] += 1
            out.append((counter[0], tag, bs, be, block))

    rec(span[0], span[1])
    return out


def _block_name(block):
    m = re.search(r'<p:cNvPr\b[^>]*\bname="([^"]*)"', block)
    return m.group(1) if m else ""


def _block_texts(block):
    return [unescape(m.group(1)) for m in re.finditer(r"<a:t>([^<]*)</a:t>",
                                                      block)]


def _block_text(block):
    return "".join(_block_texts(block))


def _block_rid(block):
    m = re.search(r'<a:blip\b[^>]*\br:embed="([^"]+)"', block)
    return m.group(1) if m else ""


def _attr_int(tag_text, name, default=None):
    m = re.search(r'\b%s="(-?\d+)"' % re.escape(name), tag_text)
    return int(m.group(1)) if m else default


def _block_rect(block):
    """块自身的 (x, y, w, h) EMU；取自块内第一个 a:off / a:ext。"""
    m = re.search(r"<a:off\b[^>]*/>", block)
    if not m:
        return None
    off = m.group(0)
    ext_m = re.search(r"<a:ext\b[^>]*/>", block)
    if not ext_m:
        return None
    ext = ext_m.group(0)
    x = _attr_int(off, "x")
    y = _attr_int(off, "y")
    cx = _attr_int(ext, "cx")
    cy = _attr_int(ext, "cy")
    if None in (x, y, cx, cy):
        return None
    return (x, y, cx, cy)


def _rect_overlap_ratio(a, b):
    """a 被 b 覆盖的面积比例（0—1）。"""
    if not a or not b:
        return 0.0
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    if aw <= 0 or ah <= 0:
        return 0.0
    ix = max(0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0, min(ay + ah, by + bh) - max(ay, by))
    return (ix * iy) / float(aw * ah)


def _max_shape_id(xml):
    ids = [int(v) for v in re.findall(r'<p:cNvPr\b[^>]*\bid="(\d+)"', xml)]
    return max(ids) if ids else 1


def _image_ext(data):
    """按魔数判断图片扩展名（只支持 PNG/JPEG，够用且不依赖 PIL）。"""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if data[:3] == b"\xff\xd8\xff":
        return "jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if data[:2] == b"BM":
        return "bmp"
    raise EditError("picture_bytes 不是可识别的图片（支持 PNG/JPEG/GIF/BMP，"
                    "实际头部 %r）" % data[:8])


def _image_size(data):
    """PNG/JPEG/GIF 的像素尺寸；读不出返回 (None, None)。"""
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            import struct
            return struct.unpack(">II", data[16:24])
        if data[:3] == b"\xff\xd8\xff":
            i, n = 2, len(data)
            while i < n - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                mk = data[i + 1]
                if mk in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                          0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    h = (data[i + 5] << 8) | data[i + 6]
                    w = (data[i + 7] << 8) | data[i + 8]
                    return int(w), int(h)
                seg = (data[i + 2] << 8) | data[i + 3]
                i += 2 + seg
        if data[:6] in (b"GIF87a", b"GIF89a"):
            return (data[6] | (data[7] << 8), data[8] | (data[9] << 8))
    except Exception:
        pass
    return None, None


def _guess_content_type(part):
    """按部件路径猜 Content-Type；返回 None 表示交给 Default Extension。"""
    if part.startswith("ppt/slides/") and part.endswith(".xml"):
        return CT_SLIDE
    if part.startswith("ppt/notesSlides/") and part.endswith(".xml"):
        return CT_NOTES_SLIDE
    if part.startswith("ppt/slideLayouts/") and part.endswith(".xml"):
        return CT_LAYOUT
    if part.startswith("ppt/slideMasters/") and part.endswith(".xml"):
        return CT_MASTER
    if part.startswith("ppt/notesMasters/") and part.endswith(".xml"):
        return CT_NOTES_MASTER
    if part.startswith("ppt/theme/") and part.endswith(".xml"):
        return CT_THEME
    fixed = {
        "ppt/presentation.xml": CT_PRES,
        "ppt/presProps.xml": CT_PRES_PROPS,
        "ppt/viewProps.xml": CT_VIEW_PROPS,
        "ppt/tableStyles.xml": CT_TABLE_STYLES,
        "docProps/core.xml": CT_CORE,
        "docProps/app.xml": CT_APP,
        "docProps/custom.xml": CT_CUSTOM,
    }
    if part in fixed:
        return fixed[part]
    if part.endswith(".rels"):
        return CT_RELS
    return None


def _guess_media_type(part):
    ext = part.rsplit(".", 1)[-1].lower()
    return {
        "png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg",
        "gif": "image/gif", "bmp": "image/bmp", "tif": "image/tiff",
        "tiff": "image/tiff", "svg": "image/svg+xml", "emf": "image/x-emf",
        "wmf": "image/x-wmf", "wdp": "image/vnd.ms-photo",
    }.get(ext)


def _normalise_layouts(items):
    """把用户给的 layout 标识统一成部件名（接受完整名/slideLayout2.xml/2）。"""
    out = []
    for item in items:
        s = str(item)
        if s.startswith("ppt/"):
            out.append(s)
            continue
        if s.isdigit():
            out.append("ppt/slideLayouts/slideLayout%s.xml" % s)
            continue
        if "/" not in s:
            out.append(posixpath.join("ppt/slideLayouts", s))
            continue
        out.append(s)
    return out


# ==========================================================================
# 文本块改写（set_text / set_notes 共用）
# ==========================================================================
_RPR_RE = re.compile(r"<a:rPr\b[^>]*/>|<a:rPr\b[^>]*>.*?</a:rPr>", re.S)
_PPR_RE = re.compile(r"<a:pPr\b[^>]*/>|<a:pPr\b[^>]*>.*?</a:pPr>", re.S)
_RUN_RE = re.compile(r"<a:r\b[^>]*>.*?</a:r>", re.S)


def _set_block_paragraphs(block, text, keep_format=True):
    """把一个形状块 txBody 里的段落整体替换为 text（按 \\n 分段）。

    keep_format=True：沿用第一段 / 第一个 run 的段落属性与字符属性（字号、
    颜色、字体、行距……），这是"改字不换皮"的默认行为。
    keep_format=False：保留段落属性，但字符属性重置为中性 rPr（字号仍取原值，
    经 units.sz_of_pt 生成，不自己拼数字）。

    返回新的形状块字符串；块内没有 txBody（如图片）时返回 None。
    """
    m = re.search(r"(<p:txBody\b[^>]*>)(.*?)(</p:txBody>)", block, re.S)
    if not m:
        return None
    inner = m.group(2)
    paras = list(re.finditer(r"<a:p(?=[\s/>])", inner))
    if paras:
        p_start = paras[0].start()
        p_end = _match_close(inner, p_start, "a:p")
        if p_end is None:
            p_end = len(inner)
        first = inner[p_start:p_end]
        head = inner[:p_start]
    else:
        p_start = p_end = len(inner)
        first = ""
        head = inner

    ppr = ""
    rpr = ""
    if first:
        pm = _PPR_RE.search(first)
        if pm:
            ppr = pm.group(0)
        rm = _RUN_RE.search(first)
        if rm:
            rp = _RPR_RE.search(rm.group(0))
            if rp:
                rpr = rp.group(0)

    # 字号：keep_format=False 时仍沿用原字号，只是重新用 units 生成
    pt = None
    if rpr:
        szm = re.search(r'\bsz="(\d+)"', rpr)
        if szm:
            pt = units.pt_of_sz(szm.group(1))
    if pt is None:
        pt = 18.0
    if not keep_format:
        rpr = '<a:rPr lang="zh-CN" altLang="en-US" sz="%s" b="0"/>' % (
            units.sz_of_pt(pt))
    elif not rpr:
        rpr = '<a:rPr lang="zh-CN" altLang="en-US" sz="%s" b="0"/>' % (
            units.sz_of_pt(pt))

    lines = text.split("\n") if isinstance(text, str) else list(text)
    new_paras = []
    for line in lines:
        if line == "":
            new_paras.append("<a:p>%s</a:p>" % ppr)
        else:
            new_paras.append('<a:p>%s<a:r>%s<a:t>%s</a:t></a:r></a:p>'
                             % (ppr, rpr, escape(line)))
    new_inner = head + "".join(new_paras)
    return block[:m.start()] + m.group(1) + new_inner + m.group(3) + \
        block[m.end():]


_PARA_RE = re.compile(r"<a:p\b[^>]*>.*?</a:p>", re.S)


def _rebuild_para(para_block, text):
    """把一段重建为单 run：保留段落 pPr 与首 run 的 rPr（refs/edit-ops.md §二）。"""
    open_m = re.match(r"<a:p\b[^>]*>", para_block)
    open_tag = open_m.group(0) if open_m else "<a:p>"
    pm = _PPR_RE.search(para_block)
    ppr = pm.group(0) if pm else ""
    rpr = ""
    rm = _RUN_RE.search(para_block)
    if rm:
        rp = _RPR_RE.search(rm.group(0))
        if rp:
            rpr = rp.group(0)
    return open_tag + ppr + "<a:r>%s<a:t>%s</a:t></a:r></a:p>" % (
        rpr, escape(text))


def _replace_texts_in_xml(xml, old, new, all_occurrences=True):
    """页内文本替换（含跨 run），返回 (新 XML, 替换次数, 重建段数)。

    2026-10 的教训：PowerPoint 会按编辑历史把一个词拆成多个 run，只在单个
    a:t 内替换会留下残片（"课程汇" + "报" 只换到一半）。做法：
      1. 先在各 a:t 内替换；
      2. 若整段文本仍含 old（说明 old 跨 run），把段内文字拼起来、按首 run 的
         格式重建为单 run；
      3. 段内有 a:fld（页码域）或 a:br（软换行）时保守处理，只做第 1 步。
    """
    state = {"done": False}
    count = [0]
    rebuilt = [0]

    def replace_plain(txt):
        if old not in txt:
            return txt, 0
        if not all_occurrences and state["done"]:
            return txt, 0
        if all_occurrences:
            n = txt.count(old)
            nt = txt.replace(old, new)
        else:
            n = 1
            nt = txt.replace(old, new, 1)
            state["done"] = True
        return nt, n

    def do_para(m):
        para = m.group(0)
        if "<a:t>" not in para:
            return para
        para_text = "".join(unescape(a.group(1)) for a in
                            re.finditer(r"<a:t>([^<]*)</a:t>", para))
        if old not in para_text:
            return para

        def rep_at(am):
            txt = unescape(am.group(1))
            nt, n = replace_plain(txt)
            if n:
                count[0] += n
                return "<a:t>%s</a:t>" % escape(nt)
            return am.group(0)

        new_para = re.sub(r"<a:t>([^<]*)</a:t>", rep_at, para)
        if "<a:fld" in para or "<a:br" in para:
            return new_para
        now_text = "".join(unescape(a.group(1)) for a in
                           re.finditer(r"<a:t>([^<]*)</a:t>", new_para))
        if old not in now_text:
            return new_para
        if not all_occurrences and state["done"]:
            return new_para
        if all_occurrences:
            n = now_text.count(old)
            final_text = now_text.replace(old, new)
        else:
            n = 1
            final_text = now_text.replace(old, new, 1)
            state["done"] = True
        count[0] += n
        rebuilt[0] += 1
        return _rebuild_para(new_para, final_text)

    out = _PARA_RE.sub(do_para, xml)
    return out, count[0], rebuilt[0]


def _remap_superscripts(xml, mapping):
    """页内上标角标重映射。

    处理两类写法：
      1. Unicode 上标字符（¹²³…）—— 2026-10 的成品稿就是这么写的；
      2. `<a:rPr baseline="30000">` 的正常数字 run —— PowerPoint 另一种上标。
    返回 (新 XML, 改动次数)。mapping 的键值已归一成上标字符。
    """
    changed = [0]

    def rep_text(txt):
        out = []
        n = 0
        for ch in txt:
            if ch in mapping:
                out.append(mapping[ch])
                n += 1
            else:
                out.append(ch)
        if n:
            changed[0] += n
        return "".join(out)

    def rep_at(m):
        txt = unescape(m.group(1))
        nt = rep_text(txt)
        if nt == txt:
            return m.group(0)
        return "<a:t>%s</a:t>" % escape(nt)

    # 先处理 baseline 上标的纯数字 run（run 内 a:t）
    def rep_run(m):
        run = m.group(0)
        bm = re.search(r'<a:rPr\b[^>]*\bbaseline="(-?\d+)"', run)
        if not bm or int(bm.group(1)) <= 0:
            return run
        return re.sub(r"<a:t>([^<]*)</a:t>", rep_at, run)

    xml = re.sub(r"<a:r\b[^>]*>.*?</a:r>", rep_run, xml, flags=re.S)
    # 再处理所有残留的上标字符
    xml = re.sub(r"<a:t>([^<]*)</a:t>", rep_at, xml)
    return xml, changed[0]


# ==========================================================================
# Editor
# ==========================================================================
class Editor(object):
    """PPTX 原位改造器。打开时把 zip 全部部件读入内存，保存时原子写出。"""

    def __init__(self, path, parts, zip_order):
        self.path = path
        self.parts = parts              # {部件名: bytes}，保留 zip 全部部件
        self._zip_order = zip_order     # 原 zip 的部件顺序（写回时尽量保持）
        self.order = []                 # 页序：ppt/slides/*.xml
        self._deck_cache = None
        self._media_hash_cache = None
        self._load_order()
        self._invalidate()

    # ------------------------------------------------------------------
    # 打开 / 页序
    # ------------------------------------------------------------------
    @classmethod
    def open(cls, path):
        path = os.path.abspath(path)
        if not os.path.isfile(path):
            raise EditError("文件不存在：%s" % path)
        try:
            z = zipfile.ZipFile(path)
        except zipfile.BadZipFile as e:
            raise EditError("不是合法的 .pptx（zip 打不开）：%s（%s）"
                            % (path, e))
        with z:
            names = [n for n in z.namelist() if not n.endswith("/")]
            parts = {}
            for n in names:
                parts[n] = z.read(n)
        if "ppt/presentation.xml" not in parts:
            raise EditError("缺少 ppt/presentation.xml，不是 presentation 包：%s"
                            % path)
        return cls(path, parts, names)

    def _load_order(self):
        """按 presentation.xml 的 sldIdLst 恢复页序。"""
        pres_raw = self.parts["ppt/presentation.xml"].decode("utf-8")
        rels = _parse_rels(self.parts.get(_rels_path("ppt/presentation.xml")))
        rel_by_id = dict((r["id"], r) for r in rels)
        order = []
        for m in re.finditer(r"<p:sldId\b([^>]*)/>", pres_raw):
            attrs = m.group(1)
            rid_m = re.search(r'\br:id="([^"]+)"', attrs)
            if not rid_m:
                continue
            rel = rel_by_id.get(rid_m.group(1))
            if not rel:
                continue
            part = _abs_target("ppt/presentation.xml", rel["target"])
            if part in self.parts:
                order.append(part)
        self.order = order

    @property
    def deck(self):
        """read.Deck（每次改后重新生成；不含未保存的中间态之外的差异）。"""
        if self._deck_cache is None:
            tmp = self._temp_path()
            try:
                self._write_zip(tmp)
                self._deck_cache = read.Deck.load(tmp)
            finally:
                try:
                    os.unlink(tmp)
                except OSError:
                    pass
        return self._deck_cache

    def _invalidate(self):
        self._deck_cache = None

    def _temp_path(self):
        fd, tmp = tempfile.mkstemp(prefix="edit_", suffix=".pptx")
        os.close(fd)
        return tmp

    # ------------------------------------------------------------------
    # 低层读写
    # ------------------------------------------------------------------
    def _check_page(self, slide_no):
        if not isinstance(slide_no, int) or isinstance(slide_no, bool):
            raise EditError("页号必须是 int（1-based），收到 %r" % (slide_no,))
        if not (1 <= slide_no <= len(self.order)):
            raise EditError("页号 %d 超出 1—%d" % (slide_no, len(self.order)))
        return self.order[slide_no - 1]

    def _slide_part(self, slide_no):
        return self._check_page(slide_no)

    def _xml(self, part):
        if part not in self.parts:
            raise EditError("部件不存在：%s" % part)
        return self.parts[part].decode("utf-8")

    def _set_xml(self, part, xml):
        self.parts[part] = xml.encode("utf-8")

    def _rels_of(self, part):
        return _parse_rels(self.parts.get(_rels_path(part)))

    def _set_rels_of(self, part, rows):
        self.parts[_rels_path(part)] = _rels_xml(rows).encode("utf-8")

    def _target_of_rid(self, part, rows, rid):
        for r in rows:
            if r["id"] == rid:
                if r["mode"] == "External":
                    return r["target"]
                return _abs_target(part, r["target"])
        return ""

    def _layout_of(self, slide_part):
        for r in self._rels_of(slide_part):
            if _short_type(r["type"]) == "slideLayout":
                return _abs_target(slide_part, r["target"])
        return ""

    def _notes_part_of(self, slide_part):
        for r in self._rels_of(slide_part):
            if _short_type(r["type"]) == "notesSlide":
                return _abs_target(slide_part, r["target"])
        return ""

    def _notes_master_part(self):
        for r in self._rels_of("ppt/presentation.xml"):
            if _short_type(r["type"]) == "notesMaster":
                return _abs_target("ppt/presentation.xml", r["target"])
        return ""

    def _next_part_name(self, directory, prefix, ext):
        i = 1
        while True:
            name = posixpath.join(directory, "%s%d%s" % (prefix, i, ext))
            if name not in self.parts:
                return name
            i += 1

    # ------------------------------------------------------------------
    # 文本：set_text / replace_text
    # ------------------------------------------------------------------
    def _find_shapes(self, part, name=None, idx=None, contains=None):
        """定位页内形状块，返回 [(s, e, block)]（绝对下标）。

        · idx 与 read.Shape.idx 完全一致（grpSp 展开后计数）；
        · contains 是模糊匹配，命中多个时只取第一个（避免误伤）；
        · name 命中多个时全部返回（同名形状在 PPT 里很常见）。
        """
        xml = self._xml(part)
        hits = []
        for i, tag, s, e, block in _walk_elements(xml):
            if idx is not None:
                if i == idx:
                    hits.append((s, e, block))
                continue
            if name is not None and _block_name(block) != name:
                continue
            if contains is not None and contains not in _block_text(block):
                continue
            hits.append((s, e, block))
            if contains is not None:
                break                       # 模糊匹配：只改第一个
        return hits

    def set_text(self, slide_no, *, name=None, idx=None, contains=None, text,
                 keep_format=True):
        """把页内某个形状的文字整体换成 text（保留原格式）。返回改动形状数。

        定位三选一：name（cNvPr@name 精确）/ idx（read.Shape.idx）/ contains
        （页面文字包含该子串的第一个形状）。找不到返回 0，不抛异常。
        """
        loc = [v is not None for v in (name, idx, contains)]
        if sum(loc) != 1:
            raise EditError("set_text 必须且只能用一种定位方式：name / idx / "
                            "contains 三选一")
        part = self._slide_part(slide_no)
        hits = self._find_shapes(part, name=name, idx=idx, contains=contains)
        if not hits:
            return 0
        xml = self._xml(part)
        changed = 0
        for s, e, block in sorted(hits, key=lambda h: h[0], reverse=True):
            new_block = _set_block_paragraphs(block, text,
                                             keep_format=keep_format)
            if new_block is None:
                continue
            xml = xml[:s] + new_block + xml[e:]
            changed += 1
        if changed:
            self._set_xml(part, xml)
            self._invalidate()
        return changed

    def replace_text(self, slide_no, old, new, all_occurrences=True):
        """页内文本替换，返回替换次数。

        · 支持"一个词被拆到多个 run"的情况（PowerPoint 常见）：先在各 run 内
          替换，剩下的跨 run 命中把整段按首 run 格式重建（refs/edit-ops.md §二）；
        · all_occurrences=False 时全页只替换第一处；
        · 段内是页码域（a:fld）或软换行（a:br）时只做单 run 替换，避免破坏结构。
        """
        if not old:
            raise EditError("replace_text 的 old 不能为空")
        part = self._slide_part(slide_no)
        out, n, rebuilt = _replace_texts_in_xml(
            self._xml(part), old, new, all_occurrences)
        if n:
            self._set_xml(part, out)
            self._invalidate()
        if rebuilt:
            print("[edit] replace_text(第 %d 页)：%d 处跨 run 命中，已按首 run "
                  "格式重建段落" % (slide_no, rebuilt))
        return n

    def renumber_superscripts(self, mapping, slides=None):
        """页内上标角标重映射，返回改动次数。

        mapping 形如 {"¹": "⁴", "²": "⁵"} 或 {1: 4, 2: 5}（普通数字也认）。
        默认处理全部页；slides=[2, 5] 可限定页。两类写法都处理：
        Unicode 上标字符、`baseline>0` 的数字 run。
        """
        if not isinstance(mapping, dict) or not mapping:
            raise EditError("renumber_superscripts 需要一个非空 dict")
        norm = {}
        for k, v in mapping.items():
            ks = str(k)
            vs = str(v)
            ks = SUP_TO_DIGIT.get(ks, ks)
            vs = SUP_TO_DIGIT.get(vs, vs)
            if ks in DIGIT_TO_SUP:
                ks = DIGIT_TO_SUP[ks]
            if vs in DIGIT_TO_SUP:
                vs = DIGIT_TO_SUP[vs]
            norm[ks] = vs
        parts = [self._slide_part(i) for i in
                 (slides or range(1, len(self.order) + 1))]
        total = 0
        for part in parts:
            xml, n = _remap_superscripts(self._xml(part), norm)
            if n:
                self._set_xml(part, xml)
                total += n
        if total:
            self._invalidate()
        return total

    # ------------------------------------------------------------------
    # 页操作：删除 / 移动 / 复制
    # ------------------------------------------------------------------
    def delete_slide(self, slide_no):
        """删除一页（1-based），并清理其独占的 notesSlide 与关系/类型登记。"""
        part = self._slide_part(slide_no)
        notes = self._notes_part_of(part)
        self.order.remove(part)
        self.parts.pop(part, None)
        self.parts.pop(_rels_path(part), None)
        # notesSlide 正常情况下 1:1；只有确认没有别的页引用它才删
        if notes and notes in self.parts:
            used = False
            for other in self.order:
                if self._notes_part_of(other) == notes:
                    used = True
                    break
            if not used:
                self.parts.pop(notes, None)
                self.parts.pop(_rels_path(notes), None)
        self._invalidate()

    def move_slide(self, src, dst):
        """把第 src 页移到第 dst 位（均为 1-based 最终位置）。dst 越界自动收敛。"""
        part = self._slide_part(src)
        if not isinstance(dst, int) or isinstance(dst, bool):
            raise EditError("move_slide 的 dst 必须是 int（1-based）")
        dst = max(1, min(dst, len(self.order)))
        self.order.remove(part)
        self.order.insert(dst - 1, part)
        self._invalidate()

    def duplicate_slide(self, slide_no):
        """复制一页并紧随原页（含备注副本、图片/版式关系）。"""
        part = self._slide_part(slide_no)
        idx = self.order.index(part)
        new_part = self._next_part_name("ppt/slides", "slide_copy", ".xml")
        self.parts[new_part] = self.parts[part]
        old_rows = self._rels_of(part)
        new_rows = []
        for r in old_rows:
            short = _short_type(r["type"])
            if short == "notesSlide":
                old_notes = _abs_target(part, r["target"])
                if old_notes in self.parts:
                    new_notes = self._next_part_name(
                        "ppt/notesSlides", "notesSlide_copy", ".xml")
                    self.parts[new_notes] = self.parts[old_notes]
                    # 备注副本内部的 slide 回指必须改到新页
                    nrows = []
                    for nr in self._rels_of(old_notes):
                        nr2 = dict(nr)
                        if _short_type(nr["type"]) == "slide":
                            nr2["target"] = posixpath.relpath(
                                new_part, posixpath.dirname(new_notes))
                        nrows.append(nr2)
                    self._set_rels_of(new_notes, nrows)
                    new_rows.append({
                        "id": r["id"], "type": r["type"],
                        "target": posixpath.relpath(
                            new_notes, posixpath.dirname(new_part)),
                        "mode": "Internal"})
                    continue
            new_rows.append(dict(r))
        self._set_rels_of(new_part, new_rows)
        self.order.insert(idx + 1, new_part)
        self._invalidate()

    # ------------------------------------------------------------------
    # 备注：set_notes（新建或更新 notesSlide）
    # ------------------------------------------------------------------
    def _template_notes_xml(self):
        for name in sorted(self.parts):
            if name.startswith("ppt/notesSlides/") and name.endswith(".xml"):
                return self._xml(name)
        return MIN_NOTES_XML

    @staticmethod
    def _make_notes_body_sp(uid, text):
        paras = "".join(
            '<a:p><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="%s"/>'
            '<a:t>%s</a:t></a:r></a:p>'
            % (units.sz_of_pt(12), escape(line))
            for line in (text or "").split("\n"))
        return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="备注占位符 1"/>'
                '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
                '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
                '<p:spPr><a:noFill/></p:spPr>'
                '<p:txBody><a:bodyPr/><a:lstStyle/>%s</p:txBody></p:sp>'
                % (uid, paras))

    def _set_notes_body(self, xml, text):
        """在 notesSlide XML 里定位 body 占位符并换文字；没有占位符就补一个。"""
        for _i, tag, s, e, block in _walk_elements(xml):
            if tag != "p:sp":
                continue
            if "<p:ph" not in block or 'type="body"' not in block:
                continue
            new_block = _set_block_paragraphs(block, text)
            if new_block is None:
                continue
            return xml[:s] + new_block + xml[e:]
        sp = self._make_notes_body_sp(_max_shape_id(xml) + 1, text)
        if "</p:spTree>" not in xml:
            raise EditError("notesSlide 缺 </p:spTree>，无法写入备注")
        return xml.replace("</p:spTree>", sp + "</p:spTree>", 1)

    def _update_notes_text(self, notes_part, text):
        xml = self._set_notes_body(self._xml(notes_part), text)
        self._set_xml(notes_part, xml)

    def _create_notes_part(self, slide_part, text):
        notes_name = self._next_part_name("ppt/notesSlides", "notesSlide", ".xml")
        xml = self._set_notes_body(self._template_notes_xml(), text)
        self._set_xml(notes_name, xml)
        master = self._notes_master_part()
        rows = []
        rid = 1
        if master:
            rows.append({
                "id": "rId%d" % rid,
                "type": REL_NOTES_MASTER,
                "target": posixpath.relpath(master, "ppt/notesSlides"),
                "mode": "Internal"})
            rid += 1
        rows.append({
            "id": "rId%d" % rid,
            "type": REL_SLIDE,
            "target": posixpath.relpath(slide_part, "ppt/notesSlides"),
            "mode": "Internal"})
        self._set_rels_of(notes_name, rows)
        if master is None:
            # 没有 notesMaster 的包理论上不允许有 notesSlide；这里仍写出备注，
            # 但明确提示（WPS/PowerPoint 可能忽略）。
            print("[edit] 警告：包内没有 notesMaster，备注可能不被 WPS 显示")
        return notes_name

    def set_notes(self, slide_no, text):
        """新建或更新该页备注（幂等：已有 notesSlide 只改文字，不重复建部件）。"""
        part = self._slide_part(slide_no)
        text = "" if text is None else str(text)
        notes = self._notes_part_of(part)
        if notes and notes in self.parts:
            self._update_notes_text(notes, text)
            self._invalidate()
            return
        notes = self._create_notes_part(part, text)
        rows = [r for r in self._rels_of(part)
                if _short_type(r["type"]) != "notesSlide"]
        rows.append({
            "id": _next_rid(rows),
            "type": REL_NOTES_SLIDE,
            "target": posixpath.relpath(notes, posixpath.dirname(part)),
            "mode": "Internal"})
        self._set_rels_of(part, rows)
        self._invalidate()

    # ------------------------------------------------------------------
    # 统一页面元素：unify_element
    # ------------------------------------------------------------------
    def _import_media(self, data):
        """把图片字节导入包内媒体（按内容哈希去重，幂等）。返回部件名。"""
        data = bytes(data)
        ext = _image_ext(data)
        sha = hashlib.sha1(data).hexdigest()
        if self._media_hash_cache is None:
            cache = {}
            for name, blob in self.parts.items():
                if name.startswith("ppt/media/") and isinstance(blob, bytes):
                    cache[hashlib.sha1(blob).hexdigest()] = name
            self._media_hash_cache = cache
        if sha in self._media_hash_cache:
            return self._media_hash_cache[sha]
        name = "ppt/media/unify_%s.%s" % (sha[:12], ext)
        i = 1
        while name in self.parts and self.parts[name] != data:
            name = "ppt/media/unify_%s_%d.%s" % (sha[:12], i, ext)
            i += 1
        self.parts[name] = data
        self._media_hash_cache[sha] = name
        return name

    def _part_pics(self, part):
        """[ (media 部件名, rect|None) ]：部件内所有图片实例（含 grpSp 内）。"""
        xml = self._xml(part)
        rows = self._rels_of(part)
        out = []
        for _i, tag, s, e, block in _walk_elements(xml):
            if tag != "p:pic":
                continue
            rid = _block_rid(block)
            media = self._target_of_rid(part, rows, rid) if rid else ""
            out.append((media, _block_rect(block)))
        return out

    def _layout_media(self, layout_part):
        return set(m for m, _r in self._part_pics(layout_part) if m)

    def _resolve_layouts(self, target_layouts):
        if target_layouts is None:
            out = []
            for part in self.order:
                lay = self._layout_of(part)
                if lay and lay in self.parts and lay not in out:
                    out.append(lay)
            return out
        if isinstance(target_layouts, (str, bytes)):
            target_layouts = [target_layouts]
        out = []
        for item in target_layouts:
            name = _normalise_layouts([item])[0]
            if name not in self.parts:
                raise EditError("target_layouts 里的版式不存在：%s" % item)
            if name not in out:
                out.append(name)
        return out

    @staticmethod
    def _make_pic_xml(rid, x, y, w, h, uid, name="统一标志"):
        return ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s"/>'
                '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
                '<p:nvPr userDrawn="1"/></p:nvPicPr>'
                '<p:blipFill><a:blip r:embed="%s"/>'
                '<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
                '<p:spPr><a:xfrm><a:off x="%d" y="%d"/>'
                '<a:ext cx="%d" cy="%d"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                '<a:noFill/><a:ln><a:noFill/></a:ln></p:spPr></p:pic>'
                % (uid, escape(name), rid, int(x), int(y), int(w), int(h)))

    def _add_pic_to_part(self, part, media_part, x, y, w, h):
        """在任意部件（layout/master/slide）的 spTree 末尾追加一张图片。"""
        xml = self._xml(part)
        if "</p:spTree>" not in xml:
            raise EditError("部件 %s 没有 </p:spTree>" % part)
        rows = self._rels_of(part)
        rid = _next_rid(rows)
        rows.append({
            "id": rid, "type": REL_IMAGE,
            "target": posixpath.relpath(media_part, posixpath.dirname(part)),
            "mode": "Internal"})
        self._set_rels_of(part, rows)
        uid = _max_shape_id(xml) + 1
        pic = self._make_pic_xml(rid, x, y, w, h, uid)
        self._set_xml(part, xml.replace("</p:spTree>", pic + "</p:spTree>", 1))
        return True

    def unify_element(self, picture_bytes=None, x=None, y=None, w=None, h=None,
                      target_layouts=None, slides=None, match="overlap"):
        """把同一张元素（校徽/标志）统一到 slideLayout，并清掉页面里的重复副本。

        语义
        ----
        · picture_bytes 给定时：图片只在 Layout 定义一次（每个目标版式一份，
          重复调用不会重复添加——按图片内容哈希去重）。图片追加到
          target_layouts 指定版式；target_layouts=None 表示"所有被页面引用
          的版式各一份"。
        · 无论是否给图：删除**页面内**与所属版式同 media 的重复副本（幂等）。
          match="overlap"（默认）要求与版式里同 media 元素矩形相交 >50% 才删
          （与 check.py 的"双图标"判定一致，避免误删页内合法同图）；
          match="media" 只要同 media 就删（严格按字面语义）。
        · 坐标 x/y/w/h 为 EMU（与 read.Shape 一致，可先用 units.cm() 换算）。

        返回 {"added": 加到版式的张数, "removed": 从页面删掉的张数,
              "media": 媒体部件名 or None}。
        """
        added = 0
        removed = 0
        media_part = None
        if picture_bytes is not None:
            if None in (x, y, w, h):
                raise EditError("unify_element 给 picture_bytes 时必须同时给 "
                                "x/y/w/h（EMU）")
            for name, val in (("x", x), ("y", y), ("w", w), ("h", h)):
                units.assert_emu(int(val), name)
            if int(w) <= 0 or int(h) <= 0:
                raise EditError("unify_element 的 w/h 必须为正（EMU）")
            media_part = self._import_media(picture_bytes)
            for layout in self._resolve_layouts(target_layouts):
                if media_part in self._layout_media(layout):
                    continue
                self._add_pic_to_part(layout, media_part, x, y, w, h)
                added += 1
        page_parts = [self._slide_part(i) for i in
                      (slides or range(1, len(self.order) + 1))]
        for part in page_parts:
            layout = self._layout_of(part)
            if not layout or layout not in self.parts:
                continue
            lay_pics = self._part_pics(layout)
            if not lay_pics:
                continue
            media_rects = {}
            for media, rect in lay_pics:
                if media:
                    media_rects.setdefault(media, []).append(rect)
            xml = self._xml(part)
            rows = self._rels_of(part)
            spans = []
            for _i, tag, s, e, block in _walk_elements(xml):
                if tag != "p:pic":
                    continue
                rid = _block_rid(block)
                if not rid:
                    continue
                media = self._target_of_rid(part, rows, rid)
                if media not in media_rects:
                    continue
                if match == "media":
                    drop = True
                else:
                    rect = _block_rect(block)
                    drop = rect is None
                    if rect:
                        for lay_rect in media_rects[media]:
                            if _rect_overlap_ratio(rect, lay_rect) > 0.5:
                                drop = True
                                break
                if drop:
                    spans.append((s, e))
            if spans:
                for s, e in sorted(spans, reverse=True):
                    xml = xml[:s] + xml[e:]
                    removed += 1
                self._set_xml(part, xml)
        if added or removed:
            self._invalidate()
        return {"added": added, "removed": removed, "media": media_part}

    # ------------------------------------------------------------------
    # 保存：原子写 + 基础校验 + check（延迟导入）
    # ------------------------------------------------------------------
    def _flush(self):
        """按 self.order 重建 sldIdLst / presentation.xml.rels /
        [Content_Types].xml（只动这三处，其余部件字节不变）。幂等。"""
        pres_part = "ppt/presentation.xml"
        pres_raw = self._xml(pres_part)
        old_rels = self._rels_of(pres_part)
        rel_by_id = dict((r["id"], r) for r in old_rels)

        old_ids = {}
        for m in re.finditer(r"<p:sldId\b([^>]*)/>", pres_raw):
            attrs = m.group(1)
            sid_m = re.search(r'\bid="(\d+)"', attrs)
            rid_m = re.search(r'\br:id="([^"]+)"', attrs)
            if not (sid_m and rid_m):
                continue
            rel = rel_by_id.get(rid_m.group(1))
            if not rel:
                continue
            old_ids[_abs_target(pres_part, rel["target"])] = int(sid_m.group(1))

        rid_of_part = {}
        # 已占用的 rId 必须包含 presentation.xml.rels 里的**全部**关系
        # （master / theme / notesMaster 也在里面）。2026-10 集成时只收集了
        # slide 类型，next_rid() 会从 rId1 起分配，正好撞上 master 的 rId1，
        # check 报"关系 Id 重复：rId1"。
        used_rids = set(r["id"] for r in old_rels)
        for r in old_rels:
            if _short_type(r["type"]) != "slide":
                continue
            tgt = _abs_target(pres_part, r["target"])
            if tgt in self.order and tgt not in rid_of_part:
                rid_of_part[tgt] = r["id"]

        def next_rid():
            i = 1
            while ("rId%d" % i) in used_rids:
                i += 1
            rid = "rId%d" % i
            used_rids.add(rid)
            return rid

        next_id = max(old_ids.values()) + 1 if old_ids else 256
        next_id = max(next_id, 256)
        items = []
        slide_rels = []
        for part in self.order:
            rid = rid_of_part.get(part) or next_rid()
            sid = old_ids.get(part)
            if sid is None:
                sid = next_id
                next_id += 1
            items.append('<p:sldId id="%d" r:id="%s"/>' % (sid, rid))
            slide_rels.append({
                "id": rid, "type": REL_SLIDE,
                "target": posixpath.relpath(part, "ppt"), "mode": "Internal"})
        block = "<p:sldIdLst>%s</p:sldIdLst>" % "".join(items)
        m = re.search(r"<p:sldIdLst\b[^>]*/>|<p:sldIdLst\b[^>]*>.*?</p:sldIdLst>",
                      pres_raw, re.S)
        if m:
            pres_new = pres_raw[:m.start()] + block + pres_raw[m.end():]
        else:
            if "</p:presentation>" not in pres_raw:
                raise EditError("presentation.xml 缺 </p:presentation>")
            pres_new = pres_raw.replace("</p:presentation>",
                                        block + "</p:presentation>", 1)
        self._set_xml(pres_part, pres_new)
        self._set_rels_of(pres_part,
                          [r for r in old_rels
                           if _short_type(r["type"]) != "slide"] + slide_rels)
        self._flush_content_types()

    def _flush_content_types(self):
        ct_part = "[Content_Types].xml"
        if ct_part not in self.parts:
            raise EditError("缺少 [Content_Types].xml")
        try:
            root = ET.fromstring(self.parts[ct_part])
        except ET.ParseError as e:
            raise EditError("[Content_Types].xml 无法解析：%s" % e)
        defaults = []
        overrides = []
        for el in root:
            tag = el.tag.rsplit("}", 1)[-1]
            if tag == "Default":
                defaults.append((el.get("Extension") or "",
                                 el.get("ContentType") or ""))
            elif tag == "Override":
                name = (el.get("PartName") or "").lstrip("/")
                if name:
                    overrides.append((name, el.get("ContentType") or ""))
        part_names = set(p for p in self.parts if p != ct_part)
        ov_map = dict((p, c) for p, c in overrides if p in part_names)
        for part in sorted(part_names):
            if not part.endswith(".xml") or part in ov_map:
                continue
            ct = _guess_content_type(part)
            if ct:
                ov_map[part] = ct
        def_map = {}
        for ext, ct in defaults:
            if ext:
                def_map[ext.lower()] = (ext, ct)
        if "xml" not in def_map:
            def_map["xml"] = ("xml", "application/xml")
        if "rels" not in def_map:
            def_map["rels"] = ("rels", CT_RELS)
        for part in part_names:
            if part.endswith((".xml", ".rels")):
                continue
            ext = part.rsplit(".", 1)[-1]
            if ext.lower() in def_map:
                continue
            ct = _guess_media_type(part)
            if ct:
                def_map[ext.lower()] = (ext, ct)
        buf = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<Types xmlns="%s">' % PKG_CT]
        for ext, ct in sorted(def_map.values(), key=lambda x: x[0].lower()):
            buf.append('<Default Extension="%s" ContentType="%s"/>'
                       % (escape(ext), escape(ct)))
        for part in sorted(ov_map):
            buf.append('<Override PartName="/%s" ContentType="%s"/>'
                       % (escape(part), escape(ov_map[part])))
        buf.append("</Types>")
        self.parts[ct_part] = "".join(buf).encode("utf-8")

    def _write_zip(self, dst):
        self._flush()
        order = ["[Content_Types].xml"]
        seen = set(order)
        for n in self._zip_order:
            if n in self.parts and n not in seen:
                order.append(n)
                seen.add(n)
        for n in self.parts:
            if n not in seen:
                order.append(n)
                seen.add(n)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
            for n in order:
                z.writestr(n, self.parts[n])

    # ------------------------------------------------------------------
    # 图片比例：消除拉伸（check 的 IMAGE_STRETCH 对应修复）
    # ------------------------------------------------------------------
    def fit_images(self, slide_no=None, keep="height", tolerance=0.03):
        """按图片像素纵横比修正框尺寸，消除拉伸（保持框中心不动）。

        keep="height"：高度不变、按比例调宽（默认；不会向下压住后续内容）
        keep="width" ：宽度不变、按比例调高（可能压到下方元素，慎用）
        只处理偏差 > tolerance 的图片；带裁切（a:srcRect）的跳过并计数。

        返回 {"fixed": n, "skipped": m,
              "details": [{"page", "name", "before_cm", "after_cm", "dev_before"}]}
        """
        if keep not in ("height", "width"):
            raise EditError("keep 只能是 'height' 或 'width'")
        targets = [self._slide_part(slide_no)] if slide_no else list(self.order)
        fixed, skipped, details = 0, 0, []
        for part in targets:
            xml = self._xml(part)
            rels = dict((r["id"], r) for r in self._rels_of(part))
            changed = False
            for _idx, tag, s, e, block in reversed(list(_walk_elements(xml))):
                if tag != "p:pic":
                    continue
                rid = _block_rid(block)
                rel = rels.get(rid)
                if not rel:
                    skipped += 1
                    continue
                raw = self.parts.get(_abs_target(part, rel["target"]))
                if raw is None:
                    skipped += 1
                    continue
                iw, ih = _image_size(raw)
                if not iw or not ih:
                    skipped += 1
                    continue
                m_sr = re.search(r'<a:srcRect\b([^>]*)/>', block)
                eff_w, eff_h = float(iw), float(ih)
                if m_sr:
                    cut = {"l": 0, "t": 0, "r": 0, "b": 0}
                    for k in cut:
                        mm = re.search(r'\b%s="(-?\d+)"' % k, m_sr.group(1))
                        if mm:
                            cut[k] = int(mm.group(1))
                    eff_w = iw * (1.0 - (cut["l"] + cut["r"]) / 100000.0)
                    eff_h = ih * (1.0 - (cut["t"] + cut["b"]) / 100000.0)
                if eff_w <= 0 or eff_h <= 0:
                    skipped += 1
                    continue
                m_off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', block)
                m_ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', block)
                if not (m_off and m_ext):
                    skipped += 1
                    continue
                x, y = int(m_off.group(1)), int(m_off.group(2))
                cx, cy = int(m_ext.group(1)), int(m_ext.group(2))
                ar_px = eff_w / eff_h
                ar_box = cx / float(cy)
                dev = abs(ar_px / ar_box - 1.0)
                if dev <= tolerance:
                    continue
                if keep == "height":
                    ncx, ncy = int(round(cy * ar_px)), cy
                else:
                    ncx, ncy = cx, int(round(cx / ar_px))
                if ncx <= 0 or ncy <= 0:
                    skipped += 1
                    continue
                nx = x + (cx - ncx) // 2
                ny = y + (cy - ncy) // 2
                new_block = re.sub(r'<a:off x="-?\d+" y="-?\d+"/>',
                                   '<a:off x="%d" y="%d"/>' % (nx, ny), block, count=1)
                new_block = re.sub(r'<a:ext cx="\d+" cy="\d+"/>',
                                   '<a:ext cx="%d" cy="%d"/>' % (ncx, ncy),
                                   new_block, count=1)
                if new_block == block:
                    skipped += 1
                    continue
                xml = xml[:s] + new_block + xml[e:]
                fixed += 1
                changed = True
                details.append({
                    "page": self.order.index(part) + 1,
                    "name": _block_name(block),
                    "before_cm": (round(units.emu_cm(cx), 2),
                                  round(units.emu_cm(cy), 2)),
                    "after_cm": (round(units.emu_cm(ncx), 2),
                                 round(units.emu_cm(ncy), 2)),
                    "dev_before": round(dev, 4),
                })
            if changed:
                self._set_xml(part, xml)
        if fixed:
            self._invalidate()
        return {"fixed": fixed, "skipped": skipped, "details": details}

    def save(self, dst, verify=True):
        """原子写：先写 <dst>.tmp → 基础结构校验 → verify 时跑 check →
        通过后 os.replace。任何失败都保留旧文件并抛 VerifyError。

        verify 参数：
          True（默认）：跑 check；structure/units 类 ERROR 阻止落盘，
                       inherit/layout 类 ERROR（例如源文件自带的"双图标"）
                       只打印提示、不阻断——内容是可以用 unify_element 修的，
                       不该让人连中间稿都存不下来。
          "strict"   ：check 的所有 ERROR 都阻止落盘（交付前用）。
          False      ：不跑 check，只做基础结构校验（仍保证包没坏）。
        """
        dst = os.path.abspath(dst)
        parent = os.path.dirname(dst) or "."
        if not os.path.isdir(parent):
            os.makedirs(parent)
        strict = verify == "strict"
        tmp = dst + ".tmp"
        try:
            self._write_zip(tmp)
            problems = _verify_pptx(tmp)
            if verify:
                hard, soft = self._run_check(tmp, strict=strict)
                problems += hard
                if soft:
                    print("[edit] check 提示 %d 条内容类 ERROR（不阻断保存，"
                          "可用 unify_element 等方法修复）：" % len(soft))
                    for s in soft[:8]:
                        print("  - " + s)
            if problems:
                raise VerifyError(
                    "保存前校验失败（%d 项），已保留旧文件：\n  - %s"
                    % (len(problems), "\n  - ".join(problems[:20])))
            os.replace(tmp, dst)
        except Exception:
            try:
                if os.path.exists(tmp):
                    os.unlink(tmp)
            except OSError:
                pass
            raise
        return dst

    @staticmethod
    def _run_check(path, strict=False):
        """延迟导入 check（队友模块可能未就绪）：未就绪时降级并打印提示。

        返回 (hard, soft)：hard 阻止保存，soft 只提示。按 check.Finding.check
        分类：structure/units 属"包坏了"必须硬失败；inherit/layout 是内容/
        排版问题，默认只提示（strict=True 时也阻断）。
        """
        try:
            import check as _check                     # noqa: F401  (延迟导入)
        except Exception as e:                        # 语法错误等同样降级
            print("[edit] check 未就绪，跳过校验（%s）" % e)
            return [], []
        check_fn = getattr(_check, "check", None)
        if check_fn is None:
            print("[edit] check 未就绪，跳过校验")
            return [], []
        try:
            rep = check_fn(path, level="all")
        except Exception as e:                        # check 自身异常不影响保存
            print("[edit] check 调用失败（%s: %s），跳过校验"
                  % (type(e).__name__, e))
            return [], []
        findings = list(getattr(rep, "findings", []) or [])
        if not findings and getattr(rep, "ok", True) is False:
            # 兼容只有 errors 字符串、没有 findings 的 Report 实现
            return (["check: " + str(e)
                     for e in (getattr(rep, "errors", []) or [])[:10]], [])
        hard, soft = [], []
        for f in findings:
            if getattr(f, "level", "") != "ERROR":
                continue
            text = f.as_text() if hasattr(f, "as_text") else str(f)
            category = getattr(f, "check", "") or ""
            if strict or category in ("structure", "units"):
                hard.append("check: " + text)
            else:
                soft.append("check: " + text)
        if not hard:
            extra = "；另有 %d 条内容类 ERROR（不阻断）" % len(soft) if soft else ""
            print("[edit] check 校验通过：结构/单位 0 ERROR%s" % extra)
        return hard, soft

    # ------------------------------------------------------------------
    # diff
    # ------------------------------------------------------------------
    def diff(self, other_path):
        """页数 / 每页文本 / 图片数 / 备注差异（以 self 为基线，other 为对比）。"""
        mine = self.deck
        theirs = read.Deck.load(os.path.abspath(other_path))
        out = {
            "pages": {"self": len(mine.slides), "other": len(theirs.slides)},
            "slides": [],
            "summary": {},
        }
        same = changed = only_self = only_other = 0
        for i in range(1, max(len(mine.slides), len(theirs.slides)) + 1):
            a = mine.slide(i) if i <= len(mine.slides) else None
            b = theirs.slide(i) if i <= len(theirs.slides) else None
            if a is None:
                only_other += 1
                out["slides"].append({"index": i, "status": "only-other",
                                      "pics": [None, len(b.pictures)],
                                      "text": [None, b.text]})
                continue
            if b is None:
                only_self += 1
                out["slides"].append({"index": i, "status": "only-self",
                                      "pics": [len(a.pictures), None],
                                      "text": [a.text, None]})
                continue
            text_changed = a.text != b.text
            notes_changed = a.notes != b.notes
            pics_changed = len(a.pictures) != len(b.pictures)
            if text_changed or notes_changed or pics_changed:
                changed += 1
                status = "changed"
            else:
                same += 1
                status = "same"
            row = {"index": i, "status": status,
                   "text_changed": text_changed,
                   "notes_changed": notes_changed,
                   "pics": [len(a.pictures), len(b.pictures)]}
            if text_changed:
                row["text_self"] = a.text[:300]
                row["text_other"] = b.text[:300]
            out["slides"].append(row)
        out["summary"] = {"same": same, "changed": changed,
                          "only_self": only_self, "only_other": only_other}
        return out


# ==========================================================================
# 保存前基础校验（不依赖 check.py，保证任何情况下都不落坏包）
# ==========================================================================
def _verify_pptx(path):
    """扫描一个 .pptx 的基础结构问题，返回中文问题列表（空 = 通过）。"""
    problems = []
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return ["无法打开产出的 zip：%s" % e]
    with z:
        names = set(n for n in z.namelist() if not n.endswith("/"))
        try:
            ct = ET.fromstring(z.read("[Content_Types].xml"))
        except Exception as e:
            return ["[Content_Types].xml 无法解析：%s" % e]
        defaults = set()
        ovs = set()
        for el in ct:
            tag = el.tag.rsplit("}", 1)[-1]
            if tag == "Default":
                defaults.add((el.get("Extension") or "").lower())
            elif tag == "Override":
                ovs.add((el.get("PartName") or "").lstrip("/"))
        for n in sorted(names):
            if n == "[Content_Types].xml":
                continue
            ext = n.rsplit(".", 1)[-1].lower()
            if n not in ovs and ext not in defaults:
                problems.append("Content_Types 未覆盖部件：%s" % n)

        # 关系断链（含 slide/notes/layout/master/theme/image）
        for n in sorted(names):
            if not n.endswith(".xml"):
                continue
            try:
                rels = read.read_rels(z, n, names)
            except Exception as e:
                problems.append("%s 的 rels 无法解析：%s" % (n, e))
                continue
            for rid, r in rels.items():
                if r["mode"] == "External":
                    continue
                if r["target"] not in names:
                    problems.append("%s 关系 %s 断链 → %s"
                                    % (n, rid, r["target"]))

        # presentation 页序链
        try:
            pres = ET.fromstring(z.read("ppt/presentation.xml"))
        except Exception as e:
            return problems + ["presentation.xml 无法解析：%s" % e]
        pres_rels = read.read_rels(z, "ppt/presentation.xml", names)
        lst = pres.find("{%s}sldIdLst" % NS_P)
        slide_count = 0
        if lst is None:
            problems.append("presentation.xml 缺 sldIdLst")
        else:
            seen_ids = set()
            seen_parts = set()
            for s in lst:
                slide_count += 1
                sid = s.get("id")
                if sid in seen_ids:
                    problems.append("sldIdLst 里 sldId id 重复：%s" % sid)
                seen_ids.add(sid)
                rid = s.get("{%s}id" % NS_R)
                rel = pres_rels.get(rid)
                if rel is None:
                    problems.append("sldIdLst 的 r:id=%s 没有对应关系" % rid)
                    continue
                target = rel["target"]
                if target not in names:
                    problems.append("sldIdLst 指向的页部件不存在：%s" % target)
                if target in seen_parts:
                    problems.append("同一页部件在 sldIdLst 出现两次：%s"
                                    % target)
                seen_parts.add(target)
            if not seen_ids:
                problems.append("sldIdLst 为空（没有页）")

        # 每个 slide 必须有 layout；每个 notesSlide 必须有回指 slide
        for n in sorted(names):
            if not n.startswith("ppt/slides/") or not n.endswith(".xml"):
                continue
            rels = read.read_rels(z, n, names)
            has_layout = any(r["type"].endswith("slideLayout")
                             for r in rels.values())
            if not has_layout:
                problems.append("%s 没有 slideLayout 关系" % n)
            for r in rels.values():
                if r["type"].endswith("notesSlide"):
                    np = r["target"]
                    nrels = read.read_rels(z, np, names)
                    back = [x for x in nrels.values() if x["type"] == "slide"]
                    if not back:
                        problems.append("%s 缺回指 slide 关系" % np)

        # 数值量级（units.guard_xml 是唯一判定入口）
        for n in sorted(names):
            if not (n.startswith("ppt/slides/")
                    or n.startswith("ppt/slideLayouts/")
                    or n.startswith("ppt/slideMasters/")):
                continue
            if not n.endswith(".xml"):
                continue
            try:
                raw = z.read(n).decode("utf-8")
            except Exception:
                problems.append("%s 不是 UTF-8 XML" % n)
                continue
            problems.extend(units.guard_xml(raw, n))
        if slide_count == 0:
            problems.append("presentation 里一页都没有")
    return problems


# ==========================================================================
# 自测：一键复现任务验收 5 条
# ==========================================================================
def _find_corpus(explicit=None):
    if explicit and os.path.isfile(explicit):
        return os.path.abspath(explicit)
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = []
    env = os.environ.get("MAKEPPT_CORPUS")
    if env:
        candidates.append(env if os.path.isfile(env)
                          else os.path.join(env, "deck-a.pptx"))
    candidates.append(os.path.join(here, "..", "tests", "corpus", "deck-a.pptx"))
    for c in candidates:
        if c and os.path.isfile(c):
            return os.path.abspath(c)
    return None


def _self_test(src=None):
    """一键回归：复制页/删页/改字/备注 + 回读断言 + 结构校验 + unify 幂等。"""
    t_start = time.time()
    src = _find_corpus(src)
    if not src:
        print("SKIP edit.py 自测：未设置 MAKEPPT_CORPUS"
              "（目录或 deck-a.pptx 路径），tests/corpus/deck-a.pptx 也不存在")
        return 0
    outdir = os.path.normpath(os.path.join(_ENGINE_DIR, "..", "out"))
    if not os.path.isdir(outdir):
        os.makedirs(outdir)
    fails = []
    passes = [0]

    def ok(cond, desc):
        if cond:
            passes[0] += 1
            print("  ✓ " + desc)
        else:
            fails.append(desc)
            print("  ✗ " + desc)

    print("=" * 72)
    print("edit.py 自测")
    print("语料：%s" % src)
    print("产物目录：%s" % outdir)
    print("=" * 72)

    # ---------- 1. 打开并记录基线 ----------
    ed = Editor.open(src)
    n0 = len(ed.order)
    deck0 = ed.deck
    pics0 = sum(len(s.pictures) for s in deck0.slides)
    print("\n[1/8] 打开与基线")
    print("  页数=%d 版式=%d 母版=%d 媒体=%d 图片实例=%d 备注页=%d"
          % (n0, len(deck0.layouts), len(deck0.masters), len(deck0.media),
             pics0, len(deck0.notes())))
    ok(n0 >= 5, "读到 %d 页（可做跨页操作）" % n0)
    with zipfile.ZipFile(src) as _z:
        n_parts = len([x for x in _z.namelist() if not x.endswith("/")])
    ok(len(ed.parts) == n_parts, "打开时保留 zip 全部 %d 个部件" % n_parts)

    # ---------- 2. 复制第 2 页 / 删除第 3 页 ----------
    print("\n[2/8] 跨页操作：复制第 2 页，删除第 3 页")
    ed.duplicate_slide(2)
    ok(len(ed.order) == n0 + 1, "复制后页数 %d = %d + 1" % (len(ed.order), n0))
    d_dup = ed.deck
    ok(d_dup.slide(3).text == d_dup.slide(2).text, "副本紧随原页（p3 == p2 文字）")
    ok(d_dup.slide(3).notes == d_dup.slide(2).notes,
       "副本备注独立复制（改副本不影响原页）")
    ed.delete_slide(3)
    ok(len(ed.order) == n0, "删除后页数 %d = %d + 1 - 1" % (len(ed.order), n0))

    # ---------- 3. 改第 1 页一个字 ----------
    print("\n[3/8] 改字：第 1 页 '汇报' → '报告'（replace_text）")
    tx = ed.deck.slide(1).text_shapes
    pt_before = (tx[0].paras[0].runs[0].pt
                 if tx and tx[0].paras and tx[0].paras[0].runs else None)
    cnt = ed.replace_text(1, "汇报", "报告")
    ok(cnt > 0, "替换命中 %d 处" % cnt)
    page1_text = ed.deck.slide(1).text
    ok("报告" in page1_text and "汇报" not in page1_text,
       "改字生效（含'报告'、不含'汇报'）")
    # set_text 也走一遍（按 name 定位后整体改写，保留格式）
    if tx:
        target = tx[0]
        n_set = ed.set_text(1, name=target.name, text="【自测·封面】" + target.text)
        ok(n_set > 0, "set_text(name=%r) 改动 %d 个形状" % (target.name, n_set))
        ok("【自测·封面】" in ed.deck.slide(1).text, "set_text 生效")
        same = [s for s in ed.deck.slide(1).shapes if s.name == target.name]
        pt_after = (same[0].paras[0].runs[0].pt
                    if same and same[0].paras and same[0].paras[0].runs
                    else None)
        ok(pt_before is None or pt_after == pt_before,
           "set_text 保留首 run 字号（%.1fpt）"
           % (pt_before if pt_before is not None else -1))
    else:
        ok(False, "第 1 页没有可定位的文字形状，set_text 未验证")
    # 跨 run 替换（refs/edit-ops.md §二：一个词被拆成多个 run）
    sample = ('<a:p><a:pPr algn="l"/><a:r><a:rPr sz="1200"/>'
              '<a:t>课程汇</a:t></a:r><a:r><a:rPr sz="900" b="1"/>'
              '<a:t>报</a:t></a:r></a:p>')
    s_out, s_n, s_rebuilt = _replace_texts_in_xml(sample, "汇报", "报告")
    ok(s_n == 1 and s_rebuilt == 1 and "课程报告" in s_out,
       "跨 run 替换：'课程汇'+'报' → '课程报告'（重建 %d 段）" % s_rebuilt)

    # ---------- 4. 写备注 ----------
    print("\n[4/8] 备注：给第 2 页写备注（该页原有备注 → 走更新分支）")
    note = "【edit.py 自测】第 2 页备注已更新。\n第二行用于验证多段写入。"
    ed.set_notes(2, note)
    ok(ed.deck.slide(2).notes == note, "备注回读一致（含多段）")
    # 找一页原本没有备注的，验证"新建 notesSlide"分支
    no_note = [i for i in range(1, len(ed.order) + 1)
               if not ed.deck.slide(i).notes]
    if no_note:
        idx = no_note[0]
        ed.set_notes(idx, "新建 notesSlide 测试")
        ok(ed.deck.slide(idx).notes == "新建 notesSlide 测试",
           "第 %d 页原无备注 → 新建 notesSlide 成功" % idx)
    else:
        print("  （本语料每页都有备注，新建分支略过）")

    # ---------- 5. 保存（原子写 + 校验） ----------
    print("\n[5/8] 保存：%s（verify=True；结构类 ERROR 硬失败、"
          "内容类 ERROR 只提示）" % os.path.join(outdir, "test_edit.pptx"))
    dst = os.path.join(outdir, "test_edit.pptx")
    ed.save(dst, verify=True)
    ok(os.path.isfile(dst), "产出存在：%s" % dst)
    ok(not os.path.exists(dst + ".tmp"), "临时文件已清理（原子写完成）")

    # ---------- 6. 回读断言（验收 2） ----------
    print("\n[6/8] read.Deck.load 回读（验收 2）")
    d = read.Deck.load(dst)
    want = n0 + 1 - 1
    ok(len(d.slides) == want,
       "页数 = %d（%d + 1 复制 - 1 删除）" % (len(d.slides), n0))
    ok("报告" in d.slide(1).text, "第 1 页改字持久化")
    ok(d.slide(2).notes.startswith("【edit.py 自测】"), "第 2 页备注持久化")
    ok(len(d.slides) == len(Editor.open(dst).order), "页序与 sldIdLst 一致")
    # 字节保真：没被操作的页必须逐字节不变（不整体重排 XML）
    ed_out = Editor.open(dst)
    untouched = ed.order[5] if len(ed.order) > 5 else None
    ok(untouched is not None
       and ed_out.parts.get(untouched) == ed.parts.get(untouched),
       "未触碰的页字节级不变（%s）" % untouched)

    # ---------- 7. 结构校验（验收 3） ----------
    print("\n[7/8] 结构校验：关系断链 / 数值量级 / Content_Types（验收 3）")
    problems = _verify_pptx(dst)
    ok(not problems, "基础结构 0 问题（断链、CT 覆盖、guard_xml）")
    if problems:
        for p in problems[:10]:
            print("      ! %s" % p)
    bad_xml = []
    n_xml = 0
    with zipfile.ZipFile(dst) as _z:
        crc_bad = _z.testzip()
        for n in _z.namelist():
            if not n.endswith(".xml"):
                continue
            n_xml += 1
            try:
                ET.fromstring(_z.read(n))
            except Exception as e:
                bad_xml.append((n, e))
    ok(not bad_xml and crc_bad is None,
       "全部 XML 良构 + zip CRC 正常（%d 个 xml 部件）" % n_xml)
    hard, soft = Editor._run_check(dst)          # check 未就绪时自动降级
    ok(not hard, "check 结构/单位 0 ERROR（断链 0、数值 0；未就绪时降级）")
    try:                                         # 直接调 check，打印覆盖对象数
        import check as _check_mod
        rep = _check_mod.check(dst, level="all")
        st = rep.stats or {}
        print("  check 报告：%d 页 | %d 版式 | %d 形状 | %d 图片 | "
              "%d ERROR / %d WARN | %.2fs"
              % (st.get("slides", 0), st.get("layouts", 0),
                 st.get("shapes", 0), st.get("pics", 0),
                 rep.count("ERROR"), rep.count("WARN"), rep.elapsed))
    except Exception as e:
        print("  （check 未就绪，跳过报告：%s）" % type(e).__name__)
    if soft:
        print("  源文件既有的内容类 ERROR %d 条（不阻断保存，如'双图标'）："
              % len(soft))
        for s in soft[:4]:
            print("    - " + s[:110])

    # ---------- 8. unify_element 幂等（验收 4） ----------
    print("\n[8/8] unify_element 幂等（验收 4）")
    ed2 = Editor.open(src)
    r1 = ed2.unify_element()                    # 只清页面内重复副本
    d_u1 = ed2.deck
    pics1 = sum(len(s.pictures) for s in d_u1.slides)
    p1 = os.path.join(outdir, "test_unify_1.pptx")
    ed2.save(p1, verify=False)
    print("  第一次 unify：added=%d removed=%d 图片实例 %d → %d"
          % (r1["added"], r1["removed"], pics0, pics1))
    if r1["removed"] > 0:
        ok(True, "第一次删除了页面内与版式重复的图片 %d 张" % r1["removed"])
    else:
        # 语料可能已被修过：此时 unify 无可删，改由 check 证明"确实没有重复"
        try:
            import check as _ck
            _dup = [e for e in _ck.check(src).errors if "双图标" in e]
            ok(not _dup, "语料已无双图标（unify 无可删，check 也报 0 条）")
        except Exception as _e:
            print("  （check 不可用，跳过无双图标校验：%s）" % _e)
    ed3 = Editor.open(p1)
    r2 = ed3.unify_element()
    d_u2 = ed3.deck
    pics2 = sum(len(s.pictures) for s in d_u2.slides)
    p2 = os.path.join(outdir, "test_unify_deck-b.pptx")
    ed3.save(p2, verify=False)
    print("  第二次 unify：added=%d removed=%d 图片实例 %d"
          % (r2["added"], r2["removed"], pics2))
    ok(r2["removed"] == 0, "第二次 removed=0（幂等）")
    ok(pics2 == pics1, "第二次页面内图片数不再减少（%d == %d）"
       % (pics2, pics1))

    # ---------- 9. 修复闭环：源文件自带双图标 → unify 修掉 → strict 校验 ----------
    print("\n[9/9] 修复闭环：unify 清掉源文件自带双图标后 strict 交付校验")
    ed9 = Editor.open(dst)
    r_fix = ed9.unify_element()
    fixed = os.path.join(outdir, "test_edit_fixed.pptx")
    print("  在 test_edit 上再删掉 %d 张与版式重复的校徽（源文件既有问题）"
          % r_fix["removed"])
    ed9.save(fixed, verify="strict")
    hard9, soft9 = Editor._run_check(fixed, strict=True)
    ok(not hard9 and not soft9,
       "strict 校验 0 ERROR（结构/单位/继承/布局全部通过）")
    d_fix = read.Deck.load(fixed)
    ok(len(d_fix.slides) == want, "修复产物页数不变（%d）" % len(d_fix.slides))

    # ---------- 10. 原子写：校验失败保留旧文件 ----------
    print("\n[10] 原子写：故意丢一个页部件 → 校验失败 → 旧文件保留")
    atomic = os.path.join(outdir, "test_atomic.pptx")
    edA = Editor.open(src)
    edA.save(atomic, verify=False)
    before = open(atomic, "rb").read()
    edA.parts.pop(edA.order[0], None)            # 模拟"丢部件"的人为事故
    try:
        edA.save(atomic, verify=False)
        ok(False, "坏包应当被拒绝落盘")
    except VerifyError:
        ok(True, "坏包被拒绝落盘（VerifyError）")
    after = open(atomic, "rb").read()
    ok(before == after, "旧文件字节未变（原子写生效，%d 字节）" % len(after))
    ok(not os.path.exists(atomic + ".tmp"), "临时文件已清理")

    # ---------- 11. check 未就绪时降级（任务硬要求：不能崩） ----------
    print("\n[11] check 未就绪降级：save(verify=True) 打印提示并继续")
    edB = Editor.open(src)
    dstB = os.path.join(outdir, "test_nocheck.pptx")
    real_check = sys.modules.get("check", "MISSING")
    sys.modules["check"] = None                 # 模拟"check.py 还没写好"
    degraded = False
    try:
        edB.save(dstB, verify=True)
        degraded = True
    except Exception as e:                      # noqa: BLE001
        print("  异常：%s: %s" % (type(e).__name__, e))
    finally:
        if real_check == "MISSING":
            sys.modules.pop("check", None)
        else:
            sys.modules["check"] = real_check
    ok(degraded and os.path.isfile(dstB),
       "check 缺失时降级保存成功（打印'check 未就绪，跳过校验'）")

    # ---------- 附加：layout 只定义一次 + 页内去重 ----------
    print("\n[附加 A] unify_element 在 Layout 定义图片：只加一次")
    ed4 = Editor.open(src)
    logo = None
    for name, blob in ed4.parts.items():
        if name.startswith("ppt/media/") and name.endswith(".png"):
            for m in re.finditer(r'<p:pic\b[^>]*>.*?</p:pic>',
                                 ed4._xml("ppt/slideLayouts/slideLayout2.xml")
                                 if "ppt/slideLayouts/slideLayout2.xml" in ed4.parts
                                 else "", re.S):
                if m.group(0).count('r:embed="rId2"'):
                    logo = blob
                    break
        if logo:
            break
    if logo is None:
        for name, blob in ed4.parts.items():
            if name.startswith("ppt/media/"):
                logo = blob
                break
    layout2 = "ppt/slideLayouts/slideLayout2.xml"
    layout4 = "ppt/slideLayouts/slideLayout4.xml"
    if logo and layout2 in ed4.parts and layout4 in ed4.parts:
        r3 = ed4.unify_element(
            picture_bytes=logo,
            x=units.cm(23.9), y=units.cm(17.8),
            w=units.cm(6.0), h=units.cm(1.26),
            target_layouts=[layout2])
        ok(r3["added"] == 0,
           "Layout2 已有同 media 图 → 不重复添加（added=0）")
        r4 = ed4.unify_element(
            picture_bytes=logo,
            x=units.cm(23.9), y=units.cm(17.8),
            w=units.cm(6.0), h=units.cm(1.26),
            target_layouts=[layout4])
        ok(r4["added"] == 1, "Layout4 无图 → 添加一次（added=1）")
        r5 = ed4.unify_element(
            picture_bytes=logo,
            x=units.cm(23.9), y=units.cm(17.8),
            w=units.cm(6.0), h=units.cm(1.26),
            target_layouts=[layout4])
        ok(r5["added"] == 0, "同一 Layout 连跑两次 → 第二次 added=0（幂等）")
        p3 = os.path.join(outdir, "test_unify_layout.pptx")
        ed4.save(p3, verify=False)
        ok(os.path.isfile(p3), "Layout 统一产物：%s" % p3)
    else:
        print("  （语料缺 layout2/layout4，附加 A 略过）")

    # ---------- 附加 B：move_slide + diff + 上标重映射 ----------
    print("\n[附加 B] move_slide / diff / renumber_superscripts")
    ed5 = Editor.open(src)
    if len(ed5.order) >= 3:
        first = ed5.order[0]
        ed5.move_slide(1, 3)
        ok(ed5.order[2] == first, "move_slide(1,3) 后原第 1 页在第 3 位")
        ed5.move_slide(3, 1)
        ok(ed5.order[0] == first, "再 move_slide(3,1) 复原")
    dd = ed5.diff(dst)
    ok(dd["pages"]["self"] == n0, "diff 页数：self=%d" % dd["pages"]["self"])
    print("  diff 摘要：%s" % dd["summary"])
    ed6 = Editor.open(src)
    hit_page = None
    for i in range(1, len(ed6.order) + 1):
        if re.search(r"[¹²³⁴⁵⁶⁷⁸⁹]", ed6.deck.slide(i).text):
            hit_page = i
            break
    if hit_page:
        n_sup = ed6.renumber_superscripts({1: 4, 2: 5})
        ok(n_sup > 0, "上标重映射：第 %d 页命中，改动 %d 处"
           % (hit_page, n_sup))
    else:
        # 语料没有上标时，构造一个再验
        ed6.set_text(1, contains="汇报", keep_format=False,
                     text="上标测试¹²")
        n_sup = ed6.renumber_superscripts({"¹": "⁴", "²": "⁵"})
        ok(n_sup == 2 and "上标测试⁴⁵" in ed6.deck.slide(1).text,
           "构造上标后重映射：改动 %d 处，结果含'⁴⁵'" % n_sup)

    # ---------- 附加 C：特殊字符转义 ----------
    print("\n[附加 C] 特殊字符（< & > 引号）写入 / 回读一致")
    edC = Editor.open(src)
    tx3 = edC.deck.slide(3).text_shapes
    if tx3:
        weird = 'A<B>C&D "引号" & <tag>'
        edC.set_text(3, name=tx3[0].name, text=weird)
        dstC = os.path.join(outdir, "test_escape.pptx")
        edC.save(dstC, verify=False)
        dC = read.Deck.load(dstC)
        ok(weird in dC.slide(3).text, "特殊字符转义正确：%r" % weird)
    else:
        print("  （第 3 页无文字形状，附加 C 略过）")

    # ---------- 附加 D：图片比例（消除拉伸） ----------
    print("\n[附加 D] fit_images：先制造一张拉伸图，再按像素比例修正")
    edD = Editor.open(src)
    page = None
    for sl in edD.deck.slides:
        if sl.pictures:
            page = sl.index
            break
    if page is None:
        print("  （语料无图片页，附加 D 略过）")
    else:
        part = edD._slide_part(page)
        xml = edD._xml(part)
        hit = False
        for _i, tag, s0, e0, block in _walk_elements(xml):
            if tag != "p:pic":
                continue
            m = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', block)
            if not m:
                continue
            cx, cy = int(m.group(1)), int(m.group(2))
            bad = re.sub(r'<a:ext cx="\d+" cy="\d+"/>',
                         '<a:ext cx="%d" cy="%d"/>' % (int(cx * 1.5), cy),
                         block, count=1)
            xml = xml[:s0] + bad + xml[e0:]
            edD._set_xml(part, xml)
            edD._invalidate()
            hit = True
            break
        if not hit:
            print("  （没有可改的图片，附加 D 略过）")
        else:
            rD = edD.fit_images(page, keep="height")
            ok(rD["fixed"] == 1,
               "fit_images 修正 1 张拉伸图片（实际 %d）" % rD["fixed"])
            if rD["details"]:
                det = rD["details"][0]
                ok(det["after_cm"][0] < det["before_cm"][0],
                   "宽度按比例收窄（%.2f → %.2f cm）"
                   % (det["before_cm"][0], det["after_cm"][0]))
                ok(abs(det["after_cm"][1] - det["before_cm"][1]) < 0.01,
                   "高度不变（keep='height'）：%.2f cm" % det["after_cm"][1])
            dstD = os.path.join(outdir, "test_fit_images.pptx")
            edD.save(dstD, verify=False)
            ok(os.path.isfile(dstD), "产物已写出：test_fit_images.pptx")

    # ---------- 汇总 ----------
    elapsed = time.time() - t_start
    print("\n" + "=" * 72)
    if fails:
        print("SELF-TEST FAILED：%d 通过 / %d 失败（%.1fs）"
              % (passes[0], len(fails), elapsed))
        for f in fails:
            print("  ✗ " + f)
        print("=" * 72)
        return 1
    print("SELF-TEST OK — %d 项断言全部通过（%.1fs）" % (passes[0], elapsed))
    print("验收产物：")
    for name in ("test_edit.pptx", "test_edit_fixed.pptx",
                 "test_unify_1.pptx", "test_unify_deck-b.pptx",
                 "test_unify_layout.pptx", "test_escape.pptx",
                 "test_nocheck.pptx", "test_fit_images.pptx"):
        p = os.path.join(outdir, name)
        if os.path.isfile(p):
            print("  %s（%d 字节）" % (p, os.path.getsize(p)))
    print("=" * 72)
    return 0


# ==========================================================================
# 命令行
# ==========================================================================
def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    src = argv[0]
    rest = argv[1:]
    if "--self-test" in rest:
        return _self_test(src)
    if "--summary" in rest:
        d = read.Deck.load(src)
        print(json.dumps(d.summary(), ensure_ascii=False, indent=1))
        return 0
    if "--diff" in rest:
        i = rest.index("--diff")
        if i + 1 >= len(rest):
            print("用法：python edit.py <src.pptx> --diff <other.pptx>")
            return 2
        ed = Editor.open(src)
        print(json.dumps(ed.diff(rest[i + 1]), ensure_ascii=False, indent=1))
        return 0
    print("未知参数。支持：--self-test / --summary / --diff <other.pptx>")
    print(__doc__)
    return 2


def _strip_self_args(argv):
    """去掉嵌入执行器塞进 sys.argv 的脚本自身路径（file/exec 模式兼容）。"""
    try:
        self_abs = os.path.abspath(__file__)
    except NameError:
        self_abs = os.path.join(_ENGINE_DIR, "edit.py")
    out = []
    for a in argv:
        try:
            if os.path.abspath(a) == self_abs:
                continue
        except (OSError, ValueError):
            pass
        out.append(a)
    return out


if __name__ == "__main__":
    sys.exit(main(_strip_self_args(sys.argv[1:])))
