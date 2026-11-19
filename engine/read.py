# -*- coding: utf-8 -*-
"""read.py — 把现有 .pptx 读成只读 IR（Deck / Slide / Shape / Para / Run）。

这是改造与检查的共同基础：**不读原文件就没有正确的旧值**。
所有上层模块（edit / check / preview / merge）一律通过本模块拿数据，
不得各自用正则扫 XML 再去猜结构。

IR 层级
-------
    Deck
     ├─ w, h            画布（EMU）
     ├─ slides[]        Slide（按 presentation 的 sldIdLst 顺序）
     │   ├─ part        ppt/slides/slideN.xml
     │   ├─ layout/master  继承部件路径
     │   ├─ shapes[]    页面自身的元素（含 grpSp 内递归展开、坐标已换算到页坐标）
     │   ├─ inherited[] 版式/母版上的固定元素（图片、装饰条等；双图标检查就看这里）
     │   └─ notes        备注纯文本
     ├─ parts           全部部件名
     └─ media           媒体部件 → 字节数

用法
----
    import read
    deck = read.Deck.load("a.pptx")
    print(deck.w, deck.h, len(deck.slides))
    for s in deck.slides:
        print(s.index, s.part, len(s.shapes), len(s.text))
    deck.summary()          # dict，可直接打印/落盘
    deck.to_json()          # 完整 IR，调试与 diff 用

自测：`python3 read.py --self-test <某个pptx>`
"""
import json
import os
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

_ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
if _ENGINE_DIR not in sys.path:
    sys.path.insert(0, _ENGINE_DIR)

try:                       # 支持 `import read` 与 `from engine import read` 两种用法
    from . import units
except ImportError:        # pragma: no cover
    import units

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def q(tag):
    """'p:sp' → '{ns}p:sp' 形式。"""
    pre, local = tag.split(":")
    return "{%s}%s" % (NS[pre], local)


RT = {
    "slide": "/slide",
    "slideLayout": "/slideLayout",
    "slideMaster": "/slideMaster",
    "notesSlide": "/notesSlide",
    "image": "/image",
    "theme": "/theme",
}


# --------------------------------------------------------------------------
# 底层：zip / rels / 部件解析
# --------------------------------------------------------------------------
def load_zip(path):
    return zipfile.ZipFile(path)


def part_names(z):
    return set(n for n in z.namelist() if not n.endswith("/"))


def read_rels(z, part, names=None):
    """读一个部件的 _rels。返回 {rId: {"target":绝对部件名, "type":短类型, "mode":...}}。"""
    d, b = posixpath.split(part)
    rp = posixpath.join(d, "_rels", b + ".rels")
    if names is not None and rp not in names:
        return {}
    try:
        raw = z.read(rp)
    except KeyError:
        return {}
    out = {}
    for rel in ET.fromstring(raw):
        t = rel.get("Target") or ""
        mode = rel.get("TargetMode") or "Internal"
        if mode != "External" and not t.startswith(("http://", "https://")):
            t = posixpath.normpath(posixpath.join(d, t))
        out[rel.get("Id")] = {
            "target": t,
            "type": (rel.get("Type") or "").rsplit("/", 1)[-1],
            "mode": mode,
        }
    return out


def rels_of_type(rels, short_type):
    return [(rid, r["target"]) for rid, r in rels.items() if r["type"] == short_type]


def slide_parts(z, names=None):
    """按 presentation 的 sldIdLst 返回页序（部件路径）。"""
    names = names or part_names(z)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = read_rels(z, "ppt/presentation.xml", names)
    out = []
    for s in pres.findall(q("p:sldIdLst") + "/" + q("p:sldId")):
        rid = s.get(q("r:id"))
        if rid in rels:
            out.append(rels[rid]["target"])
    return out


def slide_size(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    sz = pres.find(q("p:sldSz"))
    return int(sz.get("cx")), int(sz.get("cy"))


def parse_xml(raw):
    return ET.fromstring(raw)


def text_of_element(el):
    """元素内全部 a:t 拼接（段落之间用 \\n）。"""
    if el is None:
        return ""
    paras = []
    for p in el.iter(q("a:p")):
        s = "".join(t.text or "" for t in p.iter(q("a:t")))
        if s or paras:
            paras.append(s)
    return "\n".join(paras).strip()


def plain_text(raw_or_el):
    """原始 XML 字节/字符串 → 纯文本。"""
    el = parse_xml(raw_or_el) if isinstance(raw_or_el, (bytes, str)) else raw_or_el
    return text_of_element(el)


# --------------------------------------------------------------------------
# IR
# --------------------------------------------------------------------------
@dataclass
class Run:
    text: str = ""
    sz: str = ""            # 原始 sz，如 "1200"
    pt: float = 18.0        # units.pt_of_sz 解析结果
    bold: bool = False
    italic: bool = False
    color: str = ""         # "191B1F" 或 "scheme:accent1"
    font: str = ""

    def as_dict(self):
        return {"text": self.text, "pt": self.pt, "bold": self.bold,
                "color": self.color, "font": self.font}


@dataclass
class Para:
    runs: list = field(default_factory=list)
    text: str = ""
    align: str = ""
    level: int = 0
    spc_pct: str = ""        # 原始 spcPct val
    spc_ratio: float = 1.0   # 解析后的倍数
    spc_pts: str = ""        # 固定行距原始值
    has_bullet: bool = False

    def as_dict(self):
        return {"text": self.text, "align": self.align, "level": self.level,
                "spc_ratio": self.spc_ratio,
                "runs": [r.as_dict() for r in self.runs]}


@dataclass
class Shape:
    idx: int = 0
    kind: str = ""           # sp / pic / graphicFrame / grpSp / cxnSp / other
    name: str = ""
    x: int = None
    y: int = None
    w: int = None
    h: int = None
    geom: str = ""
    fill: str = ""
    line: str = ""
    style_fill: str = ""     # p:style/a:fillRef（主题引用），如 "scheme:accent1"
    style_line: str = ""     # p:style/a:lnRef
    line_w: int = 0          # a:ln/@w（EMU）
    flip_h: bool = False
    flip_v: bool = False
    anchor: str = ""         # bodyPr@anchor：t / ctr / b / just
    insets: tuple = (0, 0, 0, 0)   # bodyPr 的 l/t/r/b Ins（EMU）
    paras: list = field(default_factory=list)
    rid: str = ""
    media: str = ""
    media_size: int = 0
    crop: tuple = ()         # a:srcRect 的 (l,t,r,b)，千分之一百分比
    src: str = "slide"       # slide / layout / master
    placeholder: str = ""    # ph type / idx

    @property
    def text(self):
        return "\n".join(p.text for p in self.paras).strip()

    def pt(self):
        """本形状最大字号（用于体检时的字体基准）。"""
        vals = [r.pt for p in self.paras for r in p.runs if r.pt]
        return max(vals) if vals else 18.0

    def rect(self):
        if None in (self.x, self.y, self.w, self.h):
            return None
        return (self.x, self.y, self.w, self.h)

    def as_dict(self):
        d = {"idx": self.idx, "kind": self.kind, "name": self.name,
             "geom": self.geom, "fill": self.fill, "line": self.line,
             "line_w": self.line_w, "flip_h": self.flip_h, "flip_v": self.flip_v,
             "style_fill": self.style_fill, "style_line": self.style_line,
             "anchor": self.anchor, "insets": list(self.insets),
             "x": self.x, "y": self.y, "w": self.w, "h": self.h,
             "text": self.text, "src": self.src,
             "paras": [p.as_dict() for p in self.paras]}
        if self.rid:
            d["rid"] = self.rid
            d["media"] = self.media
            d["crop"] = list(self.crop)
        if self.placeholder:
            d["placeholder"] = self.placeholder
        return d


@dataclass
class Slide:
    index: int = 0
    part: str = ""
    layout: str = ""
    master: str = ""
    shapes: list = field(default_factory=list)
    inherited: list = field(default_factory=list)
    notes: str = ""
    notes_part: str = ""
    bg: str = ""
    bg_image: str = ""       # 背景图片部件（p:bg/p:bgPr/a:blipFill）

    @property
    def text(self):
        return "\n".join(s.text for s in self.shapes if s.text)

    @property
    def pictures(self):
        return [s for s in self.shapes if s.kind == "pic"]

    @property
    def text_shapes(self):
        return [s for s in self.shapes if s.paras and s.text.strip()]

    def shapes_named(self, name):
        return [s for s in self.shapes if s.name == name]

    def as_dict(self):
        return {"index": self.index, "part": self.part, "layout": self.layout,
                "master": self.master, "notes": self.notes,
                "bg": self.bg, "bg_image": self.bg_image,
                "shapes": [s.as_dict() for s in self.shapes],
                "inherited": [s.as_dict() for s in self.inherited]}


@dataclass
class Deck:
    path: str = ""
    w: int = units.SLIDE_W_EMU
    h: int = units.SLIDE_H_EMU
    slides: list = field(default_factory=list)
    parts: set = field(default_factory=set)
    media: dict = field(default_factory=dict)   # 部件名 → 字节数
    layouts: list = field(default_factory=list)
    masters: list = field(default_factory=list)

    # -------- 加载 --------
    @classmethod
    def load(cls, path):
        z = load_zip(path)
        names = part_names(z)
        w, h = slide_size(z)
        units.assert_canvas(w, h, path)
        d = cls(path=path, w=w, h=h, parts=names)

        # 媒体清单
        for n in names:
            if n.startswith("ppt/media/"):
                d.media[n] = z.getinfo(n).file_size
        # 版式与母版清单
        d.layouts = sorted(n for n in names if n.startswith("ppt/slideLayouts/")
                           and n.endswith(".xml"))
        d.masters = sorted(n for n in names if n.startswith("ppt/slideMasters/")
                           and n.endswith(".xml"))

        # 继承元素按 layout 缓存，避免 55 页重复解析
        layout_cache = {}
        for i, part in enumerate(slide_parts(z, names), 1):
            sl = _parse_slide(z, part, i, names, layout_cache)
            d.slides.append(sl)
        z.close()
        return d

    # -------- 便捷查询 --------
    @property
    def n(self):
        return len(self.slides)

    def slide(self, index):
        """1-based 取页。"""
        if not (1 <= index <= len(self.slides)):
            raise IndexError("页号 %s 超出 1—%d" % (index, len(self.slides)))
        return self.slides[index - 1]

    def all_shapes(self, src="slide"):
        for s in self.slides:
            for sh in (s.shapes if src == "slide" else s.inherited):
                yield s, sh

    def texts(self):
        return [(s.index, sh.name, sh.text) for s, sh in self.all_shapes() if sh.text]

    def notes(self):
        return {s.index: s.notes for s in self.slides if s.notes}

    def summary(self):
        return {
            "path": self.path,
            "canvas_emu": [self.w, self.h],
            "canvas_cm": [round(units.emu_cm(self.w), 2), round(units.emu_cm(self.h), 2)],
            "slides": len(self.slides),
            "layouts": len(self.layouts),
            "masters": len(self.masters),
            "media": len(self.media),
            "per_slide": [
                {"i": s.index, "shapes": len(s.shapes),
                 "text_chars": len(s.text), "pics": len(s.pictures),
                 "notes_chars": len(s.notes)}
                for s in self.slides
            ],
        }

    def to_json(self):
        return json.dumps({"summary": self.summary(),
                           "slides": [s.as_dict() for s in self.slides]},
                          ensure_ascii=False, indent=1)


# --------------------------------------------------------------------------
# 解析实现
# --------------------------------------------------------------------------
def _parse_color(el):
    """从 solidFill / ln 元素取颜色。"""
    if el is None:
        return ""
    srgb = el.find(".//" + q("a:srgbClr"))
    if srgb is not None:
        return (srgb.get("val") or "").upper()
    sch = el.find(".//" + q("a:schemeClr"))
    if sch is not None:
        return "scheme:" + (sch.get("val") or "")
    return ""


def _parse_paras(txBody):
    out = []
    if txBody is None:
        return out
    for p in txBody.findall(q("a:p")):
        pr = p.find(q("a:pPr"))
        para = Para()
        if pr is not None:
            para.align = pr.get("algn") or ""
            para.level = int(pr.get("lvl") or 0)
            ln = pr.find(q("a:lnSpc"))
            if ln is not None:
                pct = ln.find(q("a:spcPct"))
                pts = ln.find(q("a:spcPts"))
                if pct is not None:
                    para.spc_pct = pct.get("val") or ""
                    para.spc_ratio = units.ratio_of_spc_pct(para.spc_pct)
                if pts is not None:
                    para.spc_pts = pts.get("val") or ""
            for tag in ("a:buChar", "a:buAutoNum", "a:buBlip"):
                if pr.find(q(tag)) is not None:
                    para.has_bullet = True
        for r in p.findall(q("a:r")):
            run = Run()
            t = r.find(q("a:t"))
            run.text = t.text or "" if t is not None else ""
            rpr = r.find(q("a:rPr"))
            if rpr is not None:
                run.sz = rpr.get("sz") or ""
                run.pt = units.pt_of_sz(run.sz) if run.sz else units.pt_of_sz(None)
                run.bold = rpr.get("b") == "1"
                run.italic = rpr.get("i") == "1"
                run.color = _parse_color(rpr)
                latin = rpr.find(q("a:latin"))
                ea = rpr.find(q("a:ea"))
                if latin is not None and latin.get("typeface"):
                    run.font = latin.get("typeface")
                elif ea is not None and ea.get("typeface"):
                    run.font = ea.get("typeface")
            else:
                # 无 rPr：继承段落默认
                dr = pr.find(q("a:defRPr")) if pr is not None else None
                if dr is not None:
                    run.sz = dr.get("sz") or ""
                    run.pt = units.pt_of_sz(run.sz) if run.sz else 18.0
                else:
                    run.pt = 18.0
            para.runs.append(run)
        para.text = "".join(r.text for r in para.runs)
        if para.text or para.runs:
            out.append(para)
    return out


def _xfrm_of(el, kind):
    """取元素自己的 xfrm（EMU）；返回 (x, y, cx, cy) 或 None。"""
    if kind == "graphicFrame":
        x = el.find(q("p:xfrm"))
        pre = "p:"
    elif kind == "grpSp":
        gpr = el.find(q("p:grpSpPr"))
        x = gpr.find(q("a:xfrm")) if gpr is not None else None
        pre = "p:"
    else:
        sp = el.find(q("p:spPr"))
        x = sp.find(q("a:xfrm")) if sp is not None else None
        pre = "p:"
    if x is None:
        return None
    off = x.find(q("a:off"))
    ext = x.find(q("a:ext"))
    if off is None or ext is None:
        return None
    try:
        return (int(off.get("x")), int(off.get("y")),
                int(ext.get("cx")), int(ext.get("cy")))
    except (TypeError, ValueError):
        return None


def _grp_transform(el):
    """group 的坐标系换算参数：(chOff, chExt, off, ext)。无则返回 None。"""
    gpr = el.find(q("p:grpSpPr"))
    if gpr is None:
        return None
    x = gpr.find(q("a:xfrm"))
    if x is None:
        return None
    off = x.find(q("a:off"))
    ext = x.find(q("a:ext"))
    choff = x.find(q("a:chOff"))
    chext = x.find(q("a:chExt"))
    if None in (off, ext, choff, chext):
        return None
    try:
        return ((int(off.get("x")), int(off.get("y"))),
                (int(ext.get("cx")), int(ext.get("cy"))),
                (int(choff.get("x")), int(choff.get("y"))),
                (int(chext.get("cx")), int(chext.get("cy"))))
    except (TypeError, ValueError):
        return None


def _map_rect(x, y, w, h, tf):
    """把 group 子坐标换算到页坐标。tf = _grp_transform 结果。"""
    if tf is None:
        return x, y, w, h
    (ox, oy), (ex, ey), (cx, cy), (cex, cey) = tf
    sx = ex / cex if cex else 1.0
    sy = ey / cey if cey else 1.0
    return (int(round(ox + (x - cx) * sx)), int(round(oy + (y - cy) * sy)),
            int(round(w * sx)), int(round(h * sy)))


def _ins(bp, key, default):
    """bodyPr 内边距；缺省取 PowerPoint 默认值（EMU）。"""
    if bp is None:
        return default
    v = bp.get(key)
    if v in (None, ""):
        return default
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _shape_kind(el):
    tag = el.tag.rsplit("}", 1)[-1]
    return tag


def _walk(el, idx_counter, tf, rels, z, src):
    """递归解析 spTree 下的元素，返回 Shape 列表（坐标已是页坐标）。"""
    out = []
    for child in list(el):
        kind = _shape_kind(child)
        if kind in ("nvGrpSpPr", "grpSpPr"):
            continue
        if kind == "grpSp":
            sub = _walk(child, idx_counter, tf, rels, z, src)
            out.extend(sub)
            continue
        idx_counter[0] += 1
        sh = Shape(idx=idx_counter[0], kind=kind, src=src)
        nv = child.find(q("p:nvSpPr")) or child.find(q("p:nvPicPr")) \
            or child.find(q("p:nvGraphicFramePr")) or child.find(q("p:nvCxnSpPr"))
        if nv is not None:
            c = nv.find(q("p:cNvPr"))
            if c is not None:
                sh.name = c.get("name") or ""
            ph = nv.find(".//" + q("p:ph"))
            if ph is not None:
                sh.placeholder = ph.get("type") or ("idx:" + str(ph.get("idx")))
        rect = _xfrm_of(child, kind)
        if rect:
            x, y, w, h = _map_rect(*rect, tf)
            sh.x, sh.y, sh.w, sh.h = x, y, w, h

        sp = child.find(q("p:spPr"))
        has_explicit_fill = False
        has_explicit_line = False
        if sp is not None:
            # 显式填充优先于 p:style 的主题引用（OOXML 语义）：
            # 有 solidFill/gradFill/blipFill/pattFill/noFill 任一，就不该再吃 fillRef。
            for tag in ("a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:noFill"):
                if sp.find(q(tag)) is not None:
                    has_explicit_fill = True
                    break
            fill = sp.find(q("a:solidFill"))
            sh.fill = _parse_color(fill) if fill is not None else ""
            ln = sp.find(q("a:ln"))
            sh.line = _parse_color(ln) if ln is not None else ""
            if ln is not None:
                if ln.find(q("a:solidFill")) is not None \
                        or ln.find(q("a:noFill")) is not None:
                    has_explicit_line = True
                if ln.get("w"):
                    try:
                        sh.line_w = int(ln.get("w"))
                    except (TypeError, ValueError):
                        sh.line_w = 0
            xf = sp.find(q("a:xfrm"))
            if xf is not None:
                sh.flip_h = xf.get("flipH") == "1"
                sh.flip_v = xf.get("flipV") == "1"
            geom = sp.find(q("a:prstGeom"))
            if geom is not None:
                sh.geom = geom.get("prst") or ""
        # p:style 里的主题色引用（很多形状 spPr 里没有 fill/ln，颜色全靠这里）。
        # 只在"没有显式填充/线条"时才设：显式 a:noFill 是作者选的无填充，
        # 不能被主题色误填（合并稿里有 3 个大矩形属于这一类）。
        st = child.find(q("p:style"))
        if st is not None:
            fr = st.find(q("a:fillRef"))
            if fr is not None and not has_explicit_fill and fr.get("idx") not in ("0",):
                sh.style_fill = _parse_color(fr)
            lr = st.find(q("a:lnRef"))
            if lr is not None and not has_explicit_line and lr.get("idx") not in ("0",):
                sh.style_line = _parse_color(lr)
        if kind == "graphicFrame":
            xf = child.find(q("p:xfrm"))
            if xf is not None:
                sh.flip_h = xf.get("flipH") == "1"
                sh.flip_v = xf.get("flipV") == "1"
        if kind == "pic":
            bf = child.find(q("p:blipFill"))
            if bf is not None:
                sr = bf.find(q("a:srcRect"))
                if sr is not None:
                    try:
                        sh.crop = tuple(int(sr.get(k) or 0) for k in ("l", "t", "r", "b"))
                    except (TypeError, ValueError):
                        sh.crop = ()
            blip = child.find(".//" + q("a:blip"))
            if blip is not None:
                sh.rid = blip.get(q("r:embed")) or ""
                rel = rels.get(sh.rid)
                if rel and rel["mode"] != "External":
                    sh.media = rel["target"]
                    try:
                        sh.media_size = z.getinfo(sh.media).file_size
                    except KeyError:
                        sh.media_size = 0
        tx = child.find(q("p:txBody"))
        if tx is not None:
            bp = tx.find(q("a:bodyPr"))
            if bp is not None:
                sh.anchor = bp.get("anchor") or ""
                sh.insets = (_ins(bp, "lIns", 91440), _ins(bp, "tIns", 45720),
                             _ins(bp, "rIns", 91440), _ins(bp, "bIns", 45720))
            sh.paras = _parse_paras(tx)
        out.append(sh)
    return out


def _resolve_inherit(part, rels, want):
    for rid, r in rels.items():
        if r["type"] == want:
            return r["target"]
    return ""


def _parse_slide(z, part, index, names, layout_cache):
    sl = Slide(index=index, part=part)
    raw = z.read(part)
    root = parse_xml(raw)
    rels = read_rels(z, part, names)
    sl.layout = _resolve_inherit(part, rels, "slideLayout")
    if sl.layout:
        mrels = read_rels(z, sl.layout, names)
        sl.master = _resolve_inherit(sl.layout, mrels, "slideMaster")
    for rid, r in rels.items():
        if r["type"] == "notesSlide":
            sl.notes_part = r["target"]

    # 背景（纯色或图片）
    bg = root.find(q("p:cSld") + "/" + q("p:bg"))
    if bg is not None:
        sl.bg = _parse_color(bg)
        blip = bg.find(".//" + q("a:blip"))
        if blip is not None:
            rel = rels.get(blip.get(q("r:embed")) or "")
            if rel and rel["mode"] != "External":
                sl.bg_image = rel["target"]

    tree = root.find(q("p:cSld") + "/" + q("p:spTree"))
    if tree is not None:
        sl.shapes = _walk(tree, [0], None, rels, z, "slide")

    # 继承元素（layout + master 上的固定形状与图片）。
    # 注意：必须 list() 拷贝。2026-10 的版本直接 += 缓存 list，导致同版式的
    # 后续页不断往缓存里追加 master 元素，合并稿里每页 inherited 涨到 109。
    sl.inherited = list(_inherited_shapes(z, sl.layout, names, layout_cache, "layout"))
    if sl.master:
        sl.inherited += list(_inherited_shapes(z, sl.master, names, layout_cache, "master"))

    # 备注
    if sl.notes_part:
        try:
            nroot = parse_xml(z.read(sl.notes_part))
            sl.notes = text_of_element(nroot.find(q("p:cSld") + "/" + q("p:spTree")))
        except Exception:
            sl.notes = ""
    return sl


def _inherited_shapes(z, part, names, cache, src):
    if not part:
        return []
    if part in cache:
        return cache[part]
    try:
        root = parse_xml(z.read(part))
    except Exception:
        return []
    rels = read_rels(z, part, names)
    tree = root.find(q("p:cSld") + "/" + q("p:spTree"))
    out = []
    if tree is not None:
        out = _walk(tree, [0], None, rels, z, src)
        # 版式上的占位符只是提示位置，不算实际绘制元素；去掉空占位
        out = [s for s in out if s.kind == "pic" or (s.text and s.text.strip())
               or (s.fill and s.fill not in ("", "FFFFFF"))]
    cache[part] = out
    return out


# --------------------------------------------------------------------------
# 自测
# --------------------------------------------------------------------------
def _self_test(path):
    d = Deck.load(path)
    s = d.summary()
    print("path:", path)
    print("canvas: %d x %d EMU (%.2f x %.2f cm)" % (
        d.w, d.h, units.emu_cm(d.w), units.emu_cm(d.h)))
    print("slides:", s["slides"], "layouts:", s["layouts"],
          "masters:", s["masters"], "media:", s["media"])
    total_text = sum(x["text_chars"] for x in s["per_slide"])
    total_pics = sum(x["pics"] for x in s["per_slide"])
    total_notes = sum(x["notes_chars"] for x in s["per_slide"])
    print("总文字字符:", total_text, " 图片实例:", total_pics, " 备注字符:", total_notes)
    for x in s["per_slide"][:6]:
        sl = d.slide(x["i"])
        flag = "" if sl.layout else "  ⚠ 无 layout"
        print("  p%02d shapes=%-3d pics=%-2d text=%-4d notes=%-4d %s%s" % (
            x["i"], x["shapes"], x["pics"], x["text_chars"],
            x["notes_chars"], posixpath.basename(sl.part or ""), flag))
    if s["slides"] > 6:
        print("  ... 共 %d 页" % s["slides"])
    # 一条不变量：页序非空、每页有 rels
    problems = []
    if not d.slides:
        problems.append("没有读到任何页")
    for sl in d.slides:
        if not sl.part:
            problems.append("第 %d 页部件为空" % sl.index)
        if not sl.layout:
            problems.append("第 %d 页无版式" % sl.index)
    if problems:
        print("SELF-TEST FAILED")
        for p in problems[:10]:
            print("  -", p)
        return 1
    print("SELF-TEST OK — read.py IR 可用")
    return 0


if __name__ == "__main__":
    argv = [a for a in sys.argv[1:] if a != "--self-test"]
    if "--self-test" in sys.argv and argv:
        sys.exit(_self_test(argv[0]))
    print(__doc__)
