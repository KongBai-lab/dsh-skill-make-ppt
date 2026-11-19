# -*- coding: utf-8 -*-
"""build.py — 基于模板生成 .pptx（模板继承 / 页码 / 备注）。

与 read.py 的分工
-----------------
* read.py：读现有文件，不改。
* build.py：从"模板 + 内容"造一份新文件。模板的母版、版式、主题、图片原样继承，
  新页用绝对坐标摆放形状（不依赖占位符，避免 lstStyle 继承的连锁问题）。

两条路径
--------
1. **有模板**（推荐）：模板来自任意 .pptx（例如学校统一模板），继承其母版/版式/主题/字体。
       tpl = Template.from_pptx("deck-a.pptx")
       b = DeckBuilder(tpl)
       n = b.add_slide(layout="内容页")
       b.text(n, units.cm(2), units.cm(4), units.cm(20), units.cm(2), "标题", pt=28, bold=True)
       b.page_numbers(True)                 # 右下角页码（字段，放映时随页序）
       b.set_notes(n, "讲稿…")
       b.save("out.pptx")

2. **无模板**：内置一套最小母版（与 workspace/ppt-toolkit 生成器一致），用于快速出稿。

设计约定（与 units.py 互补）
----------------------------
* 所有写入的 sz / spcPct / 坐标一律经 units 生成，不在此文件里出现第二套常量。
* 保存前对每个 slide/layout/master 跑 units.guard_xml；有问题直接抛 UnitError。
* 保存后可用 read.Deck.load 读回、check.check 校验；build.py 不依赖二者。
"""
import os
import re
import sys
import uuid
import zipfile
from xml.sax.saxutils import escape as _xesc

_ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
if _ENGINE_DIR not in sys.path:
    sys.path.insert(0, _ENGINE_DIR)

try:
    from . import units
except ImportError:  # pragma: no cover
    import units

XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NSDECL = 'xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"' % (NS_A, NS_R, NS_P)

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument"


def esc(s):
    return _xesc(str(s), {'"': "&quot;"})


def _read_rels(z, part, names):
    """读部件的 rels，返回 {rId: 绝对部件名}（外部链接跳过）。"""
    import posixpath
    import xml.etree.ElementTree as ET
    d, b = posixpath.split(part)
    rp = posixpath.join(d, "_rels", b + ".rels")
    if rp not in names:
        return {}
    out = {}
    for rel in ET.fromstring(z.read(rp)):
        t = rel.get("Target") or ""
        if (rel.get("TargetMode") or "Internal") == "External" \
                or t.startswith(("http://", "https://")):
            continue
        out[rel.get("Id")] = posixpath.normpath(posixpath.join(d, t))
    return out


# --------------------------------------------------------------------------
# 模板
# --------------------------------------------------------------------------
class Template:
    """从现有 .pptx 抽取可复用的包骨架。"""

    def __init__(self):
        self.path = ""
        self.w = units.SLIDE_W_EMU
        self.h = units.SLIDE_H_EMU
        self.notes_w = 6858000
        self.notes_h = 9144000
        self.parts = {}          # 部件名 → bytes（母版/版式/主题/媒体/notesMaster/tableStyles）
        self.masters = []        # [(part, master_id)]
        self.layouts = []        # [(part, name)]
        self.notes_master = ""   # 部件名或空
        self.theme_parts = []    # 主题部件名

    @classmethod
    def from_pptx(cls, path):
        import xml.etree.ElementTree as ET
        t = cls()
        t.path = path
        z = zipfile.ZipFile(path)
        names = set(n for n in z.namelist() if not n.endswith("/"))

        # 画布与备注画布
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        P = "{%s}" % NS_P
        sz = pres.find(P + "sldSz")
        if sz is not None:
            t.w, t.h = int(sz.get("cx")), int(sz.get("cy"))
        nsz = pres.find(P + "notesSz")
        if nsz is not None:
            t.notes_w, t.notes_h = int(nsz.get("cx")), int(nsz.get("cy"))

        # 复制母版 / 版式 / 主题 / 备注母版（媒体按引用选择性复制）
        keep_prefix = ("ppt/slideMasters/", "ppt/slideLayouts/", "ppt/theme/",
                       "ppt/notesMasters/")
        for n in names:
            if n.startswith(keep_prefix) or n == "ppt/tableStyles.xml":
                t.parts[n] = z.read(n)
        # 只保留被母版/版式/备注母版实际引用的媒体，避免把整包图片搬进新文件
        used_media = set()
        for n in list(t.parts):
            if n.endswith(".xml") and "/_rels/" not in n:
                for _rid, tgt in _read_rels(z, n, names).items():
                    if "/media/" in tgt and tgt in names:
                        used_media.add(tgt)
        for m in sorted(used_media):
            t.parts[m] = z.read(m)

        # 母版列表（按 presentation rels）
        rels = _read_rels(z, "ppt/presentation.xml", names)
        mst = pres.find(P + "sldMasterIdLst")
        if mst is not None:
            for i, m in enumerate(mst.findall(P + "sldMasterId"), 1):
                rid = m.get("{%s}id" % NS_R)
                if rid in rels:
                    t.masters.append((rels[rid], int(m.get("id") or (2147483648 + i - 1))))
        if not t.masters:
            for n in sorted(names):
                if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n):
                    t.masters.append((n, 2147483648 + len(t.masters)))

        # 版式名
        for n in sorted(names):
            m = re.match(r"ppt/slideLayouts/slideLayout(\d+)\.xml$", n)
            if not m:
                continue
            root = ET.fromstring(z.read(n))
            cs = root.find(P + "cSld")
            t.layouts.append((n, (cs.get("name") if cs is not None else "") or n))
        t.layouts.sort(key=lambda x: int(re.search(r"slideLayout(\d+)", x[0]).group(1)))

        # 备注母版
        for n in sorted(names):
            if re.match(r"ppt/notesMasters/notesMaster\d+\.xml$", n):
                t.notes_master = n
                break
        t.theme_parts = sorted(n for n in names if n.startswith("ppt/theme/"))
        z.close()
        return t

    def layout_index(self, layout, default=0):
        """layout 可以是名称或从 0 开始的序号。找不到抛 KeyError。"""
        if layout is None:
            return default
        if isinstance(layout, int):
            if not (0 <= layout < len(self.layouts)):
                raise KeyError("版式序号 %d 超出 0—%d" % (layout, len(self.layouts) - 1))
            return layout
        for i, (part, name) in enumerate(self.layouts):
            if name == layout or os.path.basename(part) == layout:
                return i
        raise KeyError("模板里没有版式 %r；可用：%s"
                       % (layout, ", ".join(n for _, n in self.layouts)))


# --------------------------------------------------------------------------
# 最小内置骨架（无模板时）
# --------------------------------------------------------------------------
def _default_parts(w, h):
    """返回 (parts, masters, layouts, notes_master)。"""
    parts = {}
    parts["ppt/theme/theme1.xml"] = _theme_xml()
    parts["ppt/slideMasters/slideMaster1.xml"] = _master_xml()
    parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="%s/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" Type="%s/theme" Target="../theme/theme1.xml"/>'
        "</Relationships>" % (REL, REL))
    parts["ppt/slideLayouts/slideLayout1.xml"] = _layout_xml()
    parts["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="%s/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        "</Relationships>" % REL)
    parts["ppt/tableStyles.xml"] = (
        '<a:tblStyleLst xmlns:a="%s" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>' % NS_A)
    return parts, [("ppt/slideMasters/slideMaster1.xml", 2147483648)], \
        [("ppt/slideLayouts/slideLayout1.xml", "Blank")], ""


# --------------------------------------------------------------------------
# 构造器
# --------------------------------------------------------------------------
class DeckBuilder:
    def __init__(self, template=None):
        self.tpl = template
        if template is not None:
            self.w, self.h = template.w, template.h
            self._parts = dict(template.parts)
            self._masters = list(template.masters)
            self._layouts = list(template.layouts)
            self._notes_master = template.notes_master
        else:
            self.w, self.h = units.SLIDE_W_EMU, units.SLIDE_H_EMU
            self._parts, self._masters, self._layouts, self._notes_master = \
                _default_parts(self.w, self.h)
        self._slides = []       # 每页 dict
        self._media = {}        # 部件名 → bytes
        self._img_seq = 0
        self._page_number = None   # None / dict
        self._title = ""
        self._warn = []

    # ---------------- 页 ----------------
    def add_slide(self, layout=None, bg=None):
        """新增一页，返回 1-based 页号。layout 可给名称或序号。"""
        idx = 0
        if self._layouts:
            idx = self.tpl.layout_index(layout) if self.tpl else 0
            if self.tpl is None and layout not in (None, 0):
                raise KeyError("无模板时只有一个版式")
        self._slides.append({"layout": idx, "bg": bg, "shapes": [], "images": [],
                             "notes": ""})
        return len(self._slides)

    def slide_count(self):
        return len(self._slides)

    def _sl(self, n):
        if not (1 <= n <= len(self._slides)):
            raise IndexError("页号 %s 超出 1—%d" % (n, len(self._slides)))
        return self._slides[n - 1]

    def layout_name(self, n):
        idx = self._sl(n)["layout"]
        return self._layouts[idx][1] if idx < len(self._layouts) else ""

    # ---------------- 形状 ----------------
    def text(self, n, x, y, w, h, text, pt=18, bold=False, color="191B1F",
             align="l", anchor="t", spacing=1.0, font=None, name="文本框",
             insets=(0, 0, 0, 0), autofit=False):
        """加文本框。x/y/w/h 单位 EMU（用 units.cm 生成）。text 可含 \\n 分页。"""
        sl = self._sl(n)
        paras = []
        for line in str(text).split("\n"):
            paras.append(_para_xml(line, pt, bold, color, align, spacing, font))
        xml = _text_shape_xml(name, x, y, w, h, "".join(paras), anchor, insets, autofit)
        sl["shapes"].append(xml)
        return name

    def rect(self, n, x, y, w, h, fill=None, line=None, line_pt=1.0,
             geom="rect", radius=None, name="形状"):
        sl = self._sl(n)
        sl["shapes"].append(_rect_xml(name, x, y, w, h, fill, line, line_pt,
                                      geom, radius))
        return name

    def line(self, n, x1, y1, x2, y2, color="808080", width_pt=1.5, name="直线"):
        sl = self._sl(n)
        sl["shapes"].append(_line_xml(name, x1, y1, x2, y2, color, width_pt))
        return name

    def image(self, n, src, x, y, w=None, h=None, fit="contain", name=None):
        """src 为文件路径或 bytes。fit=contain 时按像素比例收缩，不拉伸。"""
        sl = self._sl(n)
        data, ext, iw, ih = _load_image(src)
        if fit == "contain" and w and h and iw and ih:
            r_img = iw / float(ih)
            r_box = w / float(h)
            if r_img > r_box:
                h = int(w / r_img)
            else:
                w = int(h * r_img)
        if w is None and h is None:
            if iw and ih:                    # 按 96 dpi 换算，避免把像素当 EMU
                w = int(iw * units.EMU_PER_INCH / 96.0)
                h = int(ih * units.EMU_PER_INCH / 96.0)
            else:
                w, h = units.inch(2), units.inch(1.5)
        self._img_seq += 1
        part = "ppt/media/image%d.%s" % (self._img_seq, ext)
        self._media[part] = data
        sid = 100 + self._img_seq
        sl["images"].append({"part": part, "x": x, "y": y, "w": w, "h": h,
                             "id": sid, "name": name or ("图片 %d" % self._img_seq)})
        return name or ("图片 %d" % self._img_seq)

    def page_numbers(self, enable=True, pt=10, color="808080", margin=0.35,
                     align="ctr"):
        """页脚页码（a:fld 字段，放映时按实际页序）。位置在页面底部居中。"""
        if not enable:
            self._page_number = None
            return
        self._page_number = {"pt": pt, "color": color,
                             "margin": units.cm(margin), "align": align}

    def set_notes(self, n, text):
        self._sl(n)["notes"] = str(text or "")

    def set_title(self, title):
        self._title = title

    # ---------------- 保存 ----------------
    def save(self, path, verify=True):
        parts = dict(self._parts)
        n = len(self._slides)
        # 母版 id / 关系
        master_rels = []
        master_ids = []
        for i, (part, mid) in enumerate(self._masters):
            rid = "rIdM%d" % (i + 1)
            master_rels.append((rid, part))
            master_ids.append((mid, rid))
        # 页码占位（全局开关）
        slide_xml_parts = {}
        for i, sl in enumerate(self._slides, 1):
            shapes = list(sl["shapes"])
            if self._page_number:
                shapes.append(_page_number_xml(i, self.w, self.h, self._page_number))
            rid_map = {}
            img_rels = []
            for j, im in enumerate(sl["images"]):
                rid = "rIdImg%d" % (j + 1)
                rid_map[im["id"]] = rid
                aimg = _pic_xml(im, rid)
                shapes.append(aimg)
                img_rels.append((rid, im["part"]))
            xml = _slide_xml(shapes, sl["bg"])
            # 页内数值断言
            probs = units.guard_xml(xml, "slide%d" % i)
            if probs:
                raise units.UnitError("第 %d 页写入前校验失败：%s" % (i, probs[:3]))
            slide_xml_parts[i] = xml
            parts["ppt/slides/slide%d.xml" % i] = xml
            # slide 在 ppt/slides/ 下，target 相对该目录
            rels = [("rId1", "../slideLayouts/" + os.path.basename(self._layouts[sl["layout"]][0]))]
            for rid, part in img_rels:
                rels.append((rid, "../media/" + os.path.basename(part)))
            if sl["notes"]:
                rels.append(("rIdN%d" % i, "../notesSlides/notesSlide%d.xml" % i))
                parts["ppt/notesSlides/notesSlide%d.xml" % i] = _notes_slide_xml(
                    sl["notes"])
                parts["ppt/notesSlides/_rels/notesSlide%d.xml.rels" % i] = _notes_rels(
                    i, self._notes_master)
            parts["ppt/slides/_rels/slide%d.xml.rels" % i] = _rels_xml(rels)
        # 备注母版（模板没有就造一个）
        if any(sl["notes"] for sl in self._slides) and not self._notes_master:
            parts["ppt/notesMasters/notesMaster1.xml"] = _notes_master_xml()
            parts["ppt/notesMasters/_rels/notesMaster1.xml.rels"] = (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="%s/theme" Target="../theme/theme2.xml"/>'
                "</Relationships>" % REL)
            parts["ppt/theme/theme2.xml"] = parts.get("ppt/theme/theme1.xml") or _theme_xml()
            self._notes_master = "ppt/notesMasters/notesMaster1.xml"
            for i, sl in enumerate(self._slides, 1):
                if sl["notes"]:
                    parts["ppt/notesSlides/_rels/notesSlide%d.xml.rels" % i] = _notes_rels(
                        i, self._notes_master)
        # 演示文稿（presentation.xml 在 ppt/ 下，target 相对 ppt/）
        parts["ppt/presentation.xml"] = _presentation_xml(
            self.w, self.h, n, master_ids, self._title)
        prels = [(rid, part.split("ppt/", 1)[1]) for rid, part in master_rels]
        for i in range(1, n + 1):
            prels.append(("rIdS%d" % i, "slides/slide%d.xml" % i))
        if self._notes_master:
            prels.append(("rIdNM", self._notes_master.split("ppt/", 1)[1]))
        if "ppt/theme/theme1.xml" in parts:
            prels.append(("rIdT1", "theme/theme1.xml"))
        parts["ppt/_rels/presentation.xml.rels"] = _rels_xml(prels)
        parts["_rels/.rels"] = (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="%s/officeDocument" Target="ppt/presentation.xml"/>'
            "</Relationships>" % REL)
        # 内容类型
        parts["[Content_Types].xml"] = _content_types(parts)
        # 媒体
        for part, data in self._media.items():
            parts[part] = data
        # 写盘（先 tmp 再替换）
        out = os.path.abspath(path)
        os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
        tmp = out + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in parts.items():
                if isinstance(data, bytes):
                    z.writestr(name, data)
                else:
                    z.writestr(name, _ensure_head(data))
        # 全包断言：所有 slide/layout/master 的数值
        z = zipfile.ZipFile(tmp)
        probs = []
        for name in z.namelist():
            if re.match(r"ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml$", name):
                probs += units.guard_xml(z.read(name).decode("utf-8"), name)
        z.close()
        if probs:
            os.remove(tmp)
            raise units.UnitError("整包数值校验失败：%s" % probs[:5])
        os.replace(tmp, out)
        self._path = out
        return out


# --------------------------------------------------------------------------
# XML 片段
# --------------------------------------------------------------------------
def _ensure_head(s):
    return s if s.lstrip().startswith("<?xml") else XML_HEAD + s


def _rels_xml(pairs):
    body = "".join(
        '<Relationship Id="%s" Type="%s/%s" Target="%s"/>'
        % (rid, REL, "image" if "/media/" in tgt else
           ("notesSlide" if "notesSlide" in tgt else
            ("notesMaster" if "notesMaster" in tgt else
             ("slideLayout" if "slideLayout" in tgt else
              ("slideMaster" if "slideMaster" in tgt else
               ("theme" if "theme" in tgt else "slide"))))), tgt)
        for rid, tgt in pairs)
    return ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + body + "</Relationships>")


def _para_xml(text, pt, bold, color, align, spacing, font=None):
    sz = units.sz_of_pt_safe(pt)
    spc = units.spc_pct_of_ratio(spacing)
    algn = {"l": "l", "left": "l", "c": "ctr", "ctr": "ctr", "center": "ctr",
            "r": "r", "right": "r", "j": "just", "just": "just"}.get(align, "l")
    fnt = ""
    if font:
        fnt = ('<a:latin typeface="%s"/><a:ea typeface="%s"/>' % (esc(font), esc(font)))
    rpr = ('<a:rPr lang="zh-CN" altLang="en-US" sz="%s"%s dirty="0">'
           '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</a:rPr>'
           % (sz, ' b="1"' if bold else "", color, fnt))
    return ('<a:p><a:pPr algn="%s"><a:lnSpc><a:spcPct val="%s"/></a:lnSpc>'
            '<a:spcBef><a:spcPts val="0"/></a:spcBef>'
            '<a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>'
            '<a:r>%s<a:t>%s</a:t></a:r></a:p>' % (algn, spc, rpr, esc(text)))


def _xfrm(x, y, w, h, flip_h=False, flip_v=False):
    attrs = ""
    if flip_h:
        attrs += ' flipH="1"'
    if flip_v:
        attrs += ' flipV="1"'
    return ('<a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            % (attrs, int(x), int(y), int(w), int(h)))


def _text_shape_xml(name, x, y, w, h, paras, anchor="t", insets=(0, 0, 0, 0),
                    autofit=False):
    l, t, r, b = insets          # 单位 EMU
    an = {"t": "t", "top": "t", "c": "ctr", "ctr": "ctr", "center": "ctr",
          "b": "b", "bottom": "b"}.get(anchor, "t")
    fit = "<a:normAutofit/>" if autofit else ""
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/>'
            '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="square" lIns="%d" tIns="%d" rIns="%d" bIns="%d" '
            'anchor="%s" rtlCol="0">%s</a:bodyPr><a:lstStyle/>%s</p:txBody></p:sp>'
            % (_sid(), esc(name), _xfrm(x, y, w, h), int(l), int(t), int(r), int(b),
               an, fit, paras))


def _rect_xml(name, x, y, w, h, fill, line, line_pt, geom="rect", radius=None):
    prst = geom
    av = ""
    if prst == "roundRect" and radius is not None:
        av = '<a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst>' % int(radius * 100000)
    fill_xml = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill
                if fill else "<a:noFill/>")
    ln_xml = ""
    if line:
        ln_xml = ('<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>'
                  % (units.pt_emu(line_pt), line))
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>%s<a:prstGeom prst="%s">%s</a:prstGeom>%s%s</p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            % (_sid(), esc(name), _xfrm(x, y, w, h), prst,
               av or "<a:avLst/>", fill_xml, ln_xml))


def _line_xml(name, x1, y1, x2, y2, color, width_pt):
    x, y = min(x1, x2), min(y1, y2)
    w, h = max(1, abs(x2 - x1)), max(1, abs(y2 - y1))
    flip_h = "1" if x2 < x1 else "0"
    flip_v = "1" if y2 < y1 else "0"
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>%s<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
            '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            % (_sid(), esc(name),
               '<a:xfrm flipH="%s" flipV="%s"><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
               % (flip_h, flip_v, int(x), int(y), int(w), int(h)),
               units.pt_emu(width_pt), color))


def _pic_xml(im, rid):
    return ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s"/>'
            '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
            % (im["id"], esc(im["name"]), rid, _xfrm(im["x"], im["y"], im["w"], im["h"])))


_sid_counter = [1000]


def _sid():
    _sid_counter[0] += 1
    return _sid_counter[0]


def _page_number_xml(page_no, w, h, cfg):
    # 页码用字段，放映/导出时按实际页序；text 只是缓存值
    margin = cfg.get("margin", units.cm(0.35))
    bw, bh = units.cm(1.6), units.cm(0.6)
    x = w - margin - bw
    y = h - margin - bh
    algn = {"l": "l", "ctr": "ctr", "r": "r"}.get(cfg.get("align", "r"), "r")
    color = cfg.get("color", "808080")
    pt = cfg.get("pt", 10)
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="页码"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="b"/>'
            '<a:lstStyle/><a:p><a:pPr algn="%s"/><a:fld id="{%s}" type="slidenum">'
            '<a:rPr lang="zh-CN" sz="%s"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:rPr>'
            '<a:t>%d</a:t></a:fld></a:p></p:txBody></p:sp>'
            % (_sid(), x, y, bw, bh, algn, str(uuid.uuid4()).upper(),
               units.sz_of_pt(pt), color, page_no))


def _slide_xml(shapes, bg=None):
    bgxml = ""
    if bg:
        bgxml = ('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
                 '<a:effectLst/></p:bgPr></p:bg>' % bg)
    return ('<p:sld %s>%s<p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            "%s</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
            % (NSDECL, bgxml, "".join(shapes)))


def _presentation_xml(w, h, n, master_ids, title=""):
    mids = "".join('<p:sldMasterId id="%d" r:id="%s"/>' % (mid, rid)
                   for mid, rid in master_ids)
    ids = "".join('<p:sldId id="%d" r:id="rIdS%d"/>' % (255 + i, i)
                  for i in range(1, n + 1))
    return ('<p:presentation %s saveSubsetFonts="1">'
            '<p:sldMasterIdLst>%s</p:sldMasterIdLst>'
            "<p:sldIdLst>%s</p:sldIdLst>"
            '<p:sldSz cx="%d" cy="%d"/>'
            '<p:notesSz cx="6858000" cy="9144000"/>'
            "</p:presentation>" % (NSDECL, mids, ids, w, h))


def _notes_slide_xml(text):
    # 每行一段；空行保留
    paras = "".join(_para_xml(line, 12, False, "000000", "l", 1.0)
                    for line in str(text).split("\n"))
    return ('<p:notes %s><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="备注文本框"/><p:cNvSpPr txBox="1"/>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="685800" y="4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/>%s</p:txBody></p:sp>'
            "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>"
            % (NSDECL, paras))


def _notes_rels(slide_no, notes_master_part):
    pairs = [("rId1", "../slides/slide%d.xml" % slide_no),
             ("rId2", "../notesMasters/" + os.path.basename(notes_master_part))]
    return _rels_xml(pairs)


def _content_types(parts):
    defaults = {"rels": "application/vnd.openxmlformats-package.relationships+xml",
                "xml": "application/xml"}
    for n in parts:
        ext = n.rsplit(".", 1)[-1].lower()
        if ext in ("png", "jpg", "jpeg", "gif", "bmp", "webp", "emf", "wmf", "svg") \
                and ext not in defaults:
            defaults[ext] = "image/" + ("jpeg" if ext == "jpg" else ext)
    overrides = {
        "ppt/presentation.xml": CT + ".presentationml.presentation.main+xml",
        "ppt/tableStyles.xml": CT + ".presentationml.tableStyles+xml",
    }
    for n in parts:
        if re.match(r"ppt/slides/slide\d+\.xml$", n):
            overrides[n] = CT + ".presentationml.slide+xml"
        elif re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n):
            overrides[n] = CT + ".presentationml.slideLayout+xml"
        elif re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n):
            overrides[n] = CT + ".presentationml.slideMaster+xml"
        elif re.match(r"ppt/theme/theme\d+\.xml$", n):
            overrides[n] = CT + ".theme+xml"
        elif re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n):
            overrides[n] = CT + ".presentationml.notesSlide+xml"
        elif re.match(r"ppt/notesMasters/notesMaster\d+\.xml$", n):
            overrides[n] = CT + ".presentationml.notesMaster+xml"
    dflt = "".join('<Default Extension="%s" ContentType="%s"/>' % (k, v)
                   for k, v in sorted(defaults.items()))
    over = "".join('<Override PartName="/%s" ContentType="%s"/>' % (k, v)
                   for k, v in sorted(overrides.items()))
    return ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            + dflt + over + "</Types>")


def _theme_xml(accent="0C4994", accent2="DB1F1F", font="微软雅黑"):
    parts = []
    for i, c in enumerate([accent, "2E75B6", accent2, "1F7A3D", "7F5AA8", "C99A00"], 1):
        parts.append('<a:accent%d><a:srgbClr val="%s"/></a:accent%d>' % (i, c, i))
    return ('<a:theme xmlns:a="%s" name="make-ppt"><a:themeElements>'
            '<a:clrScheme name="make-ppt">'
            '<a:dk1><a:srgbClr val="000000"/></a:dk1>'
            '<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
            '<a:dk2><a:srgbClr val="191B1F"/></a:dk2>'
            '<a:lt2><a:srgbClr val="F5F7FA"/></a:lt2>'
            "%s"
            '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
            '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
            '<a:fontScheme name="make-ppt">'
            '<a:majorFont><a:latin typeface="Times New Roman"/><a:ea typeface="%s"/>'
            '<a:cs typeface=""/></a:majorFont>'
            '<a:minorFont><a:latin typeface="%s"/><a:ea typeface="%s"/>'
            '<a:cs typeface=""/></a:minorFont></a:fontScheme>'
            '<a:fmtScheme name="make-ppt">%s</a:fmtScheme>'
            "</a:themeElements></a:theme>"
            % (NS_A, "".join(parts), font, font, font, _fmt_scheme()))


def _fmt_scheme():
    return ('<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"><a:lumMod val="110000"/>'
            '<a:satMod val="105000"/></a:schemeClr></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"><a:lumMod val="150000"/>'
            '<a:satMod val="35000"/></a:schemeClr></a:solidFill></a:fillStyleLst>'
            '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln></a:lnStyleLst>'
            '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
            '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"><a:lumMod val="95000"/>'
            '<a:satMod val="170000"/></a:schemeClr></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"><a:lumMod val="90000"/>'
            '<a:satMod val="120000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst>')


def _master_xml():
    return ('<p:sldMaster %s><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree></p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
            'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
            'accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
            "<p:txStyles>"
            '<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>'
            '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr></p:bodyStyle>'
            '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>'
            "</p:txStyles></p:sldMaster>" % NSDECL)


def _layout_xml():
    return ('<p:sldLayout %s type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
            % NSDECL)


def _notes_master_xml():
    return ('<p:notesMaster %s><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree></p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
            'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
            'accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            '<p:notesStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:notesStyle>'
            "</p:notesMaster>" % NSDECL)


# --------------------------------------------------------------------------
# 图片
# --------------------------------------------------------------------------
def _load_image(src):
    """返回 (bytes, 扩展名, 宽, 高)。宽高未知返回 (None, None)。"""
    if isinstance(src, (bytes, bytearray)):
        data = bytes(src)
        name = ""
    else:
        with open(src, "rb") as f:
            data = f.read()
        name = src
    ext = "png"
    if data[:3] == b"\xff\xd8\xff":
        ext = "jpeg"
    elif data[:6] in (b"GIF87a", b"GIF89a"):
        ext = "gif"
    elif data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        ext = "webp"
    elif data[:8] == b"\x89PNG\r\n\x1a\n":
        ext = "png"
    w, h = _image_size(data)
    return data, ext, w, h


def _image_size(data):
    """PNG/JPEG/GIF 的像素尺寸；读不出返回 (None, None)。"""
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            import struct
            w, h = struct.unpack(">II", data[16:24])
            return int(w), int(h)
        if data[:3] == b"\xff\xd8\xff":
            i = 2
            n = len(data)
            while i < n - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                              0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    h = (data[i + 5] << 8) | data[i + 6]
                    w = (data[i + 7] << 8) | data[i + 8]
                    return int(w), int(h)
                seg = (data[i + 2] << 8) | data[i + 3]
                i += 2 + seg
        if data[:6] in (b"GIF87a", b"GIF89a"):
            w = data[6] | (data[7] << 8)
            h = data[8] | (data[9] << 8)
            return int(w), int(h)
    except Exception:
        pass
    return None, None


# --------------------------------------------------------------------------
# 自测
# --------------------------------------------------------------------------
def _self_test():
    """有模板/无模板各建一份；没有设置 MAKEPPT_CORPUS 时跳过有模板分支。"""
    import read as rmod

    corpus = os.environ.get("MAKEPPT_CORPUS")
    if corpus and os.path.isdir(corpus):
        corpus = os.path.join(corpus, "deck-a.pptx")
    if corpus and not os.path.isfile(corpus):
        print("SKIP 有模板分支：MAKEPPT_CORPUS 指向的 %s 不存在" % corpus)
        corpus = None
    if not corpus:
        print("SKIP 有模板分支：未设置 MAKEPPT_CORPUS"
              "（可指向一份 .pptx 或含 deck-a.pptx 的目录）")

    out = os.path.join(_ENGINE_DIR, "..", "out", "selftest_build.pptx")
    out = os.path.abspath(out)

    # 1) 有模板
    path = None
    if corpus:
        b = DeckBuilder(Template.from_pptx(corpus))
        n1 = b.add_slide(layout=b.tpl.layouts[0][1])
        b.text(n1, units.cm(2), units.cm(3), units.cm(20), units.cm(2.4),
               "示例标题", pt=40, bold=True, align="ctr", anchor="ctr")
        b.text(n1, units.cm(2), units.cm(6), units.cm(20), units.cm(1.4),
               "科学技术方法论考察", pt=20, align="ctr")
        n2 = b.add_slide(layout=b.tpl.layouts[0][1])
        b.text(n2, units.cm(2), units.cm(1.6), units.cm(24), units.cm(1.6),
               "三条件同时具备，能力才提升", pt=28, bold=True)
        b.rect(n2, units.cm(2), units.cm(4), units.cm(10), units.cm(3),
               fill="F5F7FA", line="0C4994")
        b.text(n2, units.cm(2.4), units.cm(4.4), units.cm(9), units.cm(2),
               "算力：GPU 并行把训练周期从数周缩到数天", pt=14)
        b.page_numbers(True)
        b.set_notes(n2, "第 2 页讲稿：先给结论，再展开两步推导。")
        path = b.save(out, verify=True)

        d = rmod.Deck.load(path)
        assert d.n == 2, d.n
        assert "示例标题" in d.slide(1).text
        assert d.slide(2).notes, "备注没有写进去"
        print("有模板：2 页 OK，备注 %d 字" % len(d.slide(2).notes))

    # 2) 无模板
    b2 = DeckBuilder()
    b2.add_slide()
    b2.text(1, units.cm(2), units.cm(2), units.cm(20), units.cm(2), "无模板页", pt=30)
    b2.page_numbers(True)
    b2.set_notes(1, "备注")
    p2 = b2.save(out.replace(".pptx", "_plain.pptx"), verify=True)
    d2 = rmod.Deck.load(p2)
    assert d2.n == 1 and d2.slide(1).notes
    print("无模板：1 页 OK")
    print("SELF-TEST OK — build.py 可用：", path or p2)
    return 0


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        sys.exit(_self_test())
    print(__doc__)
