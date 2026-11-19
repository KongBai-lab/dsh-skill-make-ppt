# -*- coding: utf-8 -*-
"""check.py — PPTX 交付前校验器（make-ppt engine）。

为什么需要这个文件
------------------
2026-10 课程 PPT 事故里，**自查脚本本身是坏的**：
把 `sz="750"`（7.5 pt）当成 750 pt，连续三轮报“0 处溢出/0 处重叠”，
用户打开 WPS 却看到满屏叠字；行距单位写成 0.1 倍（`spcPct=100`），
7 页 96 处文字压在一起。教训是：**“0 问题”必须能被坏样本证伪**。

因此本模块输出四个级别明确、可复核的结论：
  * ERROR = 结构性问题，OOWriter/WPS 打开一定异常，或数值量级一定错；
  * WARN  = 估算性判断（文字溢出/文字相交/图片拉伸），必须给出
            “估算值 vs 阈值”的数字，禁止只说“可能重叠”。

四类检查（与 refs/ir.md 契约一致）
----------------------------------
1. structure（ERROR）：xml/rels 良构、关系断链、[Content_Types].xml 覆盖、
   sldIdLst 指向、notesSlide 关系、部件名唯一、r:embed 等引用可解析。
2. units（ERROR）：ppt/slides|slideLayouts|slideMasters/*.xml 逐个过
   units.guard_xml（字号 1/100pt、行距 1/1000%、坐标 EMU 的量级扫描）。
3. inherit（ERROR）：页面元素与 layout/master 继承的固定元素重复 ——
   同一 media 且矩形相交 > 50% 判“双图标”；页面内同 media 出现两次同判。
4. layout（WARN）：元素越界（容差 2% 画布；完全在画布外升为 ERROR）、
   文字溢出（units.text_height_emu，内含 1.5 倍保守行高）、
   文字与文字相交（用**估算的文字矩形**而非框矩形）、
   图片拉伸（像素纵横比 vs 框纵横比，误差 > 3%）。

接口
----
    rep = check(path, level="all")     # level: structure|units|inherit|layout|all
    rep.ok / rep.errors / rep.warnings / rep.stats
    print(rep.as_text())

CLI
---
    python check.py <pptx> [--json] [-o report.txt] [--level all]
    python check.py --self-test [--corpus DIR]

`--self-test` 做两件事（缺一不可）：
  A. 回归语料：deck-a.pptx / deck-b.pptx / deck-merged.pptx 三份真实文件。
     deck-a.pptx 是半成品，预期恰好 4 条 inherit/DUP_ICON（示例缺陷样本），
     deck-b.pptx 与 合并稿预期 0 ERROR；
  B. 坏样本证伪：在 out/selftest/ 下注入 8 个缺陷（spcPct=100、元素部分越界、
     元素完全出画布、layout 与页面同 media 双图标、文字溢出、图片拉伸、
     关系断链、Content_Types 缺覆盖），断言每个都让报告比基线多报一条。
     校验器只有被坏样本打穿，才配说“0 问题”。

注意
----
* 只依赖标准库 + engine 内的 units/read；不引 python-pptx / PIL。
* 不复制第二套单位常量：字号/行距/坐标一律走 units；bodyPr 内边距直接用
  read.Shape.insets（read.py 已按 OOXML 规范补缺省 91440/45720 EMU）。
* 文字坐标依赖 read.Shape 的矩形与 anchor/insets；read IR 不带旋转角（rot），
  回测三份语料 xfrm@rot 均为 0；遇到旋转图片需先扩展 read.py 再放宽本检查。
* 版式/母版上的文字（Slide.inherited）不做溢出/相交检查：它们多是占位提示，
  页面上通常已被实际内容覆盖，检查会制造大量噪声。
* 表格内文字（graphicFrame）不在 read IR 里，本模块不检查其溢出。
"""

import argparse
import json
import os
import posixpath
import re
import struct
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from urllib.parse import unquote

_ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
if _ENGINE_DIR not in sys.path:
    sys.path.insert(0, _ENGINE_DIR)

try:                       # 支持 `import check` 与 `from engine import check`
    from . import units, read
except ImportError:        # pragma: no cover
    import units
    import read

q = read.q

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------
#: 画布越界容差（2%），出血元素允许轻微超出
BOUNDS_TOL = 0.02
#: 图片像素纵横比与框纵横比的最大允许偏差（3%）
ASPECT_TOL = 0.03
#: “双图标”判定：同 media 且矩形相交面积 / 较小矩形面积 > 50%
ICON_OVERLAP = 0.5
#: 文字溢出判定：估算所需高度 > 框内可用高度 × TEXT_OVERFLOW_TOL
TEXT_OVERFLOW_TOL = 1.0

#: OOXML `a:bodyPr` 缺省内边距（EMU）。ECMA-376 21.1.2.1.1：
#: lIns/rIns = 0.1in = 91440；tIns/bIns = 0.05in = 45720。
#: 实际取值由 read.Shape.insets 提供（read.py 已按规范补缺省），此处仅作文档。

OUT_DIR = os.path.join(os.path.dirname(_ENGINE_DIR), "out")
BASE_CORPUS = os.environ.get("MAKEPPT_CORPUS") or None

#: 回归语料：文件名 → 页数（refs/ir.md 的“回归语料”表）
CORPUS = [
    ("deck-a.pptx", 20),
    ("deck-b.pptx", 16),
    ("deck-merged.pptx", 55),
]

#: 2026-10-28：deck-a.pptx 的双图标已用 edit.unify_element 清掉（p14/15/17/18 各删 1 张），
#: 现在它应当是 0 ERROR。旧缺陷基线改用修复前备份文件（若存在），见 self_test 的 [A2]：
#: “deck-a-before-fix.pptx” 仍应被检出恰好 4 条 DUP_ICON / inherit。
KNOWN_BASELINE = {}

#: 修复前备份（真实双图标缺陷）——检查器灵敏度回归用
LEGACY_DEFECT_BACKUP = "deck-a-before-fix.pptx"

#: 检查级别顺序（报告与 --level 用）
LEVELS = ("structure", "units", "inherit", "layout")

_RE_RREF = re.compile(r'\br:(embed|link|id|pict|href)="([^"]+)"')
_RE_CNVPR_NAME = re.compile(r'<p:cNvPr\b[^>]*\bname="([^"]*)"')
_RE_OFF = re.compile(r'<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"')
_RE_EXT = re.compile(r'<a:ext\b[^>]*\bcx="(-?\d+)"[^>]*\bcy="(-?\d+)"')


# --------------------------------------------------------------------------
# 报告数据结构
# --------------------------------------------------------------------------
@dataclass
class Finding:
    """一条问题（ERROR 或 WARN）。"""
    check: str          # structure / units / inherit / layout
    level: str          # ERROR / WARN
    code: str           # 稳定短代码，坏样本断言用
    msg: str            # 中文可读描述；WARN 必须含“估算值 vs 阈值”
    slide: int = 0      # 1-based 页号；0 = 全文件

    def as_text(self):
        loc = "p%02d " % self.slide if self.slide else ""
        return "[%s] %s%s: %s" % (self.level, loc, self.check, self.msg)

    def as_dict(self):
        return {"check": self.check, "level": self.level, "code": self.code,
                "slide": self.slide, "msg": self.msg}


@dataclass
class Report:
    """一次校验的完整结果。"""
    path: str = ""
    level: str = "all"
    elapsed: float = 0.0
    findings: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)

    # -------- 写入 --------
    def add(self, check, level, code, msg, slide=0):
        assert level in ("ERROR", "WARN"), level
        self.findings.append(Finding(check, level, code, msg, slide))
        return self.findings[-1]

    # -------- 查询 --------
    @property
    def ok(self):
        """无 ERROR 即通过（WARN 不阻断）。"""
        return not any(f.level == "ERROR" for f in self.findings)

    @property
    def error_findings(self):
        return [f for f in self.findings if f.level == "ERROR"]

    @property
    def warn_findings(self):
        return [f for f in self.findings if f.level == "WARN"]

    @property
    def errors(self):
        return [f.as_text() for f in self.error_findings]

    @property
    def warnings(self):
        return [f.as_text() for f in self.warn_findings]

    def count(self, level=None):
        return len([f for f in self.findings if level is None or f.level == level])

    def codes(self):
        return {f.code for f in self.findings}

    # -------- 导出 --------
    def to_dict(self):
        return {
            "path": self.path,
            "level": self.level,
            "ok": self.ok,
            "elapsed_s": round(self.elapsed, 3),
            "stats": self.stats,
            "errors": self.errors,
            "warnings": self.warnings,
            "findings": [f.as_dict() for f in self.findings],
        }

    def as_text(self):
        n_err, n_warn = self.count("ERROR"), self.count("WARN")
        mark = "PASS" if self.ok else "FAIL"
        st = self.stats or {}
        lines = []
        lines.append("=" * 72)
        lines.append("PPTX 交付前校验报告")
        lines.append("=" * 72)
        lines.append("文件    : %s" % (self.path or "-"))
        lines.append("结论    : %s — %d ERROR / %d WARN" % (mark, n_err, n_warn))
        lines.append("耗时    : %.3f s" % self.elapsed)
        lines.append("级别    : %s" % self.level)
        lines.append("规模    : %d 页 | %d 版式 | %d 母版 | %d 部件 | %d 形状 | "
                     "%d 图片 | %d 文字 | %d 页备注（%d 字）"
                     % (st.get("slides", 0), st.get("layouts", 0),
                        st.get("masters", 0), st.get("parts", 0),
                        st.get("shapes", 0), st.get("pics", 0),
                        st.get("text_chars", 0), st.get("notes", 0),
                        st.get("notes_chars", 0)))

        # ERROR 全列
        errs = self.error_findings
        lines.append("")
        lines.append("[ERROR] %d 条（结构性问题，交付前必须清零）" % len(errs))
        if not errs:
            lines.append("  （无）")
        for f in errs:
            lines.append("  [%s] %s" % (f.code, f.as_text()))

        # WARN 按检查项分组全列
        warns = self.warn_findings
        lines.append("")
        lines.append("[WARN] %d 条（估算性判断，含估算值 vs 阈值，请按需复核）"
                     % len(warns))
        lines.append("  说明：文字类 WARN 按 units 的 1.5 倍保守行高估算"
                     "（WPS 实测行高 1.32—1.42），")
        lines.append("        轻微超出（<10%）可能只是保守余量；"
                     "请结合 preview 渲染复核再决定是否改版面。")
        if not warns:
            lines.append("  （无）")
        for check_name in LEVELS:
            group = [f for f in warns if f.check == check_name]
            if not group:
                continue
            by_code = {}
            for f in group:
                by_code.setdefault(f.code, []).append(f)
            for code, items in by_code.items():
                lines.append("  %s/%s — %d 条" % (check_name, code, len(items)))
                for f in items:
                    lines.append("    %s" % f.as_text())

        # 覆盖对象数：证明“真的检查了多少对象”，否则 0 问题无意义
        checks = st.get("checks") or {}
        lines.append("")
        lines.append("覆盖对象数（证明未漏检）:")
        if "structure" in checks:
            c = checks["structure"]
            lines.append("  structure: parts=%d xml=%d rels=%d rel_targets=%d "
                         "content_types=%d(default %d/override %d) sld_ids=%d "
                         "layout_rels=%d notes_rels=%d rel_refs=%d"
                         % (c.get("parts", 0), c.get("xml_parts", 0),
                            c.get("rels_files", 0), c.get("rel_targets", 0),
                            c.get("content_type_total", 0),
                            c.get("content_type_defaults", 0),
                            c.get("content_type_overrides", 0),
                            c.get("slide_ids", 0), c.get("slide_layout_rels", 0),
                            c.get("notes_rels", 0), c.get("rel_refs", 0)))
        if "units" in checks:
            c = checks["units"]
            lines.append("  units    : parts=%d xml=%d KB findings=%d"
                         % (c.get("parts", 0), c.get("xml_bytes", 0) // 1024,
                            c.get("findings", 0)))
        if "inherit" in checks:
            c = checks["inherit"]
            lines.append("  inherit  : slides=%d with_inherited=%d "
                         "inherited_shapes=%d slide_pics=%d media_pairs=%d "
                         "skipped_show_master=%d"
                         % (c.get("slides", 0), c.get("slides_with_inherited", 0),
                            c.get("inherited_shapes", 0), c.get("slide_pics", 0),
                            c.get("media_pairs", 0),
                            c.get("skipped_show_master", 0)))
        if "layout" in checks:
            c = checks["layout"]
            lines.append("  layout   : rect_shapes=%d bounds=%d text_shapes=%d "
                         "text_paras=%d text_pairs=%d pics=%d aspect_checked=%d "
                         "aspect_skipped=%d"
                         % (c.get("shapes_with_rect", 0),
                            c.get("bounds_checked", 0),
                            c.get("text_shapes", 0), c.get("text_paras", 0),
                            c.get("text_pairs", 0), c.get("pics", 0),
                            c.get("aspect_checked", 0),
                            c.get("aspect_skipped", 0)))
        lines.append("=" * 72)
        return "\n".join(lines)


# --------------------------------------------------------------------------
# 通用小工具
# --------------------------------------------------------------------------
def _part_names(z):
    return [n for n in z.namelist() if not n.endswith("/")]


def _rels_part_of(part):
    """部件名 → 它的 .rels 部件名。"""
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def _part_of_rels(rels_name):
    """`ppt/slides/_rels/slide1.xml.rels` → `ppt/slides/slide1.xml`。"""
    d, b = posixpath.split(rels_name)
    base = posixpath.dirname(d)
    return posixpath.join(base, b[:-5])


def _resolve_target(base, target):
    """把关系 Target 解析成候选部件名（处理 URI 转义与绝对路径）。"""
    cands = []
    for t in (target, unquote(target)):
        if not t:
            continue
        if t.startswith("/"):
            cands.append(t.lstrip("/"))
        else:
            cands.append(posixpath.normpath(posixpath.join(base, t)))
    seen, out = set(), []
    for c in cands:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def _inter_rect(a, b):
    """两个矩形 (x, y, w, h) 的相交矩形；无相交返回 None。"""
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[0] + a[2], b[0] + b[2])
    y2 = min(a[1] + a[3], b[1] + b[3])
    if x2 <= x1 or y2 <= y1:
        return None
    return (x1, y1, x2 - x1, y2 - y1)


def _inter_area(a, b):
    r = _inter_rect(a, b)
    return r[2] * r[3] if r else 0


def _cm(v):
    return units.emu_cm(v)


# --------------------------------------------------------------------------
# 图片像素尺寸解析（PNG / JPEG / GIF / BMP；读不出返回 None）
# --------------------------------------------------------------------------
def _png_size(raw):
    if len(raw) >= 24 and raw[:8] == b"\x89PNG\r\n\x1a\n" and raw[12:16] == b"IHDR":
        w, h = struct.unpack(">II", raw[16:24])
        return (w, h)
    return None


_JPEG_SOF = (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
             0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF)


def _jpeg_size(raw):
    if len(raw) < 4 or raw[:2] != b"\xff\xd8":
        return None
    i = 2
    n = len(raw)
    while i + 9 < n:
        if raw[i] != 0xFF:
            i += 1
            continue
        m = raw[i + 1]
        if m == 0xFF:          # 填充字节
            i += 1
            continue
        if m in (0xD8, 0xD9, 0x01) or 0xD0 <= m <= 0xD7:
            i += 2
            continue
        try:
            seg = struct.unpack(">H", raw[i + 2:i + 4])[0]
        except struct.error:
            return None
        if m in _JPEG_SOF:
            h, w = struct.unpack(">HH", raw[i + 5:i + 9])
            return (w, h)
        if seg < 2:
            return None
        i += 2 + seg
    return None


def _gif_size(raw):
    if len(raw) >= 10 and raw[:6] in (b"GIF87a", b"GIF89a"):
        return struct.unpack("<HH", raw[6:10])
    return None


def _bmp_size(raw):
    if len(raw) >= 26 and raw[:2] == b"BM":
        w, h = struct.unpack("<ii", raw[18:26])
        return (abs(w), abs(h))
    return None


def image_size(raw):
    """图片字节 → (宽, 高) 像素；未知格式返回 None（调用方跳过并计数）。"""
    for fn in (_png_size, _jpeg_size, _gif_size, _bmp_size):
        r = fn(raw)
        if r and r[0] > 0 and r[1] > 0:
            return r
    return None


# --------------------------------------------------------------------------
# 补充 read IR 未暴露的画布属性
# --------------------------------------------------------------------------
#: read.Shape 已提供 anchor（bodyPr@anchor）/ insets（l,t,r,b）/ crop /
#: flip_h / flip_v / line_w，布局检查直接用，不再二次解析 XML。
#: 这里只补两样 IR 没有、又直接影响渲染的东西：
#:   1) Slide 根元素 showMasterSp="0"（页面不绘制版式元素，则不构成双图标）；
#:   2) bodyPr/normAutofit@fontScale（WPS 会把文字自动缩小到框内）。
def _walk_font_scale(el, out, counter):
    """按 read._walk 的顺序遍历 spTree，收集 normAutofit@fontScale。"""
    for child in list(el):
        kind = child.tag.rsplit("}", 1)[-1]
        if kind in ("nvGrpSpPr", "grpSpPr"):
            continue
        if kind == "grpSp":
            _walk_font_scale(child, out, counter)
            continue
        counter[0] += 1
        idx = counter[0]
        if kind not in ("sp", "cxnSp"):
            continue
        body = child.find(q("p:txBody") + "/" + q("a:bodyPr"))
        if body is None:
            continue
        norm = body.find(q("a:normAutofit"))
        if norm is None:
            continue
        fs = norm.get("fontScale")
        if fs:
            try:
                out[idx] = int(fs) / 100000.0
            except (TypeError, ValueError):
                pass


def _slide_extras(deck, z):
    """{页号(1-based): {"show_master_sp": bool, "font_scale": {shape_idx: ratio}}}"""
    out = {}
    for sl in deck.slides:
        entry = {"show_master_sp": True, "font_scale": {}}
        try:
            root = ET.fromstring(z.read(sl.part))
        except Exception:
            out[sl.index] = entry
            continue
        if root.get("showMasterSp") in ("0", "false"):
            entry["show_master_sp"] = False
        tree = root.find(q("p:cSld") + "/" + q("p:spTree"))
        if tree is not None:
            _walk_font_scale(tree, entry["font_scale"], [0])
        out[sl.index] = entry
    return out


# --------------------------------------------------------------------------
# 检查 1：structure（ERROR）
# --------------------------------------------------------------------------
def _check_structure(z, rep, stats):
    from collections import Counter
    st = stats["checks"]["structure"] = {
        "parts": 0, "xml_parts": 0, "rels_files": 0, "rel_targets": 0,
        "content_type_total": 0, "content_type_defaults": 0,
        "content_type_overrides": 0, "slide_ids": 0, "slide_layout_rels": 0,
        "notes_rels": 0, "rel_refs": 0,
    }
    names = _part_names(z)
    nameset = set(names)
    st["parts"] = len(names)

    # ---- 1. 部件名唯一（OPC 要求；zip 允许重复，WPS 会取任意一条）----
    for n, c in Counter(names).items():
        if c > 1:
            rep.add("structure", "ERROR", "DUP_PART",
                    "ZIP 内同名部件出现 %d 次：%s" % (c, n))
    folded = {}
    for n in names:
        folded.setdefault(n.lower(), []).append(n)
    for low, lst in folded.items():
        if len(lst) > 1:
            rep.add("structure", "ERROR", "DUP_PART",
                    "部件名仅大小写不同（OPC 冲突）：%s" % " / ".join(lst))

    # ---- 2. 每个 xml/rels 良构 ----
    roots = {}          # 只留 rels / presentation / content-types，控制内存
    for n in names:
        if not (n.endswith(".xml") or n.endswith(".rels")):
            continue
        st["xml_parts"] += 1
        try:
            root = ET.fromstring(z.read(n))
        except Exception as e:
            rep.add("structure", "ERROR", "XML_PARSE",
                    "XML 解析失败 %s：%s: %s" % (n, type(e).__name__, e))
            continue
        if n.endswith(".rels") or n in ("ppt/presentation.xml",
                                        "[Content_Types].xml"):
            roots[n] = root

    # ---- 3. 关系断链 + 部件级 rId 表 ----
    rels_by_part = {}
    for n in names:
        if not n.endswith(".rels"):
            continue
        st["rels_files"] += 1
        root = roots.get(n)
        if root is None:      # 解析失败已报 ERROR
            continue
        base = posixpath.dirname(posixpath.dirname(n))
        part = _part_of_rels(n)
        relmap = {}
        for rel in root:
            if rel.tag.rsplit("}", 1)[-1] != "Relationship":
                continue
            rid = rel.get("Id") or ""
            if not rid:
                rep.add("structure", "ERROR", "DUP_REL_ID",
                        "%s 中存在缺少 Id 的关系" % n)
                continue
            if rid in relmap:
                rep.add("structure", "ERROR", "DUP_REL_ID",
                        "%s 中关系 Id 重复：%s" % (n, rid))
            mode = rel.get("TargetMode") or "Internal"
            tgt_raw = rel.get("Target") or ""
            relmap[rid] = {
                "type": (rel.get("Type") or "").rsplit("/", 1)[-1],
                "mode": mode,
                "target": tgt_raw,
            }
            if mode == "External" or tgt_raw.startswith(("http://", "https://")):
                continue
            st["rel_targets"] += 1
            cands = _resolve_target(base, tgt_raw)
            hit = next((c for c in cands if c in nameset), None)
            if hit is None:
                rep.add("structure", "ERROR", "BROKEN_REL",
                        "关系断链 %s → %s（期望部件 %s）"
                        % (n, tgt_raw, cands[0] if cands else "?"))
            else:
                relmap[rid]["resolved"] = hit
        rels_by_part[part] = relmap

    # ---- 4. [Content_Types].xml 覆盖所有部件 ----
    ct = roots.get("[Content_Types].xml")
    defaults, overrides = {}, []
    if ct is None:
        rep.add("structure", "ERROR", "MISSING_CT",
                "缺少 [Content_Types].xml（WPS 直接打不开）")
    else:
        for el in ct:
            tag = el.tag.rsplit("}", 1)[-1]
            ctype = el.get("ContentType") or ""
            if tag == "Default":
                ext = (el.get("Extension") or "").lower()
                if not ext:
                    rep.add("structure", "ERROR", "CT_EMPTY",
                            "[Content_Types].xml 存在无 Extension 的 Default")
                    continue
                if ext in defaults:
                    rep.add("structure", "ERROR", "DUP_CT_DEFAULT",
                            "[Content_Types].xml 重复 Default 扩展名：%s" % ext)
                defaults[ext] = ctype
                if not ctype:
                    rep.add("structure", "ERROR", "CT_EMPTY",
                            "Default 扩展名 %s 的 ContentType 为空" % ext)
            elif tag == "Override":
                pn = (el.get("PartName") or "").lstrip("/")
                if not pn:
                    rep.add("structure", "ERROR", "CT_EMPTY",
                            "[Content_Types].xml 存在无 PartName 的 Override")
                    continue
                overrides.append(pn)
                if not ctype:
                    rep.add("structure", "ERROR", "CT_EMPTY",
                            "Override %s 的 ContentType 为空" % pn)
        st["content_type_defaults"] = len(defaults)
        st["content_type_overrides"] = len(overrides)
        st["content_type_total"] = len(defaults) + len(overrides)
        for pn, c in Counter(overrides).items():
            if c > 1:
                rep.add("structure", "ERROR", "DUP_CT_OVERRIDE",
                        "[Content_Types].xml 重复 Override：%s（%d 次）" % (pn, c))
        ovset = set(overrides)
        for n in names:
            if n == "[Content_Types].xml" or n in ovset:
                continue
            bn = n.rsplit("/", 1)[-1]
            ext = bn.rsplit(".", 1)[-1].lower() if "." in bn else ""
            if ext not in defaults:
                rep.add("structure", "ERROR", "MISSING_CT_PART",
                        "部件既无 Override 也无对应 Default 扩展名：%s" % n)

    # ---- 5. sldIdLst 每项指向存在的 slide 部件 ----
    slide_parts = []
    pres = roots.get("ppt/presentation.xml")
    if pres is None:
        if "ppt/presentation.xml" not in names:
            rep.add("structure", "ERROR", "MISSING_PRES",
                    "缺少 ppt/presentation.xml（WPS 直接打不开）")
    else:
        pres_rels = rels_by_part.get("ppt/presentation.xml", {})
        lst = pres.find(q("p:sldIdLst"))
        if lst is None:
            rep.add("structure", "ERROR", "NO_SLDIDLST",
                    "presentation.xml 缺少 sldIdLst（没有页序）")
        else:
            seen_ids = set()
            for sid in list(lst):
                if sid.tag != q("p:sldId"):
                    continue
                st["slide_ids"] += 1
                iid = sid.get("id")
                if iid in seen_ids:
                    rep.add("structure", "ERROR", "DUP_SLD_ID",
                            "sldIdLst 中 sldId/@id 重复：%s" % iid)
                seen_ids.add(iid)
                rid = sid.get(q("r:id"))
                rel = pres_rels.get(rid)
                if rel is None:
                    rep.add("structure", "ERROR", "SLDID_RID_MISSING",
                            "sldIdLst 引用了 relationships 里不存在的关系 %s" % rid)
                    continue
                if rel["type"] != "slide":
                    rep.add("structure", "ERROR", "SLDID_TYPE",
                            "sldId 关系 %s 类型为 %s（应为 slide）"
                            % (rid, rel["type"]))
                tgt = rel.get("resolved")
                if not tgt or tgt not in nameset:
                    rep.add("structure", "ERROR", "SLDID_TARGET",
                            "sldId 指向不存在的 slide 部件：%s"
                            % (tgt or rel["target"]))
                    continue
                slide_parts.append(tgt)

    # ---- 6. 每页 layout 关系 + notesSlide 关系 ----
    for part in slide_parts:
        relmap = rels_by_part.get(part, {})
        layouts = [r for r in relmap.values() if r["type"] == "slideLayout"]
        if not layouts:
            rep.add("structure", "ERROR", "NO_LAYOUT",
                    "%s 没有 slideLayout 关系（read/preview 无法继承版式）" % part)
        else:
            st["slide_layout_rels"] += 1
            if not any(r.get("resolved") in nameset for r in layouts):
                rep.add("structure", "ERROR", "LAYOUT_TARGET",
                        "%s 的 slideLayout 关系目标缺失：%s"
                        % (part, layouts[0]["target"]))
        for r in relmap.values():
            if r["type"] != "notesSlide":
                continue
            st["notes_rels"] += 1
            if r.get("resolved") not in nameset:
                rep.add("structure", "ERROR", "NOTES_TARGET",
                        "%s 的 notesSlide 关系目标缺失：%s"
                        % (part, r["target"]))

    # ---- 7. XML 内 r:embed/r:id 等引用必须能在本部件 relationships 找到 ----
    for n in names:
        if not n.endswith(".xml"):
            continue
        if not n.startswith(("ppt/slides/", "ppt/slideLayouts/",
                             "ppt/slideMasters/")):
            continue
        raw = z.read(n).decode("utf-8", "replace")
        relmap = rels_by_part.get(n, {})
        for m in _RE_RREF.finditer(raw):
            rid = m.group(2)
            st["rel_refs"] += 1
            if rid not in relmap:
                rep.add("structure", "ERROR", "BROKEN_RREF",
                        "%s 引用了不存在的关系 %s（r:%s）"
                        % (n, rid, m.group(1)))
    return names, rels_by_part


# --------------------------------------------------------------------------
# 检查 2：units（ERROR）
# --------------------------------------------------------------------------
def _check_units(z, rep, stats):
    st = stats["checks"]["units"] = {
        "parts": 0, "xml_bytes": 0, "findings": 0,
    }
    names = _part_names(z)
    prefixes = ("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/")
    parts = sorted(n for n in names
                   if n.endswith(".xml") and n.startswith(prefixes))
    for n in parts:
        raw = z.read(n).decode("utf-8", "replace")
        st["parts"] += 1
        st["xml_bytes"] += len(raw.encode("utf-8"))
        for msg in units.guard_xml(raw, n):
            st["findings"] += 1
            rep.add("units", "ERROR", "UNIT_SCALE", msg)
    return parts


# --------------------------------------------------------------------------
# 检查 3：inherit（ERROR）—— 双图标
# --------------------------------------------------------------------------
def _check_inherit(deck, extras, rep, stats):
    st = stats["checks"]["inherit"] = {
        "slides": len(deck.slides), "slides_with_inherited": 0,
        "inherited_shapes": 0, "slide_pics": 0, "media_pairs": 0,
        "skipped_show_master": 0,
    }
    for sl in deck.slides:
        st["inherited_shapes"] += len(sl.inherited)
        if sl.inherited:
            st["slides_with_inherited"] += 1
        page_pics = [s for s in sl.shapes if s.kind == "pic" and s.media]
        st["slide_pics"] += len(page_pics)

        # (a) 页面内同 media 出现两次 → 同素材重复绘制
        by_media = {}
        for s in page_pics:
            by_media.setdefault(s.media, []).append(s)
        for media, lst in by_media.items():
            if len(lst) < 2:
                continue
            for extra in lst[1:]:
                rep.add(
                    "inherit", "ERROR", "DUP_MEDIA",
                    "页面内同一素材重复出现：%s 共 %d 次（%s）；同页不应"
                    "引用同一 media 多次，会叠出双份"
                    % (media, len(lst),
                       "、".join("「%s」" % s.name for s in lst)), sl.index)

        if not extras.get(sl.index, {}).get("show_master_sp", True):
            st["skipped_show_master"] += 1
            continue

        # (b) layout/master 继承下来的固定图片（排除图片占位符：
        #     占位符只是提示位置，由页面填充，不构成重复绘制）
        inherited_pics = [s for s in sl.inherited
                          if s.kind == "pic" and s.media and not s.placeholder]
        inh_by_media = {}
        for s in inherited_pics:
            inh_by_media.setdefault(s.media, []).append(s)

        for media, plist in by_media.items():
            ilist = inh_by_media.get(media, [])
            for ps in plist:
                rp = ps.rect()
                if not rp:
                    continue
                for ish in ilist:
                    ri = ish.rect()
                    if not ri:
                        continue
                    st["media_pairs"] += 1
                    area = _inter_area(rp, ri)
                    small = min(rp[2] * rp[3], ri[2] * ri[3])
                    if small <= 0 or area <= 0:
                        continue
                    ratio = area / float(small)
                    if ratio <= ICON_OVERLAP:
                        continue
                    where = "版式" if ish.src == "layout" else "母版"
                    rep.add(
                        "inherit", "ERROR", "DUP_ICON",
                        "双图标：页面图片「%s」与%s固定元素「%s」使用同一素材 "
                        "%s，矩形相交 %.0f%%（>%.0f%%）（页面 %.2f×%.2f cm @ "
                        "(%.2f, %.2f)，%s %.2f×%.2f cm @ (%.2f, %.2f)）"
                        % (ps.name, where, ish.name, media, ratio * 100,
                           ICON_OVERLAP * 100,
                           _cm(rp[2]), _cm(rp[3]), _cm(rp[0]), _cm(rp[1]),
                           where, _cm(ri[2]), _cm(ri[3]), _cm(ri[0]),
                           _cm(ri[1])), sl.index)

        # (c) layout 与 master 同时绘制同一素材（也属重复继承）
        for media, ilist in inh_by_media.items():
            if len(ilist) < 2:
                continue
            for i in range(len(ilist)):
                for j in range(i + 1, len(ilist)):
                    a, b = ilist[i].rect(), ilist[j].rect()
                    if not a or not b:
                        continue
                    area = _inter_area(a, b)
                    small = min(a[2] * a[3], b[2] * b[3])
                    if small <= 0 or area <= 0 or area / float(small) <= ICON_OVERLAP:
                        continue
                    rep.add(
                        "inherit", "ERROR", "DUP_ICON_INHERIT",
                        "双图标：版式「%s」与母版「%s」重复绘制同一素材 %s，"
                        "矩形相交 %.0f%%（>%.0f%%）"
                        % (ilist[i].name, ilist[j].name, media,
                           area / float(small) * 100, ICON_OVERLAP * 100),
                        sl.index)


# --------------------------------------------------------------------------
# 检查 4：layout（WARN）—— 越界 / 文字溢出 / 文字相交 / 图片拉伸
# --------------------------------------------------------------------------
def _para_metrics(para, avail_w, scale):
    """估算一个段落：返回 (所需高度 EMU, 行数, 字号 pt, 行距倍数, 最大行宽 EMU)。

    字号/高度一律走 units（units.text_height_emu 内含 1.5 倍保守行高）。
    固定行距（a:spcPts）换算成倍数后交给 units 的 spacing_ratio。
    """
    pt = max([r.pt for r in para.runs if r.pt] or [18.0]) * scale
    ratio = para.spc_ratio if para.spc_ratio else 1.0
    if para.spc_pts:
        try:
            fixed = units.pt_of_spc_pts(para.spc_pts)
            if fixed > 0 and pt > 0:
                ratio = fixed / pt
        except (TypeError, ValueError):
            pass
    if para.text:
        wrapped = units.wrap_lines(para.text, avail_w, pt)
        n = len(wrapped) if wrapped else 0
        h = units.text_height_emu(para.text, avail_w, pt, ratio)
        wmax = max([units.text_width_emu(line, pt) for line in wrapped] or [0])
    else:
        # 空段落（有 run 无文字）仍占一行；用 units.LINE_HEIGHT，不自造系数
        n = 1 if para.runs else 0
        h = int(round(pt * units.EMU_PER_PT * units.LINE_HEIGHT * ratio)) if n else 0
        wmax = 0
    return h, n, pt, ratio, wmax


def _estimate_text(sh, font_scale=1.0):
    """估算一个文本形状（avail_* 已按 read.Shape 的内边距扣除）。

    read.Shape 提供的 anchor/insets 直接使用：
      avail_w = w - lIns - rIns；avail_h = h - tIns - bIns；
      anchor  t（""/"t"）→ 文字从顶部排；"ctr" → 上下居中；"b" → 底部对齐。
    字号一律走 units（run.pt 由 read 用 units.pt_of_sz 解析过）；
    高度一律走 units.text_height_emu（内含 1.5 倍保守行高）。
    """
    l, t, r, b = sh.insets or (0, 0, 0, 0)
    avail_w = max(0, (sh.w or 0) - l - r)
    avail_h = max(0, (sh.h or 0) - t - b)
    scale = font_scale or 1.0
    need, lines_total, pt_max, wmax = 0, 0, 0.0, 0
    ratio_first, align_first = None, ""
    for para in sh.paras:
        h, n, pt, ratio, w = _para_metrics(para, max(1, avail_w), scale)
        need += h
        lines_total += n
        pt_max = max(pt_max, pt)
        wmax = max(wmax, w)
        if ratio_first is None:
            ratio_first = ratio
        if not align_first and para.align:
            align_first = para.align
    return {
        "need": need, "lines": lines_total,
        "pt": pt_max, "ratio": ratio_first if ratio_first else 1.0,
        "wmax": wmax, "avail_w": avail_w, "avail_h": avail_h,
        "align": align_first or "l",
        "anchor": sh.anchor or "t",
        "l": l, "t": t, "r": r, "b": b,
        "font_scale": scale,
    }


def _text_rect(sh, est):
    """文字内容的估算矩形（不是文本框矩形）——用于“文字与文字相交”。"""
    if est["need"] <= 0 or est["avail_w"] <= 0:
        return None
    block_w = min(est["wmax"], est["avail_w"]) if est["wmax"] > 0 else 0
    if block_w <= 0:
        return None
    block_h = est["need"]
    x = (sh.x or 0) + est["l"]
    y = (sh.y or 0) + est["t"]
    align = est["align"]
    if align == "ctr":
        x += (est["avail_w"] - block_w) // 2
    elif align in ("r", "just"):
        x += est["avail_w"] - block_w
    if est["anchor"] == "ctr":
        y += (est["avail_h"] - block_h) // 2
    elif est["anchor"] == "b":
        y += est["avail_h"] - block_h
    return (int(x), int(y), int(block_w), int(block_h))


def _check_layout(deck, extras, z, rep, stats,
                  bounds_tol=BOUNDS_TOL, aspect_tol=ASPECT_TOL,
                  text_tol=TEXT_OVERFLOW_TOL):
    st = stats["checks"]["layout"] = {
        "shapes_with_rect": 0, "bounds_checked": 0,
        "text_shapes": 0, "text_paras": 0, "text_pairs": 0,
        "pics": 0, "aspect_checked": 0, "aspect_skipped": 0,
        "text_rects": 0, "font_scale_applied": 0,
    }
    W, H = deck.w, deck.h
    for sl in deck.slides:
        entry = extras.get(sl.index) or {}
        fs_map = entry.get("font_scale") or {}

        # ---------- 4.1 元素越界（容差 2%）----------
        for sh in sl.shapes:
            r = sh.rect()
            if not r or r[2] <= 0 or r[3] <= 0:
                continue
            st["shapes_with_rect"] += 1
            x, y, w, h = r
            st["bounds_checked"] += 1
            # 完全落在画布外：内容一定看不到 → ERROR
            if x + w <= 0 or y + h <= 0 or x >= W or y >= H:
                rep.add("layout", "ERROR", "OUT_OF_CANVAS",
                        "「%s」%s 完全在画布外（元素 %.2f×%.2f cm @ (%.2f, %.2f)；"
                        "画布 %.2f×%.2f cm）"
                        % (sh.name, sh.kind, _cm(w), _cm(h), _cm(x), _cm(y),
                           _cm(W), _cm(H)), sl.index)
                continue
            over = []
            if x < -W * bounds_tol:
                over.append("左溢出 %.2f cm" % _cm(-x))
            if y < -H * bounds_tol:
                over.append("上溢出 %.2f cm" % _cm(-y))
            if x + w > W * (1 + bounds_tol):
                over.append("右溢出 %.2f cm（%.1f%% 画布宽）"
                            % (_cm(x + w - W), (x + w - W) / float(W) * 100))
            if y + h > H * (1 + bounds_tol):
                over.append("下溢出 %.2f cm（%.1f%% 画布高）"
                            % (_cm(y + h - H), (y + h - H) / float(H) * 100))
            if over:
                rep.add("layout", "WARN", "OUT_OF_CANVAS",
                        "「%s」超出画布：%s（容差 %.0f%%：%.2f×%.2f cm；"
                        "元素 %.2f×%.2f cm @ (%.2f, %.2f)）"
                        % (sh.name, "、".join(over), bounds_tol * 100,
                           _cm(W), _cm(H), _cm(w), _cm(h), _cm(x), _cm(y)),
                        sl.index)

        # ---------- 4.2 文字溢出 / 4.3 文字相交 ----------
        text_rects = []
        for sh in sl.shapes:
            if not sh.paras or not sh.text.strip():
                continue
            r = sh.rect()
            if not r or r[2] <= 0 or r[3] <= 0:
                continue
            scale = fs_map.get(sh.idx, 1.0)
            est = _estimate_text(sh, scale)
            if scale < 1.0:
                st["font_scale_applied"] += 1
            st["text_shapes"] += 1
            st["text_paras"] += len(sh.paras)
            if est["need"] <= 0:
                continue
            # 文字溢出：need > 框内可用高
            if est["avail_h"] > 0 and est["need"] > est["avail_h"] * text_tol:
                rep.add(
                    "layout", "WARN", "TEXT_OVERFLOW",
                    "「%s」文字需 %.2f cm，框高 %.2f cm（已扣上下内边距 %.2f cm，"
                    "可用 %.2f cm）；估算 %d 行 × %.1f pt × 行距 %.2f 倍"
                    "（含 1.5 倍保守行高），超出 %.1f%%%s"
                    % (sh.name, _cm(est["need"]), _cm(sh.h),
                       _cm(est["t"] + est["b"]), _cm(est["avail_h"]),
                       est["lines"], est["pt"], est["ratio"],
                       (est["need"] / float(est["avail_h"]) - 1) * 100,
                       "；已按 autofit fontScale=%.1f%% 缩字估算" % (scale * 100)
                       if scale < 1.0 else ""),
                    sl.index)
            # 文字矩形（用于相交）
            tr = _text_rect(sh, est)
            if tr:
                text_rects.append((sh, tr, est))
                st["text_rects"] += 1

        for i in range(len(text_rects)):
            sh_a, ra, ea = text_rects[i]
            for j in range(i + 1, len(text_rects)):
                sh_b, rb, eb = text_rects[j]
                st["text_pairs"] += 1
                inter = _inter_rect(ra, rb)
                if not inter:
                    continue
                rep.add(
                    "layout", "WARN", "TEXT_OVERLAP",
                    "「%s」与「%s」文字矩形相交 %.2f×%.2f cm（面积 %.3f cm²）："
                    "A 文字 %.2f×%.2f cm @ (%.2f, %.2f)，B 文字 %.2f×%.2f cm @ "
                    "(%.2f, %.2f)（按 %.0f%% 保守行高估算，请结合预览复核）"
                    % (sh_a.name, sh_b.name,
                       _cm(inter[2]), _cm(inter[3]),
                       _cm(inter[2]) * _cm(inter[3]),
                       _cm(ra[2]), _cm(ra[3]), _cm(ra[0]), _cm(ra[1]),
                       _cm(rb[2]), _cm(rb[3]), _cm(rb[0]), _cm(rb[1]),
                       units.LINE_HEIGHT * 100),
                    sl.index)

        # ---------- 4.4 图片拉伸（像素纵横比 vs 框纵横比）----------
        for sh in sl.pictures:
            r = sh.rect()
            if not r or r[2] <= 0 or r[3] <= 0:
                continue
            st["pics"] += 1
            if not sh.media:
                st["aspect_skipped"] += 1
                continue
            try:
                raw = z.read(sh.media)
            except KeyError:
                st["aspect_skipped"] += 1
                continue
            px = image_size(raw)
            if not px:
                st["aspect_skipped"] += 1
                continue
            pw, ph = px
            crop = sh.crop or ()
            if len(crop) == 4:
                l, t, rr, bb = crop
            else:
                l = t = rr = bb = 0
            if l + rr >= 100000 or t + bb >= 100000:
                st["aspect_skipped"] += 1
                continue
            crop_w = pw * (100000 - l - rr) / 100000.0
            crop_h = ph * (100000 - t - bb) / 100000.0
            if crop_w <= 0 or crop_h <= 0:
                st["aspect_skipped"] += 1
                continue
            px_ar = crop_w / crop_h
            box_ar = r[2] / float(r[3])
            if box_ar <= 0:
                st["aspect_skipped"] += 1
                continue
            st["aspect_checked"] += 1
            err = abs(px_ar - box_ar) / box_ar
            if err > aspect_tol:
                rep.add(
                    "layout", "WARN", "IMAGE_STRETCH",
                    "「%s」图片被拉伸：像素 %d×%d%s（裁切后 AR %.3f）vs "
                    "框 %.2f×%.2f cm（AR %.3f），偏差 %.1f%% > %.0f%%"
                    % (sh.name, pw, ph,
                       "（含 srcRect 裁切 %d/%d/%d/%d）" % (l, t, rr, bb)
                       if len(crop) == 4 else "", px_ar,
                       _cm(r[2]), _cm(r[3]), box_ar, err * 100,
                       aspect_tol * 100),
                    sl.index)


# --------------------------------------------------------------------------
# 主入口
# --------------------------------------------------------------------------
def _load_deck(path):
    try:
        return read.Deck.load(path), None
    except Exception as e:
        return None, "%s: %s" % (type(e).__name__, e)


def _summarize_deck(deck, stats):
    stats["slides"] = len(deck.slides)
    stats["layouts"] = len(deck.layouts)
    stats["masters"] = len(deck.masters)
    stats["media"] = len(deck.media)
    stats["shapes"] = sum(len(s.shapes) for s in deck.slides)
    stats["pics"] = sum(len(s.pictures) for s in deck.slides)
    stats["text_chars"] = sum(len(s.text) for s in deck.slides)
    stats["notes"] = sum(1 for s in deck.slides if s.notes)
    stats["notes_chars"] = sum(len(s.notes) for s in deck.slides)
    stats["inherited_shapes"] = sum(len(s.inherited) for s in deck.slides)


def check(path, level="all", **kw):
    """校验一个 .pptx，返回 Report。

    level: "all" 或 structure / units / inherit / layout 之一。
    可选关键字（估算阈值，一般用默认）：
      bounds_tol（画布越界容差，默认 2%）
      aspect_tol（图片纵横比容差，默认 3%）
      text_tol  （文字溢出倍数阈值，默认 1.0）
    """
    if level not in ("all",) + LEVELS:
        raise ValueError("level 必须是 all 或 %s 之一，收到 %r"
                         % ("/".join(LEVELS), level))
    t0 = time.time()
    rep = Report(path=path, level=level)
    stats = {"path": path, "level": level, "checks": {}}
    rep.stats = stats
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        rep.add("structure", "ERROR", "OPEN_FAIL",
                "无法作为 ZIP 打开：%s: %s" % (type(e).__name__, e))
        rep.elapsed = time.time() - t0
        stats["elapsed_s"] = round(rep.elapsed, 3)
        return rep

    levels = set(LEVELS) if level == "all" else {level}
    with z:
        if "structure" in levels:
            _check_structure(z, rep, stats)
            # 规模指标统一放 stats 顶层，供 as_text 直接读取
            stats["parts"] = stats["checks"]["structure"]["parts"]
        else:
            stats["parts"] = len(_part_names(z))
        if "units" in levels:
            _check_units(z, rep, stats)
        if levels & {"inherit", "layout"}:
            deck, err = _load_deck(path)
            if deck is None:
                rep.add("structure", "ERROR", "IR_LOAD",
                        "IR 加载失败（结构损坏或非法 XML）：%s" % err)
            else:
                _summarize_deck(deck, stats)
                extras = _slide_extras(deck, z)
                if "inherit" in levels:
                    _check_inherit(deck, extras, rep, stats)
                if "layout" in levels:
                    _check_layout(deck, extras, z, rep, stats,
                                  bounds_tol=kw.get("bounds_tol", BOUNDS_TOL),
                                  aspect_tol=kw.get("aspect_tol", ASPECT_TOL),
                                  text_tol=kw.get("text_tol", TEXT_OVERFLOW_TOL))
        # 未加载 IR 时至少给出页数（从 structure 的 sldIdLst 计数）
        if "slides" not in stats:
            sld = stats.get("checks", {}).get("structure", {}).get("slide_ids")
            stats["slides"] = sld or 0
    rep.elapsed = time.time() - t0
    stats["elapsed_s"] = round(rep.elapsed, 3)
    return rep


# --------------------------------------------------------------------------
# 坏样本构造（--self-test 用；产物写到 out/selftest/）
# --------------------------------------------------------------------------
def _read_parts(path):
    with zipfile.ZipFile(path) as z:
        return [(n, z.read(n)) for n in z.namelist() if not n.endswith("/")]


def _write_parts(path, items):
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in items:
            z.writestr(n, data)
    os.replace(tmp, path)
    return path


def _find_slide_in(by_name, order, needle):
    """返回第一个包含 needle 的页部件名。"""
    for part in order:
        if needle in by_name.get(part, b"").decode("utf-8", "replace"):
            return part
    return order[0] if order else None


def _pic_xml(rid, name, x, y, cx, cy):
    return (
        '<p:pic><p:nvPicPr><p:cNvPr id="9001" name="%s"/><p:cNvPicPr>'
        '<a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/>'
        '</a:stretch></p:blipFill><p:spPr><a:xfrm>'
        '<a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        % (name, rid, x, y, cx, cy))


def _add_rel(xml_text, rid, rtype, target):
    rel = ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
           'officeDocument/2006/relationships/%s" Target="%s"/>'
           % (rid, rtype, target))
    if "</Relationships>" not in xml_text:
        return xml_text
    return xml_text.replace("</Relationships>", rel + "</Relationships>", 1)


def _insert_pic(xml_text, pic_fragment):
    if "</p:spTree>" not in xml_text:
        return None
    return xml_text.replace("</p:spTree>", pic_fragment + "</p:spTree>", 1)


def _first_shape_ext(raw, min_cx=0):
    """在 XML 里找第一个 <a:off .../><a:ext cx=".." cy=".."/>，返回 (start,end,vals)。"""
    for m_off in _RE_OFF.finditer(raw):
        seg = raw[m_off.end():m_off.end() + 400]
        m_ext = _RE_EXT.search(seg)
        if not m_ext:
            continue
        cx, cy = int(m_ext.group(1)), int(m_ext.group(2))
        if cx < min_cx or cy <= 0:
            continue
        return m_off, m_ext, cx, cy
    return None, None, 0, 0


def build_bad_samples(src_pptx, outdir):
    """构造坏样本，返回 [ {name, path, code, level, why} ]。

    每个样本只注入一处缺陷，改动都发生在 ZIP 部件层面（真实 PPTX 文件）。
    """
    os.makedirs(outdir, exist_ok=True)
    samples = []
    base = _read_parts(src_pptx)
    by_name = dict(base)

    with zipfile.ZipFile(src_pptx) as z:
        names = read.part_names(z)
        order = read.slide_parts(z, names)
    slide3 = order[2] if len(order) > 2 else order[0]
    layout1 = next((n for n in names
                    if n.startswith("ppt/slideLayouts/")
                    and n.endswith(".xml")), None)
    media = next((n for n in sorted(names)
                  if n.startswith("ppt/media/") and n.lower().endswith(".png")), None)
    if not media:
        media = next((n for n in sorted(names) if n.startswith("ppt/media/")), None)

    # ---- 坏样本 1：spcPct=100（0.1 倍行距）→ units ERROR ----
    part = _find_slide_in(by_name, order, "<a:lnSpc")
    items = list(base)
    done = False
    for i, (n, d) in enumerate(items):
        if n != part:
            continue
        raw = d.decode("utf-8")
        m = re.search(r'<a:lnSpc\b[^>]*>\s*<a:spcPct\b[^>]*\bval="(\d+)"', raw)
        if m:
            raw = raw[:m.start(1)] + "100" + raw[m.end(1):]
            done = True
        items[i] = (n, raw.encode("utf-8"))
    if not done:
        raise RuntimeError("坏样本构造失败：找不到可改的 lnSpc")
    samples.append({
        "name": "bad_spcPct100.pptx",
        "path": _write_parts(os.path.join(outdir, "bad_spcPct100.pptx"), items),
        "code": "UNIT_SCALE", "level": "ERROR",
        "expect_msg": "spcPct",
        "why": "把 a:lnSpc/a:spcPct 的 100000 改成 100（=0.1 倍行距，2026-10 事故）",
    })

    # ---- 坏样本 2：元素部分越界 → layout WARN ----
    items = list(base)
    for i, (n, d) in enumerate(items):
        if n != slide3:
            continue
        raw = d.decode("utf-8")
        m_off, m_ext, cx, cy = _first_shape_ext(raw, min_cx=1000000)
        if m_ext is None:
            raise RuntimeError("坏样本构造失败：找不到足够大的元素")
        new_x = 12192000 - cx // 4          # 右边界超出画布 3/4 个元素宽
        raw = raw[:m_off.start(1)] + str(new_x) + raw[m_off.end(1):]
        items[i] = (n, raw.encode("utf-8"))
    samples.append({
        "name": "bad_bounds_partial.pptx",
        "path": _write_parts(os.path.join(outdir, "bad_bounds_partial.pptx"), items),
        "code": "OUT_OF_CANVAS", "level": "WARN",
        "expect_msg": "超出画布",
        "why": "把页面元素 x 移到画布右缘外（部分越界，容差 2% 之外）",
    })

    # ---- 坏样本 3：元素完全在画布外 → layout ERROR ----
    items = list(base)
    for i, (n, d) in enumerate(items):
        if n != slide3:
            continue
        raw = d.decode("utf-8")
        m_off, m_ext, cx, cy = _first_shape_ext(raw)
        if m_ext is None:
            raise RuntimeError("坏样本构造失败：找不到可移元素")
        raw = raw[:m_off.start(1)] + "13000000" + raw[m_off.end(1):]
        items[i] = (n, raw.encode("utf-8"))
    samples.append({
        "name": "bad_bounds_full.pptx",
        "path": _write_parts(os.path.join(outdir, "bad_bounds_full.pptx"), items),
        "code": "OUT_OF_CANVAS", "level": "ERROR",
        "expect_msg": "完全在画布外",
        "why": "把页面元素 x 移到画布外（x >= 画布宽，内容一定看不到）",
    })

    # ---- 坏样本 4：layout 与页面同 media 双图标 → inherit ERROR ----
    if not (layout1 and media):
        raise RuntimeError("坏样本构造失败：缺 layout 或 media 部件")
    layout_rels = _rels_part_of(layout1)
    slide_rels = _rels_part_of(slide3)
    media_target = posixpath.relpath(media, posixpath.dirname(layout1))
    slide_media_target = posixpath.relpath(media, posixpath.dirname(slide3))
    items = list(base)
    hit = {"layout": False, "layout_rels": False,
           "slide": False, "slide_rels": False}
    for i, (n, d) in enumerate(items):
        if n not in (layout1, layout_rels, slide3, slide_rels):
            continue          # 只解码 xml/rels 部件（media 是二进制）
        raw = d.decode("utf-8")
        if n == layout1:
            frag = _insert_pic(raw, _pic_xml("rIdLay900", "LayoutLogo",
                                             500000, 500000, 1000000, 1000000))
            if frag is not None:
                items[i] = (n, frag.encode("utf-8"))
                hit["layout"] = True
        elif n == layout_rels:
            items[i] = (n, _add_rel(raw, "rIdLay900", "image",
                                    media_target).encode("utf-8"))
            hit["layout_rels"] = True
        elif n == slide3:
            frag = _insert_pic(raw, _pic_xml("rIdSld901", "PageLogo",
                                             500000, 500000, 1000000, 1000000))
            if frag is not None:
                items[i] = (n, frag.encode("utf-8"))
                hit["slide"] = True
        elif n == slide_rels:
            items[i] = (n, _add_rel(raw, "rIdSld901", "image",
                                    slide_media_target).encode("utf-8"))
            hit["slide_rels"] = True
    if not all(hit.values()):
        raise RuntimeError("坏样本构造失败：双图标部件注入不完整 %r" % hit)
    samples.append({
        "name": "bad_dup_icon.pptx",
        "path": _write_parts(os.path.join(outdir, "bad_dup_icon.pptx"), items),
        "code": "DUP_ICON", "level": "ERROR",
        "expect_msg": "LayoutLogo",
        "why": "在 slideLayout 与页面里各放一份同一 media 的图片（相交 100%），"
               "复刻 2026-10 的双图标",
    })

    # ---- 坏样本 5：文字溢出 → layout WARN ----
    # 选一个基线**不溢出**的文本框，注入后才能证明是新报出来的
    target_name = "文本 503"
    items = list(base)
    done = False
    for i, (n, d) in enumerate(items):
        if n != slide3:
            continue
        raw = d.decode("utf-8")
        pos = raw.find('name="%s"' % target_name)
        if pos < 0:
            continue
        seg = raw[pos:pos + 4000]
        m_ext = _RE_EXT.search(seg)
        if not m_ext:
            continue
        raw = raw[:pos + m_ext.start(2)] + "100000" + raw[pos + m_ext.end(2):]
        items[i] = (n, raw.encode("utf-8"))
        done = True
    if not done:
        raise RuntimeError("坏样本构造失败：找不到文本框 %s" % target_name)
    samples.append({
        "name": "bad_text_overflow.pptx",
        "path": _write_parts(os.path.join(outdir, "bad_text_overflow.pptx"), items),
        "code": "TEXT_OVERFLOW", "level": "WARN",
        "expect_msg": target_name,
        "why": "把「%s」的框高压到 0.28 cm，文字需要多行" % target_name,
    })

    # ---- 坏样本 6：图片拉伸 → layout WARN ----
    pic_names = []
    with zipfile.ZipFile(src_pptx) as z:
        raw3 = z.read(slide3).decode("utf-8")
    for m in _RE_CNVPR_NAME.finditer(raw3):
        pic_names.append(m.group(1))
    target_pic = None
    for nm in pic_names:
        pos = raw3.find('name="%s"' % nm)
        # 只认 p:pic 里的名字（p:pic 元素前有 <p:pic>）
        if 0 <= pos and raw3.rfind("<p:pic>", 0, pos) > raw3.rfind("<p:sp>", 0, pos):
            target_pic = nm
            break
    if not target_pic:
        raise RuntimeError("坏样本构造失败：slide3 里找不到图片")
    items = list(base)
    for i, (n, d) in enumerate(items):
        if n != slide3:
            continue
        raw = d.decode("utf-8")
        pos = raw.find('name="%s"' % target_pic)
        seg = raw[pos:pos + 6000]
        m_ext = _RE_EXT.search(seg)
        if not m_ext:
            raise RuntimeError("坏样本构造失败：找不到图片 xfrm")
        cx, cy = int(m_ext.group(1)), int(m_ext.group(2))
        raw = raw[:pos + m_ext.start(2)] + str(cx * 2) + raw[pos + m_ext.end(2):]
        items[i] = (n, raw.encode("utf-8"))
    samples.append({
        "name": "bad_image_stretch.pptx",
        "path": _write_parts(os.path.join(outdir, "bad_image_stretch.pptx"), items),
        "code": "IMAGE_STRETCH", "level": "WARN",
        "expect_msg": target_pic,
        "why": "把图片「%s」的框宽改成 2 倍（像素纵横比与框纵横比必然偏差）"
               % target_pic,
    })

    # ---- 坏样本 7：关系断链（删除被页面引用的 media 部件）→ structure ERROR ----
    used_media = None
    with zipfile.ZipFile(src_pptx) as z:
        for part in order:
            raw = z.read(part).decode("utf-8", "replace")
            m_embed = re.search(r'r:embed="([^"]+)"', raw)
            if not m_embed:
                continue
            rels_name = _rels_part_of(part)
            if rels_name not in names:
                continue
            try:
                rels_root = ET.fromstring(z.read(rels_name))
            except Exception:
                continue
            for rel in rels_root:
                if (rel.get("Id") == m_embed.group(1)
                        and (rel.get("TargetMode") or "Internal") != "External"):
                    used_media = posixpath.normpath(posixpath.join(
                        posixpath.dirname(part), rel.get("Target") or ""))
                    break
            if used_media:
                break
    if not used_media:
        raise RuntimeError("坏样本构造失败：找不到被页面引用的 media")
    items = [(n, d) for (n, d) in base if n != used_media]
    samples.append({
        "name": "bad_broken_rel.pptx",
        "path": _write_parts(os.path.join(outdir, "bad_broken_rel.pptx"), items),
        "code": "BROKEN_REL", "level": "ERROR",
        "expect_msg": "关系断链",
        "why": "删除被页面图片引用的 media 部件 %s，制造关系断链" % used_media,
    })

    # ---- 坏样本 8：部件同时失去 Default 与 Override → structure ERROR ----
    # deck-b.pptx 的部件几乎都有单独 Override；把 png 的 Default 和一个 png 的
    # Override 同时删掉，就有一个 media 部件“既无 Override 也无 Default”。
    items = []
    hit_ct = False
    for n, d in base:
        if n == "[Content_Types].xml":
            raw = d.decode("utf-8")
            new = re.sub(r'<Default\b[^>]*\bExtension="png"[^>]*/>', "",
                         raw, count=1)
            new = re.sub(
                r'<Override\b[^>]*\bPartName="/ppt/media/[^"]*\.png"[^>]*/>',
                "", new, count=1)
            if new != raw:
                d = new.encode("utf-8")
                hit_ct = True
        items.append((n, d))
    if not hit_ct:
        raise RuntimeError("坏样本构造失败：找不到可删的 png Default/Override")
    samples.append({
        "name": "bad_content_types.pptx",
        "path": _write_parts(os.path.join(outdir, "bad_content_types.pptx"), items),
        "code": "MISSING_CT_PART", "level": "ERROR",
        "expect_msg": "扩展名",
        "why": "同时删掉 [Content_Types].xml 里 png 的 Default 和某个 png 的 "
               "Override，使该 media 部件无内容类型覆盖",
    })
    return samples


# --------------------------------------------------------------------------
# 自测：语料回归 + 坏样本证伪
# --------------------------------------------------------------------------
def _fmt_rep(rep):
    st = rep.stats or {}
    return ("%d ERROR / %d WARN  %.3fs  slides=%s shapes=%s pics=%s "
            "inherited=%s" % (rep.count("ERROR"), rep.count("WARN"),
                              rep.elapsed, st.get("slides"),
                              st.get("shapes"), st.get("pics"),
                              st.get("inherited_shapes")))


def self_test(corpus=BASE_CORPUS, outdir=None):
    """返回 0 = 全部通过；1 = 有失败；没有语料时打印 SKIP 并返回 0。

    两部分：
      A. 三份示例语料：deck-b.pptx / 合并稿必须 0 ERROR；deck-a.pptx 必须报出
         4 条已知双图标（基线白名单见 KNOWN_BASELINE）；
      B. 坏样本证伪：每个样本只注入一处缺陷，断言比基线多报出该 code，
         且消息里能定位到注入点。
    """
    if not corpus:
        print("SKIP check.py 自测：未设置 MAKEPPT_CORPUS（示例语料目录）")
        return 0
    if os.path.isfile(corpus):
        corpus = os.path.dirname(os.path.abspath(corpus))
    if not os.path.isdir(corpus):
        print("SKIP check.py 自测：MAKEPPT_CORPUS 指向的目录 %s 不存在" % corpus)
        return 0
    outdir = outdir or os.path.join(OUT_DIR, "selftest")
    os.makedirs(outdir, exist_ok=True)
    fails = []
    log = []

    def say(msg=""):
        print(msg)
        log.append(str(msg))

    def finish(rc):
        try:
            os.makedirs(OUT_DIR, exist_ok=True)
            with open(os.path.join(OUT_DIR, "check-selftest.txt"),
                      "w", encoding="utf-8") as fh:
                fh.write("\n".join(log) + "\n")
        except Exception as e:      # 报告写不进去不应打断自测
            print("（自测日志写入失败：%s）" % e)
        return rc

    say("=" * 72)
    say("check.py 自测：A) 回归语料  B) 坏样本证伪")
    say("语料目录: %s" % corpus)
    say("=" * 72)

    # ---------- A. 语料回归 ----------
    say("\n[A] 回归语料（refs/ir.md 表）")
    reps = {}
    for name, pages in CORPUS:
        path = os.path.join(corpus, name)
        if not os.path.isfile(path):
            fails.append("[A] 缺少语料 %s（--corpus 指向别处？）" % path)
            say("  [SKIP] %-26s 文件不存在" % name)
            continue
        rep = check(path)
        reps[name] = rep
        say("  %-26s %s" % (name, _fmt_rep(rep)))
        # 顺手落盘完整报告，便于 Lead 复核
        try:
            fname = "check-report-%s.txt" % os.path.splitext(name)[0]
            with open(os.path.join(OUT_DIR, fname), "w", encoding="utf-8") as fh:
                fh.write(rep.as_text() + "\n")
        except Exception as e:
            say("  （%s 报告写入失败：%s）" % (name, e))
        if rep.stats.get("slides") != pages:
            fails.append("[A] %s 页数 %s != 期望 %d"
                         % (name, rep.stats.get("slides"), pages))
        base = KNOWN_BASELINE.get(name)
        if base:
            errs = rep.error_findings
            codes = {f.code for f in errs}
            kinds = {f.check for f in errs}
            if (len(errs) != base["error_count"] or codes != base["error_codes"]
                    or kinds != base["error_checks"]):
                fails.append(
                    "[A] %s 已知缺陷基线不符：期望 %d 条 %s（检查项 %s），"
                    "实际 %d 条 %s（检查项 %s）"
                    % (name, base["error_count"], sorted(base["error_codes"]),
                       sorted(base["error_checks"]), len(errs), sorted(codes),
                       sorted(kinds)))
            else:
                say("      ↳ 已知基线：%d 条 %s / 检查项 %s（%s）——"
                    "真实缺陷，如实报 ERROR"
                    % (len(errs), "、".join(sorted(codes)),
                       "、".join(sorted(kinds)), base["why"]))
        elif rep.count("ERROR"):
            fails.append("[A] %s 出现 %d 条 ERROR（期望 0）"
                         % (name, rep.count("ERROR")))
        # 覆盖数不为 0，否则“0 ERROR”没有说服力
        checks = rep.stats.get("checks") or {}
        for lv in LEVELS:
            if lv in checks and not checks[lv]:
                fails.append("[A] %s 的 %s 检查对象计数为空" % (name, lv))
        if checks.get("layout", {}).get("text_shapes", 0) <= 0:
            fails.append("[A] %s 没有检查到任何文本形状" % name)
        if checks.get("units", {}).get("parts", 0) <= 0:
            fails.append("[A] %s 没有扫描到任何单位检查部件" % name)

    # ---------- A2. 旧缺陷回归（备份存在时验证“真缺陷仍能被抓到”） ----------
    bak = os.path.join(corpus, LEGACY_DEFECT_BACKUP)
    if os.path.isfile(bak):
        say("\n[A2] 旧缺陷回归：%s（2026-10-28 修复前的真实双图标）"
            % LEGACY_DEFECT_BACKUP)
        rep_bak = check(bak)
        errs = rep_bak.error_findings
        codes = {f.code for f in errs}
        if len(errs) != 4 or codes != {"DUP_ICON"}:
            fails.append("[A2] 备份应报 4 条 DUP_ICON，实际 %d 条 %s"
                         % (len(errs), sorted(codes)))
            say("  ✗ 期望 4 条 DUP_ICON，实际 %d 条 %s" % (len(errs), sorted(codes)))
        else:
            say("  ✓ 检出 %d 条 DUP_ICON / inherit（真实双图标），检查器灵敏度正常"
                % len(errs))
    else:
        say("\n[A2] 旧缺陷回归：%s 不存在，跳过（不影响本轮结论）"
            % LEGACY_DEFECT_BACKUP)

    # ---------- B. 坏样本证伪 ----------
    say("\n[B] 坏样本证伪（每个样本只注入一处缺陷，断言被报出）")
    from collections import Counter
    baseline = reps.get("deck-b.pptx")
    baseline_codes = Counter(f.code for f in baseline.findings) if baseline else Counter()
    src = os.path.join(corpus, "deck-b.pptx")
    if not os.path.isfile(src):
        fails.append("[B] 缺少坏样本基底 deck-b.pptx")
        say("  [SKIP] 基底不存在")
    else:
        try:
            samples = build_bad_samples(src, outdir)
        except Exception as e:
            samples = []
            fails.append("[B] 坏样本构造失败：%s" % e)
        for s in samples:
            rep = check(s["path"])
            hits = [f for f in rep.findings if f.code == s["code"]]
            same_level = [f for f in hits if f.level == s["level"]]
            extra = len(hits) - baseline_codes.get(s["code"], 0)
            kw = s.get("expect_msg") or ""
            kw_hits = [f for f in same_level if kw and kw in f.msg]
            ok = extra >= 1 and bool(kw_hits)
            say("  %-24s %-16s 新增 %d 条 %s 级 %s%s  %s"
                % (s["name"], s["code"], extra, s["level"],
                   "，含注入点「%s」" % kw if kw else "",
                   "" if extra >= 1 else "（未新增）",
                   "PASS" if ok else "FAIL"))
            say("      ↳ 注入：%s" % s["why"])
            if not hits:
                fails.append("[B] %s 未报出 %s" % (s["name"], s["code"]))
            elif not same_level:
                got = "、".join(sorted({f.level for f in hits}))
                fails.append("[B] %s 的 %s 级别为 %s，期望 %s"
                             % (s["name"], s["code"], got, s["level"]))
            elif extra < 1:
                fails.append("[B] %s 的 %s 在基线里已存在（%d 条），无法证明"
                             "注入被检出" % (s["name"], s["code"],
                                          baseline_codes.get(s["code"], 0)))
            elif not kw_hits:
                fails.append("[B] %s 报出了 %s 但没有定位到注入点「%s」"
                             % (s["name"], s["code"], kw))
            else:
                say("      ↳ 首条：%s" % kw_hits[0].as_text())

    # ---------- 结论 ----------
    say("\n" + "=" * 72)
    if fails:
        say("SELF-TEST FAILED（%d 项）" % len(fails))
        for f in fails:
            say("  - %s" % f)
        say("=" * 72)
        return finish(1)
    say("SELF-TEST OK — 语料回归 + 坏样本证伪全部通过")
    say("坏样本: %s" % outdir)
    say("报告  : %s" % OUT_DIR)
    say("=" * 72)
    return finish(0)


def _self_test(*args, **kwargs):
    """tests/run_all.py 约定入口（与 edit/preview 的 --self-test 一致）。"""
    return self_test(*args, **kwargs)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="check.py",
        description="PPTX 交付前校验器：结构/单位/继承/布局四类检查")
    ap.add_argument("pptx", nargs="?", help="要校验的 .pptx")
    ap.add_argument("--level", default="all",
                    choices=["all"] + list(LEVELS),
                    help="只跑某一类检查（默认 all）")
    ap.add_argument("--json", action="store_true", help="输出结构化 JSON")
    ap.add_argument("-o", "--out", help="把文本报告写到文件")
    ap.add_argument("--self-test", action="store_true",
                    help="语料回归 + 坏样本证伪")
    ap.add_argument("--corpus", default=BASE_CORPUS,
                    help="自测语料目录，缺省读环境变量 MAKEPPT_CORPUS")
    ap.add_argument("--out-dir", default=None,
                    help="自测产物目录（默认 out/selftest/）")
    args = ap.parse_args(argv)

    if args.self_test:
        return self_test(corpus=args.corpus, outdir=args.out_dir)

    if not args.pptx:
        ap.error("需要给出 .pptx 路径（或 --self-test）")
    rep = check(args.pptx, level=args.level)
    text = rep.as_text()
    if args.json:                       # 机器可读：只输出 JSON
        print(json.dumps(rep.to_dict(), ensure_ascii=False, indent=1))
    else:
        print(text)
    if args.out:
        d = os.path.dirname(os.path.abspath(args.out))
        if d:
            os.makedirs(d, exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
        if not args.json:
            print("\n报告已写入: %s" % args.out)
    return 0 if rep.ok else 1


if __name__ == "__main__":
    sys.exit(main())
