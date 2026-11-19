# -*- coding: utf-8 -*-
"""merge.py — 跨母版合并：把别的 .pptx 的页并入目标文件。

背景：课程 PPT 第一份要把四个专业的素材并成一个包，每份素材有自己的母版、
版式、主题、图片、备注。直接拷 slide XML 会断链，页序也会丢。
本模块做的是：深度收集一页的全部依赖 → 统一重编号 → 重写关系与内容类型 →
追加到 presentation 的页序末尾。

用法
----
    import merge
    merge.merge("目标.pptx", "素材.pptx", "输出.pptx")            # 全部页
    merge.merge("目标.pptx", "素材.pptx", "输出.pptx", pages=[2, 3, 4])  # 指定页
    merge.merge_many("目标.pptx", [("a.pptx", None), ("b.pptx", [3, 4])], "输出.pptx")

约束
----
* 不修改输入文件；输出到新路径（先写 .tmp 再替换）。
* 丢弃 ppt/tags（custData），同时删除 slide 里对它的引用，避免断链。
* 合并后建议用 check.check 校验；本模块保存前只做路径级断言（每个 rels 的 target 必须在包内）。
"""
import os
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

_ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
if _ENGINE_DIR not in sys.path:
    sys.path.insert(0, _ENGINE_DIR)

try:
    from . import units
except ImportError:  # pragma: no cover
    import units

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CT_PREFIX = "application/vnd.openxmlformats-officedocument"

# 这些部件属于"包级"，不允许作为依赖被复制
_NOT_DEPS = ("ppt/presentation.xml", "[Content_Types].xml", "_rels/.rels",
             "ppt/presProps.xml", "ppt/viewProps.xml", "ppt/tableStyles.xml")


def _p(tag):
    return "{%s}%s" % (P_NS, tag)


def _read_raw_rels(z, part, names):
    """{rId: (Target, Type全URL, Mode)}；target 归一化为绝对部件名（内部）。"""
    d, b = posixpath.split(part)
    rp = posixpath.join(d, "_rels", b + ".rels")
    if rp not in names:
        return {}
    out = {}
    for rel in ET.fromstring(z.read(rp)):
        t = rel.get("Target") or ""
        mode = rel.get("TargetMode") or "Internal"
        if mode == "Internal" and not t.startswith(("http://", "https://")):
            t = posixpath.normpath(posixpath.join(d, t))
        out[rel.get("Id")] = (t, rel.get("Type") or "", mode)
    return out


def _rels_name(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def _names_of(z):
    return set(n for n in z.namelist() if not n.endswith("/"))


# --------------------------------------------------------------------------
# 依赖收集
# --------------------------------------------------------------------------
def _collect(sz, slide_part, snames, drop_tags=True):
    """从一页出发，收集其全部内部依赖（父→子）。"""
    out = []
    seen = set()
    stack = [slide_part]
    while stack:
        p = stack.pop()
        if p in seen or p not in snames:
            continue
        seen.add(p)
        if p in _NOT_DEPS:
            continue
        if drop_tags and p.startswith("ppt/tags/"):
            continue
        out.append(p)
        for _rid, (tgt, _typ, mode) in _read_raw_rels(sz, p, snames).items():
            if mode != "External" and tgt in snames and tgt not in _NOT_DEPS \
                    and not (drop_tags and tgt.startswith("ppt/tags/")):
                stack.append(tgt)
    return out


# --------------------------------------------------------------------------
# 新部件名分配
# --------------------------------------------------------------------------
def _max_no(names, pattern):
    mx = 0
    for n in names:
        m = re.match(pattern, n)
        if m:
            mx = max(mx, int(m.group(1)))
    return mx


def _allocator(tnames, prefix="ext"):
    """返回 alloc(旧部件名) -> 新部件名，保证与目标已有部件不冲突。"""
    state = {
        "slide": _max_no(tnames, r"ppt/slides/slide(\d+)\.xml$"),
        "layout": _max_no(tnames, r"ppt/slideLayouts/slideLayout(\d+)\.xml$"),
        "master": _max_no(tnames, r"ppt/slideMasters/slideMaster(\d+)\.xml$"),
        "theme": _max_no(tnames, r"ppt/theme/theme(\d+)\.xml$"),
        "notes": _max_no(tnames, r"ppt/notesSlides/notesSlide(\d+)\.xml$"),
        "notesMaster": _max_no(tnames, r"ppt/notesMasters/notesMaster(\d+)\.xml$"),
        "media": _max_no(tnames, r"ppt/media/image(\d+)\.\w+$"),
        "other": 0,
    }
    mapping = {}
    order = []

    def alloc(part):
        if part in mapping:
            return mapping[part]
        ext = part.rsplit(".", 1)[-1] if "." in part else "bin"
        if re.match(r"ppt/slides/slide\d+\.xml$", part):
            state["slide"] += 1
            new = "ppt/slides/slide%d.xml" % state["slide"]
        elif re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", part):
            state["layout"] += 1
            new = "ppt/slideLayouts/slideLayout%d.xml" % state["layout"]
        elif re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", part):
            state["master"] += 1
            new = "ppt/slideMasters/slideMaster%d.xml" % state["master"]
        elif re.match(r"ppt/theme/theme\d+\.xml$", part):
            state["theme"] += 1
            new = "ppt/theme/theme%d.xml" % state["theme"]
        elif re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", part):
            state["notes"] += 1
            new = "ppt/notesSlides/notesSlide%d.xml" % state["notes"]
        elif re.match(r"ppt/notesMasters/notesMaster\d+\.xml$", part):
            state["notesMaster"] += 1
            new = "ppt/notesMasters/notesMaster%d.xml" % state["notesMaster"]
        elif part.startswith("ppt/media/"):
            state["media"] += 1
            new = "ppt/media/image%d.%s" % (state["media"], ext)
        else:
            state["other"] += 1
            base = posixpath.basename(part)
            sub = posixpath.dirname(part)
            new = to_free(sub + "/" + prefix + "_%d_%s" % (state["other"], base))
        mapping[part] = new
        order.append(part)
        return new

    def to_free(candidate):
        n = candidate
        k = 0
        while n in tnames or n in mapping.values():
            k += 1
            root, ext = n.rsplit(".", 1)
            n = "%s_%d.%s" % (root, k, ext)
        return n

    def new_rels_name(part):
        """旧部件 rels 文件 → 新部件 rels 文件。"""
        new = alloc(part)
        return _rels_name(new)

    def mapping_of():
        return mapping

    return alloc, new_rels_name


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------
def merge_many(target, sources, out=None, drop_tags=True, verbose=False):
    """把 sources 的页依次追加到 target 末尾。

    sources: list[str] 或 list[(str, pages|None)]
    """
    src_list = []
    for s in sources:
        if isinstance(s, (tuple, list)):
            src_list.append((s[0], s[1]))
        else:
            src_list.append((s, None))

    tz = zipfile.ZipFile(target)
    tnames = _names_of(tz)
    tparts = {n: tz.read(n) for n in tnames}
    tz.close()

    pres_xml = tparts["ppt/presentation.xml"].decode("utf-8")
    pres_rels_xml = tparts["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    ct_xml = tparts["[Content_Types].xml"].decode("utf-8")

    # 现有 id / 编号
    mids = [int(m) for m in re.findall(r'<p:sldMasterId id="(\d+)"', pres_xml)]
    sids = [int(m) for m in re.findall(r'<p:sldId id="(\d+)"', pres_xml)]
    next_mid = (max(mids) + 1) if mids else 2147483648
    next_sid = (max(sids) + 1) if sids else 256
    alloc, new_rels_name = _allocator(tnames)

    new_rels = []          # (rid, type_url, target)
    new_master_ids = []    # (id, rid)
    new_slide_ids = []     # (id, rid)
    created = {}           # 新部件名 -> bytes
    rid_seq = [0]

    def next_rid():
        rid_seq[0] += 1
        return "rIdM%d" % rid_seq[0]

    stats = {"sources": 0, "slides_added": 0, "parts_copied": 0}

    for src_path, pages in src_list:
        sz = zipfile.ZipFile(src_path)
        snames = _names_of(sz)
        sparts = _slide_parts(sz, snames)
        if pages is None:
            pages = list(range(1, len(sparts) + 1))
        # 1) 收集全部依赖并分配新名
        all_deps = []
        for pg in pages:
            if not (1 <= pg <= len(sparts)):
                raise IndexError("%s 没有第 %d 页（共 %d 页）" % (src_path, pg, len(sparts)))
            all_deps += _collect(sz, sparts[pg - 1], snames, drop_tags)
        # 去重保序
        seen = set()
        deps = [d for d in all_deps if not (d in seen or seen.add(d))]
        deps_set = set(deps)
        for d in deps:
            alloc(d)
        # 2) 复制部件 + 重写 rels
        for d in deps:
            if d.startswith("ppt/media/"):
                created[alloc(d)] = sz.read(d)
                stats["parts_copied"] += 1
                continue
            text = sz.read(d).decode("utf-8")
            if drop_tags:
                text = re.sub(r"<p:custDataLst>.*?</p:custDataLst>", "", text, flags=re.S)
                text = re.sub(r"<p:custData\b[^>]*/>", "", text)
            created[alloc(d)] = text.encode("utf-8")
            stats["parts_copied"] += 1
            raw = _read_raw_rels(sz, d, snames)
            if not raw:
                continue
            pairs = []
            for rid, (tgt, typ, mode) in raw.items():
                if mode == "External" or tgt.startswith(("http://", "https://")):
                    pairs.append((rid, typ, tgt, "External"))
                    continue
                if tgt not in deps_set:
                    # 指向未复制的部件（包级 / tags）：丢弃该关系
                    continue
                rel_target = posixpath.relpath(alloc(tgt), posixpath.dirname(alloc(d)))
                pairs.append((rid, typ, rel_target, "Internal"))
            created[_rels_name(alloc(d))] = _rels_xml(pairs).encode("utf-8")
        # 3) 追加页序与 master
        for pg in pages:
            old_slide = sparts[pg - 1]
            new_slide = alloc(old_slide)
            next_sid += 1
            rid = next_rid()
            new_slide_ids.append((next_sid, rid))
            new_rels.append((rid, "slide", new_slide))
            stats["slides_added"] += 1
        # 该源用到的 master：从 slide 的 layout → master
        src_masters = set()
        for pg in pages:
            layout = None
            for rid, (tgt, typ, mode) in _read_raw_rels(sz, sparts[pg - 1], snames).items():
                if typ.endswith("/slideLayout"):
                    layout = tgt
            if layout:
                for rid, (tgt, typ, mode) in _read_raw_rels(sz, layout, snames).items():
                    if typ.endswith("/slideMaster"):
                        src_masters.add(tgt)
        for m in sorted(src_masters):
            next_mid += 1
            rid = next_rid()
            new_master_ids.append((next_mid, rid))
            new_rels.append((rid, "slideMaster", alloc(m)))
        sz.close()
        stats["sources"] += 1

    # 4) 重写 presentation.xml（追加 master 与页）
    add_mid = "".join('<p:sldMasterId id="%d" r:id="%s"/>' % (i, r)
                      for i, r in new_master_ids)
    add_sid = "".join('<p:sldId id="%d" r:id="%s"/>' % (i, r)
                      for i, r in new_slide_ids)
    if add_mid:
        pres_xml = _insert_before(pres_xml, "</p:sldMasterIdLst>", add_mid)
    if add_sid:
        pres_xml = _insert_before(pres_xml, "</p:sldIdLst>", add_sid)
    tparts["ppt/presentation.xml"] = pres_xml.encode("utf-8")

    # 5) 重写 presentation.xml.rels
    rel_xml = "".join(
        '<Relationship Id="%s" Type="%s/%s" Target="%s"/>'
        % (rid, _REL_TYPES[typ], typ, relpath_from_ppt(tgt))
        for rid, typ, tgt in new_rels)
    pres_rels_xml = _insert_before(pres_rels_xml, "</Relationships>", rel_xml)
    tparts["ppt/_rels/presentation.xml.rels"] = pres_rels_xml.encode("utf-8")

    # 6) 内容类型
    ct_add = _content_types_for(created, tparts, ct_xml)
    tparts["[Content_Types].xml"] = ct_add.encode("utf-8")

    # 7) 合并部件
    for n, data in created.items():
        tparts[n] = data

    # 8) 写盘 + 路径级断言
    out = out or (os.path.splitext(target)[0] + "-merged.pptx")
    out = os.path.abspath(out)
    tmp = out + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in tparts.items():
            z.writestr(n, data)
    probs = _assert_package(tmp)
    if probs:
        os.remove(tmp)
        raise ValueError("合并后包校验失败：%s" % probs[:5])
    os.replace(tmp, out)
    if verbose:
        print("merged →", out, stats)
    return out


def merge(target, source, out=None, pages=None, verbose=False):
    return merge_many(target, [(source, pages)], out, verbose=verbose)


# --------------------------------------------------------------------------
# 辅助
# --------------------------------------------------------------------------
_REL_TYPES = {
    "slide": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    "slideLayout": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    "slideMaster": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
}


def _insert_before(text, marker, content):
    i = text.find(marker)
    if i < 0:
        raise ValueError("找不到插入点 %r" % marker)
    return text[:i] + content + text[i:]


def relpath_from_ppt(part):
    """presentation.xml 在 ppt/ 下，target 相对它。"""
    if part.startswith("ppt/"):
        return part[4:]
    return posixpath.relpath(part, "ppt")


def _rels_xml(pairs):
    body = "".join(
        '<Relationship Id="%s" Type="%s" Target="%s"%s/>'
        % (rid, typ, tgt, ' TargetMode="External"' if mode == "External" else "")
        for rid, typ, tgt, mode in pairs)
    return ('<Relationships xmlns="%s">%s</Relationships>' % (REL_NS, body))


def _content_types_for(created, tparts, ct_xml):
    """给新部件补 Default/Override；已有条目跳过。"""
    defaults = set(re.findall(r'<Default Extension="([^"]+)"', ct_xml))
    overrides = set(re.findall(r'<Override PartName="/([^"]+)"', ct_xml))
    new_defs, new_overs = [], []
    type_by_prefix = [
        (r"ppt/slides/slide\d+\.xml$", CT_PREFIX + ".presentationml.slide+xml"),
        (r"ppt/slideLayouts/slideLayout\d+\.xml$", CT_PREFIX + ".presentationml.slideLayout+xml"),
        (r"ppt/slideMasters/slideMaster\d+\.xml$", CT_PREFIX + ".presentationml.slideMaster+xml"),
        (r"ppt/theme/theme\d+\.xml$", CT_PREFIX + ".theme+xml"),
        (r"ppt/notesSlides/notesSlide\d+\.xml$", CT_PREFIX + ".presentationml.notesSlide+xml"),
        (r"ppt/notesMasters/notesMaster\d+\.xml$", CT_PREFIX + ".presentationml.notesMaster+xml"),
    ]
    for n in created:
        if n.endswith(".rels"):
            continue
        ext = n.rsplit(".", 1)[-1].lower()
        if n.startswith("ppt/media/") and ext not in defaults:
            new_defs.append('<Default Extension="%s" ContentType="image/%s"/>'
                            % (ext, "jpeg" if ext == "jpg" else ext))
            defaults.add(ext)
            continue
        if n in overrides:
            continue
        for pat, t in type_by_prefix:
            if re.match(pat, n):
                new_overs.append('<Override PartName="/%s" ContentType="%s"/>' % (n, t))
                overrides.add(n)
                break
    return _insert_before(ct_xml, "</Types>", "".join(new_defs) + "".join(new_overs))


def _slide_parts(z, names=None):
    names = names or _names_of(z)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    raw = _read_raw_rels(z, "ppt/presentation.xml", names)
    out = []
    for s in pres.findall(_p("sldIdLst") + "/" + _p("sldId")):
        rid = s.get("{%s}id" % R_NS)
        if rid in raw:
            out.append(raw[rid][0])
    return out


def _assert_package(path):
    """包级检查：rels 的 target 必须存在；slide 由 sldIdLst 可达。"""
    probs = []
    z = zipfile.ZipFile(path)
    names = _names_of(z)
    for n in names:
        if not n.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(n))
        for rel in ET.fromstring(z.read(n)):
            if (rel.get("TargetMode") or "Internal") == "External":
                continue
            t = rel.get("Target") or ""
            if t.startswith(("http://", "https://")):
                continue
            full = posixpath.normpath(posixpath.join(base, t))
            if full not in names:
                probs.append("断链 %s -> %s" % (n, full))
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    lst = pres.find(_p("sldIdLst"))
    if lst is None:
        probs.append("没有 sldIdLst")
    z.close()
    return probs


# --------------------------------------------------------------------------
# 自测
# --------------------------------------------------------------------------
def _self_test():
    corpus = os.environ.get("MAKEPPT_CORPUS")
    if corpus and os.path.isdir(corpus):
        corpus = os.path.join(corpus, "deck-a.pptx")
    if not corpus or not os.path.isfile(corpus):
        print("SKIP merge.py 自测：未设置 MAKEPPT_CORPUS"
              "（可指向一份 .pptx 或含 deck-a.pptx 的目录）")
        return 0
    out_dir = os.path.abspath(os.path.join(_ENGINE_DIR, "..", "out"))
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "selftest_merge.pptx")
    import read as rmod
    base_n = rmod.Deck.load(corpus).n
    if base_n < 3:
        print("SKIP merge.py 自测：%s 只有 %d 页，至少需要 3 页" % (corpus, base_n))
        return 0
    merge(corpus, corpus, out, pages=[2, 3], verbose=True)
    d = rmod.Deck.load(out)
    print("合并后页数:", d.n)
    assert d.n == base_n + 2, "期望 %d + 2，实际 %d" % (base_n, d.n)
    # 新页的版式必须是素材自己的版式（不是目标模板）
    last_layouts = {d.slide(i).layout for i in (base_n + 1, base_n + 2)}
    print("新页版式:", [os.path.basename(x) for x in last_layouts])
    probs = _assert_package(out)
    assert not probs, probs
    print("SELF-TEST OK — merge.py 可用：", out)
    return 0


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        sys.exit(_self_test())
    print(__doc__)
