# -*- coding: utf-8 -*-
"""units.py — OOXML 单位与量级的唯一真源（make-ppt engine）。

为什么必须有这个文件
--------------------
2026-10 的课程 PPT 任务里，所有低级错误都出在单位上：

* `a:spcPct/@val` 写成 100（应为 100000），七页 96 处文字行距变成 0.1 倍，
  行与行直接压在一起，连续三轮被用户指出"字重合"；
* 自查脚本把 `sz="750"`（7.5 pt）当成 750 pt，连续三轮报"0 处溢出"；
* 正则替换把 y 值拼进 cx，元素跑出画布。

此后规定：生成与检查共用本模块，任何模块不得自行解析 sz / spcPct / EMU。

单位速查（DrawingML / PresentationML）
--------------------------------------
| 属性                        | 单位                        | 例                          |
|-----------------------------|-----------------------------|-----------------------------|
| `a:off/@x`、`a:ext/@cx`     | EMU（1 cm=360000，1 in=914400） | 12192000 = 33.87 cm      |
| `a:rPr/@sz`                 | 1/100 pt                    | 12 pt = "1200"              |
| `a:spcPct/@val`             | 1/1000 %                    | 100% = "100000"             |
| `a:spcPts/@val`             | 1/100 pt（固定行距）        | 20 pt = "2000"              |
| `a:ln/@w`                   | EMU                         | 1 pt = 12700                |
| `a:bodyPr/@lIns` 等内边距   | EMU                         | 0.25 cm = 90000             |

用法
----
    from units import cm, pt_emu, sz_of_pt, pt_of_sz, spc_pct_of_ratio, guard_xml

    sz_of_pt(12)                 # "1200"
    pt_of_sz("1200")             # 12.0
    spc_pct_of_ratio(1.0)        # "100000"
    guard_xml(slide_xml)         # 返回问题列表（空列表 = 通过）

自测：`python3 units.py --self-test`
"""
import re

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------
EMU_PER_CM = 360000
EMU_PER_MM = 36000
EMU_PER_INCH = 914400
EMU_PER_PT = 12700

# 16:9 默认画布（PowerPoint 2013+ 宽屏 13.333in x 7.5in）
SLIDE_W_EMU = 12192000
SLIDE_H_EMU = 6858000
# 4:3 画布（老模板，跨母版合并时会遇到）
SLIDE_W_43 = 9144000
SLIDE_H_43 = 6858000

# 合理量级范围（用于断言，不是排版建议）
SZ_MIN, SZ_MAX = 100, 40000          # 1 pt — 400 pt
SPC_PCT_MIN, SPC_PCT_MAX = 10000, 400000   # 10 % — 400 %
SPC_PTS_MAX = 200000                 # 2000 pt，够任何标题用
COORD_LIMIT = 2 * SLIDE_W_EMU        # 坐标/尺寸不超过两倍画布

# 中文排版估算：单倍行高系数（保守取 1.5，WPS 实测 1.32—1.42）
LINE_HEIGHT = 1.5

# 文字宽度估算：以 em 为单位（微软雅黑 / 思源黑体近似）
EM_CJK = 1.0        # 汉字、全角标点
EM_ASCII = 0.52     # 小写字母、数字
EM_ASCII_UPPER = 0.62
EM_SPACE = 0.28
EM_NARROW = 0.35    # 逗号句号等半角标点


class UnitError(ValueError):
    """单位量级不合法（生成前拦截用）。"""


# --------------------------------------------------------------------------
# 坐标换算
# --------------------------------------------------------------------------
def cm(v):
    """厘米 → EMU。"""
    return int(round(v * EMU_PER_CM))


def mm(v):
    """毫米 → EMU。"""
    return int(round(v * EMU_PER_MM))


def inch(v):
    """英寸 → EMU。"""
    return int(round(v * EMU_PER_INCH))


def pt_emu(v):
    """磅 → EMU。"""
    return int(round(v * EMU_PER_PT))


def emu_cm(v):
    return v / float(EMU_PER_CM)


def emu_mm(v):
    return v / float(EMU_PER_MM)


def emu_inch(v):
    return v / float(EMU_PER_INCH)


def emu_pt(v):
    return v / float(EMU_PER_PT)


# --------------------------------------------------------------------------
# 字号：唯一解析 / 生成入口
# --------------------------------------------------------------------------
def sz_of_pt(pt_value, default=18.0):
    """磅值 → `a:rPr/@sz`（1/100 pt）。12 pt → "1200"。"""
    v = default if pt_value is None else float(pt_value)
    return str(int(round(v * 100)))


def pt_of_sz(sz_value, default=18.0):
    """`a:rPr/@sz` → 磅值。 "1200" → 12.0。

    注意：`sz` 是 1/100 pt，不是 pt；也不是 1/10 pt。
    2026-10 的自查脚本在这里多除了一次 100，导致连续三轮误报"无溢出"。
    """
    if sz_value is None or sz_value == "":
        return float(default)
    return float(sz_value) / 100.0


def sz_of_pt_safe(pt_value):
    """带范围断言的版本，生成器写 XML 前调用。"""
    s = sz_of_pt(pt_value)
    assert_sz(s)
    return s


# --------------------------------------------------------------------------
# 行距 / 段间距
# --------------------------------------------------------------------------
def spc_pct_of_ratio(ratio):
    """倍数 → `a:spcPct/@val`（1/1000 %）。1.0（100%）→ "100000"。

    这是本轮事故的直接原因：写成 "100" 表示 0.1 倍行距。
    """
    return str(int(round(float(ratio) * 100000)))


def ratio_of_spc_pct(val, default=1.0):
    """`a:spcPct/@val` → 倍数。 "100000" → 1.0。"""
    if val is None or val == "":
        return float(default)
    return float(val) / 100000.0


def spc_pts_of_pt(pt_value):
    """固定行距（磅）→ `a:spcPts/@val`（1/100 pt）。20 → "2000"。"""
    return str(int(round(float(pt_value) * 100)))


def pt_of_spc_pts(val, default=0.0):
    if val is None or val == "":
        return float(default)
    return float(val) / 100.0


# --------------------------------------------------------------------------
# 文本尺寸估算（用于"会不会溢出"的判断，不用于生成）
# --------------------------------------------------------------------------
def char_em(ch):
    """单字符宽度（em）。"""
    o = ord(ch)
    if ch in " \t":
        return EM_SPACE
    if o < 128:
        if ch.isupper() or ch.isdigit():
            return EM_ASCII_UPPER if ch.isupper() else EM_ASCII
        if ch in ",.;:'\"!|":
            return EM_NARROW
        return EM_ASCII
    # 中日韩统一表意文字、全角标点、全角字母数字
    if 0x4E00 <= o <= 0x9FFF or 0x3000 <= o <= 0x303F or 0xF900 <= o <= 0xFAFF:
        return EM_CJK
    if 0xFF00 <= o <= 0xFF60 or 0xFFE0 <= o <= 0xFFE6:
        return EM_CJK
    if 0x3040 <= o <= 0x30FF:
        return EM_CJK
    return EM_CJK


def text_width_emu(text, pt_value):
    """估算一行文字宽度（EMU）。"""
    if not text:
        return 0
    return int(round(sum(char_em(c) for c in text) * float(pt_value) * EMU_PER_PT))


def wrap_lines(text, box_w_emu, pt_value):
    """按框宽切行（估算，不处理英文断词）；返回行列表。"""
    if not text:
        return []
    lines, cur, cur_w = [], "", 0.0
    limit = float(box_w_emu)
    for ch in text:
        w = char_em(ch) * float(pt_value) * EMU_PER_PT
        if ch == "\n":
            lines.append(cur)
            cur, cur_w = "", 0.0
            continue
        if cur_w + w > limit and cur:
            lines.append(cur)
            cur, cur_w = ch, w
        else:
            cur += ch
            cur_w += w
    if cur:
        lines.append(cur)
    return lines


def text_height_emu(text, box_w_emu, pt_value, spacing_ratio=1.0,
                    line_height=LINE_HEIGHT):
    """估算文本所需高度（EMU）。

    line_height 默认 1.5：比 WPS 微软雅黑实测行高（1.32—1.42）更保守。
    任何"是否溢出"的判断都应走这里，禁止各模块自带系数。
    """
    n = len(wrap_lines(text, box_w_emu, pt_value))
    if n == 0:
        return 0
    per_line = float(pt_value) * EMU_PER_PT * line_height * (
        spacing_ratio if spacing_ratio else 1.0)
    return int(round(per_line * n))


def fits(text, box_w_emu, box_h_emu, pt_value, spacing_ratio=1.0,
         line_height=LINE_HEIGHT):
    """文本是否能装进给定框（保守判断）。"""
    return text_height_emu(text, box_w_emu, pt_value, spacing_ratio,
                           line_height) <= box_h_emu


# --------------------------------------------------------------------------
# 断言：生成前拦截非法量级
# --------------------------------------------------------------------------
def assert_sz(sz_value, where=""):
    v = int(float(sz_value))
    if not (SZ_MIN <= v <= SZ_MAX):
        raise UnitError("字号 sz=%s 不在 %d—%d（1/100 pt）之间%s"
                        % (sz_value, SZ_MIN, SZ_MAX, ("：" + where) if where else ""))
    return v


def assert_spc_pct(val, where=""):
    v = int(float(val))
    if not (SPC_PCT_MIN <= v <= SPC_PCT_MAX):
        msg = ("行距 spcPct=%s 不在 %d—%d（1/1000 %%）之间%s"
               % (val, SPC_PCT_MIN, SPC_PCT_MAX, ("：" + where) if where else ""))
        if abs(v) <= 1000:
            msg += "；这多半是把倍数直接写进来了（100%% 应写 100000）"
        raise UnitError(msg)
    return v


def assert_spc_pts(val, where=""):
    v = int(float(val))
    if not (0 <= v <= SPC_PTS_MAX):
        raise UnitError("段距 spcPts=%s 超范围（1/100 pt）%s"
                        % (val, ("：" + where) if where else ""))
    return v


def assert_emu(v, name="EMU", limit=COORD_LIMIT):
    v = int(v)
    if abs(v) > limit:
        raise UnitError("%s=%s 超范围（|v| 应 ≤ %d）" % (name, v, limit))
    return v


def assert_canvas(w_emu, h_emu, where=""):
    """拒绝 0 尺寸与倒置画布。"""
    if w_emu <= 0 or h_emu <= 0:
        raise UnitError("画布尺寸非法：%s x %s%s" % (
            w_emu, h_emu, ("：" + where) if where else ""))
    if w_emu > COORD_LIMIT or h_emu > COORD_LIMIT:
        raise UnitError("画布尺寸超范围：%s x %s" % (w_emu, h_emu))
    return w_emu, h_emu


# --------------------------------------------------------------------------
# 产物扫描：生成后/检查前扫描 XML 里的危险数值
# --------------------------------------------------------------------------
_RE_SZ = re.compile(r'<a:rPr[^>]*\bsz="(-?\d+)"')
_RE_SZ_END = re.compile(r'<a:endParaRPr[^>]*\bsz="(-?\d+)"')
# 行距与段距分开：spcPct 在 lnSpc 里是行距（0 异常），在 spcBef/spcAft 里是段距（0 正常）
_RE_BLOCK = {
    "lnSpc": re.compile(r"<a:lnSpc\b[^>]*>(.*?)</a:lnSpc>", re.S),
    "spcBef": re.compile(r"<a:spcBef\b[^>]*>(.*?)</a:spcBef>", re.S),
    "spcAft": re.compile(r"<a:spcAft\b[^>]*>(.*?)</a:spcAft>", re.S),
}
_RE_PCT = re.compile(r'<a:spcPct\b[^>]*\bval="(-?\d+)"')
_RE_PTS = re.compile(r'<a:spcPts\b[^>]*\bval="(-?\d+)"')
_RE_OFF = re.compile(r'<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"')
_RE_EXT = re.compile(r'<a:ext\b[^>]*\bcx="(-?\d+)"[^>]*\bcy="(-?\d+)"')
_RE_LNW = re.compile(r'<a:ln\b[^>]*\bw="(-?\d+)"')


def guard_xml(xml, where="", limit=COORD_LIMIT):
    """扫描一段 slide/layout XML，返回问题列表（中文描述）。

    检查五类：字号量级、行距量级、段距量级、坐标/尺寸范围、线宽。
    生成器写完一页后调用；check.py 对整份 deck 的每个部件调用。

    注意：`spcPct` 出现在两处，含义不同——
      * `<a:lnSpc><a:spcPct val="100000"/>` 是行距，100000 = 100%；
      * `<a:spcBef><a:spcPct val="0"/>` 是段前距，0 合法（无段距）。
    2026-10 的第一次扫描把后者也当行距报错，属于误报，已按块区分。
    """
    problems = []

    def add(msg):
        problems.append(("[%s] " % where if where else "") + msg)

    for m in _RE_SZ.finditer(xml):
        v = int(m.group(1))
        if not (SZ_MIN <= v <= SZ_MAX):
            add("字号 sz=\"%s\"（=%.2f pt）不在 %d—%d 之间"
                % (m.group(1), v / 100.0, SZ_MIN, SZ_MAX))
    for m in _RE_SZ_END.finditer(xml):
        v = int(m.group(1))
        if not (SZ_MIN <= v <= SZ_MAX):
            add("段落末字号 sz=\"%s\"（=%.2f pt）不合量级"
                % (m.group(1), v / 100.0))

    # 行距：只在 lnSpc 块内检查
    for m in _RE_BLOCK["lnSpc"].finditer(xml):
        body = m.group(1)
        for pm in _RE_PCT.finditer(body):
            v = int(pm.group(1))
            if not (SPC_PCT_MIN <= v <= SPC_PCT_MAX):
                hint = ("；100% 行距应写 100000，写 100 就是 0.1 倍"
                        if 0 < abs(v) < 10000 else "")
                add("行距 spcPct val=\"%s\"（=%.3f 倍）不在 %d—%d 之间%s"
                    % (pm.group(1), v / 100000.0, SPC_PCT_MIN, SPC_PCT_MAX, hint))
        for pm in _RE_PTS.finditer(body):
            v = int(pm.group(1))
            if v <= 0 or v > SPC_PTS_MAX:
                add("固定行距 spcPts val=\"%s\" 不合量级" % pm.group(1))

    # 段距：spcBef / spcAft，0 合法
    for key in ("spcBef", "spcAft"):
        for m in _RE_BLOCK[key].finditer(xml):
            body = m.group(1)
            for pm in _RE_PCT.finditer(body):
                v = int(pm.group(1))
                if v < 0 or v > SPC_PCT_MAX:
                    add("%s spcPct val=\"%s\" 超范围" % (key, pm.group(1)))
            for pm in _RE_PTS.finditer(body):
                v = int(pm.group(1))
                if v < 0 or v > SPC_PTS_MAX:
                    add("%s spcPts val=\"%s\" 超范围" % (key, pm.group(1)))

    for m in _RE_OFF.finditer(xml):
        x, y = int(m.group(1)), int(m.group(2))
        if abs(x) > limit or abs(y) > limit:
            add("坐标越界 x=%d y=%d" % (x, y))
    for m in _RE_EXT.finditer(xml):
        cx, cy = int(m.group(1)), int(m.group(2))
        if cx < 0 or cy < 0 or cx > limit or cy > limit:
            add("尺寸越界 cx=%d cy=%d" % (cx, cy))
    for m in _RE_LNW.finditer(xml):
        v = int(m.group(1))
        if v < 0 or v > EMU_PER_PT * 100:   # 100 pt 以内的线宽
            add("线宽 w=\"%s\" 超范围" % m.group(1))
    return problems


# --------------------------------------------------------------------------
# 自测
# --------------------------------------------------------------------------
def _self_test():
    errs = []

    def eq(got, want, label):
        if got != want:
            errs.append("%s: got %r, want %r" % (label, got, want))

    eq(cm(1), 360000, "cm")
    eq(pt_emu(12), 152400, "pt_emu")
    eq(sz_of_pt(12), "1200", "sz_of_pt")
    eq(pt_of_sz("1200"), 12.0, "pt_of_sz")
    eq(pt_of_sz("750"), 7.5, "pt_of_sz 7.5pt")
    eq(spc_pct_of_ratio(1.0), "100000", "spc_pct 100%")
    eq(ratio_of_spc_pct("100000"), 1.0, "ratio_of_spc_pct")
    eq(spc_pts_of_pt(20), "2000", "spc_pts_of_pt")
    eq(pt_of_spc_pts("2000"), 20.0, "pt_of_spc_pts")

    # 宽度估算：4 个汉字 12pt ≈ 4.8em = 4.8*12pt
    w = text_width_emu("科学技术", 12)
    eq(w, int(round(4 * 12 * EMU_PER_PT)), "text_width 汉字")
    # 溢出判断
    h = text_height_emu("科技", cm(4), 12)
    eq(h, int(round(12 * EMU_PER_PT * LINE_HEIGHT * 1)), "text_height 单行")

    # 断言：正确量级不抛
    assert_sz("1200", "t")
    assert_sz("750", "t")          # 7.5 pt 是合法小字，不能误杀
    assert_spc_pct("100000", "t")
    # 错误量级必须抛
    for fn, arg in ((assert_sz, "50"), (assert_sz, "50000"),
                    (assert_spc_pct, "100"), (assert_spc_pct, "1000"),
                    (assert_emu, "99999999")):
        try:
            fn(arg)
            errs.append("%s(%r) 应抛 UnitError 却没有" % (fn.__name__, arg))
        except UnitError:
            pass

    # guard_xml：0.1 倍行距必须被扫出
    bad_xml = '<a:p><a:pPr><a:lnSpc><a:spcPct val="100"/></a:lnSpc></a:pPr></a:p>'
    if not guard_xml(bad_xml, "t"):
        errs.append("guard_xml 未检出 spcPct=100")
    # 行距 0 也异常
    zero_ln = '<a:pPr><a:lnSpc><a:spcPct val="0"/></a:lnSpc></a:pPr>'
    if not guard_xml(zero_ln, "t"):
        errs.append("guard_xml 未检出 lnSpc spcPct=0")
    # 段前距 0 是正常写法，必须不报
    bef_xml = '<a:pPr><a:spcBef><a:spcPct val="0"/></a:spcBef></a:pPr>'
    if guard_xml(bef_xml, "t"):
        errs.append("guard_xml 把段前距 0 误报为行距错误：%r" % guard_xml(bef_xml, "t"))
    good_xml = '<a:rPr sz="1200"/><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>'
    if guard_xml(good_xml, "t"):
        errs.append("guard_xml 对正常 XML 误报：%r" % guard_xml(good_xml, "t"))

    if errs:
        print("SELF-TEST FAILED (%d)" % len(errs))
        for e in errs:
            print("  -", e)
        return 1
    print("SELF-TEST OK — units.py 全部通过")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_self_test())
