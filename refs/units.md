# OOXML 单位速查与六个坑

本文件只讲"写 PPT 时最容易错的那几个数"。代码里一律用 `engine/units.py`，
不要在任何地方重新写一套常量——2026-10 的教训是：两套实现必然分叉，
一套错了另一套还报"没问题"。

## 一、单位表

| 属性 | 单位 | 换算 | 例子 |
|---|---|---|---|
| `a:off/@x`、`a:ext/@cx` | EMU | 1 cm = 360000；1 in = 914400；1 pt = 12700 | `12192000` = 33.87 cm（16:9 宽） |
| `a:rPr/@sz` | 1/100 pt | 12 pt = `"1200"`；7.5 pt = `"750"` | |
| `a:spcPct/@val` | 1/1000 % | 100% = `"100000"`；0.1 倍 = `"100"` | |
| `a:spcPts/@val` | 1/100 pt | 20 pt = `"2000"` | |
| `a:ln/@w` | EMU | 1 pt = `"12700"` | |
| `a:bodyPr/@lIns` 等 | EMU | 0.25 cm = 90000 | |
| `p:sldSz/@cx` | EMU | 16:9 = 12192000 × 6858000；4:3 = 9144000 × 6858000 | |

## 二、六个坑（都真实踩过）

### 坑 1：行距写成倍数（最致命）

```xml
<!-- 错：0.1 倍行距，行与行直接压在一起 -->
<a:lnSpc><a:spcPct val="100"/></a:lnSpc>
<!-- 对：100% -->
<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
```

2026-10 的课程 PPT 里，生成器把 7 页 96 处行距写成 `val="100"`，
WPS 打开后满屏叠字，前后四轮被用户指出"字重合"，调坐标、加高框全都没用。
生成时用 `units.spc_pct_of_ratio(1.0)`，写完扫 `units.guard_xml`。

### 坑 2：字号除以 100 两次

```python
# 错：sz="750" 已经是 7.5pt，再除 100 变成 0.075
pt = int(sz) / 100 / 100
# 对
pt = units.pt_of_sz(sz)      # → 7.5
```

自查脚本用错这个，连续三轮报"0 处文字溢出"，但 WPS 里字一直挤在一起。

### 坑 3：`spcBef` / `spcAft` 里的 `spcPct val="0"` 是正常的

```xml
<a:spcBef><a:spcPct val="0"/></a:spcBef>   <!-- 段前距 0，合法 -->
<a:lnSpc><a:spcPct val="100000"/></a:lnSpc> <!-- 行距 100%，必须在这个范围 -->
```

`guard_xml` 按块区分：只检查 `lnSpc` 里的 `spcPct` 量级（10%—400%），
`spcBef/spcAft` 只查负值与超上限。第一版扫描没区分，把正常的段距 0 当错误报了。

### 坑 4：行高不能按"字号 × 行数"算

微软雅黑的实际行高约为字号的 1.32—1.42 倍（WPS/Office 略有差异），
再加段前段后。判断"会不会溢出"时用 `units.text_height_emu()`，
它固定取 **1.5 倍**（比实际更保守）。各模块不许自己乘 1.35 / 1.42 / 1.55 ——
系数一多，同一页在不同模块里的结论就会互相矛盾。

### 坑 5：坐标进了 cx

用正则改 x/y 时，把 `y="123"` 的值拼进了 `cx="..."`，元素宽度变成画布两倍。
保存前后各跑一次 `units.guard_xml`，它会断言 `|cx| ≤ 2×画布宽`。

### 坑 6：图片尺寸别当 EMU

图片像素（如 763×160）不是 EMU。如果直接把像素塞进 `a:ext cx`，
图片会缩成针尖大小。按 96 dpi 换算：`emu = px × 914400 / 96`；
或按框 `contain`：保持像素纵横比缩放，别两边都硬填。

## 三、快速自检

```python
import units
print(units._self_test())          # 单元自测（含坏样本）
probs = units.guard_xml(xml, "slide3")
if probs:
    print("\n".join(probs))
```

整包扫描：

```python
import zipfile, units
z = zipfile.ZipFile("a.pptx")
for n in z.namelist():
    if n.startswith(("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/")) \
            and n.endswith(".xml"):
        for p in units.guard_xml(z.read(n).decode("utf-8"), n):
            print(p)
```

`check.py` 已经把这个扫描做成固定检查项（units 类别，ERROR 级）。
