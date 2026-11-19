---
name: make-ppt
description: 生成、修改、合并 PowerPoint：照模板造新 PPT、改已有 PPTX（换文字/移页/删页/统一校徽/写备注/重编号）、几份素材合成一包、交付前结构校验与排版体检、无渲染器时的 PNG 预览；含教学示意图/电路图（原生形状绘制）拓扑自检清单与课程课件场景。纯 Python 直接操作 OOXML，不依赖 python-pptx / Office / 网络。
whenToUse: 用户要给课程/报告/汇报做 PPT，或要改一份已有 .pptx，或说"照这个模板做一份""把这几份 PPT 合起来""PPT 里有叠字/断链/打不开/校徽重复""帮我检查一下这份 PPT"时。只做数据图（柱状/流程/时间轴）用 diagram，不做页面排版。
---

# make-ppt：PPTX 生成 · 原位改造 · 交付前自检

三件事：**造新页**（照模板）、**改旧页**（原位改造）、**交付前查错**（结构 + 排版）。
全部能力由 `engine/` 下的模块提供：核心只依赖标准库，
不依赖 python-pptx / Office / 网络。

## 0. 先选路径

| 任务 | 用什么 | 入口 |
|---|---|---|
| 照一份模板造新 PPT | `build.py` | `Template.from_pptx` + `DeckBuilder` |
| 改已有 PPT：换字、移页、删页、复制页、统一校徽、写备注 | `edit.py` | `Editor` |
| 把几份素材的页合成一个包（跨母版） | `merge.py` | `merge.merge` / `merge_many` |
| 交付前查错（断链、单位、重复元素、溢出） | `check.py` | `check.check` |
| 看见排版（PNG 渲染） | `preview.py` | `preview.render` |
| 读现有文件的结构与文字 | `read.py` | `Deck.load` |
| 只要数据图（柱/线/饼/流程/时间轴） | 不是本 skill | 用 `diagram` skill |

**判断口诀**：原文件要保留、只换内容 → 改造（edit）；要多份并一份 → 合并（merge）；
从零/照模板排版 → 生成（build）。三者都要在最后跑一遍 check。

## 1. 三十秒上手

```python
import sys
E = "/path/to/make-ppt/engine"
sys.path.insert(0, E)
import units, read, edit, check, preview, merge, build
```

**必须平铺导入**（`import edit`），不要写 `from engine import edit`：
`engine` 是常见包名，容易与别的项目撞名。

```python
# 读
d = read.Deck.load("a.pptx")
print(d.summary())                       # 页数、每页形状/文字/图片/备注量

# 改
ed = edit.Editor.open("a.pptx")
ed.set_text(2, contains="旧标题", text="新标题")
ed.move_slide(5, 2)
ed.set_notes(2, "讲稿…")
ed.unify_element()                       # 顺带清掉"页面里重复的版式元素"（双图标）
ed.save("b.pptx", verify="strict")       # 交付用 strict：结构/单位/继承/布局全硬失败

# 查
rep = check.check("b.pptx")
print(rep.as_text())                     # ERROR 必须为 0；WARN 逐条看

# 看
pngs = preview.render("b.pptx", "out/preview", pages=[1, 2], width=1400)

# 交付
# 把 b.pptx 交给用户，或按你的发布方式上传
```

每个模块都有 `--self-test`，换环境后先跑一遍。**写法注意**：若宿主执行器不注入
`__file__`（例如把源码交给 `exec`），engine 模块级的路径定位会报
`NameError: name '__file__' is not defined`；用 `runpy.run_path` 可以正确设置
`__file__`：

```python
# 推荐：runpy 会正确设置 __file__
import runpy, sys
E = "/path/to/make-ppt/engine"
sys.argv = ["check.py", "--self-test"]
runpy.run_path(E + "/check.py", run_name="__main__")
# 其它模块同理：sys.argv = ["units.py", "--self-test"]; runpy.run_path(E + "/units.py", run_name="__main__")
```

## 2. 三条铁律（全部来自 2026-10 的返工）

**铁律一：单位只走 `units.py`。**
`a:rPr/@sz` 是 1/100 pt（7.5pt = `"750"`）；`a:spcPct/@val` 是 1/1000 %
（100% = `"100000"`，写成 `"100"` 就是 0.1 倍行距，行与行直接叠在一起）。
生成用 `units.sz_of_pt / spc_pct_of_ratio`，解析用 `units.pt_of_sz / ratio_of_spc_pct`，
写完一页立刻 `units.guard_xml(xml)`。**任何模块不许自己写第二套常量。**

**铁律二：改完必查，查出必看。**
`check.py` 的 ERROR 清零才算结构可信；`check` 的 WARN 是估算，必须配合
`preview.py` 渲染 + `read_image` 真看一眼。只看数字不看图，等于没看。

**铁律三：原子写 + 先备份。**
`save()` 写 `<dst>.tmp`，校验通过才 `os.replace`；改造前把原文件复制一份
`<name>.bak.pptx`。生成器与修补器分离，不要让"重建"覆盖"手修"。
交付用 `save(dst, verify="strict")`（结构/单位/继承/布局全部硬失败）；
`verify=True` 只对结构/单位硬失败，用于"源文件本身就有内容缺陷"的场景。

## 3. 模块速查

### units.py
```python
units.cm(2), units.pt_emu(12), units.inch(1)      # → EMU
units.sz_of_pt(12)          # "1200"
units.pt_of_sz("750")       # 7.5
units.spc_pct_of_ratio(1.0) # "100000"
units.text_height_emu("一段中文", 框宽EMU, 14, spacing_ratio=1.0)   # 保守 1.5 倍行高
units.guard_xml(xml, "slide3")   # 返回问题列表
units.assert_sz / assert_spc_pct / assert_canvas
```

### read.py
`Deck.load(path)` → `deck.slides[i]`（`Slide.index` 从 1 起，列表下标从 0 起）、
`slide.shapes`（页面自身元素）、`slide.inherited`（**版式/母版上的固定元素**，
查"双图标"就看这里）、`shape.paras/runs`、`slide.notes`。

### edit.py（原位改造）
`set_text` / `replace_text` / `move_slide` / `delete_slide` / `duplicate_slide` /
`set_notes` / `unify_element` / `fit_images` / `renumber_superscripts` /
`save(dst, verify="strict")` / `diff(other)`。详见 `refs/edit-ops.md`。

### build.py（照模板生成）
```python
tpl = build.Template.from_pptx("模板.pptx")
b = build.DeckBuilder(tpl)
n = b.add_slide(layout="内容页")        # 版式名或序号；模板没有名字就用序号
b.text(n, units.cm(2), units.cm(3), units.cm(20), units.cm(2), "标题", pt=28, bold=True)
b.rect(n, units.cm(2), units.cm(6), units.cm(10), units.cm(3), fill="F5F7FA")
b.image(n, "图.png", units.cm(22), units.cm(6), units.cm(9), units.cm(6))  # contain，不拉伸
b.page_numbers(True)                    # 右下角页码（字段，放映按实际页序）
b.set_notes(n, "讲稿…")
b.save("out.pptx")                      # 写入前逐页 guard_xml
```
模板的母版、版式、主题、字体、版式里的图片全部继承；新页形状用绝对坐标，
不走占位符（避免 lstStyle 继承的连锁问题）。

### merge.py（跨母版合并）
```python
merge.merge("目标.pptx", "素材.pptx", "输出.pptx", pages=[3, 4, 5])
merge.merge_many("目标.pptx", [("a.pptx", None), ("b.pptx", [2, 3])], "输出.pptx")
```
深度收集依赖 → 重编号 → 重写 rels 与内容类型 → 追加页序。
丢弃 `ppt/tags`（同时删 slide 里的 `custDataLst`，否则断链）。

### check.py（交付前校验）
```python
rep = check.check("a.pptx")             # level="all"
rep.ok / rep.errors / rep.warnings / rep.stats
print(rep.as_text())
```
四类检查：结构（XML/断链/内容类型/页序）、单位（guard_xml）、继承（版式与页面
重复元素 = 双图标）、排版（越界/溢出/文字相交/图片拉伸）。
ERROR = 打开一定异常或量级一定错；WARN = 估算，必须带"估算值 vs 阈值"数字。

### preview.py（无渲染器时看图）
```python
pngs = preview.render("a.pptx", "out/preview", pages=None, width=1400)
preview.contact_sheet(pngs, "out/overview.png", cols=2)
```
PNG 能解码；JPEG 画等比占位框并标注尺寸（不依赖 PIL，不硬解 JPEG）。
**定位问题是够用的，最终版面以 WPS 打开为准。**

## 4. 标准流程（改一份已有 PPT）

1. **读**：`read.Deck.load`，把 `summary()` 和每页标题列出来，确认页序与目标页；
2. **备份**：`cp a.pptx a.bak.pptx`；
3. **列操作**：把要改的页/元素/新文字写成清单（页号 + 元素名 + 新值），一次改完；
4. **改**：`Editor` 逐个操作，`save(dst, verify="strict")`；
5. **查**：`check.check(dst)`，ERROR 清零；WARN 逐条判断要不要动；
6. **看**：`preview.render` 出图，`read_image` 抽查首页、最密的一页、最满的一页；
7. **交付**：`publish` 到下载目录 + `present`。

全量回归（换环境/改过 engine 后先跑；八项：units / read / build / merge / edit / check / preview / end2end）。
示例语料可选：设置环境变量 `MAKEPPT_CORPUS` 指向语料目录，文件不齐时会打印 SKIP：

```python
import runpy, sys
sys.argv = ["run_all.py"]          # 只跑单元自测传 ["run_all.py", "--quick"]
runpy.run_path("/path/to/make-ppt/tests/run_all.py", run_name="__main__")
```

## 4b. 教学示意图 / 电路图（用原生形状画，可编辑）

教学示意图——尤其**电路图**——不要用 `diagram` 出图凑，也不要用图片：用 `build.py` 的**原生形状**画，用户/同事能在 WPS 里直接拖动修改。

可用图元（2026-11-07 实测）：`b.line(n, x1, y1, x2, y2, color=, width_pt=)` 画任意两点线段（导线，可斜线）；`b.rect(n, x, y, w, h, fill=, line=, geom=)` 的 `geom` 直通 OOXML `prstGeom`——`"rect"` / `"ellipse"` / `"triangle"` / `"rightArrow"` / `"roundRect"` 都可用；`b.text` 写标号。**没有箭头端（`headEnd/tailEnd`）支持**，箭头用 `triangle` 拼，或干脆不画箭头（变阻器用四接线柱画法已够）。

**固定符号库（约定俗成，别自创画法）**：

| 元件 | 画法约定 |
|---|---|
| 电源（电池组） | 长短线成对：**长线为正极、画在上/左**；多节串联重复该对 |
| 开关 | **串在干路**（除非题意就是支路开关）；断开状态画成抬起的斜线 |
| 电流表 | **串联**接入被测支路，圈内写 A；量程标在旁边 |
| 电压表 | **并联**在被测元件两端，圈内写 V |
| 滑动变阻器 | **四个接线柱画全**（上下各两），注清实际接入的是"一上一下"哪两个柱 |

**拓扑自检清单（必须逐条打勾，不是"看一眼"）**：
1. 每条导线的**两端都落在符号端点上**（不悬空、不越过端点）；
2. 不存在**无元件的闭合回路**（= 短路）；电源两极不经负载直接相连更是硬错；
3. **并联母线不得画成闭合矩形**（矩形四边都连就是短路环）；
4. 开关在**干路还是支路**符合题意，且与文字描述一致；
5. 电表接入方式（电流表串联 / 电压表并联）与量程正确、正负接线柱方向对；
6. 文字不压图形/导线，元件标号（R1、E、K）齐全。

**结论以"节点-边表"为准**：先在代码里列 `节点 → 边（元件）` 表，再照表画；检查也照表查。`preview` + `read_image` **只能看排版（叠字/出界/美观），看不出电气错误**——导线连错、短路环、电表接反在图上看不出来（本轮电路图**返工 6 轮**的教训）。改完必须回到节点-边表逐条复核 + 上面 6 条打勾，再交付。

## 4c. 课程课件（16:9 教学 PPT）

- **画布 16:9 = 33.87 × 19.05 cm**（12192000 × 6858000 EMU）。画布由**模板**决定：`build.Template.from_pptx(模板)` 继承模板的 `sldSz`，不要手改画布；选模板前先验证：
  ```python
  d = read.Deck.load("模板.pptx")
  print(d.summary()["canvas_cm"])      # 16:9 应为 [33.87, 19.05]
  ```
- **公式**：课件页不经过 md2docx、没有 OMML 转换，**禁止把 LaTeX 源码（`$...$`、`\frac`、`\alpha`）留在页面上**，会原样显示。用普通文本 + Unicode 数学符号：`α β γ Δ Σ ∫ × ≥ ≈ ± ° ½`。注意 `DeckBuilder.text()` 目前**没有 italic 参数**（2026-11-07 实测），需要变量斜体时用后处理给对应 `a:r` 补 `<a:i val="1"/>`（改 slideN.xml，同 md2docx「后处理」章的原子替换做法）；复杂公式建议先 md2docx 生成 docx 再截图嵌图（保真、可读），备注里写明公式来源。
- **教师备注**：每页 `b.set_notes(n, "讲稿要点…")` 必写（含时间提示、提问点、易错点）；交付前 `read.Deck.load` 抽查备注都在。
- **页数节奏**：**25–30 页 / 2 小时课**（含封面、目录、小结、练习）；每 5–7 页一个"停顿点"（小结/提问页），一页一个知识点，别把一节课塞进两页。
- **交付前**：`check.check` ERROR=0 + `preview.render` 抽看封面/最密页/公式页 + `read_image` 确认公式没显示成 `\frac{}{}`。

## 5. 常见坑

| 现象 | 根因 | 处理 |
|---|---|---|
| 行与行叠在一起 | `spcPct` 写成 100（应 100000） | `units.guard_xml` 会报；全包扫一遍 |
| 自查报 0 问题但 WPS 有病 | 自查脚本自己解析单位错了 | 只用 `units.pt_of_sz`，禁止各写各的 |
| 右下角两个校徽 | 版式里自带一个，页面里又放一个 | 图标只在版式定义一次；`check` 的 inherit 项会报 |
| 文字溢出框外 | 用框高当文字高度 | 用 `units.text_height_emu`（1.5 倍保守） |
| 元素跑出画布 | 正则把 y 拼进 cx | 保存前后各跑一次 `guard_xml` |
| 图片被拉长/压扁 | 框尺寸与像素纵横比不一致 | `check` 的 layout 项报 IMAGE_STRETCH；`edit.fit_images()` 按像素比例修正（支持裁切） |
| 修好的页被重建覆盖 | 生成器与修补器混在一起 | 生成到 tmp → 校验 → 原子替换 |
| 合并后 WPS 打不开 | rels 断链 / 缺 Content_Types | `merge` 保存前有路径断言，仍要跑 `check` |
| 讲稿写进备注但放映看不到 | 没建 notesMaster | `build`/`edit` 的 set_notes 会补 notesMaster |

## 6. 参考文档

- `refs/ir.md`：IR 与模块接口契约（改接口时必须同步改它）
- `refs/units.md`：OOXML 单位速查与六个坑
- `refs/edit-ops.md`：改造操作手册（每个操作改哪些部件）
- `refs/checklist.md`：交付前 12 项检查

## 7. 依赖与限制

- 依赖：Python 3.10+；`build` / `edit` / `merge` / `check` / `read` / `units`
  仅用标准库。
- `preview.py` 的 PNG 文字渲染需要可选依赖 `fontTools`（`pip install fonttools`）
  和一个中文 TTF 字体：默认使用 `/system/fonts/FZHeiT-SC-Regular.ttf`
  （HarmonyOS 等系统自带），其它环境用环境变量 `MAX_PPT_FONT` 指定
  （常规/粗体用 `os.pathsep` 分隔两个路径）。
- 不依赖 LibreOffice / WPS CLI / PIL / numpy / python-pptx，不联网。
  渲染是近似实现，精度以"能定位问题"为准。
- 不支持：SmartArt、动画时间轴、渐变、3D、图表对象（graphicFrame 只按占位处理）。
