# 原位改造操作手册

改造 = 不改母版体系，只动内容与页序。改错一次整份打不开，所以每个操作都要知道
"改了哪些部件"。本文按操作列清楚。

## 一、先定位，再动刀

三种定位方式，优先用文本（模板里形状名常是"文本框 7"这种无意义的名字）：

| 方式 | 写法 | 适用 |
|---|---|---|
| 按文本包含 | `contains="旧标题"` | 最稳，人看得懂 |
| 按形状名 | `name="标题"` | 形状名被规范命名过时 |
| 按页内序号 | `idx=3` | 批量脚本内部用，人不好核对 |

改之前先 `read.Deck.load` 打出该页的 `(idx, name, text)` 清单核对一遍：
页码错一页，后面全是白工。

## 二、换文字

**必须保留格式。** 做法是取该段落第一个 run 的 `a:rPr` 作为母版，
把新文字塞进同一个 `a:t`，再更新段落文本。

两种情况要特殊处理：

1. **一个词被拆成多个 run**（PowerPoint 会按编辑历史切分，很常见）。
   直接替换第一个 run 会留下残片。正确做法：把该段所有 run 的文本拼起来匹配，
   命中后把整段重建为单 run（保留首 run 的格式），并记录"该段原本有 N 个 run"，
   便于复核。
2. **文字来自版式占位符**：改的是页面里的占位符副本，不影响版式；
   如果要改的是版式上的固定文字，必须改 layout 部件（会影响所有用它的页），
   这属于"统一元素"，不要当普通改字做。

## 三、移页 / 删页 / 复制页

页序的唯一真相是 `ppt/presentation.xml` 里的 `<p:sldIdLst>`，
**不是** `ppt/slides/slideN.xml` 的文件名顺序。操作必须同步四处：

| 部件 | 改什么 |
|---|---|
| `ppt/presentation.xml` | `sldIdLst` 里的顺序或增减 |
| `ppt/_rels/presentation.xml.rels` | `rId → slides/slideN.xml` 的对应关系 |
| `[Content_Types].xml` | 新页的 Override（复制页时） |
| `ppt/slides/_rels/slideN.xml.rels` | 新页的版式、图片、备注关系 |

复制页的细节：

- 新 slide 部件名顺延（`slide{k}.xml`），不要覆盖；
- 备注：复制 notesSlide 部件并让新页 rels 指向它；两页共享同一个 notesSlide 是错的
  （改一处另一处跟着变）；
- 图片：rels 可以共享同一 `ppt/media/xxx.png`（只读），不必复制字节；
- 复制后再改文字，不要先改再复制。

删页的细节：**只删 slide 与它的 notesSlide**，不要删 layout / master / media ——
别的页可能还在用。删完把 `sldIdLst`、`presentation.xml.rels`、`Content_Types`
里对应的条目一起去掉（不删也行，但会留下孤儿部件，check 的 structure 项会警告）。

## 四、统一元素（校徽、页脚、页眉）

2026-10 的"双图标"事故：版式里自带一个校徽，页面里又放了一个，位置还不同，
WPS 里就是两个图标叠在右下角。修了三次才对。

**正解：图标只定义一次，放在版式里。**

1. 把图片加到要用的 `slideLayoutN.xml`（以及它的 rels），位置/尺寸给绝对值；
2. 把 55 页页面里的校徽图片**全部删掉**（`p:pic` 元素）；
3. 其它版式里如果有旧图标，也清掉。

**幂等**要求：重复执行不能越删越多或越加越多。做法是每次先扫当前状态，
按"同 media + 矩形重叠 > 50%"判定重复，再决定删哪个。

`edit.Editor.unify_element()` 与 `check` 的 inherit 项覆盖这两件事（写/查）。

## 五、写备注（讲稿）

备注不是"在 slide 里加个文本框"，它是独立部件，最少要四样东西：

1. `ppt/notesMasters/notesMaster1.xml` + 它的 rels（指向 theme2）；
2. 每页 `ppt/notesSlides/notesSlideN.xml`；
3. 每页 `ppt/notesSlides/_rels/notesSlideN.xml.rels`：指向 notesMaster 与对应 slide；
4. 每张 slide 的 rels 加一条 `notesSlide` 关系；
5. `[Content_Types].xml` 加 notesSlide / notesMaster 的 Override。

备注正文用 `<p:sp>` + `p:ph type="body"`，每段一个 `<a:p>`。
讲稿一行一段，别塞进一个超长段落（演讲者视图里会连成一片）。

## 六、重编号角标与参考文献

页内上标（`<a:rPr baseline="30000">`）与参考页的编号必须双向一致。
批量做法：

1. 扫全篇，收集每个上标 run 的文本与所属页；
2. 建立"旧编号 → 新编号"映射（按页序或按文献首次出现顺序）；
3. 先改参考页正文，再改页内上标；
4. 改完再扫一遍，确认没有孤立编号（页内有、参考页没有；或反之）。

## 七、修图片拉伸

2026-10 的 合并稿 p38 有三张图被拉变形（像素 423×555 的竖图塞进 4.12×3.09 cm 的横框，
偏差 42.9%）。判断标准是"像素纵横比 vs 框纵横比"，`check` 的 layout 项会报 IMAGE_STRETCH。

修法用 `fit_images`，不是手工改坐标：

```python
ed = edit.Editor.open("a.pptx")
res = ed.fit_images(38, keep="height")   # 只修第 38 页
# res["details"] -> [{page, name, before_cm, after_cm, dev_before}]
ed.save("b.pptx", verify="strict")
```

- `keep="height"`（默认）：高度不动、按比例调宽。**优先用它**——宽度只会变窄，
  不会向下压住图注或落点句。
- `keep="width"`：宽度不动、按比例调高；图变高可能压到下方内容，先看该页的
  `y + h` 有没有越过下方元素的 `y`。
- 保持框中心不动（`x + (old_w-new_w)/2`），所以图上不会偏。
- 带裁切（`a:srcRect`）的图按**裁切后**的有效比例计算，不会被跳过。
- 幂等：修过的图再跑一次 `fixed=0`。

修完渲染该页看一眼：图应该比原来的框窄，但不再变形。

## 八、保存

- **原子写**：写 `<dst>.tmp`，用 `check` 校验通过后 `os.replace`；
  失败时保留原文件，不要留下半个包。
- **备份**：改造前 `cp a.pptx a.bak.pptx`。生成器（build）与修补器（edit）
  分开，别让重建覆盖手修。
- **先小后大**：拿 2—3 页试一遍，`read` 读回核对页数与文字，再批量跑全篇。

## 九、验证清单（改造后）

```python
import read, check
d = read.Deck.load("out.pptx")
print(d.summary())          # 页数、每页形状/文字/备注量
rep = check.check("out.pptx")
print(rep.as_text())        # ERROR 必须 0
```

再抽 2 页 `preview.render` 看一眼，确认改的位置对、没压到别的东西。
