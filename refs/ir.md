# IR 与模块契约（make-ppt engine）

本文件是 engine 各模块之间的接口约定。**改动接口必须同步改本文件**。
所有模块只依赖标准库；不依赖 python-pptx / lxml / PIL，也不联网。

## 目录

```
make-ppt/
  SKILL.md              # 使用者入口（决策表 + 验收流程）
  refs/                 # 参考文档（units / edit-ops / checklist / 本文件）
  engine/
    units.py            # 单位与量级唯一真源（Lead 维护）
    read.py             # pptx → 只读 IR（Lead 维护）
    edit.py             # 原位改造：改字、移页、删页、复制页、备注、重编号
    check.py            # 校验：结构 / 数值 / 继承 / 布局
    preview.py          # PPTX → PNG 近似渲染
    merge.py            # 跨母版合并（Lead 维护）
    build.py            # 生成内核（由 workspace/ppt-toolkit 迁入，Lead 维护）
  tests/                # 回归测试 + 语料清单
  out/                  # 运行产物
```

## 导入与运行

```python
import sys
ENGINE = "/path/to/make-ppt/engine"
sys.path.insert(0, ENGINE)
import units, read
```

- 用普通 Python 3.10+ 解释器运行。若宿主执行器不注入 `__file__`（例如把源码
  交给 `exec`），改用 `runpy.run_path(..., run_name="__main__")`。
- 同一解释器进程内重复导入：改完模块必须 `importlib.reload(module)`，
  否则读到的还是旧代码（会被误当成 bug 排查）。
- 禁止各模块自己用正则解析 `sz` / `spcPct` / 坐标：一律走 `units`。

## units.py（关键 API）

| 函数 | 说明 |
|---|---|
| `cm(v)` `mm(v)` `inch(v)` `pt_emu(v)` | 长度 → EMU |
| `emu_cm(v)` `emu_pt(v)` | EMU → 长度 |
| `sz_of_pt(12)` → `"1200"` | 磅值 → `a:rPr/@sz`（1/100 pt） |
| `pt_of_sz("1200")` → `12.0` | **唯一**字号解析；传 None 返回 18.0 |
| `spc_pct_of_ratio(1.0)` → `"100000"` | 倍数 → `a:spcPct/@val`（1/1000 %） |
| `ratio_of_spc_pct("100000")` → `1.0` | 行距倍数解析 |
| `spc_pts_of_pt(20)` → `"2000"` | 固定行距 |
| `text_width_emu(text, pt)` | 估算一行宽度（EMU） |
| `wrap_lines(text, box_w, pt)` | 估算换行结果 |
| `text_height_emu(text, box_w, pt, spacing_ratio=1.0)` | 估算文本高度（默认 1.5 倍行高） |
| `fits(text, w, h, pt)` | 是否能装进框 |
| `guard_xml(xml, where="")` | 扫描 XML 数值问题，返回中文问题列表 |
| `assert_sz / assert_spc_pct / assert_emu / assert_canvas` | 生成前断言，抛 `UnitError` |

常量：`EMU_PER_CM=360000`、`SLIDE_W_EMU=12192000`、`SLIDE_H_EMU=6858000`、
`LINE_HEIGHT=1.5`、`SZ_MIN/SZ_MAX`、`SPC_PCT_MIN/SPC_PCT_MAX`。

## read.py（IR）

```python
deck = read.Deck.load(path)
deck.w, deck.h            # 画布 EMU
deck.slides               # list[Slide]，下标从 0 开始，Slide.index 从 1 开始
deck.slide(3)             # 1-based 取页
deck.summary()            # dict：页数 / 每页 shape 数 / 文本量 / 图片数 / 备注量
deck.texts()              # [(页号, 形状名, 文本)]
deck.notes()              # {页号: 备注文本}
deck.media                # {部件名: 字节数}
```

`Slide`：`index`、`part`、`layout`、`master`、`shapes`、`inherited`、`notes`、`bg`、
`text`（页面全部文字）、`pictures`、`text_shapes`、`shapes_named(name)`。

`Shape`：`idx`（页内顺序）、`kind`（`sp`/`pic`/`graphicFrame`/`grpSp`/`cxnSp`）、`name`
（`cNvPr@name`）、`x/y/w/h`（EMU，group 内已换算到页坐标）、`geom`（`prstGeom@prst`）、
`fill`/`line`（颜色 HEX 或 `scheme:xxx`）、`style_fill`/`style_line`（`p:style` 的
`a:fillRef`/`a:lnRef`，很多形状 spPr 里没有填充、颜色全靠主题引用）、`line_w`
（`a:ln/@w`）、`flip_h`/`flip_v`、`anchor`（`bodyPr@anchor`：`t`/`ctr`/`b`，空串按 `t`）、
`insets`（`l/t/r/b` 四个 EMU，缺省 91440/45720/91440/45720）、`paras`、`rid`/`media`
（图片）、`crop`（`a:srcRect` 的 `(l,t,r,b)`，千分之一百分比，无裁切为 `()`）、
`src`（`slide`/`layout`/`master`）、`placeholder`；方法 `text`、`rect()`、`pt()`。

`Slide` 另有 `bg`（纯色）与 `bg_image`（背景图片部件）。

`Shape.paras[]` → `Para`：`runs`、`text`、`align`、`level`、`spc_pct`、`spc_ratio`、
`spc_pts`、`has_bullet`；`Para.runs[]` → `Run`：`text`、`sz`、`pt`、`bold`、`italic`、
`color`、`font`。

`Slide.inherited` 是版式/母版上的**固定元素**（图片、装饰条、有文字的样式框），
用于检查"页面里又放一遍"造成的重复（2026-10 的"双图标"事故就出在这里）。

## 各模块公开接口（集成时按此调用）

```python
# edit.py
ed = edit.Editor.open(src_pptx)              # 保留 zip 全部部件
ed.set_text(slide_no, *, name=None, idx=None, contains=None, text, keep_format=True)
ed.replace_text(slide_no, old, new, all_occurrences=True)   # 返回替换次数
ed.delete_slide(slide_no) / ed.move_slide(src, dst) / ed.duplicate_slide(slide_no)
ed.set_notes(slide_no, text)                 # 新建或更新 notesSlide
ed.unify_element(picture_bytes=None, x=None, y=None, w=None, h=None,
                 target_layouts=None, slides=None, match="overlap")
                                             # 图标只在 layout 定义一次，清页面重复（幂等）
ed.fit_images(slide_no=None, keep="height", tolerance=0.03)
                                             # 按像素比例修图片框，消除拉伸（保持中心）
ed.renumber_superscripts(mapping, slides=None)
ed.save(dst, verify=True)                    # True: 结构/单位硬失败，内容类只提示
                                             # "strict": 全部 ERROR 硬失败（交付用）
                                             # False: 只做包级校验
ed.diff(other_pptx) -> dict

# check.py
rep = check.check(path, level="all")         # level: structure|units|inherit|layout|all
rep.ok                                       # 无 ERROR 即 True
rep.errors / rep.warnings                    # 人类可读字符串列表
rep.findings                                 # Finding(check, level, code, msg, slide)
rep.stats                                    # 每类检查的"检查对象数量"，用于证明没漏检
print(rep.as_text()) / rep.to_dict()

# preview.py
pngs = preview.render(pptx, outdir, pages=None, width=1400,
                      with_layout_images=True, verbose=False)
preview.contact_sheet(pngs, out_png, cols=2, cell_width=None)
```

CLI 约定：每个模块都支持 `python <module>.py <pptx> ...`，输出可读文本；
`--json` 输出结构化结果（check）。自测入口：`--self-test`（语料可选，
见下节）；没有语料时打印 SKIP，不报错。
全量回归：`python tests/run_all.py`（units/read/build/merge/edit/check/preview/end2end 八项）。

## 回归语料（可选）

示例语料不随仓库分发。把 `deck-a.pptx`（20 页，半成品）、`deck-b.pptx`（16 页，成品）、
`deck-merged.pptx`（55 页，跨母版合并）放进 `tests/corpus/`，或设置环境变量
`MAKEPPT_CORPUS` 指向语料目录（也可指向其中一份 .pptx）。文件不齐时自测只跑
不依赖语料的部分。

## 共享不变式

1. **单位**：任何写 XML 的地方都用 `units` 生成字符串；任何读的地方都用 `units` 解析。
2. **页号**：对外（CLI、报告、文档）一律 1-based；IR 的 `slides` 列表 0-based。
3. **原子写**：`save()` 先写 `<dst>.tmp`，校验通过后 `os.replace`；校验失败保留旧文件。
4. **幂等**：改造操作重复执行结果一致（不重复叠加图片、不重复加备注）。
5. **不改源文件**：所有改造输出到新路径，除非显式 `inplace=True`。

