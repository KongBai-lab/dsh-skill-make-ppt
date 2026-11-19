# make-ppt

纯 Python 直接读写 PPTX（OOXML）的工具包：**照模板造新页**、**原位改造已有文件**、
**交付前查错**。核心模块只用标准库，不依赖 python-pptx / Office / 网络。

## 三件事

| 任务 | 入口 | 说明 |
|---|---|---|
| 造页：照一份模板生成新 PPT | `engine/build.py` | `Template.from_pptx` + `DeckBuilder`，继承模板的母版、版式、主题与字体 |
| 改页：换字、移页、删页、复制页、统一图标、写备注 | `engine/edit.py` | `Editor`，只改局部字节，原子写 |
| 查错：断链、单位、重复元素、越界/溢出 | `engine/check.py` | `check.check`，ERROR / WARN 分级 |
| 合并：把几份素材的页并成一包 | `engine/merge.py` | 跨母版，自动重编号 rels 与内容类型 |
| 预览：没有 Office 时看排版 | `engine/preview.py` | PPTX → PNG 近似渲染 |
| 读取：看现有文件的结构与文字 | `engine/read.py` | `Deck.load` |

判断口诀：原文件要保留、只换内容 → `edit`；多份并一份 → `merge`；
从零/照模板排版 → `build`。三者最后都跑一遍 `check`。

## 目录结构

```
make-ppt/
├── SKILL.md              # 面向 AI agent 的完整说明书
├── README.md
├── LICENSE
├── engine/
│   ├── units.py          # OOXML 单位唯一真源
│   ├── read.py           # pptx → 只读 IR
│   ├── build.py          # 模板生成
│   ├── edit.py           # 原位改造
│   ├── merge.py          # 跨母版合并
│   ├── check.py          # 结构/单位/继承/布局校验
│   └── preview.py        # 近似 PNG 渲染
├── refs/
│   ├── ir.md             # 模块接口契约
│   ├── units.md          # 单位速查与六个坑
│   ├── edit-ops.md       # 改造操作手册
│   └── checklist.md      # 交付前 12 项检查
└── tests/
    └── run_all.py        # 回归测试
```

## 安装

- Python 3.10+。
- 把 `engine/` 放到任意位置即可；`build` / `edit` / `merge` / `check` / `read` /
  `units` 只用标准库。
- `preview.py` 的文字渲染需要可选依赖 `fontTools`（`pip install fonttools`）
  和一个中文 TTF 字体；字体路径可用环境变量 `MAX_PPT_FONT` 指定。
- 把 `engine/` 加入 `sys.path` 后**平铺导入**（`import edit`），不要写
  `from engine import edit`：`engine` 是常见包名，容易与别的项目撞名。

## 最小示例（照模板造页）

```python
import sys
sys.path.insert(0, "/path/to/make-ppt/engine")

import units, build

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

改已有文件与交付前查错：

```python
import edit, check

ed = edit.Editor.open("a.pptx")
ed.set_text(2, contains="旧标题", text="新标题")
ed.set_notes(2, "讲稿…")
ed.save("b.pptx", verify="strict")      # 结构/单位/继承/布局全硬失败

rep = check.check("b.pptx")
print(rep.as_text())                    # ERROR 必须为 0；WARN 逐条看
```

## 测试

```bash
python tests/run_all.py            # 全量
python tests/run_all.py --quick    # 只跑单元自测
```

示例语料不随仓库分发。把 `deck-a.pptx`、`deck-b.pptx`、`deck-merged.pptx`
放进 `tests/corpus/`，或设置环境变量 `MAKEPPT_CORPUS` 指向语料目录（也可指向
其中一份 `.pptx`）。未设置或文件不齐时，测试会打印 `SKIP`，只跑不依赖语料的
`units` 与 `build` 无模板分支，不会报错。

## 限制

- `preview.py` 是近似渲染：不支持渐变、阴影、旋转、艺术字、JPEG 解码、EMF/WMF，
  也不画版式里的占位提示框。它能定位叠字、出界、图片拉伸等问题，最终版面以
  PowerPoint / WPS 打开为准。
- 不支持 SmartArt、动画时间轴、3D、图表对象（`graphicFrame` 只按占位处理）。
- 没有渲染器依赖：不调用 LibreOffice / WPS CLI / PIL / numpy。
- 示例语料是可选的外部文件，仓库不自带 `.pptx`。

## 运行环境（DSHM / DSH）

- **DSHM（鸿蒙端侧）**：DeepSeek Harness 的鸿蒙发行版，数据目录位于应用沙箱内，本文记为 `$DSH_HOME`。把本技能放到 `$DSH_HOME/skills/<name>.md`（单文件）或 `$DSH_HOME/skills/<name>/SKILL.md`（带 `refs/`、`scripts/` 等附属目录），新开会话后生效。
- **DSH 桌面版（macOS / Windows / Linux）**：技能目录按对应版本文档（通常为 `~/.dsh/skills/`），结构与 DSHM 相同；区别只是 `$DSH_HOME` 不在鸿蒙应用沙箱内。
- 两种环境的 `SKILL.md` 完全通用；**不要**把 DSHM 沙箱内的绝对路径写进配置或脚本，需要可移植路径时用环境变量。

## 版权、第三方依赖与商标

- **本项目**：MIT，见 [LICENSE](LICENSE)。示例图、示例数据由本项目程序生成，随 MIT 许可发布。
- **第三方依赖**：默认只用 Python 标准库（PSF License）；`preview.py` 的文字渲染可选依赖 `fontTools`（MIT）。
- **字体不随包分发**：程序读取系统已安装字体。Noto 系列为 SIL OFL 1.1；DejaVu 系列为 Bitstream Vera 许可；Windows / macOS 商业字体需你在本机自行取得授权，仓库不附带任何字体文件。
- **商标与归属**：DSH、DSHM、DeepSeek Harness 是其各自所有者的产品名或商标；本项目是独立的第三方开源工具，与官方无隶属、赞助或背书关系。GitHub 为其所有者商标。
- **代码来源**：本仓库代码为原创实现；如果后续引入第三方代码，会在此列出其来源与许可。

## 许可

MIT，见 [LICENSE](LICENSE)。Copyright (c) 2026 KongBai-lab。
