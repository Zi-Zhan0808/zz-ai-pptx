# 出处与许可（NOTICE）

本技能由原作者原创，以 MIT 许可发布（见 `LICENSE`）。本文件说明它**用到别人的什么**，
以及**没有用到什么**。

---

## 一、运行时依赖（不随仓库分发，由使用者自行安装）

| 组件 | 用途 | 许可 | 注意 |
|---|---|---|---|
| [python-pptx](https://github.com/scanny/python-pptx) | 读写 `.pptx` 的唯一底座 | MIT | 自由使用 |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | 读 PDF 文本、取页码、裁剪图区 | **AGPL-3.0 或商业许可（双许可）** | 见下方「AGPL 提醒」 |
| [Pillow](https://python-pillow.org/) | 图像尺寸、合成、拼版 | MIT-CMU | 自由使用 |
| [matplotlib](https://matplotlib.org/) | 公式与图表渲染 | Matplotlib License（BSD 风格） | 自由使用 |
| [lxml](https://lxml.de/) | OOXML 局部修改 | BSD-3-Clause | 自由使用 |
| [LibreOffice](https://www.libreoffice.org/)（外部程序） | 把 `.pptx` 渲染成 PDF/PNG 做视觉质检 | MPL-2.0 | 仅作为外部命令调用，不链接、不分发 |

**AGPL 提醒（值得认真读一遍）**：PyMuPDF 是 AGPL 或商业双许可。本技能只 import 它、
不复制它的代码，因此**本仓库的 MIT 许可不会传染到 PyMuPDF 自身**；但如果你把本技能
改造成对外提供服务的网络应用，AGPL 对 PyMuPDF 的要求（向用户提供对应源码）可能适用。
拿不准就问律师，别问 README。

---

## 二、环境内先行技能：**没有复制任何代码或文本**

本技能运行在有其他技能的 DSH 环境里。其中 `pptx`、`docx`、`pdf`、`xlsx`、
`theme-factory`、`brand-guidelines`、`canvas-design` 由 **Anthropic PBC** 提供
（© 2025 Anthropic，保留所有权利，许可条款见各自目录的 `LICENSE.txt`）。

**结论：本仓库不包含它们的任何代码、提示词或文本。** 这一点是可核查的，不是声明：

```bash
python scripts/provenance_check.py --against ~/.dsh/skills
```

它把本技能的每个实质行（≥40 字符）与其他技能的每一行逐条比对。在这台机器上的结果是：

| 对比对象 | 重合行数 | 重合内容 |
|---|---|---|
| `pptx` | 1 | `from PIL import Image, ImageDraw, ImageFont`（标准 import） |
| `skill-creator` | 1 | 一行 `# ---------` 分隔线 |
| 其余（docx / pdf / xlsx / theme-factory / brand-guidelines / canvas-design / humanizer-zh / writing-skills 等） | 0 | — |

思路层面确实受过这些技能的影响（例如「用 python-pptx 生成 + 用 LibreOffice 渲染验证」
是这类工作的通行做法），但**方法论借鉴不构成代码复制**，两者的许可义务不同，这里如实区分。

`humanizer-zh`（中文去空话）与本技能 `HELP.md` §10 语言风格的关系：§10 的规则来自
本项目使用者的逐条反馈（记录在 `references/pitfalls.md` 与 HELP §10 的对照表里），
不是从 `humanizer-zh` 抄来的；两者目标相近，实现不同。

---

## 三、模板与图片

| 资产 | 出处 | 是否分发 |
|---|---|---|
| `assets/templates/课堂pre01/template.pptx` | 使用者本人提供的学校汇报模板（已去校徽） | **不分发**（授权原因），公开仓库中标为「待更新」 |
| `assets/templates/课堂pre01/template_profile.json` / `README.md` | 由上述模板读取得出的配色、字体、页码映射与三段提示词 | 分发（不含模板本体） |
| `scripts/_smoke_sheet.png` | 开发期自制的拼版检查图 | 不分发 |
| 演示用论文 PDF、由论文原文截取的图 | 期刊版权，属引用范围 | 不出现在本仓库 |

内置模板槽的设计就是为此：**槽在，内容由使用者自己填**。

---

## 四、字体

只按名称引用系统中文字体（等线 / 微软雅黑 / Noto Sans CJK），不打包字体文件。
使用者的机器上没有对应字体时，`scripts/preflight.py` 会报出来并降级。

---

## 五、如何复核本文档的每一句

| 说的是 | 怎么查 |
|---|---|
| 没有复制其他技能的代码 | `python scripts/provenance_check.py --against <技能目录>` |
| 依赖及其许可 | 上面的表格；用 `pip show` 或包元数据核对 |
| 模板未分发 | 仓库里没有 `template.pptx`；`assets/templates/课堂pre01/README.md` 顶部有状态说明 |
| 仓库无用户材料 | `publish.py` 的 `EXCLUDE_FILES` 清单，以及每次打包的敏感信息扫描输出 |
