# 出处与许可（NOTICE）

本技能原创，以 **MIT** 发布（见 `LICENSE`）。本文件把"用到别人的什么"分成**四种关系**
逐项写清：**包含 · 引用 · 参考 · 借鉴**。

> **本文件的性质。**
> 只有一件事可自动核验：**有没有逐字复制**（§3）。其余是**作者披露**。
> 披露里凡是能指认来源的，都给出可查的位置；指认不出的标"不确定"，不用推测填空。

---

## 1. 运行依赖（不随仓库分发，使用者自行安装）

| 组件 | 用途 | 许可 | 注意 |
|---|---|---|---|
| [python-pptx](https://github.com/scanny/python-pptx) | 读写 `.pptx` 的唯一底座 | MIT | 自由使用 |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | 读 PDF 文本、取页码、裁图区 | **AGPL-3.0 或商业许可（双许可）** | 见下方提醒 |
| [Pillow](https://python-pillow.org/) | 图像尺寸、合成、拼版 | MIT-CMU | 自由使用 |
| [matplotlib](https://matplotlib.org/) | 公式与图表渲染 | Matplotlib License（BSD 风格） | 自由使用 |
| [lxml](https://lxml.de/) | OOXML 局部修改 | BSD-3-Clause | 自由使用 |
| [LibreOffice](https://www.libreoffice.org/)（外部程序） | 把 `.pptx` 渲染成 PDF/PNG 做视觉质检 | MPL-2.0 | 仅命令行调用，不链接、不分发 |

**AGPL 提醒**：PyMuPDF 是 AGPL 或商业双许可。本技能只 `import` 它、不复制其代码，
所以本仓库的 MIT 不会传染到 PyMuPDF 自身；但若你把它改造成对外提供服务的网络应用，
AGPL 对 PyMuPDF 的要求（向用户提供对应源码）可能适用。拿不准请咨询律师。

---

## 2. 四种关系一览

| 关系 | 含义 | 本技能的情况 |
|---|---|---|
| **包含** | 把别人的代码/资产打包进仓库 | **无**（唯一例外是学校模板，且**未分发**，见 §7） |
| **引用** | 逐字引用他人文字并注明 | **无**。技能文本为原创；演示用论文属引用范围，但演示文件不入仓库 |
| **参考** | 依据他人的文档/规格/做法开发 | **有，且是本技能需求的直接来源**：一份使用者在会话开始时提供的《执行手册》，见 §5.1 |
| **借鉴** | 思路层面受启发，改写为自己的实现 | 有；逐项给出对应点，见 §6 |

**"借鉴"不等于"来源"，"参考"也不等于"复制"。** 本文件把三者界限写死，避免两种错误：
把别人的功劳说成自己的，或把无关的东西硬认成来源。

---

## 3. 包含 · 引用：均为零（可核验）

逐行比对（本机，2026-11-25）：本技能 3,682 行实质文本（≥40 字符）与**本机存在的
37 个其他技能**的全部实质行比对，只剩两处巧合：

| 对比对象 | 重合 | 内容 |
|---|---|---|
| `pptx`（© 2025 Anthropic PBC） | 1 行 | `from PIL import Image, ImageDraw, ImageFont` —— 标准 import |
| `claude-api` | 1 行 | `ap.add_argument("--quiet", action="store_true")` —— argparse 套话 |
| 其余全部 | 0 | — |

自行复核：

```bash
python scripts/provenance_check.py --against ~/.dsh/skills
```

**比对的边界（必须说清）**：本节覆盖的是**此刻装在本机**的技能。§5.2 里那批技能本机并
不存在，因此**不在自动比对范围内**；本技能也从未安装过它们。对那批技能，本文件只做披露。

---

## 4. 本机有哪些别人的技能

本机装有 37 个其他技能，其中 `pptx`、`docx`、`pdf`、`xlsx`、`theme-factory`、
`brand-guidelines`、`canvas-design` 等由 **Anthropic PBC** 提供（© 2025 Anthropic，
保留所有权利，条款见各自目录的 `LICENSE.txt`）。**它们都不随本仓库分发。**

---

## 5. 参考：依据了什么

### 5.1 《学术 PPT 创作 Agent 执行手册（DSH · 原生 PPTX 版）》——本技能需求的直接来源

这是使用者在会话开始时提供的规格文档。本技能不是从零构想，而是**把这份手册落成可执行的
代码与检查**。对应关系逐条列出，便于核对：

| 手册里的要求 | 本技能落在哪 |
|---|---|
| §2 流程 S0 盘点 → S1 文献核验 → S2 大纲 → S3 生成 → S4 公式引用 → S5 去味 → S6 质检 → S7 交付 | 收敛为三阶段 A 决定 / B 成型 / C 检查（S4、S5 并入"写的时候就做"，避免同一份稿子写四遍） |
| S1 门禁：「无『凭记忆』的引用」，来源须有 DOI/PMID/文件名 + 可引用页码 | `source_ingest.py` 建索引；`refs_check.py` 强制 `[n] ↔ 文献表 ↔ 页码` 三级可追溯 |
| S6「正文引用 ↔ 文献表双向核对」 | `refs_check.py` 的双向闭合闸门 |
| S6「逐页渲染 PNG 视觉核对」 | `render_qa.py`（LibreOffice → PDF → PNG，再与声明的形状几何逐块比对） |
| S6 版式体检（溢出、重叠、字号、对比度） | `design_check.py` + `render_qa.py` + `palette_audit.py` |
| S3「图表优先用论文原图；若必须重绘，图内标注『重绘自 [n]』」 | `SKILL.md` 质量规则第 2 条 |
| S4「PPTX 默认 `math_mode=image`，规避 OMML 兼容问题」 | `math_render.py` 输出透明 PNG + `math_manifest.json` 保留 LaTeX 源；README 的 "Formulas are PNGs, not OMML" |
| S5 去味清单（「近年来，随着……」开场、「由于时间有限」、空转大词） | `references/prose.md` + 审计里的 `check_ai_smells` |
| S0「列出参考文件与演示目标（受众、场合、时长）」 | `HELP.md` §4 用途类型 + §13 需求澄清提问清单 |
| S2「每页标题陈述结论或判断，避免『背景介绍』这类中性话题词」 | `SKILL.md` 质量规则第 6 条；审计的标题检查 |
| §1.4 提到的脚本名 `audit_deck.py` / `render_deck.py` / `check_refs.py` | 本技能同名或近名保留：`audit_deck.py` / `render_qa.py` / `refs_check.py`（`sources.md`、`outline.md` 亦沿用） |

**说明**：上表是"要求 → 实现"的对应，不是"抄了它的代码"。手册给的是规格与门禁，
实现（字体度量、几何比对、色阶求解、版面库）是本项目自己写的。

### 5.2 手册 §1 列出的技能清单——本机不存在，无法比对

手册把可用的技能分成生产层 / 引用公式层 / 表达去味层 / 质检层 / 编排层。这 12 个名字
**本机全部不存在**（已全盘搜索 `~/.dsh`、DeepSeek 程序目录与会话缓存），因此：

- 本技能**没有、也无法**复制它们的代码或文本；
- 它们对本技能的影响只可能是**手册层面的规格**（已记在 §5.1）；
- 下表按手册的说法记录它们的定位，供读者理解本技能为何这样设计。

| 名称 | 手册里的定位 | 本机状态 |
|---|---|---|
| `dsh-plugin-zhuoyu-ppt` | 生产层：从零做或继承模板；逐页参考文献与浮动红线、四级字号纪律 | 未安装 |
| `dsh-notebook-studio` | 生产层：多篇 PDF（最多 30）综述，结论系在 `[n] PDF p.x` | 未安装 |
| `dsh-paper-slides` | 生产层：单篇 DOI/arXiv 精讲，按场合与时长规划节奏 | 未安装 |
| `dsh-docflow-workflow` | 引用核查：PubMed/Crossref 逐条比对；GB/T 7714-2015 格式 | 未安装 |
| `dsh-latex2office` | 公式插入：LaTeX → OMML 或高清图片 | 未安装 |
| `dsh-tool-mathfig` | 公式图，降级备用 | 未安装 |
| `dsh-humanizer-zh` | 去味（学术模式），只改"怎么说"不改主张 | 未安装 |
| `dsh-style-mirror` | 备选：从用户语料学习语气 | 未安装 |
| `dsh-ppt-studio` | 质检层，含 `ppt_audit`（版式体检）与 `ppt_beautify` | 未安装 |
| `ppt_audit` | `dsh-ppt-studio` 内的工具，非独立技能 | 不适用 |
| `ppt_beautify` | 同上 | 不适用 |
| `dsh-ppt-master` | 商务编排；手册明确说做学术时**不要叠加**它 | 未安装 |

### 5.3 本机确实读过的

| 技能 | 读的是什么 | 用在哪 |
|---|---|---|
| `zz-github`（使用者本人创建） | `SKILL.md` 与 `references/push-workflow.md` | 发布流程：代理设置、用 `$LASTEXITCODE` 判成败（不靠 stderr 有无红字）、推送前的敏感信息扫描 |

### 5.4 不确定的

早期会话的上下文已被压缩。除 §5.1 与 §5.3 能指认的以外，我无法重建当时是否逐字读过
其他技能。本文件不把"可能读过"写成"读过"。

---

## 6. 借鉴：概念对应（不等同于来源）

下表陈述"这两个东西在处理同一类问题"，**不是**"我照它做的"。给对应点，是为了让读者
自己判断相似程度。

| 来源 | 概念对应点 | 本技能落到哪 |
|---|---|---|
| `verification-before-completion`（本机 Anthropic 技能） | 宣称"完成/通过"前必须验证 | 四道质检门 + **检查器必须能在故意做坏的输入上失败**（五个回归脚本） |
| `test-driven-development` | 先写会失败的测试 | `palette_regression.py` 先证明"吵闹版必须被拒"，再谈安静版通过 |
| `systematic-debugging` | 先定位再动手，不做猜测式修改 | 审计点名 → 只改它点名的，不回炉重写 |
| `pptx` | python-pptx 生成 + LibreOffice 渲染校验 | 同路线、代码自写；额外做逐块几何比对与 CJK 度量 |
| `pdf` | PDF 取文本/页码/图区 | `source_ingest.py`（PyMuPDF），页码即 PDF 物理页 |
| `theme-factory` | 主题/配色成套管理 | 改为按亮度求解的色阶（14 色族 × 3 模式）+ 量化配色预算 |
| `brand-guidelines` | 品牌色与排版的硬约束 | 模板模式：从 `.pptx` 读出配色/字体/页码槽，不另造 |
| `canvas-design` | 设计哲学优先于套版 | "形态跟随内容"，以及"互换两块内容"的自检问题 |
| `frontend-design` | 追求有意图的视觉，不要模板感 | 明文禁止"模板 cosplay" |
| `humanizer-zh` | 中文去空话、保留事实与作者声音 | `HELP.md` §10 语言风格三档；规则来自本项目使用者的逐条反馈，未抄其文本 |
| `skill-creator`、`writing-skills` | 技能结构：触发词、渐进披露 | SKILL.md / HELP.md / references 三层；触发词写在 frontmatter |
| `brainstorming` | 创作前先发散再收敛 | A 阶段在动手前完成（需求 → 大纲 → 配色） |
| `writing-plans`、`executing-plans`、`dispatching-parallel-agents`、`subagent-driven-development` | 多步任务先计划再执行 | 三阶段流水线 + "一轮修完"的成本约束 |
| `doc-coauthoring`、`internal-comms` | 协作写作与沟通文风 | §13 需求澄清提问清单；`HELP.md` 的对话式措辞 |

---

## 7. 模板与图片

| 资产 | 出处 | 是否分发 |
|---|---|---|
| `assets/templates/课堂pre01/template.pptx` | **使用者本人**提供的学校汇报模板（已去校徽） | **不分发**（授权原因），公开仓库标为"待更新" |
| `assets/templates/课堂pre01/template_profile.json` / `README.md` | 由上述模板读出的配色、字体、页码映射与三段提示词 | 分发（不含模板本体） |
| `scripts/_smoke_sheet.png` | 开发期自制拼版检查图 | 不分发 |
| 演示用论文 PDF、论文原图截图 | 期刊版权，属引用范围 | 不出现在本仓库 |

内置模板槽的设计正为此：**槽在，内容由使用者自己填。**

---

## 8. 字体

只按名称引用系统中文字体（等线 / 微软雅黑 / Noto Sans CJK），不打包字体文件。
缺字体时 `scripts/preflight.py` 会报出并降级。

---

## 9. 每一句怎么复核

| 本文件说的是 | 怎么查 |
|---|---|
| 没有逐字复制本机其他技能 | `python scripts/provenance_check.py --against <技能目录>` |
| 那 12 个技能不在本机 | 全盘搜索 `~/.dsh`、DeepSeek 程序目录；或问作者要安装位置后再比对 |
| 依赖及其许可 | §1 表格；`pip show <包名>` |
| 学校模板未分发 | 仓库里没有 `template.pptx`；其 `README.md` 顶部有状态说明 |
| 仓库无使用者材料 | `publish.py` 的 `EXCLUDE_FILES` 清单与每次打包的敏感信息扫描输出 |
| 参考/借鉴的自述 | 属作者披露，无法自动核验；§5.1 的每条对应都可在本仓库文件中找到落点 |
