# 出处与许可（NOTICE）

本技能原创，以 **MIT** 发布（见 `LICENSE`）。本文件把"用到别人的什么"分成**四种关系**
逐项写清：**包含 · 引用 · 参考 · 借鉴**。

> **先说清楚本文件的性质。**
> 只有一件事可以自动核验：**有没有逐字复制**（§3，方法在 §9）。
> 其余三类是**自述**（disclosure）——我能指认多少写多少，指认不出的就标"不确定"，
> 不用推测填空。早期会话的上下文已被压缩，我无法重建当时逐字读过哪些技能；这一点
> 在 §5 里明确标出，而不是假装能给出完整清单。

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
| **参考** | 开发时查阅过其文档/做法 | 见 §5：本次可确认的是 `zz-github` |
| **借鉴** | 思路层面受启发，改写为自己的实现 | 见 §6，逐项写明对应点 |

**四类都要求标注；但"借鉴"不等于"来源"，"参考"不等于"复制"。** 本文件把这三者的
界限写死，避免两种错误：把别人的功劳说成自己的，或把无关的技能硬认成来源。

---

## 3. 包含 · 引用：均为零（可核验）

逐行比对结果（本机，2026-11-25）：本技能 3,682 行实质文本（≥40 字符）与环境中
其他技能的全部实质行比对，只出现两处巧合：

| 对比对象 | 重合 | 内容 |
|---|---|---|
| `pptx`（© 2025 Anthropic PBC） | 1 行 | `from PIL import Image, ImageDraw, ImageFont` —— 标准 import |
| `claude-api` | 1 行 | `ap.add_argument("--quiet", action="store_true")` —— argparse 套话 |
| 其余全部技能 | 0 | — |

自行复核：

```bash
python scripts/provenance_check.py --against ~/.dsh/skills
```

---

## 4. 环境里存在哪些别人的技能

这台机器上有 37 个其他技能，其中 `pptx`、`docx`、`pdf`、`xlsx`、`theme-factory`、
`brand-guidelines`、`canvas-design` 等由 **Anthropic PBC** 提供（© 2025 Anthropic，
保留所有权利，条款见各自目录的 `LICENSE.txt`）。**它们都不随本仓库分发。**

---

## 5. 参考：我能确认读过的

| 技能 | 读的是什么 | 用在哪 |
|---|---|---|
| `zz-github`（使用者本人创建） | 其 `SKILL.md` 与 `references/push-workflow.md` | 本次把本技能发布到 GitHub 的流程：代理设置、`$LASTEXITCODE` 判成败、推送前的敏感信息扫描 |
| `pptx`、`theme-factory`、`brand-guidelines`、`canvas-design`、`humanizer-zh`、`skill-creator`、`writing-skills`、`verification-before-completion`、`systematic-debugging`、`test-driven-development`、`frontend-design`、`pdf`、`docx`、`xlsx`、`brainstorming`、`writing-plans`、`executing-plans`、`doc-coauthoring`、`internal-comms`、`dispatching-parallel-agents`、`subagent-driven-development` | **不确定**：这些技能的上下文在早期会话中已被压缩，我无法重建当时是否逐字读过 | — |

**对第二行的处理方式**：它们与本技能在概念上确有对应（见 §6），但我**不能**声称
"读过它们"并据此写成来源声明。所以本文件把关系停在"概念对应"，并在这里标出不确定性。
如果你记得当时让我读的是哪几个，告诉我，我把它们从这一行移到上一行、写明具体落点。

---

## 6. 借鉴：概念对应（不等同于来源）

下表只陈述"这两个东西在讲同一类问题"，**不是**"我照它做的"。对应点写出来，是为了让
读者能自己判断相似程度，而不是让我用一句话概括掉。

| 技能 | 概念上的对应点 | 本技能里落到哪 |
|---|---|---|
| `verification-before-completion` | 宣称"完成/通过"之前必须先验证 | 四道质检门 + **检查器必须能在故意做坏的输入上失败**（五个回归脚本） |
| `systematic-debugging` | 先定位再动手，不做猜测式修改 | 审计报告点名 → **只改它点名的**，不回炉重写 |
| `test-driven-development` | 先写会失败的测试 | `palette_regression.py` 等先验证"吵闹版必须被拒"，再谈安静版通过 |
| `pptx` | python-pptx 生成 + LibreOffice 渲染校验 | 同路线，代码全部自写；本技能额外做**逐块几何比对**与 CJK 度量 |
| `pdf` | PDF 取文本/页码/图区 | 用 PyMuPDF 自写 `source_ingest.py`，页码即 PDF 物理页 |
| `theme-factory` | 主题/配色成套管理 | 改为**按亮度求解的色阶**（14 色族 × 3 模式），并加量化配色预算 |
| `brand-guidelines` | 品牌色与排版的硬约束 | 模板模式：从 `.pptx` 读出配色/字体/页码槽，不另造 |
| `canvas-design` | 设计哲学优先于套版 | "形态跟随内容"、互换两块内容的自检问题 |
| `frontend-design` | 追求有意图的视觉，不要模板感 | 明文禁止"模板 cosplay"；版面按内容形态选 |
| `humanizer-zh` | 中文去空话、保留事实与作者声音 | `HELP.md` §10 语言风格三档；规则来自本项目使用者的逐条反馈，未抄其文本 |
| `skill-creator`、`writing-skills` | 技能结构：触发词、渐进披露、可测 | SKILL.md / HELP.md / references 三层；触发词写在 frontmatter |
| `brainstorming` | 创作前先发散再收敛 | 流程 A 阶段（决定：收需求 → 大纲 → 选配色）在动手前完成 |
| `writing-plans`、`executing-plans`、`dispatching-parallel-agents`、`subagent-driven-development` | 多步任务先计划再执行 | 三阶段流水线与"一轮修完"的成本约束 |
| `doc-coauthoring`、`internal-comms` | 协作式写作与沟通文风 | 需求澄清提问清单（HELP §13） |

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
| 没有逐字复制其他技能 | `python scripts/provenance_check.py --against <技能目录>` |
| 依赖及其许可 | §1 表格；`pip show <包名>` 或包元数据 |
| 学校模板未分发 | 仓库里没有 `template.pptx`；其 `README.md` 顶部有状态说明 |
| 仓库无使用者材料 | `publish.py` 的 `EXCLUDE_FILES` 清单与每次打包的敏感信息扫描输出 |
| 参考/借鉴的自述 | **无法自动核验**，属作者披露；如需更完整的来源清单，请作者补充 |
