# zz-ai-pptx

一套给 **DeepSeek Harness (DSH)** 用的技能：把材料做成**可编辑、可核对**的 `.pptx`。

## 安装

对 DSH 说一句：

```
安装 skill：zz-ai-pptx
```

或手动：

```bash
git clone https://github.com/Zi-Zhan0808/zz-ai-pptx.git ~/.dsh/skills/zz-ai-pptx
# Windows: %USERPROFILE%\.dsh\skills\zz-ai-pptx
```

装好之后对 DSH 说「help」，会拿到完整的能力清单。

## 四档能力（档位决定花多少钱）

| 档 | 能做什么 | 明确不做 |
|---|---|---|
| **T1** 基础生成 | 封面 + 标题/要点页；一种字体、一套现成配色；纯文本 | 不选版式、不推配色、不画图、不跑审计 |
| **T2** 固定版式 | T1 + 流程 / 表格 / 图+文 / 对比 四类固定版式；可用你的模板 | 不按内容形态造版式 |
| **T3** 内容驱动 | T2 + 按内容形态选或造版式；四道质检门跑一遍 | 不自己画示意图、不联网 |
| **T4** 自制示意图 | T3 + 用形状画示意图（作用范围、距离、分层、矩阵）并做几何断言 | 只在内容必须画时才画 |

**低档不具备高档的能力**：说 T2 就不会按内容造版式，即使那样更好看。

## 三档自由度（能说多少材料外的话）

| 档 | 含义 |
|---|---|
| **L1** 严格贴合材料 | 材料里没有的数字、结论、文献一律不写 |
| **L2** 自主创作文字（默认） | 可写过渡与判断；数字与文献仍须可核实 |
| **L3** 联网自由发挥 | 可检索补充，但每条外部主张都要有 DOI / PMID |

**动手前会先报预算**：档位、自由度、页数、预计轮次、是否联网、几个省钱选项，
等你点头再开工。

## 仓库内容

| 内容 | 状态 |
|---|---|
| `SKILL.md` · `HELP.md` · `LICENSE` | ✅ 有 |
| `scripts/`（18 个：构建、审计、质检） | ✅ 有 |
| `references/`（authoring / template / prose / workflows / pitfalls） | ✅ 有 |
| `assets/example_outline.md` | ✅ 有 |
| `assets/templates/课堂pre01/` 的 profile 与提示词 | ✅ 有 |
| `assets/templates/课堂pre01/template.pptx`（4.2 MB 学校模板） | ⏳ **待更新**（授权原因不公开） |

内置模板的启用方式：把任意 `.pptx` 放进 `assets/templates/课堂pre01/` 并命名
`template.pptx`，或直接在模板模式（B2）下读你自己的模板。

## 第三方与出处

- 本仓库代码原创，**MIT** 许可
- 依赖：python-pptx（MIT）、**PyMuPDF（AGPL 或商业双许可，请自行确认适用性）**、
  Pillow（MIT-CMU）、matplotlib、lxml；LibreOffice 以外部命令形式调用（MPL-2.0）
- 环境中由 Anthropic 提供的 `pptx` / `docx` / `pdf` / `xlsx` 等技能：
  **未复制任何代码或文本**。可自行复核：
  `python scripts/provenance_check.py --against <技能目录>`
- 学校模板不分发

逐项说明见 [`NOTICE.md`](NOTICE.md)。

---

以下为技能自身的技术说明。

Source-grounded deck construction: reference material in, a native editable
`.pptx` out, with every claim traceable to a page and every slide's typography
verified numerically rather than eyeballed.

## Why this exists

The usual "AI makes slides" pipeline has four failure modes, and this skill
attacks each one mechanically instead of hoping:

| Failure | How it is prevented here |
|---|---|
| Invented citations | `refs_check.py` enforces `[n]` ↔ reference list ↔ locator (`S1 p.4`), including that the page exists in the real PDF |
| Text overflowing its box | `build_deck.py` measures with real font metrics and **refuses to save** text that does not fit |
| Layout looks broken | `render_qa.py` renders through LibreOffice and compares rendered text-block bounding boxes against declared shape geometry |
| Loud-but-flat colour | `palette_audit.py` counts hue families, saturation, and whether any element leads the page |
| Template-speak prose | `audit_deck.py` flags 铺陈/空转大词 patterns; `references/prose.md` explains what to cut vs. what must survive |

## Look like what?

Two modes, decided by what the user hands over:

- **A `.pptx` is given** → template mode. The file's own theme, layouts and
  finished pages are read and reused; nothing is imitated.
- **Nothing is given** → free mode. `Theme.pick(topic)` picks one of 14 hue
  families from what the deck is about, in `mono`, `duotone` or `trio`.

Colour is a **ladder over one hue**, not a bag of pretty swatches: pale for
surfaces, dark and small for emphasis, cards lighter than the page. Ask for
`help` for the full menu of palettes, built-in templates and deck types.

## Requirements

- Python 3.10+ with `python-pptx`, `pymupdf`, `Pillow`, `matplotlib`, `lxml`
- **Standard LibreOffice** for trustworthy rendering
  (Windows: `C:\Program Files\LibreOffice\program\soffice.com`;
  override with `AI_PPT_SOFFICE`)
- A CJK font (Microsoft YaHei / SimSun / Noto Sans CJK)

Check what is actually available:

```bash
python scripts/preflight.py
```

If LibreOffice is missing, visual QA degrades to metrics-only and `preflight`
says so. Do not describe a deck as visually verified in that state.

> **Note on the bundled `libreoffice-kit`:** in this environment its Impress
> renderer ignores a text box's declared width for CJK runs (laying them out
> against the full canvas) and can mis-scale type. Certifying layout from its
> output produces false negatives, so `render_qa.py` deliberately prefers system
> LibreOffice.

## Quick start

```bash
# 1. index the sources
python scripts/source_ingest.py refs/ --out work/

# 2. find a claim and get its locator
python scripts/source_ingest.py --find "m6A" --out work/
#   → [S1 p.4] ...the sentence...

# 3. pull the whole figure as the authors laid it out
python scripts/source_ingest.py paper.pdf --out work/ --figure-page 4 --list-regions
python scripts/source_ingest.py paper.pdf --out work/ --figure-page 4 --dpi 250

# 4. render formulas
python scripts/math_render.py --latex "\frac{QK^T}{\sqrt{d_k}}" \
    --out work/math --name eq1 --number 1.1

# 5. build (see SKILL.md §S3 for the API)
python build_my_deck.py

# 6. audit — hard gate
python scripts/audit_deck.py final.pptx --work work/ --out audit/ --dpi 150

# 7. end-to-end regression on a real paper
python scripts/e2e.py path/to/paper.pdf
```

## How to read the audit

`audit/audit_report.md` ends in a verdict. A `❌ FAIL` is not a suggestion:

- `metrics-overflow*` — text will not fit its box; widen, shorten, or shrink
- `text-overflow-bottom/right` — the renderer confirms it spills out
- `undefined_citations` / `uncited_references` — citation bookkeeping is broken
- `locator_problems` — a citation points at a page that does not exist
- `overlap`, `margin-*`, `low-contrast`, `font-too-small` — visual defects

Passing the gate is necessary, not sufficient: **look at the rendered PNGs.**
Numbers cannot see a card taller than its content or a figure hugging one edge.

## Design decisions worth knowing

- **Formulas are PNGs, not OMML.** OOXML equations written by non-Microsoft tools
  are a compatibility gamble (WPS and older PowerPoint drop them). The source
  LaTeX is kept in `math_manifest.json`, so fidelity is still checkable.
- **Layout is measured, not estimated.** CJK has no spaces, so "does it fit" can
  never be guessed from word counts.
- **The builder refusing to save is the feature.** A `FitError` means a real
  person would have seen broken text.
- **Content-adaptive boxes.** Cards and panels size to their content, because a
  fixed tall box leaves a hollow region that reads as an unfinished slide.

## Provenance

Scripts and reference material here are original work for this skill. No
third-party slide template or paid asset is redistributed. LibreOffice and
LibreOfficeKit are separate MPL/LGPL components invoked as external tools.