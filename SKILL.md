---
name: zz-ai-pptx
description: "Use this skill whenever a presentation must be built from source material — papers, PDFs, reports, notes, data — and delivered as a native, editable .pptx. Covers academic decks (组会汇报, 论文精讲, 课程, 开题/答辩, 综述), research and technical briefings, and any deck where claims must be traceable to a citable page. Trigger on: 做PPT, 做一个演示文稿, 学术汇报, 组会, 论文讲解, 开题答辩, 综述汇报, journal club, slide deck, presentation, lecture slides, seminar talk, or when the user supplies PDFs/DOIs and asks for slides. Also use when auditing or fixing an existing deck for text overflow, citation consistency, or Chinese typography."
license: MIT. Bundled scripts are original work; no third-party slide template is redistributed (see the 待更新 note on assets/templates).
---

# zz-ai-pptx — source-grounded deck construction

Turn reference material into a **native, editable `.pptx`** whose every claim is
traceable to a page, whose typography is verified numerically, and whose prose
sounds like a researcher wrote it.

This skill is deliberately mechanical about the things models get wrong —
invented citations, overflowing text, formula soup, template-speak — because
enthusiasm does not survive three hours of revision.

## 0. The one rule

**No claim without a locator.** Every `[n]` in the deck resolves to an entry in
the reference slide, and every entry carries a verified locator (`S1 p.4` —
source id + physical page). If you cannot point at a page, you do not get to make
the claim.

Corollary: **never cite from memory.** If the user gives only a title or DOI and
no full text is available, stop and ask — do not reconstruct a paper from
recollection. A confident wrong citation is worse than no deck.

## 1. Environment reality check (do this first, once)

The tooling here is not the usual LibreOffice-plus-pandoc stack, and the
difference changes what "verified" means:

| Capability | How it works here |
|---|---|
| Build `.pptx` | `python-pptx` (installed) |
| Read PDFs, page text, figures | `PyMuPDF` (installed) |
| Trustworthy slide render | **Standard LibreOffice** — `soffice.com`, installed at `C:\Program Files\LibreOffice\program\soffice.com` |
| Render PPTX→PNG | LibreOffice → PDF, then PyMuPDF rasterises |
| LaTeX formulas | `matplotlib` mathtext → transparent PNG |
| Text fit | `scripts/textfit.py`, real font metrics |

⚠️ **Do not use the bundled `libreoffice-kit` to judge layout.** In this
environment its Impress renderer ignores a text box's declared width and lays
CJK runs out against the full canvas, and it can also mis-scale type. A title
that badly overflows its box renders looking merely "wide", so a QA pass built
on it returns false negatives. `scripts/render_qa.py` therefore prefers system
LibreOffice and falls back to metrics-only evaluation when it is absent.

Run the self-check before building anything:

```bash
python scripts/preflight.py
```

It reports which renderer, PDF engine, and fonts are actually available, and
tells you which checks will be degraded. Do not announce visual verification as
complete if the preflight says the renderer is missing.

## 2. The pipeline

Three phases, not seven passes. The old shape ran de-slop, citation, layout and
design as separate sweeps, which meant writing the deck four times; most of that
work is avoidable by deciding the constraints before the first slide exists and
then checking everything once.

```
A  DECIDE   intake → sources indexed → outline.md → theme chosen
B  SHAPE    build the whole deck in one pass, overflow-guarded, de-slopped as written
C  CHECK    one audit run (4 gates) → fix only what it names → deliver
```

If a gate fails you fix that specific thing and re-run the same command. You do
not go back to A.

### Misreading the request

If the user says **help / 帮助 / 你会什么 / 有哪些选项 / 有哪些配色**, read
`HELP.md` and answer from it: two modes, 14 families × 3 modes, built-in
templates, deck types, the 11 layouts, the four gates. Then ask the one question
that matters — is there a template, and what is the occasion.

### A1 — Intake (ask, but do not interrogate)

Establish only what changes the deck: **audience**, **occasion**, **time
budget**, **what the user wants emphasised**. If the user is vague, pick
sensible defaults (组会, 15 min, evidence-led) and say what you assumed in one
line. Do not block on a questionnaire.

Two things decide the entire look, so settle them here rather than discovering
them later:

| Given | Mode | Consequence |
|---|---|---|
| a `.pptx` template | template mode (`B2`) | its palette, layouts and cover; nothing invented |
| no template | free mode | `Theme.pick(topic)` chooses one of 14 families |

Settle **four dials** before building anything. They are what determines cost,
and cost is the thing the user notices.

| Dial | Values | Detail |
|---|---|---|
| 能力档位 | T1 basic · T2 fixed layouts · T3 content-driven · T4 own schematics | HELP §11 |
| 自由度 | L1 strict-on-material · L2 own wording (default) · L3 online | HELP §12 |
| 配色 / 模板 | the user's file · a built-in template · chosen from the content | A4 |
| 页数 / 时长 | from the occasion | HELP §4 |

**Then send a budget card and wait** (HELP §13): tier, freedom, pages, palette
source, estimated passes and operations, whether it will go online, and two or
three cheaper options. Anything at T3 or above, or any online step, is quoted
before work starts rather than after.

Two rules that make the tiers mean something:

- **A lower tier must not borrow a higher tier's ability.** If the user picked T2,
  do not invent a layout from the content because it would look better; that is T3.
- **T4 always starts with one sample page.** Own-drawn schematics are the most
  expensive thing this skill does — each geometry change is a rebuild, a render and
  a measurement — so one page is shown, and approved, before the rest is drawn.

Default to T3 + L2 for anything that will be presented, and say which defaults you
took instead of taking them silently.

Hard stop only if: no reference material at all, and the request is academic.
Then ask for at least one citable source.

### A2 — Ingest

```bash
python scripts/source_ingest.py refs/ --out work/            # index every page
python scripts/source_ingest.py --find "attention sink" --out work/   # locate a claim
```

This writes `work/sources.json` (per-page text, sha1, page count, DOI) and
`work/sources.md` (human-readable). **Read `sources.md` before writing any
slide** — planning from the index, not from memory, is what makes the locators
real.

To cite something, search it and use the returned locator verbatim.

### A3 — Plan

Write `outline.md`: one line per slide as
`N | claim-style title | layout | locators`.

Rules that matter for academic work:

- **Titles state a claim, not a topic.** "METTL3 丢失后身份基因先塌陷再分化"
  beats "背景介绍". A reader should get the argument from titles alone.
- Every content slide names its locator(s). If you cannot fill that column, the
  slide is speculation — cut it or mark it clearly as your own inference.
- Budget slides from the time slot: ~1.5–2 min per content slide.
- Include at least one slide of **your own judgement** (limitations, competing
  explanation, what would falsify this). Pure summary is not a briefing.
- Structure is yours to choose. Do not force a template; do force one claim per
  slide.

### A4 — Choose the palette

```python
from build_deck import Deck, Theme
d = Deck(Theme.pick("细胞自噬与衰老"))       # family from the topic
d = Deck(Theme.pick(family="teal", mode="mono"))
```

The theme is a **ladder over one hue**, and the levels are ranked so that colour
gets lighter as it gets bigger:

| level | luminance | where it goes |
|---|---|---|
| canvas | .90 | the page |
| wash | .82 | table stripes, large bands |
| panel | .73 | mid emphasis |
| veil | .575 | one block per page, at most |
| card | .962 | cards, **lighter than the canvas** |
| accent | .14 | rules, numerals, chips — small and dark |
| ink | .05 | body text |

A card darker than its page reads as a hole, which is why `card` sits above
`canvas`. `secondary` is the card level, so cards come out pale by default
rather than as saturated blocks.

Do not hand-pick hexes. If the deck needs a different mood, change the family or
the mode; if it needs a different *emphasis*, change the sizes, not the colours.

### The form follows the content, not the library

Before choosing a layout, say what the content *is*:

- a judgement that fits one sentence -> `lead_statement`
- two mutually exclusive explanations -> `duel`, with the observation that decides between them in the verdict band
- a distance or a spatial relationship -> `distance_schematic`, so the argument is a length on the page
- a categorical result (2x3, sensor x condition) -> `matrix` of marks, not a table of words
- a before/after magnitude -> `delta`, one track per pair
- items with no hierarchy -> `ledger`; items where one leads -> `focus` (one big, three small)

The test for whether a layout is doing work: **swap two of its blocks. If the page
still reads the same, the layout was a container, not an argument.** Change the
form, not the wording.

Four to six different forms across an 8-page deck is normal. Forcing in a seventh
because "variety is good" is the same eight-legged essay with more legs.

### B1 — Shape it in one pass

Use `scripts/build_deck.py`. It is not a convenience wrapper — it **refuses to
emit text that does not fit**, measured with real font metrics:

```python
from build_deck import Deck, Theme

d = Deck(Theme.calm("linen"))   # linen / mist / sage / blush
```

**Vary the layout.** The single most common complaint about generated decks is
that every content slide has the same skeleton. This library has eleven
structurally different layouts; a deck should use at least six of them.

| Layout | Use for | Shape of the page |
|---|---|---|
| `cover` | opening | left-aligned headline, rule, subtitle |
| `agenda` | contents | numbered index, one or two columns |
| `section_divider` | chapter breaks | ghost numeral + section claim |
| `lead_statement` | the argument beats | one oversized sentence |
| `cards_grid` | 2–6 parallel items | icon cards in a grid |
| `stat_row` | 3–5 numbers | numerals across a divided band |
| `process_flow` | 3–6 steps | chevrons, or a vertical list |
| `compare_split` | before/after, A/B | two symmetric columns, one rule |
| `quote` | a definition or a principle | display quote with source |
| `figure` | the paper's own figure | image + optional reading guide |
| `summary` | closing the argument | 判断 / 证据 / 意味着什么 table |
| `closing` | the ask | headline + action chips |

```python
d.cover("结论式主标题", subtitle="方法学副标题", meta="组会 · 20 分钟", badge="PROJECT")
d.agenda(["问题", "方法", "结果", "判断", "下一步"])
d.section_divider(1, "第一部分要证明什么", bullets=["要点一", "要点二"])
d.lead_statement("一句话说出你的判断。", support=["支撑说明"], cite="[1] S1 p.2")
d.cards_grid("标题即结论", [("要点", "说明", "shield")], cols=2, cite="[1] S1 p.3")
d.stat_row("三个数字", [("3.2×", "下调倍数")], cite="[1] S1 p.3")
d.process_flow("流程", [("第一步", "说明")], cite="[1] S1 p.2")
d.compare_split("对照", ["左一"], ["右一"], left_title="前", right_title="后")
d.quote("要强调的一句话。", source="来源", role="补充")
d.summary("三条结论", [["判断", "证据", "含义"]], cite="[1] S1 p.4")
d.closing("你给材料，我来跑流程", actions=["给出参考文件"])
d.references(["作者. 标题. 期刊. 年. S1 p.3"])
d.save("final.pptx")
```

Icons are **drawn from native shapes** (`shield`, `target`, `layers`, `branch`,
`grid`, `check`, `alert`, `flow`, `stack`, `dot`). Nothing is downloaded, so
every glyph is licence-clean, palette-matched, and still editable in PowerPoint.

A `FitError` is the builder doing its job: widen the box, cut text, or reduce
the size. Never work around it by turning off `strict` for the final deck.

**Figures — get the whole figure, not its debris.** Extracting embedded images
from a paper usually yields scattered sub-panels, not Figure 2. Render the
figure region instead:

```bash
python scripts/source_ingest.py paper.pdf --out work/ --figure-page 4 --list-regions
python scripts/source_ingest.py paper.pdf --out work/ --figure-page 4 --dpi 250
```

Multi-panel figures are usually portrait or square; pass `side_note` so their
dead margin carries the reading guide instead of leaving a blank half-slide.

**Formulas** are PNGs (default), not OMML — see
`references/authoring.md#formulas`. Always keep the source LaTeX: `math_render.py`
writes `math_manifest.json` next to the images, and `audit_report.md` asks a
human to confirm the picture matches the LaTeX.

**Footers must name a reference, not just a page.** `[1] S2 p.2` is a
citation; `S2 p.2` alone tells the reader where to look but not which entry it
belongs to. The builder raises rather than guessing the reference number — a
guessed number is exactly the error this skill exists to prevent.

### B2 — Inside a given template

When the user hands you a `.pptx` to "follow", **do not imitate it**. An
imitation is roughly the right green with roughly the right bar and a hundred
small errors; the template is a file, and the file already contains its own
palette, chrome and finished pages. Read them.

```bash
# 1. read the template: palette, fonts, layouts, page catalogue, brand marks
python scripts/template_ingest.py 模板.pptx --out work/template

# 2. remove the author's institution marks (confirm the list first!)
python scripts/template_ingest.py 模板.pptx --out work/template \
    --strip-all-candidates
```

Step 1 prints a score for every image and writes `template_profile.md`. A logo
cannot be identified by name (`Picture 19`), only by evidence — repeated across
slides, present in a layout, parked in the header band, extreme aspect for a
wordmark. **Show the user the list and get a yes before deleting**, because a
figure legitimately used twice also scores. Step 2 also drops the media parts,
so the file actually shrinks instead of carrying someone else's stationery.

Then build:

```python
import json
from template_build import TemplatedDeck

prof = json.load(open("work/template/template_profile.json", encoding="utf-8"))
d = TemplatedDeck.open("work/template/模板-clean.pptx", prof)

d.cover(title_en="...", title_cn="...", dept="...", group="...",
        presenter="汇报人：...", date="2026年1月")
d.agenda(["基本信息", "研究背景", "研究思路", "研究结果", "结论"])
d.cards("研究背景", [("领域现状", ["...", "..."]), ("研究问题", "...")], section=2)
d.bullets("要点", ["...", "..."], section=2, kicker="领域现状")
d.figure("论文原图", "work/figs/fig1.png", caption="...", section=3, cite="[1] p.4")
d.table("结果", ["组别", "n", "P"], [["女性", "9", "0.015"]], section=4, cite="[1] p.5")
d.closing("敬请各位老师批评指正")
d.save("final.pptx")
```

Three layers, three mechanisms — this is the whole idea:

| Layer | What it gives you | How |
|---|---|---|
| theme | palette + typeface | `theme_from_profile` re-skins every block, contrast re-verified |
| layout | title bar, number badge, footer band | new pages are slides of *that* layout, so chrome is inherited, not redrawn |
| slide | a finished composition | `clone(index)` + `set_texts(...)`, for cover / agenda / closing |

Why clone the cover instead of drawing one: a cover is a composition — a field
of colour, a rule, three weights of type at fixed offsets. Re-deriving it from
rectangles yields a page that is *almost* the template, which reads as broken.
Copying keeps it pixel-identical **and** editable, because the shapes are real
shapes and the picture relationships are remapped, not flattened.

`agenda(n)` clones the template's agenda and **resizes it to n rows** at the
template's own pitch. A five-part report must not keep a sixth empty row.

**Narrow boxes.** Designers draw a box tight around their own wording. Put a
longer group name in and the line breaks mid-word. `template_profile.json` is
meant to be edited: `pages.<role>.widen` widens named boxes when the space to
the right is empty.

```json
"cover": {"index": 5, "widen": {"文本框 14": 2.6, "文本框 11": 4.6}}
```

**Two things this mode changes about the rules below.** The canvas colour comes
from the template, so the no-harsh-background rule becomes the no-*alternation*
rule: a green cover, light content, green closing is a bookend pattern and
fine; green/white/green/white is not. And the type scale is the template's, not
ours — `template_ingest.py` reports what it measured so you can stay inside it.

Detail, including how to read the profile: `references/template.md`.

### C1 — One audit run

```bash
python scripts/audit_deck.py final.pptx --work work/ --out audit/ --dpi 150
```

One command runs all four gates. Do not run the sub-checks one at a time and do
not re-render unchanged slides; the audit is the pass, not the first of several.

| Gate | Catches | Script |
|---|---|---|
| layout | overflow (metrics **and** renderer comparison), margins, overlap, font floor, contrast | `render_qa.py` |
| colour | too many hues, oversaturation, no focal point, flat rhythm, flood | `palette_audit.py` |
| citations | `[n]` ↔ entry ↔ locator, both directions | `refs_check.py` |
| structure | mixed arrow shapes, unanchored rules, title/diagram count mismatch, orphan decoration | `design_check.py` |

**Colour is judged as composition, not as taste.** The numbers that matter:

- `too-many-hues` — more than 4 hue families among visible patches. Past that the
  reader cannot tell which colour is speaking.
- `no-focal-point` — the biggest patch owns under 38% of the coloured area. Eight
  equally sized tasteful panels is a page with no first thing to look at.
- `flat-rhythm` — patch sizes with a coefficient of variation under 0.35. Equal
  parts read as a ledger.
- `area-flood` — chroma-weighted coverage above 62%: full of colour, empty of air.

These were written from a real deck that passed every contrast test and still
looked loud; contrast answers "can it be read", not "where do I look first".

**IMPORTANT — preset names.** `design_check` compares against the strings stored
in `<a:prstGeom prst="...">`, not python-pptx constant names.
`MSO_SHAPE.PENTAGON` is written as `homePlate`, `CHEVRON` as `chevron`. A checker
written against constant names matches nothing and reports a clean deck; that
exact bug caused the "第一个背景箭头" defect to ship. Verify any new geometry
check against a real file before trusting a `0 issues` result.

### C2 — Fix only what it names

Read the report and change those slides. Resist the urge to redesign while you
are in there: a second full rewrite costs more than the defect and introduces
new ones. **Then look at the PNGs** — numbers miss optical imbalance: a card
taller than its content, a figure hugging one edge, an empty lower half.

Stop condition: all four gates pass and you have looked at every slide of a new
deck. Expect one round of fixes; more than two means the outline was wrong, not
the slides.

### C3 — Deliver

```bash
present({"files": [{"path": "final.pptx"}, {"path": "audit/audit_report.md"}]})
```

Report: what the deck argues, which assumptions you made, and the
human-confirmation items from `audit_report.md` (physical-page convention,
redrawn figures, formula fidelity). Keep `sources.md`, `outline.md` and
`audit_report.md` alongside the deck.

### Writing it de-slopped the first time

De-slop is not a later pass. `references/prose.md` is a drafting constraint:

- Cut 「近年来，随着……的快速发展」openers, 空转大词（深刻变革/全面赋能）,
  「由于时间有限」apologies, 「未来可进一步研究」as a conclusion.
- Keep hedges with meaning（在 X 条件下 / 统计上不显著）, precise terms, author
  stance. Never soften a limitation into vagueness — that is a different error.
- Humanity comes from judgement and hedging, not from casual phrasing.

#### Wording, not just sentence shape

De-slopping also means choosing smaller words. The failure mode is not a wrong
sentence but a sentence that is *dressed up*:

- **Kickers are single nouns.** 回顾 / 证据 / 前沿 / 未来 / 局限 / 结论 / 总结 —
  not 「回顾一句」「模型被验证」「当下的前沿」「未来方向」「带走」. A kicker is a
  signpost, and a signpost that needs parsing is a heading.
- **Titles state a judgement**, not a topic with a colon in front of it. If the
  words before the colon can be deleted without loss, delete them.
- **Use fewer force words.** 最 / 唯一 / 一个比一个 / 一条比一条 / 彻底. Keep them
  for the case that is genuinely singular (then 最 is correct and worth using) and
  never as filler. Hierarchy that exists reads better as 层层深入 or 四层推进.
- **One 最 per slide.** Two superlatives cancel each other.
- **No small tag that repeats the title.** If the title already says the priority,
  a 「优先」chip on top of it is decoration with a word in it.
- **Statement pages stay quiet.** `lead_statement` caps long sentences to 32pt and
  three lines: the page exists to let the audience breathe, not to shout. A
  conclusion that fills half the slide cannot be read in one pass.

`HELP.md` §10 has the three registers (平实 / 学术 / 讲稿) and a before/after table.

The audit's `check_ai_smells` catches what slips through, which is cheaper than
rewriting every paragraph afterwards.

### Turning a found defect into a check

When the user reports a visual defect, do not only fix the slide. Ask which
*class* of defect it is, add a check for the class to `design_check.py`, and add
a negative test to `scripts/design_check.py`'s companion pattern (build a deck
that contains the defect and assert it is reported). Otherwise the same wall gets
hit again on the next deck, in a slide nobody happened to look at.

A check is only trustworthy after it has failed on a deliberately broken deck.

## 3. Quality rules that are not negotiable

0. **Never alternate black and white slides.** A deck that flips between a
   near-black and a pure-white full screen forces the audience's pupils to
   re-adapt on every page change — on a projector it reads as flicker and eye
   strain. Measured on a real deck: 17.8:1 between backgrounds, nine times in
   fourteen slides. Keep the whole deck on one bright side of the luminance
   range; build structure from *layering within that range* (a slightly deeper
   tint for openers, a card tint, a rule) and from type weight and whitespace.

   ```
   good : canvas #FCF8F3  →  panel #F2E4D4        jump 1.18:1
   bad  : canvas #FFFFFF  →  panel #16181D        jump 17.80:1
   ```

   The shipped palettes (`Theme.calm("linen"|"mist"|"sage"|"blush")`) are
   **淡彩 — soft colour**: saturation carries the hue (40–60%) while lightness
   stays high (89–97%), so the page is colourful without becoming heavy. Verify
   any change with `python scripts/palette_light.py`; it enforces body >= 7:1,
   captions >= 4.5:1, and a canvas→panel jump <= 1.75:1.

   Never put dark text on a mid-tone card either — that pairing is what makes a
   projected deck look muddy.
1. **Traceable citations** — locator on every claim; the page numbered as the PDF
   reader shows it, and say so if you used printed page numbers instead.
2. **Original figures** — use the paper's own artwork. If you must redraw, label
   it 「重绘自 [n]」 on the slide.
3. **Verifiable formulas** — a rendered object, never a text string wearing a hat.
   Equation numbers consistent with the body.
4. **De-slopped without damage** — terminology, statistics, and meaningful
   hedges survive rewriting.
5. **Structure chosen for the material** — no template cosplay.
6. **One claim per slide, stated in the title.**
7. **Everything that is text is editable text** — no screenshots of text.

## 4. Degraded modes

| Situation | Do this |
|---|---|
| No standard LibreOffice | audit runs metrics-only; say visual layout was not inspected and why. Do not claim it was verified. |
| No PyMuPDF | ask for text extraction another way, or for the text itself; figures degrade to "user-supplied" |
| No matplotlib | formulas become images the user provides, or plain Unicode notation with the LaTeX shown in a note |
| No DOI/full text for an outside claim | cite only user-provided sources; mark the rest 「待人工确认」 |
| User supplies no reference material | stop for an academic request; never invent content or citations |

## 5. Files

```
SKILL.md                  this file
HELP.md                   capabilities and options, for "help"
scripts/build_deck.py     overflow-guarded deck builder (Deck, Theme, legacy idioms)
scripts/theme_library.py  tonal ladders, 14 hue families, mono/duotone/trio
scripts/palette_audit.py  colour discipline: hues, focal point, rhythm, flood
scripts/layouts.py        11 structurally different slide layouts + native-shape icons
scripts/template_ingest.py read a given .pptx → profile, brand-mark removal
scripts/template_build.py  build inside that .pptx (clone pages, inherit layout)
scripts/template_e2e.py    self-contained regression for the template path
scripts/textfit.py        CJK-aware line breaking from real font metrics
scripts/render_qa.py      LibreOffice→PDF→PNG + text-box overflow comparison
scripts/audit_deck.py     full gate → audit_report.md
scripts/refs_check.py     bidirectional citation + locator audit
scripts/source_ingest.py  PDF/MD → sources.json, locator search, figure regions
scripts/math_render.py    LaTeX → transparent PNG (+ manifest)
scripts/preflight.py      environment capability report
scripts/design_check.py   structural design lint (arrow/rule/count/orphan)
references/authoring.md   layout, typography, figures, formulas, tables
references/template.md    reading a template profile; slot maps; brand marks
references/pitfalls.md   defects that already cost time, each with its check
NOTICE.md                 third-party dependencies, prior art, how to verify
scripts/provenance_check.py  line-by-line proof that no other skill's code is here
references/prose.md       de-slopging with academic-specific guardrails
references/workflows.md   worked patterns for 组会 / 论文精讲 / 综述 / 答辩
assets/example_outline.md a filled-in outline to imitate
assets/templates/课堂pre01/  template slot: profile + prompts shipped,
                          template.pptx itself is 待更新 (not distributed)
```
