# zz-ai-pptx

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
