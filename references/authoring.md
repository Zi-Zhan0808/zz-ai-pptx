# Authoring: layout, typography, figures, formulas, tables

Load this when you are actually placing content. It encodes what years of
"this deck looks wrong and I can't say why" reduces to.

## The type scale

Four sizes. Not six, not "mostly four".

| Role | Size | Notes |
|---|---|---|
| Slide title (claim) | 36–44pt bold | must survive being read from the back row |
| Section header / strong subhead | 26–30pt bold | |
| Body | 16–20pt | academic decks skew 18pt; below 16 is unreadable when projected |
| Caption, locator, footer | 10–12pt | the only small text allowed |

`build_deck.py` fixes these as `TITLE_PT/SECTION_PT/BODY_PT/CAPTION_PT`.
Deviate deliberately, not accidentally.

**Size contrast is what makes a deck look designed.** If titles are 28pt and body
is 18pt, nothing reads as hierarchy. Aim for title ≥ 1.8× body.

## Fonts

Write font names that exist on the *audience's* machine, and ideally also here so
QA is honest:

- **Safe everywhere**: Microsoft YaHei（微软雅黑）, SimHei, SimSun, Arial, Calibri,
  Cambria, Times New Roman.
- **Avoid**: Aptos (post-2023 Office default, absent on older installs), and
  decorative faces for body text.
- **Never** mix three families. One sans for display + body is enough.

Latin/CJK mixing: set the East Asian typeface explicitly if a run is mostly CJK
but starts with a Latin token, otherwise some renderers substitute the Latin
face's CJK glyphs and the weight jumps mid-line.

Numbers in Chinese text: keep a space around Latin runs (「基于 12 项实验」, not
「基于12项实验」). It is the single cheapest legibility win in mixed-script text.

## Spacing

- 0.5 in minimum margin on all sides. `build_deck.py` uses 0.9 in, which reads as
  confident rather than cramped.
- 0.3–0.5 in between blocks; pick one and use it everywhere.
- Line spacing 1.4–1.6 for Chinese body (CJK needs more leading than Latin).
- Titles: 1.1–1.2.

**Whitespace is content.** A slide that fills every square inch has no emphasis
left to give. If the layout has a large dead area, the fix is usually *less*
content on that slide or a different layout, not a bigger font.

## Where text goes wrong (and the mechanics behind it)

- **Overflowing the box.** PowerPoint wraps text to the box width; if the box is
  too short, the text spills past the bottom edge and, because it is not clipped
  by default, it overlaps whatever is below. `build_deck.py` measures with real
  font metrics and raises `FitError` instead of shipping this.
- **The `spAutoFit` trap.** A text box created with `add_textbox` defaults to
  "resize shape to fit text", which grows the *shape* rather than wrapping inside
  it. Set `auto_size = MSO_AUTO_SIZE.NONE` and `word_wrap = True` explicitly —
  `build_deck.py` does both.
- **CJK has no spaces**, so "does it fit" cannot be estimated by word count.
  Measure glyphs (`textfit.wrap_text`), and remember Chinese breaks between
  almost any two characters while keeping 「，。」 off the line start (避头尾).
- **Centred body text.** Left-align paragraphs and lists; centre only titles and
  large numbers. Centred multi-line Chinese body is nearly unreadable.
- **Text on a busy figure.** Put the label in a solid pill or a filled band, not
  directly over the artwork.

## Layout idioms (pick per slide, vary across the deck)

| Idiom | Use for | Method |
|---|---|---|
| Title / section | framing, transitions | `title_slide`, `section_slide` |
| Bullets | 3–5 parallel points | `bullets` |
| Two-column | claim + evidence | `split` |
| Stat callouts | 1–4 numbers that carry the argument | `stats` |
| Figure | the paper's own figure | `figure` (+`side_note`) |
| Formula | derivation, definition | `math` |
| Table | comparison, ablation results | `table` |

**Do not use the same idiom three slides in a row.** A deck of eight bullet
slides is a document, not a talk.

### Lay-out proportions that look intentional

- Two-column: 55/45 or 60/40. 50/50 looks like a form.
- Cards hold their content. A card whose text occupies the top third reads as an
  empty box — size the card to content (`stats` and `split` do this).
- A portrait/square figure cannot fill a 16:9 slide by height. Give it a side
  column (`side_note`) rather than a wide dead margin.

## Figures

**Use the paper's own figure.** It carries the authors' axis choices, error bars,
and panel labels — redrawing loses exactly the information a reader needs to
judge the claim.

Budget: **one figure per content slide.** If a figure has 12 panels, either show
the composite and point at the panels you care about in `side_note`, or crop to
the panel that supports your sentence.

### Getting the whole figure, not fragments

`source_ingest.py --figures` extracts *embedded* images. In most papers a figure
is assembled from many small images plus vector axes, so that yields scattered
pieces. Render the figure region instead:

```bash
python scripts/source_ingest.py paper.pdf --out work/ --figure-page 4 --list-regions
python scripts/source_ingest.py paper.pdf --out work/ --figure-page 4 --dpi 250
```

`--list-regions` prints candidate rects; if the largest is not the figure you
want, pass `--figure-rect x0,y0,x1,y1` (PDF points) explicitly.

### Attribution

- Unmodified: caption it `原图自 [n] Fig.X`.
- Redrawn/replotted: the slide itself must say `重绘自 [n]` — in the caption, not
  only in the speaker notes.
- Never crop away a panel label you still cite.

## Formulas

Default: **PNG**, rendered from LaTeX.

```bash
python scripts/math_render.py --latex "\mathrm{softmax}\left(\frac{QK^T}{\sqrt{d_k}}\right)V" \
    --out work/math --name attention --number 3.1
```

Why not native OMML: equations written into OOXML by non-Microsoft producers are
a compatibility gamble — WPS and older PowerPoint builds drop or mangle them, and
the deck silently loses its mathematics in the room where it matters. A PNG looks
the same everywhere and can be visually verified.

A picture is still *verifiable* because the source LaTeX is kept in
`math_manifest.json` beside it; `audit_report.md` lists formula fidelity as a
human-confirmation item. Say so in the delivery note.

Rules:
- One formula per slide, at the size it will be read (a formula squeezed to 1.5 in
  tall is decoration).
- Equation numbers `(3.1)` consistent with the text that refers to them.
- Explain *what the formula implies* in one sentence; do not narrate the
  symbols.
- mathtext covers most LaTeX math but not `\begin{align}`, `\text{}` with
  complex content, or packages. If a formula will not render, simplify the
  notation rather than shipping a broken image.

## Tables

- Put the values under comparison in the same column order across tables.
- Bold the winning/baseline row; do not colour-code without a legend.
- Caption above, source locator below.
- Numbers: consistent decimals, right-align. 「0.02」and「2%」in the same column is
  a defect.
- More than ~6 columns in 16:9 will not be readable — split or transpose.

## Colour

- One palette per deck, defined once (`Theme`). Every colour comes from it.
- Dominance over equality: one colour carries 60–70% of visual weight.
- Contrast for projected text: aim ≥ 4.5:1, and prefer ≥ 7:1 for body. Light-grey
  body text on white dies on a projector.
- Avoid using red/green as the only signal for "good/bad" (colour-blind readers).

### Do not alternate dark and light slides

The "dark title / light content / dark conclusion sandwich" is a common template
habit and it is uncomfortable to sit through. Every full-screen change from
near-black to near-white makes the eye re-adapt; across a deck that is a series
of flashes, worst on a projector in a dim room.

| | canvas | panel | jump |
|---|---|---|---|
| Comfortable | `#FCF8F3` | `#F2E4D4` | 1.18:1 |
| Painful | `#FFFFFF` | `#16181D` | 17.80:1 |

Keep the deck on one bright side of the range and create structure *inside* it:

- **opener / divider**: a deeper tint of the same hue (not black — no inverting)
- **cards**: a third step of the same tint
- **emphasis**: one saturated accent used for a numeral, a rule, or an icon —
  never a whole slide

Target a background jump **≤ 1.75:1** between adjacent slides, no pure `#FFFFFF`
and no near-black canvas. `scripts/palette_light.py` checks exactly this.

### 淡彩: soft colour, not grey

A "safe" palette often drifts to near-grey and looks washed out. Keep the hue
visible by holding lightness high and raising saturation:

- backgrounds / tints: lightness 89–97%, saturation **40–60%**
- accent (one element per slide): saturation 50–65%, lightness 30–48%
- text: near-black with a hint of the same hue (e.g. `#221E1A` on a warm canvas)

That band stays calm while still reading as a deliberate colour — pastel rather
than beige.

## Slide furniture

- Page numbers: yes, small, bottom outer corner.
- Source locator: bottom left, 10–12pt, on every content slide that makes a claim.
- **No decorative bars, no accent stripes under titles, no borders around cards.**
  These are the visual signature of generated slides. Use tint, shadow, or an
  icon instead.
- Speaker notes: put the *talk track* here — what you will say, not a copy of the
  slide text. A slide that reads itself aloud means the notes are empty.
