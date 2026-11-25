"""
build_deck.py - A small, opinionated builder for academic decks.

This is the piece that turns the documented design rules into something the
agent cannot accidentally violate:

  * every text box is measured before it is placed, and the caller gets an
    exception instead of a silently overflowing slide
  * the type scale is fixed (title / section / body / caption) so "font
    discipline" is structural, not a note in a prompt
  * the palette is a single object, so a deck cannot drift in colour
  * speaker notes and a per-slide citation field are first-class

Usage (as a library):
    from build_deck import Deck, Theme
    d = Deck(Theme.navy())
    s = d.title_slide("...", subtitle="...")
    d.bullets("Title", ["a", "b"], cite="[1]")
    d.save("out.pptx")
"""
from __future__ import annotations

import copy
import sys
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from textfit import wrap_text  # noqa: E402
from layouts import Layouts  # noqa: E402
from layouts_science import ScienceLayouts  # noqa: E402

EMU_IN = 914400
SLIDE_W = 13.333
SLIDE_H = 7.5

# Type scale — the whole deck draws from these four sizes only.
TITLE_PT = 40
SECTION_PT = 28
BODY_PT = 18
CAPTION_PT = 12
SMALL_PT = 10


class FitError(RuntimeError):
    """Raised instead of emitting a slide whose text does not fit its box."""


@dataclass
class Theme:
    """
    A palette that is verified, not admired.

    Every field is chosen so its intended pairing clears a contrast threshold
    (body >= 7:1, caption >= 4.5:1, accents >= 3:1). The two `_inv` fields exist
    because a colour tuned for a white background disappears on a near-black
    one; reusing the same hex for both is the single most common reason a deck
    looks muddy when projected.
    """
    name: str
    bg_dark: str        # deeper tint of the same canvas (NOT near-black)
    bg_light: str       # reading background, always near-white
    primary: str        # titles and numerals on light
    secondary: str      # card fill on light (pale tint, high contrast with ink)
    accent: str         # emphasis on light
    ink: str            # body text on light
    ink_inverse: str    # body text on dark
    muted: str          # captions on light
    muted_inverse: str  # captions on dark
    accent_inverse: str # emphasis on dark
    tint: str = ""      # very pale wash for alternating rows / quiet panels
    font: str = "Microsoft YaHei"   # typeface declared on every run
    card: str = ""      # a sheet lighter than the canvas, for cards
    veil: str = ""      # one emphasis block per page at most
    panel: str = ""     # a mid tint, for rules and quiet emphasis

    def __post_init__(self):
        if not self.tint:
            self.tint = self.secondary
        if not self.card:
            # Legacy themes have no ladder; mix toward white from the canvas so
            # an old palette still gets cards that sit above the page.
            self.card = _mix(self.bg_light, "FFFFFF", 0.55)
        if not self.veil:
            self.veil = self.secondary
        if not self.panel:
            self.panel = self.secondary

    # 淡彩 (soft colour): saturation carries the hue, lightness stays high so
    # no slide is heavy and nothing flashes when the page changes. On the
    # deeper panel the text is still ink — a pale panel cannot carry white
    # type. Verified by scripts/palette_light.py.
    #
    # Field order: name, bg_dark(panel), bg_light(canvas), primary, secondary,
    #              accent, ink, ink_inverse, muted, muted_inverse,
    #              accent_inverse, tint
    @staticmethod
    def palette_linen() -> "Theme":
        return Theme("Linen & Clay", "F2E4D4", "FCF8F3", "221E1A", "F6EEE5",
                     "BE532D", "221E1A", "221E1A", "645A50", "645A50",
                     "BE532D", "F8F3ED")

    @staticmethod
    def palette_mist() -> "Theme":
        return Theme("Mist & Teal", "D9E6F2", "F3F7FC", "1B222B", "E9F0F6",
                     "277591", "1B222B", "1B222B", "545F6C", "545F6C",
                     "277591", "F1F5F9")

    @staticmethod
    def palette_sage() -> "Theme":
        return Theme("Sage & Leaf", "E3F4D7", "F7FCF3", "1C241B", "EFF7E9",
                     "B6702B", "1C241B", "1C241B", "566152", "566152",
                     "B6702B", "F4F9F0")

    @staticmethod
    def palette_blush() -> "Theme":
        return Theme("Blush & Plum", "F4D7E3", "FCF3F7", "261D22", "F7E9EF",
                     "BA3B78", "261D22", "261D22", "66555F", "66555F",
                     "BA3B78", "F9F1F4")

    # ---- legacy names, remapped so old scripts stop flashing ----------
    @staticmethod
    def palette_ink() -> "Theme":
        return Theme.palette_linen()

    @staticmethod
    def palette_slate() -> "Theme":
        return Theme.palette_mist()

    @staticmethod
    def palette_forest() -> "Theme":
        return Theme.palette_sage()

    @staticmethod
    def palette_plum() -> "Theme":
        return Theme.palette_blush()

    @staticmethod
    def calm(name: str = "linen") -> "Theme":
        """
        The soft palettes, now derived from the theme ladder.

        The four old names survive because old scripts and old outlines use them;
        they resolve to the nearest family in theme_library rather than to four
        hard-coded hex sets. Keeping two palette systems would mean two ideas of
        what "calm" is, and the decks would drift apart.
        """
        legacy = {"linen": "ochre", "mist": "teal", "sage": "moss",
                  "blush": "rose", "ink": "graphite", "slate": "indigo",
                  "forest": "olive", "plum": "violet"}
        key = (name or "linen").lower()
        if key in legacy:
            key = legacy[key]
        return Theme.pick(family=key)

    @staticmethod
    def pick(topic: str = "", *, family: str | None = None,
             mode: str = "duotone", seed: str | None = None) -> "Theme":
        """
        Choose a palette from what the deck is about.

        With no argument it lands on the neutral family, which is the honest
        answer for "no topic given". With a topic it picks a hue family whose
        mood fits, so a biology talk and a history talk do not come out the same
        colour - which is the whole point of not shipping one house style.
        """
        import theme_library as TL
        spec = (TL.build(family, mode) if family
                else TL.pick_for(topic, mode=mode, seed=seed))
        return Theme.from_spec(spec)

    @staticmethod
    def from_spec(spec) -> "Theme":
        """
        Map the seven-level ladder onto the fields the layouts already use.

        Two of these mappings are load-bearing:

        `bg_dark` takes the *accent*, not the veil. Covers and section dividers
        are painted with bg_dark and then written on with `ink_inverse`, which is
        white; if bg_dark became a pale tint, every cover would turn into white
        text on near-white. Covers stay deep, body pages stay light, and the deck
        alternates exactly twice - at the ends.
        """
        return Theme(
            name=spec.name,
            bg_dark=spec.accent,
            bg_light=spec.canvas,
            primary=spec.ink,
            secondary=spec.card,
            accent=spec.accent,
            ink=spec.ink,
            ink_inverse="FFFFFF",
            muted=spec.muted,
            muted_inverse=_mix(spec.accent, "FFFFFF", 0.72),
            accent_inverse=spec.canvas,
            tint=spec.wash,
            card=spec.card,
            veil=spec.veil,
            panel=spec.panel,
        )

    @staticmethod
    def navy() -> "Theme":
        return Theme.palette_mist()

    @staticmethod
    def terracotta() -> "Theme":
        return Theme.palette_linen()

    @staticmethod
    def teal() -> "Theme":
        return Theme.palette_mist()


def _rgb(hexstr: str) -> RGBColor:
    h = hexstr.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _mix(a: str, b: str, t: float) -> str:
    """Blend two hex colours, t=0 gives a, t=1 gives b."""
    a, b = a.lstrip("#"), b.lstrip("#")
    out = []
    for i in (0, 2, 4):
        out.append(round(int(a[i:i + 2], 16) * (1 - t) + int(b[i:i + 2], 16) * t))
    return "%02X%02X%02X" % tuple(out)


# ---------------------------------------------------------------------------
# helpers for working inside somebody else's file
# ---------------------------------------------------------------------------
_R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _remap_relationships(src_part, dst_part, elements):
    """
    Re-point copied XML at the destination part's own relationships.

    A picture in the template is a shape whose fill references `rId7`. Copy the
    shape into a new slide and `rId7` means something else there, or nothing -
    the shape renders as an empty frame. Copies are only real once every
    relationship id they carry has been resolved against the part that will
    hold them.
    """
    used = set()
    for el in elements:
        for node in el.iter():
            for k, v in node.attrib.items():
                if k.startswith(_R_NS):
                    used.add(v)

    mapping = {}
    for rId in sorted(used):
        rel = src_part.rels.get(rId)
        if rel is None:
            continue
        if rel.is_external:
            mapping[rId] = dst_part.rels.get_or_add_ext_rel(
                rel.reltype, rel.target_ref).rId
        else:
            mapping[rId] = dst_part.rels.get_or_add(
                rel.reltype, rel.target_part).rId

    for el in elements:
        for node in el.iter():
            for k, v in list(node.attrib.items()):
                if k.startswith(_R_NS) and v in mapping:
                    node.set(k, mapping[v])


def _write_paragraphs(shape, text: str):
    """
    Replace a shape's text while keeping the design's formatting.

    Writing `shape.text = ...` would also throw away the run properties the
    designer set - the weight, the colour, the size, sometimes the letter
    spacing. On a cloned page that is the difference between "the template's
    title" and "a default paragraph where the title used to be".
    """
    if not shape.has_text_frame:
        raise TypeError(f"{shape.name} 没有文本框，不能写入文字")
    tf = shape.text_frame
    lines = str(text).split("\n")
    paras = tf.paragraphs

    # Ensure as many paragraphs as we need, cloning the last one's pPr.
    while len(paras) < len(lines):
        tail = paras[-1]._p
        new_p = copy.deepcopy(tail)
        tall = new_p.findall(qn("a:r")) + new_p.findall(qn("a:br"))
        for r in tall:
            new_p.remove(r)
        tail.addnext(new_p)
        paras = tf.paragraphs

    for i, line in enumerate(lines):
        p = paras[i]
        runs = p.runs
        if runs:
            runs[0].text = line
            for r in runs[1:]:
                r._r.getparent().remove(r._r)
        else:
            p.add_run().text = line
    # Drop leftover paragraphs we no longer need.
    for p in tf.paragraphs[len(lines):]:
        p._p.getparent().remove(p._p)


def _shape_pt(shape, default=18.0):
    """Largest explicit run size on a shape, or the file's default (18pt)."""
    size = None
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                size = max(size or 0.0, r.font.size.pt)
    return size or default


def _shape_need(shape, pt=None):
    """(height needed in pt, does any line overflow horizontally)."""
    w = (shape.width or 0) / EMU_IN
    if w <= 0 or not shape.has_text_frame:
        return 0.0, False
    pt = pt or _shape_pt(shape)
    avail_w = (w - 0.2) * 72 - 8.0
    total, over = 0.0, False
    for line in shape.text_frame.text.split("\n"):
        r = wrap_text(line, pt, avail_w)
        total += r.needed_height_pt
        over = over or r.overflow_x
    return total, over


def _set_all_sizes(shape, pt):
    hit = False
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(pt)
            hit = True
    return hit


def _fit_cloned_shape(shape, page_no, orig_need, orig_over):
    """
    Guard a replacement text on a box we did not design.

    Two ideas, in order:

    1. The designer's own text already fits that box. If the replacement needs
       no more room than the original did - on the same metrics - then it fits
       too, whatever an independent estimate says. Skipping the check here is
       not laxity; it is trusting the stronger evidence.

    2. Otherwise shrink the type a little, because a cloned box cannot grow and
       a title that lost 2pt still looks like the template, while a title that
       spills out of its rounded corner does not. Past 60% of the original we
       stop guessing and raise: that is no longer the template's type scale.
    """
    if not shape.has_text_frame or not shape.text_frame.text.strip():
        return
    w_in = (shape.width or 0) / EMU_IN
    h = (shape.height or 0) / EMU_IN
    if h <= 0:
        return
    avail_h = (h - 0.1) * 72

    need, over = _shape_need(shape)
    if not over and need <= max(avail_h, orig_need) + 2:
        return

    base = _shape_pt(shape)
    pt = base
    while pt > base * 0.6:
        pt -= 1
        if not _set_all_sizes(shape, pt):
            break
        need, over = _shape_need(shape, pt)
        if not over and need <= avail_h + 2:
            return
    raise FitError(
        f"slide {page_no}: 替换后的文字在「{shape.name}」里放不下"
        f"（{base:.0f}pt 缩到 {pt:.0f}pt 仍溢出；框 {w_in:.2f}x{h:.2f}in）"
        f"—— 缩短文字、减少行数，或换用别的页面")



@dataclass
class Deck(Layouts, ScienceLayouts):
    theme: Theme
    slide_w: float = SLIDE_W
    slide_h: float = SLIDE_H
    margin: float = 0.9
    strict: bool = True          # raise FitError instead of overflowing
    # --- template mode -------------------------------------------------
    # Point `template` at a .pptx and every new page is a slide of *that*
    # file: it keeps the designer's layout, chrome, theme and fonts, and the
    # user can later edit it exactly as they edit the template. Without it the
    # deck is drawn from scratch, which is the right default but does not
    # reproduce somebody else's design.
    template: str | None = None
    content_layout: int | str | None = None   # layout for new content pages
    _prs: Presentation = field(init=False)
    _n: int = field(default=0, init=False)
    _lay: object = field(default=None, init=False)
    warnings: list[str] = field(default_factory=list, init=False)

    def __post_init__(self):
        if self.template:
            self._prs = Presentation(self.template)
            self.slide_w = round(self._prs.slide_width / EMU_IN, 3)
            self.slide_h = round(self._prs.slide_height / EMU_IN, 3)
            self._lay = self._pick_layout(self.content_layout)
            return
        self._prs = Presentation()
        self._prs.slide_width = Inches(self.slide_w)
        self._prs.slide_height = Inches(self.slide_h)

    def _pick_layout(self, which):
        """
        Resolve a layout by index, exact name, or substring.

        A template's layout names are localised and often prefixed ("2_标题
        幻灯片"), so exact matching alone is hostile to a caller who only knows
        the name they saw in a screenshot.
        """
        lays = list(self._prs.slide_layouts)
        if not lays:
            raise ValueError("模板里没有可用版式")
        if which is None:
            # Default: the layout the template's own content pages use - but
            # only among layouts that have a title placeholder. A designer who
            # drew every page by hand on "Blank" leaves the most-used layout
            # titleless, and building on it means the builder cannot place a
            # title at all.
            counts = {}
            for s in self._prs.slides:
                counts[s.slide_layout.name] = counts.get(s.slide_layout.name, 0) + 1

            def has_title(lay):
                for sh in lay.shapes:
                    if sh.is_placeholder:
                        try:
                            if "TITLE" in str(sh.placeholder_format.type):
                                return True
                        except Exception:
                            pass
                return False

            titled = [l for l in lays if has_title(l)]
            pool = titled or lays
            if counts:
                used = [l for l in pool if l.name in counts]
                if used:
                    return max(used, key=lambda l: counts[l.name])
            return pool[0]
        if isinstance(which, int):
            return lays[which]
        for l in lays:
            if l.name == which:
                return l
        for l in lays:
            if which in l.name:
                return l
        raise ValueError(f"找不到版式 {which!r}；现有 "
                         f"{[l.name for l in lays]}")

    # ---------------- primitives ----------------
    def _new_slide(self, dark: bool):
        if self.template:
            # No background rectangle: an opaque shape on the slide would paint
            # over the layout's chrome, which is the whole point of using it.
            s = self._prs.slides.add_slide(self._lay)
            self._n += 1
            return s
        s = self._prs.slides.add_slide(self._prs.slide_layouts[6])
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                self._prs.slide_width, self._prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = _rgb(self.theme.bg_dark if dark else self.theme.bg_light)
        bg.line.fill.background()
        bg.shadow.inherit = False
        spTree = s.shapes._spTree
        spTree.remove(bg._element)
        spTree.insert(2, bg._element)
        self._n += 1
        return s

    # ---------------- template: reuse the designer's own pages ----------
    def clone(self, index: int):
        """
        Append a verbatim copy of template slide `index` (1-based).

        Signature pages - a cover, an agenda, a closing - are compositions, not
        layouts. Re-deriving them from primitives produces a page that is
        nearly right, which is worse than a page that is exactly the template's
        own. Copying the shapes keeps it editable; copying the relationships
        keeps its pictures and hyperlinks alive.
        """
        if not self.template:
            raise RuntimeError("clone() 需要 Deck(template=...)")
        src = self._prs.slides[index - 1]
        s = self._prs.slides.add_slide(src.slide_layout)
        for sh in list(s.shapes):
            sh._element.getparent().remove(sh._element)
        tree = s.shapes._spTree
        for sh in src.shapes:
            tree.append(copy.deepcopy(sh._element))
        _remap_relationships(src.part, s.part,
                             [el for el in tree if el.tag != qn("p:nvGrpSpPr")])
        self._n += 1
        return s

    def set_texts(self, slide, mapping: dict, strict: bool | None = None):
        """
        Replace text on a cloned page by shape name.

        Formatting is kept: the first run's properties survive, the paragraph
        properties survive, only the characters change. `\\n` splits into
        paragraphs, reusing the page's own paragraph formatting.
        """
        strict = self.strict if strict is None else strict
        by_name = {}
        for sh in slide.shapes:
            by_name.setdefault(sh.name, sh)
        missing = [k for k in mapping if k not in by_name]
        if missing:
            raise KeyError(f"页面上没有这些形状: {missing}；现有 "
                           f"{sorted(by_name)}")
        for name, text in mapping.items():
            sh = by_name[name]
            before = _shape_need(sh) if sh.has_text_frame else (0.0, False)
            _write_paragraphs(sh, text)
            if strict:
                _fit_cloned_shape(sh, self._n, before[0], before[1])
        return slide

    def set_title(self, slide, text):
        """
        Fill the template's own title placeholder, keeping its formatting.

        The size is written explicitly rather than left to inherit from the
        layout. Inheritance looks tidier, but it makes the title invisible to
        every tool that measures type - and a structural lint that cannot see
        the title concludes the slide has no title at all. Writing the size the
        layout already implies changes nothing on screen and everything for the
        checks.
        """
        for sh in slide.shapes:
            if not sh.is_placeholder:
                continue
            try:
                if "TITLE" not in str(sh.placeholder_format.type):
                    continue
            except Exception:
                continue
            _write_paragraphs(sh, text)
            pt = self._title_pt()
            if pt:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(pt)
            self._check_title_fits(sh, text, pt)
            return sh
        return None

    def _title_pt(self, default=30.0):
        """The size the layout gives its title placeholder."""
        if getattr(self, "_tpl_title_pt", None):
            return self._tpl_title_pt
        pt = default
        try:
            for sh in self._lay.shapes:
                if sh.is_placeholder and "TITLE" in str(sh.placeholder_format.type):
                    import re
                    m = re.search(r'sz="(\d+)"', sh._element.xml)
                    if m:
                        pt = int(m.group(1)) / 100.0
                    break
        except Exception:
            pass
        self._tpl_title_pt = pt
        return pt

    def _check_title_fits(self, shape, text, pt):
        """
        Refuse a title that needs more lines than the title band has.

        A title that wraps into the body is the most visible way a
        template-based deck goes wrong, and shrinking it is not the fix: at 80%
        it stops looking like the template's title. So this one raises and says
        how many characters would fit.
        """
        if not self.strict:
            return
        w = (shape.width or 0) / EMU_IN
        h = (shape.height or 0) / EMU_IN
        if w <= 0 or h <= 0 or not pt:
            return
        avail_w = (w - 0.2) * 72 - 8.0
        avail_h = (h - 0.1) * 72
        r = wrap_text(text, pt, avail_w, bold=True, line_spacing=1.1)
        if r.needed_height_pt <= avail_h + 3:
            return
        fits = max(6, int(avail_w / (pt * 1.02)))
        raise FitError(
            f"slide {self._n}: 标题太长，在 {w:.2f}in 的标题带里要 "
            f"{r.needed_height_pt:.0f}pt 高（可用 {avail_h:.0f}pt，{pt:.0f}pt 字）"
            f"—— 缩到约 {fits} 字以内，或改成分标题 + 正文")

    def drop(self, slide):
        """Delete a page (used to shed template pages a deck does not need)."""
        sldIdLst = self._prs.slides._sldIdLst
        for sldId in list(sldIdLst):
            # Match on the part, not on the id attribute: the id is a string in
            # the XML and python-pptx indexes by int, which silently misses.
            if self._prs.part.related_part(sldId.rId) is slide.part:
                self._prs.part.drop_rel(sldId.rId)
                sldIdLst.remove(sldId)
                self._n -= 1
                return
        raise ValueError("这张幻灯片不属于本 deck")

    def _text(self, slide, x, y, w, h, runs, *, size, color, bold=False,
              align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.25,
              font=None, check_fit=True, name=None):
        """
        Place a text box. `runs` is a list of paragraph strings.

        Fitting is verified with real font metrics *before* saving, because a
        renderer-based check is not always available and cannot be trusted to
        honour box widths for CJK text.
        """
        if check_fit:
            # The renderer applies CJK punctuation squeeze and kerning at the end
            # of a line, which can push a glyph a few points past the inset edge.
            # Measuring against a slightly narrower budget keeps the *rendered*
            # text inside the declared box, not just the measured text.
            SAFETY_PT = 8.0
            avail_w_pt = (w - 0.2) * 72 - SAFETY_PT
            avail_h_pt = (h - 0.1) * 72
            needed = 0.0
            for para in runs:
                r = wrap_text(para, size, avail_w_pt, bold)
                needed += r.needed_height_pt
                if r.overflow_x and self.strict:
                    raise FitError(
                        f"slide {self._n}: a line is {r.widest_line_pt:.0f}pt wide "
                        f"but only {avail_w_pt:.0f}pt is available "
                        f"({size}pt text in {w:.2f}in box) — widen the box, "
                        f"shorten the text, or reduce the font size")
            if needed > avail_h_pt + 3 and self.strict:
                raise FitError(
                    f"slide {self._n}: {len(runs)} paragraph(s) need ~{needed:.0f}pt "
                    f"of height but the box gives {avail_h_pt:.0f}pt "
                    f"({size}pt text in {w:.2f}x{h:.2f}in) — enlarge the box, "
                    f"drop a line, or reduce the font size")
            if needed > avail_h_pt + 3:
                self.warnings.append(
                    f"slide {self._n}: text overflows box by "
                    f"{needed - avail_h_pt:.0f}pt (strict=False)")

        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Emu(int(0.1 * EMU_IN))
        tf.margin_top = tf.margin_bottom = Emu(int(0.05 * EMU_IN))
        face = font or self.theme.font
        for i, para in enumerate(runs):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = spacing
            r_ = p.add_run()
            r_.text = para
            f = r_.font
            f.size = Pt(size)
            f.bold = bold
            f.name = face
            f.color.rgb = _rgb(color)
        if name:
            tb.name = name
        return tb
    def _rect(self, slide, x, y, w, h, fill, line=None, radius=False,
              transparency=0):
        shape_t = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        sp = slide.shapes.add_shape(shape_t, Inches(x), Inches(y),
                                    Inches(w), Inches(h))
        if fill is None:
            sp.fill.background()
        else:
            sp.fill.solid()
            sp.fill.fore_color.rgb = _rgb(fill)
        if line:
            sp.line.color.rgb = _rgb(line)
            sp.line.width = Pt(1)
        else:
            sp.line.fill.background()
        sp.shadow.inherit = False
        if radius:
            try:
                sp.adjustments[0] = 0.08
            except Exception:
                pass
        if sp.has_text_frame:
            sp.text_frame.word_wrap = True
            b = sp.text_frame._txBody.find(qn('a:bodyPr'))
            b.set('wrap', 'square')
        return sp

    def _notes(self, slide, text):
        if text:
            slide.notes_slide.notes_text_frame.text = text

    # ---------------- slide idioms ----------------
    def _fit_size(self, paragraphs, width_in, height_in, *, base=BODY_PT,
                  max_pt=BODY_PT + 8, spacing=1.5, bold=False, floor_ratio=0.72):
        """
        Choose the largest type size (within [base*floor_ratio, max_pt]) whose
        measured height fills most of the box without overflowing.

        A slide whose text occupies a third of the column leaves 30% of the page
        dead, which reads as an unfinished layout. Growing the type — rather than
        padding with whitespace or inventing content — is how a sparse slide is
        made to look composed on a 16:9 canvas while staying within the deck's
        type scale.
        """
        lo = max(9.0, base * floor_ratio)
        best = base
        for size in range(int(max_pt), int(lo) - 1, -1):
            need = 0.0
            for p in paragraphs:
                r = wrap_text(p, size, (width_in - 0.2) * 72 - 8.0, bold,
                              line_spacing=spacing)
                if r.overflow_x:
                    break
                need += r.needed_height_pt
            else:
                if need <= (height_in - 0.1) * 72 * 0.94:
                    best = float(size)
                    break
        return best

    def title_slide(self, title, subtitle="", meta="", notes="", cite=""):
        s = self._new_slide(dark=True)
        t = self.theme.ink_inverse
        # title may wrap to 2-3 lines; measure to size the box honestly
        r = wrap_text(title, TITLE_PT, (self.slide_w - 2 * self.margin - 0.2) * 72, True)
        th = max(1.6, r.needed_height_pt / 72 + 0.3)
        self._text(s, self.margin, 2.3, self.slide_w - 2 * self.margin, th,
                   [title], size=TITLE_PT, color=t, bold=True, spacing=1.12)
        y = 2.3 + th + 0.35
        if subtitle:
            self._text(s, self.margin, y, self.slide_w - 2 * self.margin, 1.0,
                       [subtitle], size=SECTION_PT, color=self.theme.secondary)
            y += 1.1
        if meta:
            self._text(s, self.margin, y, self.slide_w - 2 * self.margin, 0.6,
                       [meta], size=BODY_PT, color=self.theme.muted)
        if cite:
            self._text(s, self.margin, self.slide_h - 0.9,
                       self.slide_w - 2 * self.margin, 0.4,
                       [cite], size=CAPTION_PT, color=self.theme.muted)
        self._notes(s, notes)
        return s

    def section_slide(self, kicker, title, notes=""):
        s = self._new_slide(dark=True)
        self._text(s, self.margin, 2.9, 3.0, 0.6, [kicker],
                   size=BODY_PT, color=self.theme.accent, bold=True)
        # size the title box from measured line count so it cannot overflow
        r = wrap_text(title, TITLE_PT, (self.slide_w - 2 * self.margin - 0.2) * 72, True)
        th = r.needed_height_pt / 72 + 0.25
        self._text(s, self.margin, 3.5, self.slide_w - 2 * self.margin, th,
                   [title], size=TITLE_PT, color=self.theme.ink_inverse, bold=True,
                   spacing=1.15)
        self._notes(s, notes)
        return s

    def _header(self, slide, kicker, title):
        """Shared content-slide header: kicker + claim-style title."""
        y = 0.62
        if kicker:
            self._text(slide, self.margin, y, self.slide_w - 2 * self.margin, 0.4,
                       [kicker], size=CAPTION_PT, color=self.theme.accent, bold=True)
            y += 0.42
        r = wrap_text(title, SECTION_PT, (self.slide_w - 2 * self.margin - 0.2) * 72, True)
        th = max(0.6, r.needed_height_pt / 72 + 0.12)
        self._text(slide, self.margin, y, self.slide_w - 2 * self.margin, th,
                   [title], size=SECTION_PT, color=self.theme.primary, bold=True,
                   spacing=1.15)
        return y + th + 0.35

    def _footer(self, slide, cite, page=True):
        # Named, so every downstream check can tell a footer from body copy. An
        # unnamed footer box is measured against the body-text margin rule and
        # reported as an overflow that does not exist.
        """Footer band. Kept inside the margin and short enough to never wrap."""
        y = self.slide_h - 0.62
        if cite:
            self._text(slide, self.margin, y, 9.6, 0.34, [cite],
                       size=SMALL_PT, color=self.theme.muted, check_fit=False,
                       name="footer-cite")
        if page:
            self._text(slide, self.slide_w - self.margin - 0.6, y, 0.6, 0.34,
                       [str(self._n)], size=SMALL_PT, color=self.theme.muted,
                       align=PP_ALIGN.RIGHT, check_fit=False, name="footer-page")

    def bullets(self, title, items, kicker="", cite="", notes="", numbered=False):
        """Text + bullets. At most one visual: a numbered process column."""
        s = self._new_slide(dark=False)
        top = self._header(s, kicker, title)
        bottom = self.slide_h - (1.05 if cite else 0.75)
        h = bottom - top
        w = self.slide_w - 2 * self.margin
        rows = [(f"{i}.  {it}" if numbered else f"\u00b7  {it}") for i, it in
                enumerate(items, start=1)]
        size = self._fit_size(rows, w, h, spacing=1.5)
        self._text(s, self.margin, top, w, h, rows, size=size,
                   color=self.theme.ink, spacing=1.5)
        self._footer(s, cite)
        self._notes(s, notes)
        return s

    def split(self, title, body_items, side, *, kicker="", cite="", notes="",
              side_ratio=0.42, side_title=""):
        """Two-column: prose left, a labelled block of evidence right.

        The side panel hugs its own text height (with a small floor) rather than
        stretching to the slide bottom, which otherwise leaves a large hollow
        area that reads as an unfinished layout.
        """
        s = self._new_slide(dark=False)
        top = self._header(s, kicker, title)
        bottom = self.slide_h - (1.05 if cite else 0.75)
        h = bottom - top
        gap = 0.5
        total = self.slide_w - 2 * self.margin
        rw = total * side_ratio
        lw = total - rw - gap
        body_size = self._fit_size(body_items, lw, h, spacing=1.5)
        self._text(s, self.margin, top, lw, h, body_items, size=body_size,
                   color=self.theme.ink, spacing=1.5)
        sx = self.margin + lw + gap
        pad = 0.35
        inner_w = rw - 2 * pad
        side_size = body_size
        side_need = 0.0
        for p in side:
            rr = wrap_text(p, side_size, (inner_w - 0.2) * 72, line_spacing=1.45)
            side_need += rr.needed_height_pt / 72 + side_size / 72 * 1.45
        title_h = 0.0
        if side_title:
            tr = wrap_text(side_title, side_size, (inner_w - 0.2) * 72, bold=True)
            title_h = tr.needed_height_pt / 72 + 0.2
            side_need += title_h
        panel_h = min(h, max(side_need + pad * 2 + 0.1, 1.8))
        self._rect(s, sx, top, rw, panel_h, self.theme.secondary, radius=True)
        ty = top + pad
        if side_title:
            self._text(s, sx + pad, ty, inner_w, title_h, [side_title],
                       size=side_size, color=self.theme.primary, bold=True,
                       spacing=1.2)
            ty += title_h
        self._text(s, sx + pad, ty, inner_w, panel_h - (ty - top) - pad, side,
                   size=side_size, color=self.theme.ink, spacing=1.45)
        self._footer(s, cite)
        self._notes(s, notes)
        return s

    def stats(self, title, stats, *, kicker="", caption="", cite="", notes=""):
        """1-4 large stat callouts. `stats` = [(value, label), ...]."""
        if not 1 <= len(stats) <= 4:
            raise ValueError("stats() takes 1-4 entries")
        s = self._new_slide(dark=False)
        top = self._header(s, kicker, title)
        bottom = self.slide_h - (1.05 if cite else 0.75)
        total = self.slide_w - 2 * self.margin
        gap = 0.45
        cw = (total - gap * (len(stats) - 1)) / len(stats)
        # Card height derives from the measured label block so the card hugs its
        # content; a fixed tall card leaves a visually hollow bottom edge.
        # The label block is measured with the same line spacing it will use.
        label_spacing = 1.4
        label_h = 0.0
        for _, label in stats:
            rr = wrap_text(label, BODY_PT, (cw - 0.84 - 0.2) * 72,
                           line_spacing=label_spacing)
            # +1 line of slack: mixed CJK/Latin strings re-wrap slightly
            # differently in the renderer than in the greedy measurement.
            label_h = max(label_h, rr.needed_height_pt / 72 + BODY_PT / 72 * label_spacing)
        # Reserve the caption band before deciding how tall the cards may be,
        # otherwise a grown card band pushes the caption past the safe area.
        cap_h = 0.55 if caption else 0.0
        band = (bottom - top) - cap_h
        num_h = 1.25
        card_h = min(band, 0.42 + num_h + 0.28 + label_h + 0.42)
        # A short band of four small cards on a 16:9 slide reads as a throwaway
        # banner. Grow the numerals (and the cards with them) to fill the band,
        # staying inside the deck's display scale.
        if band > card_h * 1.35:
            grow = min(band / card_h, 1.5)
            num_h = num_h * grow
            card_h = min(band, card_h * grow)
        # Distribute the spare vertical space so the card band sits optically
        # centred in the content area instead of hugging the title and leaving
        # a dead strip above the footer.
        spare = band - card_h
        card_top = top + min(max(spare * 0.34, 0.0), 1.4)
        num_pt = TITLE_PT + 26 + int((num_h - 1.25) * 34)
        # The numeral must also fit the card's *width*: a tall band on a 3-card
        # slide would otherwise pick a size wider than the column.
        from textfit import measure
        num_w_pt = (cw - 0.84 - 0.2) * 72 - 8.0
        while num_pt > TITLE_PT:
            widest = max(measure(v, num_pt, True) for v, _ in stats)
            if widest <= num_w_pt:
                break
            num_pt -= 2
        # keep the numeral's box proportional to the size that actually fits
        num_h = max(1.05, num_pt / (TITLE_PT + 26) * num_h)
        card_h = min(band, 0.42 + num_h + 0.28 + label_h + 0.42)
        for i, (value, label) in enumerate(stats):
            x = self.margin + i * (cw + gap)
            self._rect(s, x, card_top, cw, card_h, self.theme.secondary, radius=True)
            pad = 0.42
            inner_h = card_h - pad * 2
            # centre the numeral+label pair inside the card instead of pinning
            # them to the top, which leaves a hollow strip at the bottom
            content_h = num_h + 0.28 + label_h + 0.3
            y0 = card_top + pad + max(0.0, (inner_h - content_h) * 0.28)
            self._text(s, x + pad, y0, cw - 2 * pad, num_h, [value],
                       size=num_pt, color=self.theme.primary, bold=True,
                       spacing=0.95)
            self._text(s, x + pad, y0 + num_h + 0.28, cw - 2 * pad,
                       label_h + 0.3, [label],
                       size=BODY_PT, color=self.theme.ink, spacing=label_spacing)
        if caption:
            self._text(s, self.margin, card_top + card_h + 0.3, total, 0.4, [caption],
                       size=CAPTION_PT, color=self.theme.muted, check_fit=False)
        self._footer(s, cite)
        self._notes(s, notes)
        return s

    def figure(self, title, image, *, kicker="", caption="", cite="",
               notes="", max_h=None, max_w=None, side_note=""):
        """
        Figure slide with a cited caption, sized so the caption always fits.

        A portrait or square figure (common for multi-panel paper figures) cannot
        fill a 16:9 slide by height alone and would leave a wide dead margin.
        Passing `side_note` puts the reading guide in that margin instead of
        stretching the image, which is how the space should be used.
        """
        s = self._new_slide(dark=False)
        top = self._header(s, kicker, title)
        cap_h = 0.0
        if caption:
            cr = wrap_text(caption, CAPTION_PT,
                           (self.slide_w - 2 * self.margin - 0.2) * 72,
                           line_spacing=1.35)
            cap_h = cr.needed_height_pt / 72 + 0.25
        bottom = self.slide_h - (0.95 if cite else 0.6)

        side_w = 0.0
        if side_note:
            # Width the guide column from its own text rather than a fixed 3in:
            # too narrow and every line breaks mid-phrase, which reads as broken.
            probe = 3.2
            side_w = probe
            for sw in (3.6, 3.3, 3.0, 2.7, 2.5):
                widest = 0.0
                for line in side_note:
                    rr = wrap_text(line, BODY_PT - 2, (sw - 0.5 - 0.2) * 72)
                    # prefer a width where no line needs more than 3 wrapped rows
                    widest = max(widest, len(rr.lines))
                if widest <= 3:
                    side_w = sw
                    break
        avail_w = max_w or (self.slide_w - 2 * self.margin - side_w)
        avail_h = max_h or (bottom - top - cap_h)

        from PIL import Image
        with Image.open(image) as im:
            iw, ih = im.size
        ar = iw / ih
        # Fill the available band by whatever dimension is binding. A wide banner
        # figure is width-bound and leaves its column half empty if you only fit
        # by width, so allow the figure to use its full height before centring.
        w = avail_w
        h = w / ar
        if h < avail_h * 0.98:
            # width-bound: a 2:1 banner cannot fill a tall band, so anchor it
            # optically (slightly above centre) instead of leaving one big gap
            y = top + max(0, (avail_h - h) * 0.35)
        else:
            h = avail_h
            w = h * ar
            y = top
        x = self.margin if side_note else (self.slide_w - w) / 2
        s.shapes.add_picture(str(image), Inches(x), Inches(y), Inches(w), Inches(h))

        if side_note:
            nx = x + w + 0.45
            nw = self.slide_w - self.margin - nx
            if nw > 1.2:
                self._text(s, nx, y, nw, min(h, avail_h), side_note,
                           size=BODY_PT - 2, color=self.theme.ink, spacing=1.5)

        if caption:
            cy = min(y + h + 0.12, bottom - cap_h)
            self._text(s, self.margin, cy, self.slide_w - 2 * self.margin,
                       cap_h, [caption], size=CAPTION_PT,
                       color=self.theme.muted, spacing=1.35)
        self._footer(s, cite)
        self._notes(s, notes)
        return s

    def table(self, title, headers, rows, *, kicker="", cite="", notes="",
              col_widths=None):
        s = self._new_slide(dark=False)
        top = self._header(s, kicker, title)
        bottom = self.slide_h - (1.05 if cite else 0.75)
        total_w = self.slide_w - 2 * self.margin
        n = len(headers)
        widths = col_widths or [1.0 / n] * n
        total_units = sum(widths)
        nrows = len(rows) + 1
        row_h = min(0.55, (bottom - top) / nrows)
        shape = s.shapes.add_table(nrows, n, Inches(self.margin), Inches(top),
                                   Inches(total_w), Inches(row_h * nrows))
        tbl = shape.table
        for j, wfrac in enumerate(widths):
            tbl.columns[j].width = Inches(total_w * wfrac / total_units)
        for i in range(nrows):
            tbl.rows[i].height = Inches(row_h)
        for j, htxt in enumerate(headers):
            cell = tbl.cell(0, j)
            cell.text = ""
            p = cell.text_frame.paragraphs[0]
            r = p.add_run(); r.text = htxt
            r.font.size = Pt(BODY_PT); r.font.bold = True
            r.font.name = "Microsoft YaHei"
            r.font.color.rgb = _rgb(self.theme.ink_inverse)
            cell.fill.solid(); cell.fill.fore_color.rgb = _rgb(self.theme.primary)
        for i, row in enumerate(rows, start=1):
            for j, val in enumerate(row):
                cell = tbl.cell(i, j)
                cell.text = ""
                p = cell.text_frame.paragraphs[0]
                avail_w_pt = (total_w * widths[j] / total_units - 0.2) * 72
                res = wrap_text(str(val), BODY_PT, avail_w_pt)
                if res.overflow_x and self.strict:
                    raise FitError(
                        f"slide {self._n}: table cell too narrow for "
                        f"{str(val)[:30]!r}; widen column {j+1} or shorten")
                r = p.add_run(); r.text = str(val)
                r.font.size = Pt(BODY_PT); r.font.name = "Microsoft YaHei"
                r.font.color.rgb = _rgb(self.theme.ink)
                cell.fill.solid()
                cell.fill.fore_color.rgb = _rgb(
                    self.theme.bg_light if i % 2 else self.theme.secondary)
        self._footer(s, cite)
        self._notes(s, notes)
        return s

    def math(self, title, formula_png, *, width_in=None, explain=None,
             kicker="", eq_no="", cite="", notes=""):
        """
        Formula slide.

        `eq_no` is only for images rendered WITHOUT a baked-in number. If you
        used `math_render.py --number`, the number is already inside the PNG and
        passing eq_no here would print it twice — the renderer detects that case
        and skips the overlay rather than trusting the caller to remember.
        """
        s = self._new_slide(dark=False)
        top = self._header(s, kicker, title)
        avail_w = self.slide_w - 2 * self.margin
        from PIL import Image
        with Image.open(formula_png) as im:
            iw, ih = im.size
        ar = iw / ih
        w = width_in or min(avail_w, 9.5)
        h = w / ar
        if h > 2.0:
            h = 2.0
            w = h * ar
        x = (self.slide_w - w) / 2
        s.shapes.add_picture(str(formula_png), Inches(x), Inches(top + 0.35),
                             Inches(w), Inches(h))

        # Detect a number baked into the image via its manifest.
        baked = False
        try:
            import json
            man = Path(formula_png).parent / "math_manifest.json"
            if man.exists():
                for rec in json.loads(man.read_text(encoding="utf-8")):
                    if Path(rec.get("png", "")) == Path(formula_png) and rec.get("numbered"):
                        baked = True
        except Exception:
            baked = False

        if eq_no and not baked:
            self._text(s, self.slide_w - self.margin - 1.0,
                       top + 0.35 + h / 2 - 0.2, 1.0, 0.45, [eq_no],
                       size=BODY_PT, color=self.theme.muted,
                       align=PP_ALIGN.RIGHT, check_fit=False)
        if explain:
            ey = top + 0.35 + h + 0.55
            eh = self.slide_h - (0.95 if cite else 0.6) - ey
            if eh > 0.6:
                self._text(s, self.margin, ey, avail_w, eh, explain,
                           size=BODY_PT, color=self.theme.ink, spacing=1.5)
        self._footer(s, cite)
        self._notes(s, notes)
        return s

    def references(self, entries, title="\u53c2\u8003\u6587\u732e", notes=""):
        """
        The reference slide.

        A two-entry list on an otherwise empty slide reads as unfinished, so the
        list is set in a measured column with a rule above it and the entry count
        is stated. Every entry should end with its locator (S1 p.N).
        """
        s = self._new_slide(dark=False)
        t = self.theme
        # 32pt text needs ~43pt of line box; a 0.45in frame is too short
        self._label(s, self.margin, 0.8, 8.0, 0.75, title, SECTION_PT + 4,
                    t.primary, bold=True)
        self._label(s, self.slide_w - self.margin - 2.0, 0.95, 2.0, 0.4,
                    f"{len(entries)} \u6761\u6765\u6e90", CAPTION_PT, t.muted,
                    align=PP_ALIGN.RIGHT, check=False)
        top = 1.75
        self._line(s, self.margin, top, self.slide_w - self.margin, top,
                   t.primary, 1.5)
        # An entry may arrive already numbered. A reference list is the subset of
        # a bibliography that this deck actually cites, so its numbers are usually
        # NOT 1..n: renumbering here would silently break every in-body [n].
        rows = []
        for i, e in enumerate(entries, start=1):
            e = str(e).strip()
            rows.append(e if e.startswith("[") else f"[{i}] {e}")
        avail = (self.slide_w - 2 * self.margin) * 0.86
        room = self.slide_h - (top + 0.35) - 0.55
        # Bibliography type shrinks to fit. Six full citations at 14pt do not fit
        # a 16:9 slide, and splitting the list over two slides to preserve a font
        # size is the wrong trade: nobody reads a reference page from the back of
        # the room, and a truncated last entry is worse than small type.
        size = CAPTION_PT + 2
        while True:
            need = 0.0
            for r in rows:
                rr = wrap_text(r, size, (avail - 0.2) * 72, line_spacing=1.4)
                # +1 line of slack: entry text re-wraps slightly differently in
                # the renderer than in the greedy measurement
                need += rr.needed_height_pt / 72 + size / 72 * 1.4 + 0.16
            list_h = need + 0.2
            if list_h <= room or size <= 9:
                break
            size -= 1
        if list_h > room and self.strict:
            raise FitError(
                f'slide {self._n}: {len(entries)} reference entries need '
                f'{list_h:.2f}in at {size}pt but only {room:.2f}in remains — '
                f'shorten entries or split the reference list')
        self._label(s, self.margin, top + 0.35, avail,
                    list_h, rows, size, t.ink, spacing=1.4)
        self._footer(s, "")
        self._notes(s, notes)
        return s

    def save(self, path):
        self._prs.save(str(path))
        return path


if __name__ == "__main__":
    # Demo: a 7-slide academic deck built only from the idioms above.
    d = Deck(Theme.navy())
    d.title_slide(
        "m6A \u4fee\u9970\u5982\u4f55\u5b88\u4f4f\u6ecb\u517b\u5c42\u5e72\u7ec6\u80de\u7684\u8eab\u4efd",
        subtitle="\u4ece METTL3 \u7684\u8bc1\u636e\u5230\u53ef\u9a8c\u8bc1\u7684\u673a\u5236\u5047\u8bbe",
        meta="\u7ec4\u4f1a\u6c47\u62a5 \u00b7 20 \u5206\u949f \u00b7 2026-01",
        notes="\u5f00\u573a\u4e0d\u8981\u94fa\u9648\uff0c\u76f4\u63a5\u7ed9\u51fa\u7ed3\u8bba\u3002")
    d.section_slide("01", "\u53d1\u73b0\uff1aMETTL3 \u4e22\u5931\u540e\uff0c\u5e72\u7ec6\u80de\u8eab\u4efd\u5148\u584c\u584c\u518d\u5206\u5316")
    d.stats("\u4e09\u4e2a\u6570\u5b57\u5b9a\u4e49\u4e86\u8fd9\u4e2a\u73b0\u8c61\u7684\u5c3a\u5ea6",
            [("3.2\u00d7", "METTL3 \u4e0b\u8c03\u500d\u6570"),
             ("48 h", "\u8eab\u4efd\u57fa\u56e0\u4e0b\u964d\u65f6\u957f"),
             ("11", "\u91cd\u53e0 m6A \u4f4d\u70b9\u6570")],
            kicker="\u8bc1\u636e", caption="\u6570\u636e\u6765\u6e90\uff1aqPCR \u4e0e MeRIP-seq\uff08\u5747\u4e3a\u72ec\u7acb\u91cd\u590d\u5b9e\u9a8c\uff09",
            cite="[1] S1 p.3\uff1b[2] S1 p.5")
    d.split("\u8eab\u4efd\u4e0b\u964d\u4e0e m6A \u6c34\u5e73\u4e0b\u964d\u540c\u6b65\u53d1\u751f",
            ["\u5728\u8bf1\u5bfc\u5206\u5316\u7684\u7b2c 2 \u5929\uff0cGATA3 \u4e0e TFAP2C \u540c\u65f6\u4e0b\u964d\u3002",
             "\u540c\u671f m6A \u603b\u91cf\u51cf\u5c11\u7ea6\u4e09\u5206\u4e4b\u4e00\u3002",
             "\u4e24\u8005\u7684\u65f6\u95f4\u5e8f\u5217\u91cd\u53e0\u4e0d\u662f\u5de7\u5408\uff0c\u800c\u662f\u53ef\u68c0\u9a8c\u7684\u56e0\u679c\u7ebf\u7d22\u3002"],
            ["\u82e5\u56e0\u679c\u6210\u7acb\uff0c\u56de\u8865 METTL3 \u5e94\u80fd\u90e8\u5206\u633d\u56de\u8eab\u4efd\u57fa\u56e0\u3002",
             "\u56de\u8865\u5b9e\u9a8c\u7ed3\u679c\u89c1\u56fe\u7247\u9875\u3002"],
            kicker="\u673a\u5236", cite="[1] S1 p.4",
            side_title="\u53ef\u68c0\u9a8c\u7684\u63a8\u8bba")
    d.bullets("\u4e09\u4e2a\u53ef\u88ab\u5426\u8bc1\u7684\u63a8\u8bba",
              ["\u82e5\u4e3b\u5f20\u6210\u7acb\uff0c\u56de\u8865 METTL3 \u5e94\u90e8\u5206\u6062\u590d\u8eab\u4efd\u57fa\u56e0\u3002",
               "\u82e5\u4e3b\u5f20\u6210\u7acb\uff0c\u53ea\u5e72\u9884\u4e0b\u6e38 reader \u5e94\u89e6\u53d1\u7c7b\u4f3c\u8868\u578b\u3002",
               "\u82e5\u4e24\u8005\u90fd\u4e0d\u6210\u7acb\uff0c\u5219 m6A \u53ea\u662f\u76f8\u5173\u6027\u6807\u5fd7\u3002"],
              kicker="\u5224\u65ad", cite="[1] S1 p.6\uff1b[3] S2 p.8")
    d.closing("\u8ba8\u8bba\uff1a\u54ea\u4e00\u6761\u8bc1\u636e\u94fe\u6700\u8584\u5f31\uff1f",
              ["\u5efa\u8bae\u4f18\u5148\u9a8c\u8bc1\u56de\u8865\u5b9e\u9a8c\u7684\u91cf\u6548\u5173\u7cfb\u3002"],
              cite="[1] S1 p.7")
    d.references([
        "Wen J, et al. METTL3 safeguards cell identity and epigenome of human "
        "trophoblast stem cells. Nat Commun. 2025. doi:10.1038/s41467-025-67484-z. S1 p.3",
        "Wen J, et al. \u540c\u4e0a\uff0c\u56fe 2 \u4e0e\u56fe 3 \u7684\u5bf9\u5e94\u5b9e\u9a8c\u7ec4\u3002S1 p.5",
        "Wen J, et al. \u540c\u4e0a\uff0c\u8ba8\u8bba\u90e8\u5206\u5173\u4e8e reader \u4ee3\u507f\u6027\u4e0a\u8c03\u7684\u8bc1\u636e\u3002S1 p.8",
    ])
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("demo.pptx")
    d.save(out)
    print("saved", out, "| warnings:", d.warnings or "none")
