"""
template_build.py - build a deck *inside* somebody else's design.

The rule this file follows
--------------------------
Copy, do not imitate. A template is a .pptx, and everything that makes it look
like itself is already in the file: the palette is in the theme, the chrome is
in the layout, the finished compositions are slides. All three are reused
here by reading them, so none of them has to be guessed at.

What that means per page type
-----------------------------
  cover / agenda / closing   cloned verbatim, then their text replaced. These
                             are compositions - a green field, a rule, a stack
                             of small type in three weights - and re-deriving
                             one from rectangle primitives gives a page that is
                             almost right, which reads as broken.
  content pages              built on the template's own content layout, so
                             they inherit its title bar, number badge and
                             footer band, then filled with blocks that follow
                             the template's own idioms (white card, green
                             hairline, solid green header strip).
  everything numeric         checked with the same font-metric fitter as the
                             rest of the skill, because inheriting somebody's
                             design is no excuse for overflowing their boxes.

Usage
-----
    from template_build import TemplatedDeck, theme_from_profile
    import json

    prof = json.load(open("work/template/template_profile.json"))
    d = TemplatedDeck("work/template/xxx-clean.pptx", prof)
    d.cover(title_en=..., title_cn=..., dept=..., presenter=..., date=...)
    d.agenda(["基本信息", "研究背景", ...])
    d.cards("研究背景", section=2, items=[...])
    d.closing()
    d.save("out.pptx")
"""
from __future__ import annotations

import copy
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_deck import (  # noqa: E402
    BODY_PT, CAPTION_PT, SECTION_PT, SMALL_PT, TITLE_PT, Deck, FitError,
    Theme, _rgb,
    _fit_cloned_shape, _shape_need, _write_paragraphs,
)
from textfit import measure, wrap_text  # noqa: E402

EMU_IN = 914400

# Type scale lifted from the template itself (not invented): the cover sets
# 40pt English over 32pt Chinese, the agenda 54pt heading over 36pt numerals
# over 28pt entries, body copy 16pt, captions 14pt.
COVER_EN_PT = 40
COVER_CN_PT = 32
AGENDA_HEAD_PT = 54
BODY_SMALL_PT = 14
CARD_HEAD_PT = 22


def _contrast(a: str, b: str) -> float:
    def lum(h):
        h = h.lstrip("#")
        c = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
        c = [(v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4) for v in c]
        return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]
    la, lb = lum(a), lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def TL_hex_hue(hexstr: str) -> float:
    """Hue of a hex colour, so a ladder can be grown at the brand's own angle."""
    import colorsys
    h = hexstr.lstrip("#")
    r, g, b = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    hh, _l, _s = colorsys.rgb_to_hls(r, g, b)
    return hh * 360.0


def TL_hex_mix(a: str, b: str, t: float) -> str:
    a, b = a.lstrip("#"), b.lstrip("#")
    out = []
    for i in (0, 2, 4):
        out.append(round(int(a[i:i + 2], 16) * (1 - t) + int(b[i:i + 2], 16) * t))
    return "%02X%02X%02X" % tuple(out)


def theme_from_profile(prof: dict) -> Theme:
    """
    Turn the template's colour scheme into our Theme object.

    Only slots carry over; the *pairings* are then re-verified, because a
    template author picks colours to look good and this skill also insists the
    body copy clears 7:1. A template that fails is reported rather than
    silently repainted, so the caller decides which side to give way.
    """
    c = prof["theme"]["colors"]
    bg = "#" + c.get("lt1", "FFFFFF")
    brand = "#" + c.get("accent1", "006334")
    fonts = prof["theme"].get("fonts", {})
    face = (fonts.get("minorFont", {}).get("latin")
            or fonts.get("majorFont", {}).get("latin") or "Microsoft YaHei")

    # The template fixes two things: its canvas and its brand colour. Everything
    # else - how pale a card is, how far the wash sits from the page - is not a
    # decision the template makes, it is a decision about hierarchy. So derive
    # the rest of the ladder at the brand's own hue, instead of inventing three
    # hexes that happen to look plausible next to it.
    import theme_library as TL
    hue = TL_hex_hue(brand)
    ladder = TL.build("graphite", "duotone")      # placeholder, replaced below
    ladder.hue = hue
    ladder.family = "template"
    ladder.canvas = bg.lstrip("#")
    ladder.accent = c.get("accent1", "006334")
    ladder.wash = TL._level(hue, "wash")
    ladder.panel = TL._level(hue, "panel")
    ladder.veil = TL._level(hue, "veil")
    ladder.ink = TL._level(hue, "ink", sat_scale=0.35)
    ladder.muted = TL._level(hue, "muted", sat_scale=0.45)
    # A tinted canvas gets cards that are lighter than it; a white canvas has
    # nothing lighter to be, so its cards take the pale wash instead. Applying
    # the "cards must be lighter" rule to a white page yields invisible cards.
    if TL._lum(ladder.canvas) > 0.955:
        ladder.card = TL._level(hue, "wash")
    else:
        ladder.card = TL._level(hue, "card")

    t = Theme(
        name=f"template:{prof.get('source', '?')}",
        bg_dark=brand.lstrip("#"),
        bg_light=bg.lstrip("#"),
        primary=brand.lstrip("#"),
        secondary=ladder.card,
        accent=c.get("accent2", ladder.accent).lstrip("#"),
        ink=ladder.ink,
        ink_inverse="FFFFFF",
        muted=ladder.muted,
        muted_inverse=TL_hex_mix(brand, "#FFFFFF", 0.72),
        accent_inverse="FFFFFF",
        tint=ladder.wash,
        card=ladder.card,
        panel=ladder.panel,
        veil=ladder.veil,
        font=face,
    )
    checks = [
        # (标签, 前景, 背景, 硬下限, 期望值)
        # The hard floor is WCAG's rule *for that text size*, not one number for
        # everything. An earlier version demanded 7:1 for the cover title and so
        # rejected every ordinary blue brand: 7:1 is the AAA body-text rule, and
        # a 40pt bold headline is "large text", where 3:1 is the standard. The
        # preferred value is still reported, so the distance between "legal" and
        # "comfortable" stays visible instead of being collapsed to a pass.
        ("正文 / 画布", t.ink, t.bg_light, 7.0, 7.0),
        ("注释 / 画布", t.muted, t.bg_light, 4.5, 4.5),
        ("标题 / 画布", t.primary, t.bg_light, 3.0, 4.5),
        ("强调 / 画布", t.accent, t.bg_light, 3.0, 4.5),
        ("反白 / 品牌底", t.ink_inverse, t.bg_dark, 3.0, 4.5),
    ]
    bad, soft = [], []
    for label, fg, bgn, need, want in checks:
        got = _contrast(fg, bgn)
        if got < need:
            bad.append(f"{label} {got:.1f}:1 < {need}:1")
        elif got < want:
            soft.append(f"{label} {got:.1f}:1（合规，未达 {want}:1）")
    if bad:
        raise FitError(
            "模板配色达不到可读性下限：" + "；".join(bad)
            + "。若这是品牌色，请显式换用同色系更深的一档，不要关掉检查。")
    t.notes = soft
    return t


@dataclass
class TemplatedDeck(Deck):
    """
    A Deck that draws on a given .pptx instead of a blank canvas.

    `profile` is what template_ingest.py wrote: canvas, theme, layout geometry
    and the page catalogue. It is passed in rather than re-derived so that a
    build is reproducible from two files and nothing else.
    """
    profile: dict = field(default_factory=dict)
    keep_pages: tuple = ()      # template page numbers to keep as reference
    _orig: int = field(default=0, init=False)

    @classmethod
    def open(cls, template: str, profile) -> "TemplatedDeck":
        """
        The normal way in: a cleaned .pptx plus the profile written beside it.

        Builds the palette from the template rather than from a constant, so a
        green template stays green and a blue one stays blue without anyone
        editing this file.
        """
        if isinstance(profile, (str, Path)):
            profile = json.loads(Path(profile).read_text(encoding="utf-8"))
        theme = theme_from_profile(profile)
        return cls(theme=theme, template=str(template), profile=profile)


    def __post_init__(self):
        super().__post_init__()
        self._orig = len(self._prs.slides)
        lay = next((l for l in self.profile.get("layouts", [])
                    if l["name"] == self._lay.name), None)
        self.top = (lay or {}).get("content_top", 1.31)
        self.left = 0.56
        self.right = 12.83
        self.bottom = 6.72
        self._slots = {}

    # ------------------------------------------------------------------
    # geometry
    # ------------------------------------------------------------------
    @property
    def width(self) -> float:
        return self.right - self.left

    def _shift(self, shape, dy):
        shape.top = Emu(int(shape.top + dy * EMU_IN))

    def _badge(self, slide, no):
        """
        The section numeral that sits inside the roundel the layout draws.

        The circle belongs to the layout; only the number belongs to the page,
        which is why the template's own pages put a bare text box there. Its
        colour is read from the template rather than chosen: on a green disc the
        numeral is white, and picking the brand colour instead produces dark
        green on dark green - the one defect a reader notices instantly and
        cannot articulate.
        """
        if no is None:
            return
        colour = self.profile.get("badge_color") or self.theme.ink_inverse
        self._text(slide, 0.43, 0.42, 0.72, 0.62, [str(no)],
                   size=SECTION_PT, color=colour, bold=True,
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0,
                   name="page_number_badge")

    def _footer(self, slide, cite="", page=True):
        if cite:
            self._text(slide, self.left, 7.04, 8.6, 0.3, [cite],
                       size=SMALL_PT, color=self.theme.muted, spacing=1.0,
                       name="cite_footer")
        if page:
            self._text(slide, self.right - 1.3, 7.04, 1.3, 0.3, [str(self._n)],
                       size=SMALL_PT, color=self.theme.muted,
                       align=PP_ALIGN.RIGHT, spacing=1.0, name="page_footer")

    def _kicker(self, slide, x, y, label, w=2.2, color=None):
        """
        The small label the template builds as a thin vertical bar plus a word.

        Reproduced rather than replaced by a bullet, because it is the one
        ornament that appears on nearly every content page of the template.
        """
        color = color or self.theme.primary
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x),
                                     Inches(y + 0.08), Inches(0.055),
                                     Inches(0.26))
        bar.fill.solid()
        bar.fill.fore_color.rgb = _rgb(color)
        bar.line.fill.background()
        bar.shadow.inherit = False
        self._text(slide, x + 0.13, y, w, 0.4, [label], size=BODY_PT,
                   color=color, bold=True, spacing=1.0, name="kicker")

    def _card(self, slide, x, y, w, h, *, fill=None, line=True, radius=0.055):
        sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x),
                                    Inches(y), Inches(w), Inches(h))
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
        if fill is None:
            sp.fill.background()
        else:
            sp.fill.solid()
            sp.fill.fore_color.rgb = _rgb(fill)
        if line:
            sp.line.color.rgb = _rgb(self.theme.primary)
            sp.line.width = Pt(1.25)
        else:
            sp.line.fill.background()
        sp.shadow.inherit = False
        return sp

    def _rule(self, slide, x, y, w, color=None, weight=1.0):
        ln = slide.shapes.add_connector(1, Inches(x), Inches(y),
                                        Inches(x + w), Inches(y))
        ln.line.color.rgb = _rgb(color or self.theme.primary)
        ln.line.width = Pt(weight)
        return ln

    # ------------------------------------------------------------------
    # content pages (built on the template's own layout)
    # ------------------------------------------------------------------
    def content_page(self, title, section=None, cite="", page=True):
        s = self._new_slide(False)
        if self.set_title(s, title) is None:
            # A layout with no title placeholder. Some designers draw every page
            # by hand on a blank layout, so inheriting "their" title is not
            # possible; drawing one at the template's own title position is still
            # better than refusing to make the page.
            self._fallback_title(s, title)
        self._badge(s, section)
        self._footer(s, cite, page)
        return s

    def _fallback_title(self, slide, title):
        box = self._title_box()
        pt = self._title_pt()
        self._text(slide, box[0], box[1], box[2], box[3], [title], size=pt,
                   color=self.theme.primary, bold=True, spacing=1.1,
                   name="title_fallback")
        w, h = box[2], box[3]
        avail_w = (w - 0.2) * 72 - 8.0
        r = wrap_text(title, pt, avail_w, bold=True, line_spacing=1.1)
        if r.needed_height_pt > (h - 0.1) * 72 + 3 and self.strict:
            fits = max(6, int(avail_w / (pt * 1.02)))
            raise FitError(
                f"slide {self._n}: 标题太长，在 {w:.2f}in 的标题带里放不下"
                f"—— 缩到约 {fits} 字以内")

    def _title_box(self, default=(1.32, 0.38, 10.58, 0.72)):
        """Where the template puts a title, from the profile if it knows."""
        lay = next((l for l in self.profile.get("layouts", [])
                    if l["name"] == self._lay.name), None)
        if lay and lay.get("title_box"):
            return tuple(lay["title_box"])
        for l in self.profile.get("layouts", []):
            if l.get("title_box"):
                return tuple(l["title_box"])
        return default

    def cards(self, title, items, *, section=None, cols=None, variant="plain",
              headers=None, cite="", notes=""):
        """
        Rows of white cards. The workhorse, because most of a literature report
        is "here are two/three/four things, and what each one means".
        """
        s = self.content_page(title, section, cite)
        n = len(items)
        cols = cols or (1 if n == 1 else 2 if n <= 4 else 3)
        rows = (n + cols - 1) // cols
        gap = 0.32
        cw = (self.width - gap * (cols - 1)) / cols
        ch = (self.bottom - self.top - gap * (rows - 1)) / rows

        for i, item in enumerate(items):
            r, c = divmod(i, cols)
            x = self.left + c * (cw + gap)
            y = self.top + r * (ch + gap)
            head = item[0] if isinstance(item, (tuple, list)) else None
            body = item[1] if isinstance(item, (tuple, list)) else item
            if variant in ("strip", "bar") and head:
                self._card(s, x, y, cw, ch, fill=self.theme.secondary)
                strip_h = 0.72
                # A pale header band, not a saturated block. The original
                # template fills this bar with the brand colour; at full
                # strength three of them across a page become three dark
                # rectangles competing with the text, which is the complaint
                # that started this. The band keeps the structure and drops the
                # weight: a tint behind, an accent rule under, dark type on top.
                band = s.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                    Inches(cw), Inches(strip_h))
                try:
                    band.adjustments[0] = 0.14
                except Exception:
                    pass
                band.fill.solid()
                band.fill.fore_color.rgb = _rgb(self.theme.tint)
                band.line.fill.background()
                band.shadow.inherit = False
                self._rule(s, x + 0.02, y + strip_h, cw - 0.04,
                           self.theme.accent, 2.0)
                self._text(s, x + 0.24, y, cw - 0.48, strip_h, [head],
                           size=CARD_HEAD_PT, color=self.theme.accent,
                           bold=True, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
                body_top = y + strip_h + 0.22
                body_h = ch - strip_h - 0.3
            elif head:
                self._card(s, x, y, cw, ch, fill=self.theme.secondary)
                self._text(s, x + 0.24, y + 0.18, cw - 0.48, 0.44, [head],
                           size=CARD_HEAD_PT, color=self.theme.accent,
                           bold=True, spacing=1.0)
                body_top = y + 0.78
                body_h = ch - 0.95
            else:
                self._card(s, x, y, cw, ch, fill=self.theme.secondary)
                body_top = y + 0.22
                body_h = ch - 0.44
            lines = body if isinstance(body, (list, tuple)) else [body]
            self._text(s, x + 0.24, body_top, cw - 0.48, body_h, list(lines),
                       size=BODY_PT, color=self.theme.ink, spacing=1.35)
        self._notes(s, notes)
        return s

    def flow(self, title, steps, *, section=None, cite="", notes="",
             numbered=True):
        """
        A horizontal chain of chevrons: 3-6 steps, head over body.

        Chevrons rather than boxes because a chain has to read as a direction;
        boxes read as a list. Every step is the same shape, the first included,
        because a flat-sided first step next to pointed ones looks like a broken
        arrow rather than a beginning.
        """
        s = self.content_page(title, section, cite)
        n = len(steps)
        if not 2 <= n <= 6:
            raise FitError(f"流程支持 2-6 步，收到 {n} 步")
        gap = 0.10
        cw = (self.width - gap * (n - 1)) / n
        ch = min(1.62, (self.bottom - self.top) * 0.46)
        y = self.top + 0.20
        for i, step in enumerate(steps):
            head = step[0]
            body = step[1] if len(step) > 1 else ""
            x = self.left + i * (cw + gap)
            shape = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y),
                                       Inches(cw), Inches(ch))
            shape.adjustments[0] = 0.22
            # Alternating depth along the chain: the same pale tint six times in
            # a row is a row of identical tabs, and the eye stops reading it as
            # a sequence.
            shape.fill.solid()
            shape.fill.fore_color.rgb = _rgb(
                self.theme.tint if i % 2 == 0 else self.theme.panel)
            shape.line.fill.background()
            shape.shadow.inherit = False
            if numbered:
                self._text(s, x + 0.36, y + 0.14, 0.5, 0.32, [str(i + 1)],
                           size=BODY_SMALL_PT, color=self.theme.accent,
                           bold=True, spacing=1.0)
            self._text(s, x + 0.36, y + 0.44, cw - 0.66, 0.38, [head],
                       size=BODY_PT, color=self.theme.ink, bold=True,
                       spacing=1.0)
            if body:
                self._text(s, x + 0.36, y + 0.84, cw - 0.66, ch - 0.90, [body],
                           size=BODY_SMALL_PT, color=self.theme.muted,
                           spacing=1.25)
        self._notes(s, notes)
        return s

    def stats(self, title, figures, *, section=None, cite="", notes="",
              caption=""):
        """
        Two to four numbers with captions.

        The numeral carries the accent and the caption stays quiet. A numeral in
        the body colour competes with the two beside it, and three equal voices
        is the thing this whole palette is built to avoid.
        """
        s = self.content_page(title, section, cite)
        n = len(figures)
        if not 1 <= n <= 4:
            raise FitError(f"数字排支持 1-4 个，收到 {n} 个")
        cap_h = 0.5 if caption else 0.0
        band_top = self.top + 0.55
        band_h = min(2.6, self.bottom - band_top - cap_h)
        gap = 0.34
        cw = (self.width - gap * (n - 1)) / n

        # One size for the whole row. Shrinking each numeral to fit its own cell
        # gives 54pt beside 48pt beside 40pt - three sizes in one row read as a
        # mistake, not as emphasis. Take the smallest size that fits every cell.
        avail = (cw - 0.2) * 72 - 8.0
        row_size = 54
        for value, _label in figures:
            sz = 54
            r = wrap_text(str(value), sz, avail, True)
            while (r.overflow_x or len(r.lines) > 1) and sz > 24:
                sz -= 2
                r = wrap_text(str(value), sz, avail, True)
            row_size = min(row_size, sz)

        for i, (value, label) in enumerate(figures):
            x = self.left + i * (cw + gap)
            if i:
                bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                         Inches(x - gap / 2),
                                         Inches(band_top + 0.15),
                                         Inches(0.012), Inches(band_h - 0.35))
                bar.fill.solid()
                bar.fill.fore_color.rgb = _rgb(self.theme.panel)
                bar.line.fill.background()
                bar.shadow.inherit = False
            # The budget must match _text exactly (0.2in insets plus an 8pt
            # safety margin) and must require a single line: a numeral is one
            # mark, and two lines of 54pt type is a headline that lost its ruler.
            size = row_size
            self._text(s, x, band_top, cw, 1.15, [str(value)], size=size,
                       color=self.theme.accent, bold=True, spacing=1.0)
            self._text(s, x, band_top + 1.18, cw - 0.1, band_h - 1.18, [label],
                       size=BODY_SMALL_PT, color=self.theme.muted, spacing=1.35)
        if caption:
            self._text(s, self.left, self.bottom - cap_h, self.width, cap_h,
                       [caption], size=CAPTION_PT, color=self.theme.muted,
                       align=PP_ALIGN.CENTER, spacing=1.2)
        self._notes(s, notes)
        return s

    def statement(self, lines, *, section=None, cite="", notes="",
                  lead=None, size=None):
        """
        One to three sentences, set large, with air around them.

        The page that ends an argument. The template has no equivalent, but it
        does have a 72pt closing line, so the idiom is its own: brand rule, big
        type, nothing else. Setting the three sentences that must be remembered
        as a bullet list is how they stop being memorable.
        """
        s = self._new_slide(False)
        self._badge(s, section)
        self._footer(s, cite)
        top = self.top + 0.45
        n = len(lines)
        sz = size or (40 if n == 1 else 34 if n == 2 else 30)
        # Same width budget _text will use: (w - 0.2in insets) * 72 - 8pt.
        # Measuring with a looser budget leaves a line that fits the estimate and
        # wraps in reality, which _text then correctly refuses.
        avail = (self.width - 1.1 - 0.2) * 72 - 8.0
        while True:
            rs = [wrap_text(str(t), sz, avail, False,
                            line_spacing=1.35) for t in lines]
            need = sum(r.needed_height_pt for r in rs) / 72 + 0.42 * (n - 1)
            if need <= self.bottom - top or sz <= 18:
                break
            sz -= 2
        if need > self.bottom - top and self.strict:
            raise FitError(
                f"slide {self._n}: {n} 句话在 {self.bottom - top:.1f}in 里放不下"
                f"（缩到 {sz}pt 仍需 {need:.2f}in）")
        self._rule(s, self.left, top - 0.18, 1.6, self.theme.accent, 1.5)
        y = top
        for i, (r, t) in enumerate(zip(rs, lines)):
            h = r.needed_height_pt / 72
            emph = (lead is not None and i == lead)
            self._text(s, self.left + 0.05, y, self.width - 1.1, h + 0.16,
                       [str(t)], size=sz,
                       color=self.theme.accent if emph else self.theme.ink,
                       bold=emph, spacing=1.35)
            y += h + 0.42
        self._notes(s, notes)
        return s

    def bullets(self, title, items, *, section=None, kicker=None, cite="",
                notes="", body_size=BODY_PT):
        """
        A single wide card of running text, the template's block for argument
        rather than enumeration.

        Heights are allocated from what each item actually needs, not split
        evenly: a list whose first line is four times longer than the rest looks
        broken if every line gets an equal band, and overflows if every line gets
        the same fixed one.
        """
        s = self.content_page(title, section, cite)
        top = self.top
        if kicker:
            self._kicker(s, self.left, top, kicker)
            top += 0.55
        h = self.bottom - top
        self._card(s, self.left, top, self.width, h,
                   fill=self.theme.secondary)
        inner_x = self.left + 0.32
        inner_w = self.width - 0.64
        avail_in = h - 0.6

        # Shrink until the block fits, then give each item its measured share.
        size = body_size
        avail_pt = avail_in * 72
        while True:
            needs = [wrap_text(str(t), size, (inner_w - 0.24 - 0.2) * 72 - 8.0,
                               line_spacing=1.3).needed_height_pt for t in items]
            if sum(needs) + 6.0 * (len(items) - 1) <= avail_pt or size <= 11:
                break
            size -= 1
        if sum(needs) + 6.0 * (len(items) - 1) > avail_pt and self.strict:
            raise FitError(
                f"slide {self._n}: {len(items)} 条要点在 {avail_in:.1f}in 里放不下"
                f"（缩到 {size}pt 仍需 "
                f"{(sum(needs) + 6.0 * (len(items) - 1)) / 72:.2f}in）"
                f"—— 减条数或删字")

        total_pt = sum(needs) + 6.0 * (len(items) - 1)
        slack = max(0.0, avail_pt - total_pt)
        pad = min(slack / max(len(items), 1), 14.0)     # spread leftover evenly

        y = top + 0.3
        for need, line in zip(needs, items):
            dot = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(inner_x),
                                     Inches(y + 0.11), Inches(0.1), Inches(0.1))
            dot.fill.solid()
            dot.fill.fore_color.rgb = _rgb(self.theme.accent)
            dot.line.fill.background()
            dot.shadow.inherit = False
            box_h = need / 72.0 + 0.08
            self._text(s, inner_x + 0.24, y, inner_w - 0.24, box_h,
                       [str(line)], size=size, color=self.theme.ink, spacing=1.3)
            y += box_h + (pad / 72.0)
        self._notes(s, notes)
        return s

    def figure(self, title, image, *, caption="", section=None, cite="",
               notes="", width=None, max_h=None):
        s = self.content_page(title, section, cite)
        cap_h = 0.5 if caption else 0.0
        box_top = self.top
        box_bottom = self.bottom - cap_h
        avail_h = box_bottom - box_top - 0.2
        avail_w = self.width - 0.4
        max_h = max_h or avail_h
        pic = s.shapes.add_picture(str(image), Inches(self.left + 0.2),
                                   Inches(box_top + 0.1),
                                   width=Inches(width) if width else None)
        scale = min(1.0, avail_h / (pic.height / EMU_IN),
                    avail_w / (pic.width / EMU_IN), max_h / (pic.height / EMU_IN))
        if scale < 1.0:
            pic.width = Emu(int(pic.width * scale))
            pic.height = Emu(int(pic.height * scale))
        pic.left = Emu(int(((self.slide_w - pic.width / EMU_IN) / 2) * EMU_IN))
        if caption:
            self._text(s, self.left, box_bottom + 0.04, self.width,
                       cap_h + 0.12, [caption], size=CAPTION_PT,
                       color=self.theme.muted, align=PP_ALIGN.CENTER,
                       spacing=1.15, name="figure_caption")
        self._notes(s, notes)
        return s

    def table(self, title, headers, rows, *, section=None, cite="", notes="",
              widths=None):
        s = self.content_page(title, section, cite)
        nrows, ncols = len(rows) + 1, len(headers)
        top = self.top
        h = self.bottom - top - 0.1
        row_h = min(0.66, h / nrows)
        gfx = s.shapes.add_table(nrows, ncols, Inches(self.left), Inches(top),
                                 Inches(self.width), Inches(row_h * nrows))
        tbl = gfx.table
        frac = widths or [1.0] * ncols
        total = sum(frac)
        for j, f in enumerate(frac):
            tbl.columns[j].width = Inches(self.width * f / total)
        for i in range(nrows):
            tbl.rows[i].height = Inches(row_h)

        def cell_fill(cell, hexcol):
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(hexcol)

        for j, htxt in enumerate(headers):
            cell = tbl.cell(0, j)
            cell.text = ""
            # A pale header band. Filling it with the brand colour turns the
            # table into a dark block that outweighs the data it describes.
            cell_fill(cell, self.theme.tint)
            cell.margin_left = cell.margin_right = Emu(int(0.14 * EMU_IN))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = str(htxt)
            r.font.size = Pt(BODY_SMALL_PT + 2)
            r.font.bold = True
            r.font.name = self.theme.font
            r.font.color.rgb = _rgb(self.theme.accent)

        for i, row in enumerate(rows, start=1):
            for j, val in enumerate(row):
                cell = tbl.cell(i, j)
                cell.text = ""
                cell_fill(cell, self.theme.secondary if i % 2
                          else self.theme.bg_light)
                cell.margin_left = cell.margin_right = Emu(int(0.14 * EMU_IN))
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                avail = (tbl.columns[j].width / EMU_IN - 0.32) * 72
                sz = BODY_SMALL_PT
                while sz > 9 and measure(str(val), sz, j == 0) > avail:
                    sz -= 1
                if measure(str(val), sz, j == 0) > avail and self.strict:
                    raise FitError(
                        f"slide {self._n}: 表格第 {j+1} 列放不下 {str(val)[:20]!r}")
                p = cell.text_frame.paragraphs[0]
                r = p.add_run()
                r.text = str(val)
                r.font.size = Pt(sz)
                r.font.bold = (j == 0)
                r.font.name = self.theme.font
                r.font.color.rgb = _rgb(self.theme.primary if j == 0
                                        else self.theme.ink)
        self._notes(s, notes)
        return s

    def references(self, entries, *, title="参考文献", section=None, notes="",
                   size=BODY_SMALL_PT):
        """
        The reference list, as its own page type rather than a bullet list.

        Two things matter and both are structural: each entry is its own
        paragraph, so a parser can split on the number, and the page carries no
        footer citation - a footer on the reference slide reads as another entry
        and desynchronises the [n] audit.
        """
        s = self.content_page(title, section, "", page=False)
        self._card(s, self.left, self.top, self.width,
                   self.bottom - self.top, fill=self.theme.secondary)
        inner_x = self.left + 0.34
        inner_w = self.width - 0.68
        avail_h = (self.bottom - self.top) - 0.44

        n = max(len(entries), 1)
        sz = size

        def needs_at(pt):
            return [wrap_text(str(e), pt, (inner_w - 0.2) * 72 - 8.0,
                              line_spacing=1.4).needed_height_pt for e in entries]

        needs = needs_at(sz)
        while sz > 9 and sum(needs) + 6.0 * (n - 1) > avail_h * 72:
            sz -= 1
            needs = needs_at(sz)
        if sum(needs) + 6.0 * (n - 1) > avail_h * 72 and self.strict:
            raise FitError(
                f"slide {self._n}: 参考文献 {n} 条在 {avail_h:.1f}in 里放不下"
                f"（缩到 {sz}pt 仍需 "
                f"{(sum(needs) + 6.0 * (n - 1)) / 72:.2f}in）")

        y = self.top + 0.24
        for need, e in zip(needs, entries):
            box_h = need / 72.0 + 0.06
            self._text(s, inner_x, y, inner_w, box_h, [str(e)], size=sz,
                       color=self.theme.ink, spacing=1.4)
            y += box_h + 6.0 / 72.0
        self._notes(s, notes)
        return s

    # ------------------------------------------------------------------
    # cloned signature pages
    # ------------------------------------------------------------------
    def cover(self, *, title_en="", title_cn="", eyebrow="", dept="", group="",
              presenter="", date="", page_index=None):
        """
        Clone the template's cover and refill it.

        Slot names are read from the profile so a different template only needs
        a different map, not a different method.
        """
        idx, slots, widen = self._sig("cover", page_index, {
            "eyebrow": "文本框 1", "title_en": "文本框 9", "title_cn": "文本框 7",
            "dept": "文本框 4", "group": "文本框 14",
            "presenter": "文本框 10", "date": "文本框 11",
        })
        s = self.clone(idx)
        self._apply_widen(s, widen)
        self._fill(s, slots, dict(eyebrow=eyebrow, title_en=title_en,
                                  title_cn=title_cn, dept=dept, group=group,
                                  presenter=presenter, date=date))
        return s

    def closing(self, text=None, *, dept="", group="", presenter="", date="",
                page_index=None):
        idx, slots, widen = self._sig("closing", page_index, {
            "text": "文本框 7", "dept": "文本框 4", "group": "文本框 14",
            "presenter": "文本框 10", "date": "文本框 11",
        })
        s = self.clone(idx)
        self._apply_widen(s, widen)
        self._fill(s, slots, dict(text=text, dept=dept, group=group,
                                  presenter=presenter, date=date))
        return s

    def agenda(self, items, *, heading="目录", page_index=None, notes="",
               pitch=None):
        """
        Clone the template's agenda and resize it to `len(items)`.

        The template ships six rows; a five-part report must not keep a sixth
        empty row, and the rows must redistribute or the page reads as if a
        section went missing. Rows keep the template's own pitch and centre on
        the heading block, so three items look deliberate rather than stretched.
        """
        idx, slots, widen = self._sig("agenda", page_index, {})
        s = self.clone(idx)
        self._apply_widen(s, widen)

        by_name = {sh.name: sh for sh in s.shapes}
        head_name = slots.get("heading")
        if head_name in by_name and heading:
            _write_paragraphs(by_name[head_name], heading)

        rows = self._agenda_rows(s)
        n = len(items)
        if n > len(rows):
            raise FitError(f"这一页目录只排得下 {len(rows)} 条，收到 {n} 条")

        step = pitch or (rows[1][0] - rows[0][0] if len(rows) > 1 else 1.09)
        centre = (rows[0][0] + rows[-1][0]) / 2 if len(rows) > 1 else rows[0][0]
        first = centre - step * (n - 1) / 2
        for i, (base, group, num, label) in enumerate(rows):
            if i >= n:
                for g in group:
                    g._element.getparent().remove(g._element)
                continue
            target = first + i * step
            for g in group:
                self._shift(g, target - base)
            _write_paragraphs(num, f"{i + 1:02d}")
            _write_paragraphs(label, items[i])
        self._notes(s, notes)
        return s

    def _agenda_rows(self, slide):
        """
        Find the agenda's rows by geometry: numeral, label, optional rule.

        An earlier version hard-coded one template's shape names
        ("文本框 5", "直接连接符 29"). That worked exactly once. The structure is
        recognisable without names - a short all-digits text box is a numeral,
        the text box to its right on the same line is its label, a connector on
        that line is its rule - so read it that way and any agenda page works.
        """
        import re as _re
        rows = []
        numerals = []
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text.strip()
            if _re.fullmatch(r"\d{1,2}", t):
                numerals.append((sh, (sh.top or 0) / EMU_IN))
        numerals.sort(key=lambda x: x[1])
        if not numerals:
            raise FitError(
                "目录页上找不到纯数字序号（01/02/…）；"
                "请在 profile 的 pages.agenda.rows 里显式指定")

        texts = [sh for sh in slide.shapes
                 if sh.has_text_frame and sh.text_frame.text.strip()
                 and sh not in [x[0] for x in numerals]]
        connectors = [sh for sh in slide.shapes
                      if "CONNECTOR" in str(sh.shape_type)]
        for num, top in numerals:
            n_right = (num.left or 0) + (num.width or 0) * 0.8
            cands = [(abs((sh.top or 0) / EMU_IN - top), sh) for sh in texts
                     if (sh.left or 0) >= n_right
                     and abs((sh.top or 0) / EMU_IN - top) < 0.6]
            if not cands:
                raise FitError(f"目录序号「{num.text_frame.text}」右侧找不到标题文字")
            label = min(cands, key=lambda c: c[0])[1]
            rule = None
            for c in connectors:
                if abs((c.top or 0) / EMU_IN - top) < 0.35:
                    rule = c
                    break
            group = [g for g in (num, label, rule) if g is not None]
            rows.append((min((g.top or 0) / EMU_IN for g in group),
                         group, num, label))
        return rows

    def section_page(self, no, title, subtitle="", *, notes=""):
        """
        A divider, which the template does not ship.

        Deliberately restrained: the numeral in the roundel plus a large title,
        so a deck with dividers still looks like this template rather than like
        a second design pasted into it.
        """
        s = self._new_slide(False)
        self._badge(s, no)
        self._text(s, self.left, 2.5, self.width, 0.5, [f"{no:02d}"],
                   size=COVER_CN_PT, color=self.theme.accent, bold=True,
                   spacing=1.0)
        r = wrap_text(title, COVER_EN_PT,
                      (self.width - 0.2) * 72, True)
        size = COVER_EN_PT
        while r.overflow_x and size > 24:
            size -= 2
            r = wrap_text(title, size, (self.width - 0.2) * 72, True)
        self._text(s, self.left, 3.1, self.width, 1.0, [title], size=size,
                   color=self.theme.primary, bold=True, spacing=1.1)
        self._rule(s, self.left, 4.35, 2.2, self.theme.accent, 1.5)
        if subtitle:
            self._text(s, self.left, 4.6, self.width, 0.9, [subtitle],
                       size=BODY_PT, color=self.theme.muted, spacing=1.3)
        self._notes(s, notes)
        return s

    # ------------------------------------------------------------------
    # internals
    # ------------------------------------------------------------------
    def _sig(self, role, page_index, slots):
        """
        Locate a signature page, its slot map, and any box that must be widened.

        The profile may carry an explicit `pages` map (written once, by a human
        looking at the template); otherwise fall back to the index the caller
        gave. Nothing is guessed from shape names, because they are localised
        and differ per file.
        """
        pages = self.profile.get("pages", {}) or {}
        entry = pages.get(role) or {}
        idx = page_index or entry.get("index")
        if idx is None:
            raise ValueError(
                f"没有 {role!r} 页的索引：在 profile 里加 pages.{role}.index，"
                f"或直接传 page_index=")
        merged = dict(entry.get("slots") or {})
        for k, v in slots.items():
            merged.setdefault(k, v)
        return int(idx), merged, dict(entry.get("widen") or {})

    def _apply_widen(self, slide, widen):
        """
        Give an under-sized text box the room it needs.

        Designers draw a box tight around their own wording; put a longer group
        name in it and the line breaks in the middle of a word, which is how a
        borrowed page starts to look broken. Widening is safe when the space to
        the right is empty, and only that case is declared in the profile.
        """
        if not widen:
            return
        by_name = {sh.name: sh for sh in slide.shapes}
        for name, width in widen.items():
            sh = by_name.get(name)
            if sh is None or sh.width is None:
                continue
            right = (sh.left or 0) / EMU_IN + float(width)
            if right > self.slide_w - 0.3:
                continue
            sh.width = Emu(int(float(width) * EMU_IN))

    def _fill(self, slide, slots, values):
        by_name = {sh.name: sh for sh in slide.shapes}
        for key, text in values.items():
            if not text:
                continue
            name = slots.get(key)
            if not name or name not in by_name:
                continue
            sh = by_name[name]
            before = _shape_need(sh)
            _write_paragraphs(sh, text)
            if self.strict:
                _fit_cloned_shape(sh, self._n, before[0], before[1])

    def save(self, path):
        """
        Write the deck, first shedding the template's own sample pages.

        The originals are kept until now on purpose: `clone()` reads them, so a
        page cannot be removed before it has been copied.
        """
        if self._orig:
            drop = list(self._prs.slides)[:self._orig]
            keep = set(self.keep_pages)
            for i, s in enumerate(drop, 1):
                if i not in keep:
                    self.drop(s)
            self._orig = 0
        super().save(path)
        return path
