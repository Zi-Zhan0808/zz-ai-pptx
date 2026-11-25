"""
layouts.py - A library of structurally different slide layouts.

WHY THIS EXISTS
---------------
The first release had one honest flaw: almost every content slide was
"text on the left, tinted panel on the right". Repeating one skeleton eight
times makes a deck feel machine-made no matter how good each individual slide
is. Variety has to come from *structure* — where the reading eye travels —
not from swapping colours or adding decoration.

Every layout here obeys the same hard rules as the rest of the skill:
  * text is measured before placement (FitError instead of silent overflow)
  * four-step type scale, no ad-hoc sizes
  * ink-on-light for anything meant to be read; dark slides only frame
  * no decorative bars, no accent stripes
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from textfit import measure, wrap_text  # noqa: E402

EMU_IN = 914400

TITLE_PT = 40
SECTION_PT = 28
BODY_PT = 18
CAPTION_PT = 12
SMALL_PT = 10


def rgb(hexs: str) -> RGBColor:
    h = hexs.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


class FitError(RuntimeError):
    """Raised instead of emitting a slide whose text does not fit its box."""


class Layouts:
    """
    Mixed into Deck. Keeps geometry code out of the content API so a caller can
    read `deck.timeline(...)` and not care about EMUs.
    """

    # ------------------------------------------------------------------ utils
    def _noautofit(self, tf):
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        bodyPr = tf._txBody.find(qn("a:bodyPr"))
        for c in list(bodyPr):
            bodyPr.remove(c)
        bodyPr.set("wrap", "square")
        bodyPr.append(bodyPr.makeelement(qn("a:noAutofit"), {}))

    def _fit(self, paragraphs, w_in, h_in, *, base=BODY_PT, max_pt=None,
             spacing=1.45, bold=False, floor_ratio=0.7):
        """Largest size in [base*floor, max] that fills but does not overflow."""
        max_pt = max_pt or base + 10
        lo = max(9, int(base * floor_ratio))
        for size in range(int(max_pt), lo - 1, -1):
            need = 0.0
            ok = True
            for p in paragraphs:
                r = wrap_text(p, size, (w_in - 0.2) * 72 - 8, bold,
                              line_spacing=spacing)
                if r.overflow_x:
                    ok = False
                    break
                need += r.needed_height_pt
            if ok and need <= (h_in - 0.1) * 72 * 0.95:
                return float(size)
        return float(lo)

    def _shape(self, slide, kind, x, y, w, h, fill=None, line=None,
               line_w=None, adj=None):
        sp = slide.shapes.add_shape(kind, Inches(x), Inches(y),
                                    Inches(w), Inches(h))
        if fill is None:
            sp.fill.background()
        else:
            sp.fill.solid()
            sp.fill.fore_color.rgb = rgb(fill)
        if line:
            sp.line.color.rgb = rgb(line)
            sp.line.width = Pt(line_w or 1.25)
        else:
            sp.line.fill.background()
        sp.shadow.inherit = False
        if adj is not None:
            try:
                for i, v in enumerate(adj):
                    sp.adjustments[i] = v
            except Exception:
                pass
        self._noautofit(sp.text_frame)
        return sp

    def _line(self, slide, x1, y1, x2, y2, color, width=1.5, dash=None):
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                        Inches(x1), Inches(y1),
                                        Inches(x2), Inches(y2))
        ln.line.color.rgb = rgb(color)
        ln.line.width = Pt(width)
        if dash:
            from pptx.enum.dml import MSO_LINE_DASH_STYLE
            ln.line.dash_style = dash
        return ln

    def _label(self, slide, x, y, w, h, text, size, color, bold=False,
               align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.35,
               check=True, name=None):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        self._noautofit(tb.text_frame)
        tf = tb.text_frame
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Emu(int(0.1 * EMU_IN))
        tf.margin_top = tf.margin_bottom = Emu(int(0.05 * EMU_IN))
        lines = [text] if isinstance(text, str) else list(text)
        if check:
            need = 0.0
            for p in lines:
                r = wrap_text(p, size, (w - 0.2) * 72 - 8, bold, line_spacing=spacing)
                if r.overflow_x and self.strict:
                    raise FitError(
                        f"slide {self._n}: {size}pt line {r.widest_line_pt:.0f}pt "
                        f"> {w:.2f}in box — widen, shorten, or shrink")
                need += r.needed_height_pt
            if need > (h - 0.1) * 72 + 3 and self.strict:
                raise FitError(
                    f"slide {self._n}: {len(lines)} line(s) need {need:.0f}pt, "
                    f"box gives {(h-0.1)*72:.0f}pt")
        for i, ln in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = spacing
            r = p.add_run()
            r.text = ln
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.name = "Microsoft YaHei"
            r.font.color.rgb = rgb(color)
        if name:
            tb.name = name
        return tb

    def _chip(self, slide, x, y, w, h, text, fill, color, size=CAPTION_PT):
        """Small pill label. Used for step numbers and tags."""
        sp = self._shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
                         fill=fill, adj=[0.5])
        self._label(slide, x, y, w, h, text, size, color, bold=True,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        return sp


    def _cite_label(self, cite: str) -> str:
        """
        Validate a footer citation: it must name a reference, not just a page.

        "S2 p.2" says where to look but not which entry it belongs to; mapping it
        to "[2] S2 p.2" would be a guess, and inventing reference numbers is the
        exact failure mode this skill guards against. So this raises instead,
        and the caller writes "[1] S2 p.2".
        """
        import re as _re
        if not cite:
            return cite
        if _re.search(r"\[\s*\d", cite):
            return cite
        if _re.search(r"\bS\d+\s*p", cite):
            raise FitError(
                f"slide {self._n}: footer {cite!r} names a locator but no "
                f"reference marker — write it as \"[n] {cite}\" so the reader "
                f"can tell which entry it belongs to")
        return cite

    # ------------------------------------------------------------- intro pages
    def cover(self, title, subtitle="", meta="", badge="", notes="", cite=""):
        """
        Cover with a structural anchor instead of a centred block of text:
        a large left-aligned title, an optional corner badge, and a hairline
        rule that carries the deck's accent.
        """
        s = self._new_slide(dark=True)
        t = self.theme
        if badge:
            self._label(s, self.margin, 0.85, 6.0, 0.4, badge, CAPTION_PT,
                        t.accent_inverse, bold=True)
        # A headline may carry its own line break; respect it and size the block
        # from the wrapped height of everything that will actually be drawn.
        lines = title.split("\n")
        size = TITLE_PT + 6
        avail_w = (self.slide_w - self.margin * 2 - 0.2) * 72
        need = 0.0
        for ln in lines:
            rr = wrap_text(ln, size, avail_w, True, line_spacing=1.12)
            need += rr.needed_height_pt
        th = need / 72 + 0.3
        self._label(s, self.margin, 1.55, self.slide_w - self.margin * 2, th,
                    lines, size, t.ink_inverse, bold=True, spacing=1.12)
        y = 1.55 + th + 0.45
        self._line(s, self.margin, y, self.margin + 1.1, y, t.accent_inverse, 3)
        y += 0.42
        if subtitle:
            sr = wrap_text(subtitle, SECTION_PT,
                           (self.slide_w - self.margin * 2 - 0.6) * 72)
            sh = sr.needed_height_pt / 72 + 0.2
            self._label(s, self.margin, y, self.slide_w - self.margin * 2, sh,
                        subtitle, SECTION_PT, t.muted_inverse, spacing=1.4)
            y += sh + 0.25
        if meta:
            self._label(s, self.margin, y, self.slide_w - self.margin * 2, 0.5,
                        meta, BODY_PT, t.muted_inverse)
        if cite:
            self._label(s, self.margin, self.slide_h - 0.85,
                        self.slide_w - self.margin * 2, 0.35, cite,
                        CAPTION_PT, t.muted_inverse, check=False)
        self._notes(s, notes)
        return s

    def agenda(self, items, title="目录", kicker="", notes=""):
        """
        Numbered agenda as a two-column index with big numerals — reads as a
        map of the deck rather than another bullet list.
        """
        s = self._new_slide(dark=False)
        t = self.theme
        top = self._header(s, kicker or "CONTENTS", title)
        bottom = self.slide_h - 0.85
        n = len(items)
        cols = 2 if n > 4 else 1
        rows = math.ceil(n / cols)
        gap = 0.6
        col_w = (self.slide_w - 2 * self.margin - gap) / cols
        # Cap the row pitch: stretching a 3-row list across the full height
        # produces gaps so large the items stop reading as one list.
        avail = bottom - top
        row_h = min(avail / rows, 1.05)
        v_pad = max(0.0, (avail - row_h * rows) * 0.22)
        for i, item in enumerate(items):
            c, r_ = divmod(i, rows)
            x = self.margin + c * (col_w + gap)
            y = top + v_pad + r_ * row_h
            num = f"{i+1:02d}"
            self._label(s, x, y, 1.0, 0.62, num, SECTION_PT, t.accent, bold=True,
                        check=False)
            txt = item if isinstance(item, str) else item[0]
            sz = self._fit([txt], col_w - 1.1, row_h - 0.15, base=BODY_PT,
                           max_pt=BODY_PT + 4, spacing=1.3)
            self._label(s, x + 1.05, y + 0.02, col_w - 1.1, row_h - 0.2,
                        txt, sz, t.ink, spacing=1.3)
            # No divider between entries: a rule that stops halfway across the
            # page divides nothing and reads as leftover furniture. Whitespace
            # and the numerals already separate the rows.
        self._footer(s, "")
        self._notes(s, notes)
        return s

    def section_divider(self, index, title, bullets=(), notes=""):
        """
        Section opener with a big ghost numeral. Keeps dark framing slides from
        all looking like the same centred sentence.
        """
        s = self._new_slide(dark=True)
        t = self.theme
        # ghost numeral: a浅 wash of the ink, not the canvas colour — on a pale
        # panel the canvas hue would disappear entirely
        self._label(s, self.margin, 1.5, 3.2, 2.6, f"{index:02d}",
                    TITLE_PT + 96, t.secondary, bold=True, check=False)
        # ghost numeral: recolour to a low-contrast tone so it reads as texture
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text == f"{index:02d}":
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.color.rgb = rgb(self._mix(t.bg_dark, t.ink_inverse, 0.14))
        r = wrap_text(title, TITLE_PT, (self.slide_w - 2 * self.margin - 3.8) * 72,
                      True)
        th = r.needed_height_pt / 72 + 0.3
        self._label(s, self.margin + 3.6, 2.35,
                    self.slide_w - self.margin * 2 - 3.6, th, title, TITLE_PT,
                    t.ink_inverse, bold=True, spacing=1.12)
        if bullets:
            by = 2.35 + th + 0.4
            rows = ["\u2014 " + b for b in bullets]
            need = 0.0
            for b_ in rows:
                rr = wrap_text(b_, BODY_PT,
                               (self.slide_w - self.margin * 2 - 3.6 - 0.2) * 72,
                               line_spacing=1.6)
                need += rr.needed_height_pt / 72
            bh = min(need + 0.45, self.slide_h - by - 0.7)
            self._label(s, self.margin + 3.6, by,
                        self.slide_w - self.margin * 2 - 3.6, bh,
                        rows, BODY_PT, t.muted_inverse, spacing=1.6)
        self._notes(s, notes)
        return s

    @staticmethod
    def _mix(a: str, b: str, t: float) -> str:
        a, b = a.lstrip("#"), b.lstrip("#")
        out = []
        for i in (0, 2, 4):
            av, bv = int(a[i:i+2], 16), int(b[i:i+2], 16)
            out.append(f"{int(av + (bv - av) * t):02x}")
        return "".join(out)

    # --------------------------------------------------------------- body pages
    def lead_statement(self, statement, support=(), kicker="", cite="", notes=""):
        """
        One large statement with a short support line. Deliberately text-only and
        oversized: it gives the deck a beat between dense slides instead of a
        panel that exists only to fill a column.
        """
        s = self._new_slide(dark=False)
        t = self.theme
        top = 1.35
        if kicker:
            self._label(s, self.margin, top, 8.0, 0.4, kicker, CAPTION_PT,
                        t.accent, bold=True)
            top += 0.5
        avail_h = self.slide_h - top - (1.7 if support else 1.1)
        # A sentence of two clauses set at 46pt fills half the slide and cannot be
        # read in one pass. Keep the beat, drop the shout: long statements are
        # capped to three lines and a smaller ceiling, short ones may stay large.
        cap = TITLE_PT + 6 if len(statement) <= 14 else 32
        sz = self._fit([statement], self.slide_w - 2 * self.margin, avail_h,
                       base=min(TITLE_PT, cap), max_pt=cap, spacing=1.28,
                       bold=True)
        while sz > 26:
            rr = wrap_text(statement, sz, (self.slide_w - 2 * self.margin - 0.2) * 72,
                           True, line_spacing=1.28)
            if len(rr.lines) <= 3:
                break
            sz -= 1
        r = wrap_text(statement, sz, (self.slide_w - 2 * self.margin - 0.2) * 72,
                      True, line_spacing=1.28)
        h = r.needed_height_pt / 72 + 0.2
        self._label(s, self.margin, top, self.slide_w - 2 * self.margin, h,
                    statement, sz, t.primary, bold=True, spacing=1.28)
        if support:
            sy = top + h + 0.30
            self._label(s, self.margin, sy, self.slide_w - 2 * self.margin,
                        self.slide_h - sy - (1.0 if cite else 0.6),
                        list(support), BODY_PT, t.muted, spacing=1.5)
        self._footer(s, self._cite_label(cite))
        self._notes(s, notes)
        return s

    def cards_grid(self, title, cards, *, cols=3, kicker="", cite="", notes="",
                   icon_set=None):
        """2x2 / 1x3 / 2x3 grid of icon cards. Different rhythm from `split`."""
        s = self._new_slide(dark=False)
        t = self.theme
        top = self._header(s, kicker, title)
        bottom = self.slide_h - (1.0 if cite else 0.75)
        n = len(cards)
        rows = math.ceil(n / cols)
        gap = 0.34
        total_w = self.slide_w - 2 * self.margin
        cw = (total_w - gap * (cols - 1)) / cols
        ch = (bottom - top - gap * (rows - 1)) / rows
        icon_size = min(0.62, ch * 0.18)
        for i, card in enumerate(cards):
            label, body = card[0], card[1]
            icon = card[2] if len(card) > 2 else (icon_set[i] if icon_set and i < len(icon_set) else None)
            r_, c_ = divmod(i, cols)
            x = self.margin + c_ * (cw + gap)
            y = top + r_ * (ch + gap)
            self._shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, cw, ch,
                        fill=t.secondary, adj=[0.06])
            ty = y + 0.34
            if icon:
                self._draw_icon(s, icon, x + 0.34, ty, icon_size, t.accent)
                ty += icon_size + 0.24
            self._label(s, x + 0.34, ty, cw - 0.68, 0.5, label, BODY_PT,
                        t.primary, bold=True, spacing=1.25)
            ty += 0.46
            body_sz = self._fit([body], cw - 0.68, ch - (ty - y) - 0.3,
                                base=BODY_PT - 4, max_pt=BODY_PT - 2, spacing=1.45)
            body_h = y + ch - ty - 0.26
            self._label(s, x + 0.34, ty, cw - 0.68, max(0.34, body_h),
                        body, body_sz, t.ink, spacing=1.45)
        self._footer(s, self._cite_label(cite))
        self._notes(s, notes)
        return s

    def stat_row(self, title, stats, *, kicker="", caption="", cite="",
                 notes="", divide=True):
        """Numbers across a full-width band, separated by rules (no cards)."""
        s = self._new_slide(dark=False)
        t = self.theme
        top = self._header(s, kicker, title)
        bottom = self.slide_h - (1.0 if cite else 0.75)
        cap_h = 0.5 if caption else 0.0
        band = bottom - top - cap_h
        n = len(stats)
        total = self.slide_w - 2 * self.margin
        cw = total / n
        num_h = min(band * 0.46, 1.75)
        num_pt = 78
        while num_pt > 40:
            if max(measure(v, num_pt, True) for v, _ in stats) <= (cw - 0.5) * 72 - 8:
                break
            num_pt -= 4
        y0 = top + max(0.0, (band - (num_h + 0.75)) * 0.28)
        for i, (value, label) in enumerate(stats):
            x = self.margin + i * cw
            self._label(s, x, y0, cw - 0.3, num_h, value, num_pt, t.accent,
                        bold=True, spacing=1.0)
            self._label(s, x, y0 + num_h + 0.12, cw - 0.35, 0.75, label,
                        BODY_PT, t.ink, spacing=1.35)
            if divide and i > 0:
                self._line(s, x - 0.15, y0 + 0.1, x - 0.15, y0 + num_h + 0.7,
                           t.panel, 1.25)
        if caption:
            self._label(s, self.margin, bottom - cap_h + 0.05, total, 0.4,
                        caption, CAPTION_PT, t.muted, check=False)
        self._footer(s, self._cite_label(cite))
        self._notes(s, notes)
        return s

    def process_flow(self, title, steps, *, kicker="", cite="", notes="",
                     horizontal=True):
        """
        Horizontal or vertical chevron flow. Native shapes, fully editable,
        with the step count driving the geometry instead of a fixed template.
        """
        s = self._new_slide(dark=False)
        t = self.theme
        top = self._header(s, kicker, title)
        bottom = self.slide_h - (1.0 if cite else 0.75)
        n = len(steps)
        if horizontal:
            gap = 0.16
            total = self.slide_w - 2 * self.margin
            cw = (total - gap * (n - 1)) / n
            ch = min(1.45, (bottom - top) * 0.30)
            y = top + 0.15
            for i, step in enumerate(steps):
                head, body = step[0], (step[1] if len(step) > 1 else "")
                x = self.margin + i * (cw + gap)
                # Every step is a CHEVRON, including the first. A PENTAGON
                # (homePlate) has a flat right edge, so step 1 next to a chevron
                # row reads as a broken arrow rather than the start of a flow.
                self._shape(s, MSO_SHAPE.CHEVRON, x, y, cw, ch,
                            fill=t.secondary, adj=[0.22])
                # white on the accent fill: the accent is the one saturated
                # element on the page, so white type on it stays legible
                self._chip(s, x + 0.22, y + 0.2, 0.42, 0.34, str(i + 1),
                           t.accent, "FFFFFF", size=CAPTION_PT - 1)
                self._label(s, x + 0.24, y + 0.6, cw - 0.55, ch - 0.7, head,
                            BODY_PT - 2, t.primary, bold=True, spacing=1.2)
                if body:
                    by = y + ch + 0.22
                    self._label(s, x, by, cw - 0.15,
                                bottom - by, body, BODY_PT - 5, t.ink, spacing=1.45)
        else:
            row_h = (bottom - top) / n
            bar_w = 0.06
            for i, step in enumerate(steps):
                head, body = step[0], (step[1] if len(step) > 1 else "")
                y = top + i * row_h
                self._label(s, self.margin, y, 0.7, 0.5, f"{i+1:02d}", SECTION_PT - 4,
                            t.accent, bold=True, check=False)
                self._label(s, self.margin + 0.75, y, 3.6, 0.5, head, BODY_PT,
                            t.primary, bold=True, spacing=1.2)
                if body:
                    self._label(s, self.margin + 4.5, y, 
                                self.slide_w - 2 * self.margin - 4.5,
                                row_h - 0.15, body, BODY_PT - 3, t.ink, spacing=1.45)
                if i < n - 1:
                    self._line(s, self.margin, y + row_h - 0.12,
                               self.slide_w - self.margin, y + row_h - 0.12,
                               t.panel, 1)
        self._footer(s, self._cite_label(cite))
        self._notes(s, notes)
        return s

    def compare_split(self, title, left, right, *, left_title="", right_title="",
                      kicker="", cite="", notes=""):
        """
        Two opposing columns with a vertical rule between them. The visual
        argument is the symmetry, so it uses no cards at all.
        """
        s = self._new_slide(dark=False)
        t = self.theme
        top = self._header(s, kicker, title)
        bottom = self.slide_h - (1.0 if cite else 0.75)
        h = bottom - top
        gap = 0.8
        col_w = (self.slide_w - 2 * self.margin - gap) / 2
        lx = self.margin
        rx = self.margin + col_w + gap
        for x, head, items, col in ((lx, left_title, left, t.primary),
                                    (rx, right_title, right, t.accent)):
            hy = top
            if head:
                self._label(s, x, hy, col_w, 0.5, head, BODY_PT, col, bold=True)
                self._line(s, x, hy + 0.52, x + min(1.4, col_w), hy + 0.52, col, 2)
                hy += 0.75
            sz = self._fit(list(items), col_w, bottom - hy, base=BODY_PT,
                           max_pt=BODY_PT + 3, spacing=1.55)
            self._label(s, x, hy, col_w, bottom - hy, list(items), sz, t.ink,
                        spacing=1.55)
        self._line(s, self.margin + col_w + gap / 2, top + 0.1,
                   self.margin + col_w + gap / 2, bottom - 0.1, t.panel, 1.25)
        self._footer(s, self._cite_label(cite))
        self._notes(s, notes)
        return s

    def quote(self, text, source="", role="", kicker="", cite="", notes=""):
        """
        Pull quote for a definition, a principle, or a paper's own words.

        The quote mark is a fixed anchor in its own column; the text starts to
        its right. Placing the glyph behind the text (the usual shortcut) makes
        it collide with the first line at every wrap width.
        """
        s = self._new_slide(dark=True)
        t = self.theme
        gutter = 1.25                     # column reserved for the quote mark
        text_x = self.margin + gutter
        text_w = self.slide_w - text_x - self.margin

        if kicker:
            self._label(s, text_x, 1.15, text_w, 0.4, kicker, CAPTION_PT,
                        t.accent_inverse, bold=True)

        # measure first so the block can be centred vertically
        sz = self._fit([text], text_w, 3.4, base=SECTION_PT, max_pt=TITLE_PT,
                       spacing=1.38, bold=True)
        r = wrap_text(text, sz, (text_w - 0.2) * 72, True, line_spacing=1.38)
        qh = r.needed_height_pt / 72 + 0.2
        extra = (0.5 if source else 0) + (0.42 if role else 0)
        block_h = qh + extra
        top = max(1.5, (self.slide_h - block_h) / 2 - 0.15)

        # quote mark sits in the gutter, aligned to the text's cap height
        q_pt = TITLE_PT + 34
        q_h = q_pt / 72 * 1.35
        self._label(s, self.margin - 0.05, top - 0.30, gutter, q_h, "\u201c",
                    q_pt, self._mix(t.bg_dark, t.ink_inverse, 0.30),
                    bold=True, check=False)

        self._label(s, text_x, top, text_w, qh, text, sz, t.ink_inverse,
                    bold=True, spacing=1.38)
        y = top + qh + 0.3
        if source:
            self._label(s, text_x, y, text_w, 0.45, source, BODY_PT,
                        t.accent_inverse, bold=True)
            y += 0.5
        if role:
            self._label(s, text_x, y, text_w, 0.4, role, CAPTION_PT,
                        t.muted_inverse)
        if cite:
            self._label(s, self.margin, self.slide_h - 0.85,
                        self.slide_w - 2 * self.margin, 0.35, cite,
                        CAPTION_PT, t.muted_inverse, check=False)
        self._notes(s, notes)
        return s

    def summary(self, title, rows, *, kicker="小结", cite="", notes=""):
        """
        A closing summary table: claim / evidence / what to do. The point of a
        summary is the third column — a deck that ends on restated bullets has
        not summarised anything.
        """
        s = self._new_slide(dark=False)
        t = self.theme
        top = self._header(s, kicker, title)
        bottom = self.slide_h - (1.0 if cite else 0.75)
        headers = ["判断", "证据", "意味着什么"]
        n = len(rows) + 1
        total = self.slide_w - 2 * self.margin
        widths = [0.34, 0.30, 0.36]
        # Header row stays shallow; the body rows carry the reading weight.
        head_h = 0.46
        body_h = min(0.86, (bottom - top - head_h) / max(1, n - 1))
        tbl = s.shapes.add_table(n, 3, Inches(self.margin), Inches(top),
                                 Inches(total),
                                 Inches(head_h + body_h * (n - 1))).table
        for j, wf in enumerate(widths):
            tbl.columns[j].width = Inches(total * wf)
        tbl.rows[0].height = Inches(head_h)
        for i in range(1, n):
            tbl.rows[i].height = Inches(body_h)
        for j, htxt in enumerate(headers):
            cell = tbl.cell(0, j)
            cell.text = ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(t.primary)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = htxt
            r.font.size = Pt(CAPTION_PT + 2)
            r.font.bold = True
            r.font.name = "Microsoft YaHei"
            r.font.color.rgb = rgb(t.bg_light)
            cell.margin_left = cell.margin_right = Emu(int(0.14 * EMU_IN))
        for i, row in enumerate(rows, start=1):
            for j, val in enumerate(row):
                cell = tbl.cell(i, j)
                cell.text = ""
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(t.bg_light if i % 2 else t.tint
                                               if hasattr(t, "tint") else t.secondary)
                cell.margin_left = cell.margin_right = Emu(int(0.14 * EMU_IN))
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                sz = BODY_PT - 4
                avail = (total * widths[j] - 0.3) * 72 - 8
                while sz > 10 and measure(str(val), sz) > avail:
                    sz -= 1
                if measure(str(val), sz) > avail and self.strict:
                    raise FitError(
                        f"slide {self._n}: summary cell {j+1} too narrow for "
                        f"{str(val)[:24]!r}")
                p = cell.text_frame.paragraphs[0]
                r = p.add_run()
                r.text = str(val)
                r.font.size = Pt(sz)
                r.font.name = "Microsoft YaHei"
                r.font.bold = (j == 0)
                r.font.color.rgb = rgb(t.primary if j == 0 else t.ink)
        self._footer(s, self._cite_label(cite))
        self._notes(s, notes)
        return s

    def closing(self, title, lines=(), actions=(), cite="", notes=""):
        """Closing slide with an explicit ask, optionally as action chips."""
        s = self._new_slide(dark=True)
        t = self.theme
        r = wrap_text(title, TITLE_PT, (self.slide_w - 2 * self.margin) * 72, True)
        th = r.needed_height_pt / 72 + 0.3
        self._label(s, self.margin, 2.15, self.slide_w - 2 * self.margin, th,
                    title, TITLE_PT, t.ink_inverse, bold=True, spacing=1.15)
        y = 2.15 + th + 0.4
        if lines:
            self._label(s, self.margin, y, self.slide_w - 2 * self.margin, 1.3,
                        list(lines), BODY_PT, t.muted_inverse, spacing=1.55)
            y += 1.3
        if actions:
            x = self.margin
            for a in actions:
                w = measure(a, CAPTION_PT + 2, True) / 72 + 0.7
                if x + w > self.slide_w - self.margin:
                    break
                self._chip(s, x, y + 0.1, w, 0.46, a, t.accent_inverse, t.bg_dark,
                           size=CAPTION_PT + 2)
                x += w + 0.22
        if cite:
            self._label(s, self.margin, self.slide_h - 0.85,
                        self.slide_w - 2 * self.margin, 0.35, cite,
                        CAPTION_PT, t.muted_inverse, check=False)
        self._notes(s, notes)
        return s

    # ------------------------------------------------------------------ icons
    def _draw_icon(self, slide, name, x, y, size, color):
        """
        Draw a small geometric icon from native shapes.

        Icons are built here rather than downloaded so every glyph is
        licence-clean, palette-matched, and editable in PowerPoint.
        """
        s = slide
        st = MSO_SHAPE
        cx, cy = x + size / 2, y + size / 2
        half = size / 2
        thin = max(0.05, size * 0.09)
        if name == "shield":
            self._shape(s, st.PENTAGON, x, y, size, size, fill=color, adj=[0.4])
        elif name == "target":
            self._shape(s, st.DONUT, x, y, size, size, fill=color, adj=[0.32])
            self._shape(s, st.OVAL, cx - size * 0.11, cy - size * 0.11,
                        size * 0.22, size * 0.22, fill=color)
        elif name == "layers":
            for i, dy in enumerate((0.16, 0.42, 0.68)):
                self._shape(s, st.ROUNDED_RECTANGLE, x + dy * size * 0.5, y + dy * size,
                            size * 0.62, size * 0.2, fill=color, adj=[0.4])
        elif name == "branch":
            self._line(s, cx, y + size * 0.12, cx, y + size * 0.42, color, 2.2)
            self._line(s, cx, y + size * 0.42, x + size * 0.2, y + size * 0.72, color, 2.2)
            self._line(s, cx, y + size * 0.42, x + size * 0.8, y + size * 0.72, color, 2.2)
            for px, py in ((cx, y + size * 0.06), (x + size * 0.16, y + size * 0.78),
                           (x + size * 0.84, y + size * 0.78)):
                self._shape(s, st.OVAL, px - thin, py - thin / 2, thin * 2, thin * 2,
                            fill=color)
        elif name == "grid":
            for r_ in range(2):
                for c_ in range(2):
                    self._shape(s, st.ROUNDED_RECTANGLE,
                                x + c_ * size * 0.54, y + r_ * size * 0.54,
                                size * 0.42, size * 0.42, fill=color, adj=[0.2])
        elif name == "check":
            self._shape(s, st.OVAL, x, y, size, size, fill=color)
            self._line(s, x + size * 0.26, y + size * 0.52, x + size * 0.44,
                       y + size * 0.70, "#FFFFFF", 2.6)
            self._line(s, x + size * 0.44, y + size * 0.70, x + size * 0.76,
                       y + size * 0.32, "#FFFFFF", 2.6)
        elif name == "alert":
            self._shape(s, st.DIAMOND, x, y, size, size, fill=color)
            # the exclamation mark is drawn, not typed: text inside an icon
            # joins the deck's text blocks and gets measured as body copy.
            self._shape(s, st.ROUNDED_RECTANGLE,
                        cx - size * 0.055, y + size * 0.24,
                        size * 0.11, size * 0.30, fill="#FFFFFF", adj=[0.5])
            self._shape(s, st.OVAL, cx - size * 0.065, y + size * 0.62,
                        size * 0.13, size * 0.13, fill="#FFFFFF")
        elif name == "flow":
            self._shape(s, st.CHEVRON, x, y, size * 0.55, size, fill=color,
                        adj=[0.32])
            self._shape(s, st.CHEVRON, x + size * 0.45, y, size * 0.55, size,
                        fill=self._mix(color, "#FFFFFF", 0.45), adj=[0.32])
        elif name == "stack":
            for i, dy in enumerate((0.62, 0.34, 0.06)):
                self._shape(s, st.ROUNDED_RECTANGLE, x, y + size * dy,
                            size, size * 0.24, fill=color, adj=[0.35])
        else:  # dot: a neutral bullet marker
            self._shape(s, st.OVAL, cx - size * 0.18, cy - size * 0.18,
                        size * 0.36, size * 0.36, fill=color)
        return None
