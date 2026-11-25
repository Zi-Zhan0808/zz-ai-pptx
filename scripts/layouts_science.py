"""
layouts_science.py - page forms a scientific argument needs and a grid of boxes
cannot express.

Why a separate module
---------------------
The eleven layouts in `layouts.py` are one idea: put text into regions. That
covers "here are three points" and little else. A paper talk keeps asking for
different pictures:

  two rival models         -> draw both, then say what would decide between them
  a distance that matters  -> draw the distance, put the marks at its ends
  a 2x3 categorical result -> a matrix of marks, not a table of words
  a before/after magnitude -> two marks on one track, so the ratio is visible
  five limitations         -> a ledger of claims, not five equal cards
  one direction that matters -> one dominant tile, three satellites

None of those is a nicer card grid; they are different pictures of different
content. Implemented as a mixin so `Deck` keeps one public vocabulary.

Every method sizes its type to the room it actually has (`_shrink`) and then
passes the result to `_label`, which measures again. Guessing a size and letting
the fitter reject it is a slow way to do arithmetic.
"""
from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

EMU_IN = 914400
BODY_PT = 18
CAPTION_PT = 12


def _rgb(hexstr: str) -> RGBColor:
    h = hexstr.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


class ScienceLayouts:
    """Page forms for diagrams, matrices, deltas and ledgers."""

    # ---------------------------------------------------------------- helpers
    def _shrink(self, texts, w_in, h_in, base, floor=11, spacing=1.35):
        """Largest type size at which `texts` fits `w_in x h_in`, measured."""
        from textfit import wrap_text
        if isinstance(texts, str):
            texts = [texts]
        avail_w = (w_in - 0.2) * 72 - 8.0
        avail_h = (h_in - 0.1) * 72
        size = base
        while size > floor:
            need = sum(wrap_text(str(t), size, avail_w, line_spacing=spacing
                                 ).needed_height_pt for t in texts)
            if need <= avail_h:
                break
            size -= 1
        return size

    def _bar(self, slide, x, y, w, h, color, name="measure-bar"):
        """A hairline drawn as a shape, not a connector.

        A connector is how this library draws separators, and the design lint is
        right to question an unanchored one. A measured span - a distance, a
        before/after track - is data, so it is built as a very thin rectangle and
        read as data.
        """
        bar = self._shape(slide, MSO_SHAPE.RECTANGLE, x, y, max(w, 0.008), h,
                          fill=color)
        bar.line.fill.background()
        bar.name = name
        return bar

    def _text_w(self, text, size):
        from textfit import measure
        return measure(text, size) / 72.0

    def _page(self, title, kicker="", cite="", dark=False):
        """Standard content page: optional kicker, title, footer. Returns top."""
        s = self._new_slide(dark)
        t = self.theme
        top = 1.10
        if kicker:
            self._label(s, self.margin, 0.58, self.slide_w - 2 * self.margin,
                        0.34, kicker.upper(), CAPTION_PT + 1, t.accent,
                        bold=True, check=False)
            top = 1.34
        if title:
            w = self.slide_w - 2 * self.margin
            size = self._fit([title], w - 0.2, 0.94, base=32, max_pt=34,
                             bold=True)
            self._label(s, self.margin, top - 0.26, w, 0.96, title, size,
                        t.ink, bold=True, spacing=1.1)
            top += 0.74
        self._footer(s, self._cite_label(cite))
        return s, top

    # ---------------------------------------------------------------- duel
    def duel(self, title, left, right, *, verdict="", verdict_label="判据",
             cite="", notes="", kicker=""):
        """
        Two rival mechanisms, each with a small drawing, closed by the test that
        separates them.

        A two-column split only says "there are two sides". Two mechanisms need
        three things: what each claims, what it would look like, and what
        observation decides between them. The verdict band is what turns the page
        from a menu into an argument.

        side = {"tag", "head", "lines": [...], "motif": "switch" | "reach"}
        """
        s, top = self._page(title, kicker, cite)
        t = self.theme
        gap = 0.46
        pw = (self.slide_w - 2 * self.margin - gap) / 2
        vh = 0.92 if verdict else 0.0
        ph = 6.70 - top - vh - (0.14 if verdict else 0)
        pad = 0.28

        for i, side in enumerate((left, right)):
            px = self.margin + i * (pw + gap)
            self._shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, px, top, pw, ph,
                        fill=t.tint if i == 0 else None,
                        line=None if i == 0 else t.panel, adj=[0.04])
            y_tag = top + pad
            y_head = y_tag + 0.34
            head_h = 0.58
            y_motif = y_head + head_h + 0.08
            motif_h = min(1.20, max(0.72, ph * 0.30))
            y_lines = y_motif + motif_h + 0.10
            lines_h = top + ph - pad - y_lines
            if lines_h < 0.55:
                raise FitError(
                    f"slide {self._n}: 对照页文字区只剩 {lines_h:.2f}in —— "
                    f"删一行说明或缩短标题")

            self._label(s, px + 0.30, y_tag, pw - 0.60, 0.32,
                        side.get("tag", ""), CAPTION_PT + 1, t.accent,
                        bold=True, check=False)
            self._label(s, px + 0.30, y_head, pw - 0.60, head_h,
                        side.get("head", ""), BODY_PT + 4, t.ink, bold=True,
                        spacing=1.05, check=False)
            self._motif(s, px + 0.30, y_motif, pw - 0.60, motif_h,
                        side.get("motif", "reach"), t)
            body = side.get("lines", [])
            bsz = self._shrink(body, pw - 0.60, lines_h, BODY_PT - 3, floor=12,
                               spacing=1.4)
            self._label(s, px + 0.30, y_lines, pw - 0.60, lines_h, body, bsz,
                        t.muted, spacing=1.4)

        if verdict:
            vy = 6.70 - vh
            self._shape(s, MSO_SHAPE.RECTANGLE, self.margin, vy,
                        self.slide_w - 2 * self.margin, vh, fill=t.veil)
            self._label(s, self.margin + 0.24, vy + 0.10, 0.9, 0.32,
                        verdict_label, CAPTION_PT, t.ink, bold=True,
                        check=False)
            vw = self.slide_w - 2 * self.margin - 1.36
            vsz = self._shrink([verdict], vw, vh - 0.20, BODY_PT - 3, floor=11,
                               spacing=1.25)
            self._label(s, self.margin + 1.16, vy + 0.10, vw, vh - 0.18,
                        verdict, vsz, t.ink, spacing=1.25)
        self._notes(s, notes)
        return s

    def _motif(self, slide, x, y, w, h, kind, t):
        """A three-second drawing of what each mechanism claims."""
        cx, cy = x + w / 2, y + h / 2
        unit = min(1.0, h / 1.15)
        if kind == "switch":
            for k, (dx, sz) in enumerate(((-0.18, 0.58), (0.22, 0.44))):
                blob = self._shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                                   cx + dx - sz * unit / 2,
                                   cy - sz * unit / 2,
                                   sz * unit, sz * unit,
                                   fill=(t.panel if k == 0 else t.veil),
                                   line=None, adj=[0.42])
                blob.line.fill.background()
            self._label(slide, cx - 0.30, max(y, cy - 0.62 * unit - 0.34),
                        1.2, 0.32, "构象变", CAPTION_PT - 1, t.muted,
                        align=PP_ALIGN.CENTER, check=False)
            self._label(slide, cx - 0.30,
                        min(y + h - 0.32, cy + 0.34 * unit + 0.04),
                        1.2, 0.32, "活性↓", CAPTION_PT - 1, t.accent,
                        align=PP_ALIGN.CENTER, bold=True, check=False)
        else:
            r = 0.50 * unit
            reach = self._shape(slide, MSO_SHAPE.OVAL, cx - r, cy - r, r * 2,
                                r * 2, fill=None, line=t.accent)
            reach.line.width = Pt(1.0)
            reach.line.dash_style = 4
            self._shape(slide, MSO_SHAPE.OVAL, cx - r * 0.80, cy - 0.08,
                        0.16, 0.16, fill=t.accent)
            self._shape(slide, MSO_SHAPE.OVAL, cx + r * 0.62, cy - 0.08,
                        0.16, 0.16, fill=t.bg_light, line=t.muted)
            self._line(slide, cx - r * 0.68, cy - 0.36, cx + r * 0.52,
                       cy - 0.36, t.muted, 1.0)
            self._label(slide, cx - 0.72,
                        min(y + h - 0.32, cy + r + 0.02), 1.44, 0.32, "活性不变",
                        CAPTION_PT - 1, t.muted, align=PP_ALIGN.CENTER,
                        check=False)

    # ---------------------------------------------------------------- strata
    def strata(self, title, rows, *, axis="", kicker="", caption="", cite="",
               notes="", level_w=1.62):
        """
        The same question, answered at several levels of description.

        A ledger numbers peer claims; these are not peers, they are *scales* -
        structure, cell, pathway, measurement - and reading them top to bottom is
        the argument. So the left column is an axis with ticks, not a list of
        numbers, and the rows carry their own citation at the right edge.

        rows = [(level, text, refs)]
        """
        s, top = self._page(title, kicker, cite)
        t = self.theme
        n = len(rows)
        cap_h = 0.5 if caption else 0.0
        axis_h = 0.42 if axis else 0.0
        band_top = top + axis_h
        band_h = 6.66 - cap_h - band_top
        row_h = band_h / n
        ax_x = self.margin + level_w - 0.16
        ref_w = 1.30
        body_x = self.margin + level_w + 0.22
        body_w = self.slide_w - self.margin - ref_w - body_x

        if axis:
            # the axis runs the width of the page, so its caption may too; a box
            # as narrow as the level column broke it across two lines
            self._label(s, self.margin, top - 0.02,
                        self.slide_w - 2 * self.margin, 0.30, axis,
                        CAPTION_PT, t.muted, check=False)
        self._bar(s, ax_x, band_top + 0.06, 0.014, band_h - 0.12, t.panel,
                  name="axis-bar")

        # A single body size for the whole stack. Sizing each row to its own band
        # gives 16pt above 14pt above 16pt, which reads as a mistake; rows in one
        # stack should differ by content, not by fitting accidents.
        body_size = BODY_PT - 1
        for _lv, _tx, _rf in rows:
            body_size = min(body_size, self._shrink([_tx], body_w,
                                                    row_h - 0.16, BODY_PT - 1,
                                                    floor=12, spacing=1.4))

        for i, (level, text, refs) in enumerate(rows):
            y = band_top + i * row_h
            mid = y + row_h / 2
            self._bar(s, ax_x - 0.055, mid - 0.007, 0.125, 0.014, t.accent,
                      name="tick-bar")
            lsz = self._shrink([level], level_w - 0.34, 0.42, BODY_PT - 2,
                               floor=12, spacing=1.0)
            self._label(s, self.margin, mid - 0.21, level_w - 0.34, 0.42, level,
                        lsz, t.ink, bold=True, align=PP_ALIGN.RIGHT,
                        spacing=1.0, check=False)
            self._label(s, body_x, y + 0.08, body_w, row_h - 0.16, text,
                        body_size, t.muted, spacing=1.4)
            if refs:
                self._label(s, self.slide_w - self.margin - ref_w, mid - 0.16,
                            ref_w, 0.32, refs, CAPTION_PT + 1, t.accent,
                            align=PP_ALIGN.RIGHT, check=False)
        if caption:
            csz = self._shrink([caption], self.slide_w - 2 * self.margin, cap_h,
                               CAPTION_PT, floor=10, spacing=1.25)
            self._label(s, self.margin, 6.72 - cap_h,
                        self.slide_w - 2 * self.margin, cap_h, caption, csz,
                        t.muted, spacing=1.25)
        self._notes(s, notes)
        return s

    # ---------------------------------------------------------------- matrix
    def matrix(self, title, row_labels, col_labels, marks, *, kicker="",
               caption="", legend=None, highlight=None, cite="", notes="",
               row_width=3.1):
        """
        A categorical result as marks, not as prose.

        `marks` is a grid of 1 (present), 0 (absent) or None (not tested). The
        honest shape of "two sensors read under three conditions" is a 2x3 grid:
        a table of words hides the pattern, and a bullet list cannot show that
        exactly one cell differs.
        """
        s, top = self._page(title, kicker, cite)
        t = self.theme
        nrows, ncols = len(row_labels), len(col_labels)
        left = self.margin + row_width
        grid_w = self.slide_w - 2 * self.margin - row_width
        col_w = grid_w / ncols
        head_y = top + 0.14
        cap_h = 0.58 if caption else 0.0
        leg_h = 0.46 if legend else 0.0
        row_h = min(0.92, (6.66 - cap_h - leg_h - head_y - 0.50) / max(nrows, 1))

        for j, cl in enumerate(col_labels):
            self._label(s, left + j * col_w, head_y, col_w - 0.10, 0.44, cl,
                        CAPTION_PT + 2, t.muted, bold=True,
                        align=PP_ALIGN.CENTER, spacing=1.15, check=False)

        for i, rl in enumerate(row_labels):
            y = head_y + 0.50 + i * row_h
            band = self._shape(s, MSO_SHAPE.RECTANGLE, self.margin, y - 0.05,
                               self.slide_w - 2 * self.margin, row_h - 0.04,
                               fill=(t.tint if i % 2 == 0 else None))
            if band is not None:
                band.line.fill.background()
            rsz = self._shrink([rl], row_width - 0.24, row_h - 0.12, BODY_PT - 2,
                               floor=12, spacing=1.2)
            self._label(s, self.margin + 0.12, y + 0.02, row_width - 0.24,
                        row_h - 0.12, rl, rsz, t.ink, bold=True,
                        anchor=MSO_ANCHOR.MIDDLE, spacing=1.2)
            for j, m in enumerate(marks[i]):
                cx = left + j * col_w + col_w / 2
                cy = y + row_h / 2 - 0.06
                size = 0.30
                if m is None:
                    self._line(s, cx - 0.12, cy + 0.06, cx + 0.12, cy + 0.06,
                               t.muted, 1.5)
                    continue
                self._shape(s, MSO_SHAPE.OVAL, cx - size / 2, cy - size / 2,
                            size, size, fill=(t.accent if m else None),
                            line=(None if m else t.muted))
                if highlight == (i, j):
                    ring = self._shape(s, MSO_SHAPE.OVAL, cx - 0.34, cy - 0.34,
                                       0.68, 0.68, fill=None, line=t.accent)
                    ring.line.width = Pt(1.25)

        base_y = head_y + 0.50 + nrows * row_h
        if legend:
            x = self.margin + 0.12
            for kind, text in legend:
                if kind == "on":
                    self._shape(s, MSO_SHAPE.OVAL, x, base_y + 0.10, 0.20, 0.20,
                                fill=t.accent)
                elif kind == "off":
                    self._shape(s, MSO_SHAPE.OVAL, x, base_y + 0.10, 0.20, 0.20,
                                fill=None, line=t.muted)
                else:
                    self._line(s, x, base_y + 0.20, x + 0.20, base_y + 0.20,
                               t.muted, 1.5)
                self._label(s, x + 0.30, base_y + 0.02, 3.0, 0.34, text,
                            CAPTION_PT, t.muted, check=False)
                x += 0.30 + 0.12 + self._text_w(text, CAPTION_PT) + 0.44
        if caption:
            cy2 = 6.72 - cap_h
            csz = self._shrink([caption], self.slide_w - 2 * self.margin, cap_h,
                               CAPTION_PT, floor=10, spacing=1.25)
            self._label(s, self.margin, cy2, self.slide_w - 2 * self.margin,
                        cap_h, caption, csz, t.muted, spacing=1.25)
        self._notes(s, notes)
        return s

    # ---------------------------------------------------------------- delta
    def delta(self, title, pairs, *, kicker="", unit="", caption="", cite="",
              notes="", max_value=None):
        """
        Before and after on one track, so the ratio is visible rather than stated.

        `pairs` = [(label, before, after, before_text, after_text)].
        A table would deliver the numbers; a track delivers the *size* of the
        change, which is the part an audience keeps.
        """
        s, top = self._page(title, kicker, cite)
        t = self.theme
        n = len(pairs)
        cap_h = 0.62 if caption else 0.0
        unit_h = 0.34 if unit else 0.0
        band_top = top + unit_h + 0.10
        band_h = 6.72 - cap_h - band_top
        row_h = band_h / max(n, 1)
        track_l = self.margin + 3.4
        track_w = self.slide_w - 2 * self.margin - 3.4 - 0.75
        vmax = max_value or max(max(p[1], p[2]) for p in pairs)

        if unit:
            self._label(s, self.margin, top + 0.04, 7.0, 0.30, unit,
                        CAPTION_PT, t.muted, check=False)

        for i, (label, before, after, btxt, atxt) in enumerate(pairs):
            y = band_top + i * row_h
            self._label(s, self.margin, y + 0.02, 3.2, 0.34, label, BODY_PT - 2,
                        t.ink, bold=True, check=False)
            base = y + row_h * 0.62
            self._bar(s, track_l, base - 0.012, track_w, 0.024, t.panel,
                      name="track-bar")
            bx = track_l + track_w * (before / vmax)
            ax = track_l + track_w * (after / vmax)
            span = self._shape(s, MSO_SHAPE.RECTANGLE, bx, base - 0.05,
                               max(ax - bx, 0.03), 0.10, fill=t.veil)
            span.line.fill.background()
            self._shape(s, MSO_SHAPE.OVAL, bx - 0.09, base - 0.09, 0.18, 0.18,
                        fill=t.bg_light, line=t.muted)
            self._shape(s, MSO_SHAPE.OVAL, ax - 0.11, base - 0.11, 0.22, 0.22,
                        fill=t.accent)
            self._label(s, bx - 0.55, base - 0.42, 1.1, 0.30, btxt,
                        CAPTION_PT + 1, t.muted, align=PP_ALIGN.CENTER,
                        check=False)
            self._label(s, ax - 0.30, base + 0.14, 1.4, 0.32, atxt,
                        CAPTION_PT + 2, t.accent, bold=True, check=False)

        if caption:
            cy2 = 6.72 - cap_h
            csz = self._shrink([caption], self.slide_w - 2 * self.margin, cap_h,
                               CAPTION_PT, floor=10, spacing=1.25)
            self._label(s, self.margin, cy2, self.slide_w - 2 * self.margin,
                        cap_h, caption, csz, t.muted, spacing=1.25)
        self._notes(s, notes)
        return s

    # ---------------------------------------------------------------- ledger
    def ledger(self, title, items, *, kicker="", cite="", notes="", intro=""):
        """
        A numbered list where the numerals hang in the margin and no boxes appear.

        Five limitations in five cards is a page of five equal rectangles, which
        is the "everything weighs the same" complaint in its purest form. A
        ledger carries the same information with the emphasis in the numeral and
        the rule under each row.
        """
        s, top = self._page(title, kicker, cite)
        t = self.theme
        if intro:
            self._label(s, self.margin, top - 0.08,
                        self.slide_w - 2 * self.margin, 0.36, intro,
                        BODY_PT - 3, t.muted, spacing=1.3)
            top += 0.42
        n = len(items)
        avail = 6.66 - top
        row_h = avail / n
        num_w = 0.76
        for i, item in enumerate(items):
            tag, text = (item[0], item[1]) if isinstance(item, (tuple, list)) \
                else ("", item)
            y = top + i * row_h
            self._label(s, self.margin, y, num_w, 0.44, f"{i + 1:02d}",
                        BODY_PT + 4, t.accent, bold=True, check=False)
            x = self.margin + num_w
            w = self.slide_w - 2 * self.margin - num_w
            ty = y
            if tag:
                self._label(s, x, ty, w, 0.42, tag, BODY_PT - 3, t.ink,
                            bold=True, spacing=1.0, check=False)
                ty += 0.40
            h_avail = row_h - (ty - y) - 0.10
            bsz = self._shrink([text], w, h_avail, BODY_PT - 3, floor=11)
            self._label(s, x, ty, w, h_avail, text, bsz, t.muted, spacing=1.35)
            if i < n - 1:
                self._line(s, self.margin, y + row_h - 0.04,
                           self.slide_w - self.margin, y + row_h - 0.04,
                           t.panel, 0.75)
        self._notes(s, notes)
        return s

    # ---------------------------------------------------------------- focus
    def focus(self, title, lead, satellites, *, kicker="", cite="", notes="",
              lead_tag="", sat_label=""):
        """
        One tile that dominates, three that support it.

        A 2x2 grid says "these four are equally important", which for "what to do
        next" is usually false: one direction matters and the rest are follow-ups.
        Emphasis by area is the honest way to say so.
        """
        s, top = self._page(title, kicker, cite)
        t = self.theme
        total_w = self.slide_w - 2 * self.margin
        left_w = total_w * 0.54
        right_x = self.margin + left_w + 0.42
        right_w = self.slide_w - self.margin - right_x
        h = 6.68 - top - 0.04
        pad = 0.30

        self._shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, self.margin, top, left_w, h,
                    fill=t.tint, line=None, adj=[0.045])
        y = top + 0.26
        if lead_tag:
            self._label(s, self.margin + pad, y, left_w - 2 * pad, 0.30, lead_tag,
                        CAPTION_PT + 2, t.accent, bold=True, check=False)
            y += 0.32
        hsz = self._shrink([lead[0]], left_w - 2 * pad, 0.76, 30, floor=20,
                           spacing=1.05)
        self._label(s, self.margin + pad, y, left_w - 2 * pad, 0.80, lead[0],
                    hsz, t.ink, bold=True, spacing=1.05)
        y += 0.88
        if len(lead) > 1:
            body_h = top + h - pad - y
            bsz = self._shrink([lead[1]], left_w - 2 * pad, body_h, BODY_PT,
                               floor=12, spacing=1.45)
            self._label(s, self.margin + pad, y, left_w - 2 * pad, body_h,
                        lead[1], bsz, t.muted, spacing=1.45)

        n = len(satellites)
        gap = 0.12
        rh = (h - gap * (n - 1)) / n
        for i, sat in enumerate(satellites):
            yy = top + i * (rh + gap)
            head_h = 0.34
            if sat_label and i == 0:
                self._label(s, right_x, yy, right_w, 0.26, sat_label,
                            CAPTION_PT, t.muted, check=False)
                yy += 0.24
                head_h = 0.30
            hh = self._shrink([sat[0]], right_w, head_h, BODY_PT - 1, floor=13,
                              spacing=1.0)
            self._label(s, right_x, yy, right_w, head_h, sat[0], hh, t.ink,
                        bold=True, spacing=1.0, check=False)
            body_top = yy + head_h + 0.02
            body_h = top + (i + 1) * (rh + gap) - gap - body_top - 0.02
            stxt = sat[1] if len(sat) > 1 else ""
            ssz = self._shrink([stxt], right_w, body_h, BODY_PT - 4, floor=11,
                               spacing=1.3)
            self._label(s, right_x, body_top, right_w, body_h, stxt, ssz,
                        t.muted, spacing=1.3)
            if i < n - 1:
                self._line(s, right_x, yy + rh - 0.02, right_x + right_w,
                           yy + rh - 0.02, t.panel, 0.75)
        self._notes(s, notes)
        return s

    # ---------------------------------------------------------- schematic
    def distance_schematic(self, title, states, *, kicker="", caption="",
                           cite="", notes="", unit_label="间距"):
        """
        Two states of the same machine, drawn, with the distance called out.

        This is the page a paper talk cannot do without and a layout library
        cannot fake: the argument *is* a distance, so the distance has to be on
        screen as a length you can compare between panels. Everything is made of
        ellipses, rounded bars and one dashed circle, so it stays editable.

        states: [{"label", "gap", "radius", "gap_text", "phospho", "note"}]
        """
        s, top = self._page(title, kicker, cite)
        t = self.theme
        n = len(states)
        gapx = 0.50
        panel_w = (self.slide_w - 2 * self.margin - gapx * (n - 1)) / n
        cap_h = 0.62 if caption else 0.0
        panel_top = top
        panel_bot = 6.70 - cap_h

        for i, st in enumerate(states):
            px = self.margin + i * (panel_w + gapx)
            cx = px + panel_w / 2
            gap = st.get("gap", 1.2)
            radius = st.get("radius", 1.02)
            note = st.get("note", "")
            note_h = 0.44 if note else 0.0
            dist_h = 0.62

            y_label = panel_top
            draw_top = panel_top + 0.36
            draw_bot = panel_bot - note_h - dist_h
            draw_h = draw_bot - draw_top
            if draw_h < 1.5:
                raise FitError(
                    f"slide {self._n}: 示意图可用高度只有 {draw_h:.2f}in —— "
                    f"删掉 caption 或缩短说明")
            cy = draw_top + draw_h / 2
            half = min(1.30, draw_h / 2 * 0.90)

            self._label(s, px, y_label, panel_w, 0.44, st["label"],
                        BODY_PT - 1, t.ink, bold=True,
                        align=PP_ALIGN.CENTER, check=False)

            for k, frac in enumerate((1.0, 0.70, 0.44)):
                rr = radius * 2 * frac
                ring = self._shape(s, MSO_SHAPE.OVAL, cx - rr / 2, cy - rr / 2,
                                   rr, rr, fill=(t.tint if k == 0 else t.panel))
                ring.line.fill.background()
            reach = self._shape(s, MSO_SHAPE.OVAL, cx - radius, cy - radius,
                                radius * 2, radius * 2, fill=None, line=t.accent)
            reach.line.width = Pt(1.0)
            reach.line.dash_style = 4

            for side in (-1, 1):
                kx = cx + side * gap / 2
                arm = self._shape(s, MSO_SHAPE.ROUNDED_RECTANGLE,
                                  kx - 0.16, cy - half, 0.32, half * 2,
                                  fill=t.veil, line=None, adj=[0.5])
                arm.line.fill.background()
                self._shape(s, MSO_SHAPE.ROUNDED_RECTANGLE,
                            kx - side * 0.04 - 0.05, cy - 0.28, 0.10, 0.56,
                            fill=t.ink, line=None, adj=[0.35])
                phospho = st.get("phospho", False)
                self._shape(s, MSO_SHAPE.OVAL, kx - side * 0.10 - 0.09,
                            cy - 0.09, 0.18, 0.18,
                            fill=(t.accent if phospho else t.bg_light),
                            line=(None if phospho else t.muted))
                if st.get("tension", gap > radius * 1.05):
                    ay = cy - half - 0.16
                    self._line(s, kx + side * 0.22, ay, kx + side * 0.74, ay,
                               t.accent, 1.5)
                    tip = self._shape(s, MSO_SHAPE.ISOSCELES_TRIANGLE,
                                      kx + side * 0.74 - 0.10, ay - 0.10,
                                      0.20, 0.20, fill=t.accent)
                    tip.rotation = 90 if side > 0 else 270

            dy = draw_bot + 0.12
            self._bar(s, cx - gap / 2, dy - 0.010, gap, 0.020, t.muted)
            for x0 in (cx - gap / 2, cx + gap / 2):
                self._bar(s, x0 - 0.008, dy - 0.07, 0.016, 0.14, t.muted)
            dtext = f"{unit_label} {st.get('gap_text', '')}".strip()
            self._label(s, px + 0.10, dy + 0.14, panel_w - 0.20, 0.36, dtext,
                        CAPTION_PT, t.muted, align=PP_ALIGN.CENTER, check=False)

            if note:
                nsz = self._shrink([note], panel_w - 0.20, note_h, BODY_PT - 5,
                                   floor=10, spacing=1.25)
                self._label(s, px + 0.10, panel_bot - note_h, panel_w - 0.20,
                            note_h, note, nsz, t.muted,
                            align=PP_ALIGN.CENTER, spacing=1.25)

        if caption:
            cy2 = 6.72 - cap_h
            csz = self._shrink([caption], self.slide_w - 2 * self.margin, cap_h,
                               CAPTION_PT, floor=10, spacing=1.25)
            self._label(s, self.margin, cy2, self.slide_w - 2 * self.margin,
                        cap_h, caption, csz, t.muted, spacing=1.25)
        self._notes(s, notes)
        return s
