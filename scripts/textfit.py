"""
textfit.py - Deterministic text layout metrics for PPTX QA.

WHY THIS EXISTS
---------------
The bundled LibreOffice kit in this environment does not honour a text box's
declared width when wrapping CJK text: every run is laid out against the full
canvas width instead. Rendering alone therefore cannot certify that a slide's
text fits its box. This module computes layout from real font metrics instead,
so overflow is detected from the *file*, not from a picture of the file.

The rules mirror how PowerPoint lays text out:
  * explicit line breaks  -> always break
  * spaces                -> break opportunities
  * CJK ideographs/kana   -> break allowed between any two characters
  * Latin words           -> kept whole, broken only if a single word is wider
                             than the available width
  * soft hyphens          -> break opportunity
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from functools import lru_cache

from PIL import ImageFont

EMU_PER_INCH = 914400
PT_PER_INCH = 72.0

# Characters that may start a line in Chinese typography without violating rules.
_NO_LINE_START = set("，。、；：？！）】》」』〉”’%‰℃·…—～")
_NO_LINE_END = set("（【《「『〈“‘")

_FONT_CANDIDATES = {
    "default": [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simsun.ttc",
                r"C:\Windows\Fonts\Arial.ttf"],
    "sans":    [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\Arial.ttf"],
    "serif":   [r"C:\Windows\Fonts\simsun.ttc", r"C:\Windows\Fonts\Arial.ttf"],
}


@lru_cache(maxsize=512)
def _load_font(size_px: int, bold: bool, kind: str) -> ImageFont.FreeTypeFont:
    """Load a real TrueType face so widths are measured, not guessed."""
    cands = list(_FONT_CANDIDATES.get(kind, _FONT_CANDIDATES["default"]))
    if bold:
        cands.insert(0, r"C:\Windows\Fonts\msyhbd.ttc")
    last = None
    for path in cands:
        try:
            return ImageFont.truetype(path, size_px)
        except Exception as exc:  # pragma: no cover - depends on host fonts
            last = exc
    raise RuntimeError(f"No usable font for metrics: {last}")


def measure(text: str, size_pt: float, bold: bool = False) -> float:
    """Width of `text` in points, measured with real font metrics."""
    if not text:
        return 0.0
    # Measure at a large px size then scale, to keep hinting out of the result.
    px = 200
    font = _load_font(px, bold, "default")
    width_px = font.getlength(text)
    return width_px * (size_pt / px)


def _is_cjk(ch: str) -> bool:
    code = ord(ch)
    return (
        0x4E00 <= code <= 0x9FFF      # CJK Unified Ideographs
        or 0x3400 <= code <= 0x4DBF   # Extension A
        or 0x3040 <= code <= 0x30FF   # Hiragana / Katakana
        or 0xAC00 <= code <= 0xD7AF   # Hangul
        or 0xFF00 <= code <= 0xFF60   # Fullwidth forms
        or 0x3000 <= code <= 0x303F   # CJK punctuation
        or 0xF900 <= code <= 0xFAFF   # Compatibility ideographs
    )


def _tokens(text: str):
    """Split into break-atomic units with a flag: may a line break precede it?"""
    out = []
    buf = ""
    for ch in text:
        if ch == " ":
            if buf:
                out.append((buf, True)); buf = ""
            out.append((" ", False))       # collapse later; space attaches after
        elif _is_cjk(ch):
            if buf:
                out.append((buf, True)); buf = ""
            out.append((ch, False))
        else:
            # Latin/digit run: keep together as one word
            if buf and _is_cjk(buf[-1]):
                out.append((buf, True)); buf = ""
            buf += ch
    if buf:
        out.append((buf, True))
    return out


@dataclass
class LayoutResult:
    lines: list[str] = field(default_factory=list)
    truncated: list[str] = field(default_factory=list)
    width_pt: float = 0.0
    height_pt: float = 0.0
    needed_height_pt: float = 0.0
    widest_line_pt: float = 0.0
    overflow_x: bool = False
    overflow_y: bool = False
    unbreakable_word_pt: float = 0.0

    @property
    def ok(self) -> bool:
        return not (self.overflow_x or self.overflow_y)


def wrap_text(text: str, size_pt: float, width_pt: float, bold: bool = False,
              line_spacing: float = 1.2) -> LayoutResult:
    """
    Greedy line breaking that mirrors PowerPoint's behaviour closely enough to
    catch real overflow, including the CJK case (break allowed between glyphs).
    """
    res = LayoutResult(width_pt=width_pt)
    if width_pt <= 0:
        res.overflow_x = True
        return res

    paragraphs = text.split("\n")
    lines: list[str] = []
    truncated: list[str] = []

    for para in paragraphs:
        if para == "":
            lines.append("")
            continue
        toks = _tokens(para)
        cur = ""
        for tok, breakable_before in toks:
            if tok == " " and cur.endswith(" "):
                continue
            cand = cur + tok
            if measure(cand, size_pt, bold) <= width_pt or not cur.strip():
                # single glyph wider than the box -> flag, but still place it
                if not cur.strip() and measure(tok, size_pt, bold) > width_pt:
                    res.overflow_x = True
                    res.unbreakable_word_pt = max(res.unbreakable_word_pt,
                                                  measure(tok, size_pt, bold))
                cur = cand
            else:
                # Honour Chinese line-breaking (避头尾):
                #  - closing punctuation must not START a line -> carry it down
                #  - an opening bracket must not END a line -> pull it down too
                stripped = cur.rstrip()
                carried = ""
                if stripped and stripped[-1] in _NO_LINE_START:
                    carried = stripped[-1]
                    stripped = stripped[:-1]
                if stripped and stripped[-1] in _NO_LINE_END and len(stripped) > 1:
                    carried = stripped[-1] + carried
                    stripped = stripped[:-1]
                if stripped:
                    lines.append(stripped)
                if carried:
                    truncated.append(stripped)
                cur = carried + tok.lstrip(" ")
        if cur.strip() or not cur:
            lines.append(cur.rstrip())

    res.lines = lines
    res.widest_line_pt = max((measure(l, size_pt, bold) for l in lines), default=0.0)
    if res.widest_line_pt > width_pt + 0.5:
        res.overflow_x = True
    # Height: PowerPoint uses ~1.2x font size per line (single spacing).
    res.needed_height_pt = len(lines) * size_pt * line_spacing
    return res


def check(available_w_pt: float, available_h_pt: float, text: str,
          size_pt: float, bold: bool = False, line_spacing: float = 1.2) -> LayoutResult:
    res = wrap_text(text, size_pt, available_w_pt, bold, line_spacing)
    res.height_pt = available_h_pt
    if res.needed_height_pt > available_h_pt + 0.5:
        res.overflow_y = True
    return res


if __name__ == "__main__":
    # self-check
    t = "\u6ce8\u610f\u529b\u673a\u5236\u80fd\u5426\u89e3\u91ca\u957f\u7a0b\u4f9d\u8d56\uff1f"
    for w in (2.0, 3.0, 6.0):
        r = wrap_text(t, 40, w * 72)
        print(f"width={w}in -> {len(r.lines)} lines, overflow_x={r.overflow_x}, "
              f"widest={r.widest_line_pt:.1f}pt")
