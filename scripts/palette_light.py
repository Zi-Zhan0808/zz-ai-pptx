"""
Palettes with "淡彩" (soft colour) rather than near-grey.

Adjustment: raise the *saturation* of every background and card tint into the
38-60% band while holding lightness where it was. Lightness drives both comfort
and text contrast, so holding it fixed means the colour can be strengthened
without the page getting heavier or the reading contrast dropping.

Accents (used for a single numeral or rule, never for a whole slide) sit at
55-75% saturation so they still read as the one saturated element.
"""
from __future__ import annotations

import colorsys


def _hls_to_hex(h: int, l: int, s: int) -> str:
    r, g, b = colorsys.hls_to_rgb(h / 360, l / 100, s / 100)
    return "".join(f"{round(v * 255):02X}" for v in (r, g, b))


def tune(hexs: str, *, s: int | None = None, l: int | None = None,
         dh: int = 0) -> str:
    """Return the same colour with saturation/lightness/hue nudged."""
    h = hexs.lstrip("#")
    r, g, b = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    hh, ll, ss = colorsys.rgb_to_hls(r, g, b)
    H = (round(hh * 360) + dh) % 360
    L = round(ll * 100) if l is None else l
    S = round(ss * 100) if s is None else s
    return _hls_to_hex(H, L, S)


def lum(hexs: str) -> float:
    h = hexs.lstrip("#")
    r, g, b = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]

    def f(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def ratio(a: str, b: str) -> float:
    la, lb = lum(a), lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


# Hue anchors: warm sand, cool sky, fresh leaf, soft rose
PALETTES = {
    "linen": {
        "canvas":  _hls_to_hex(33, 97, 55),
        "panel":   _hls_to_hex(33, 89, 52),
        "card":    _hls_to_hex(33, 93, 48),
        "tint":    _hls_to_hex(33, 95, 45),
        "accent":  _hls_to_hex(16, 46, 62),
        "accent2": _hls_to_hex(38, 40, 55),
    },
    "mist": {
        "canvas":  _hls_to_hex(210, 97, 55),
        "panel":   _hls_to_hex(210, 90, 50),
        "card":    _hls_to_hex(210, 94, 42),
        "tint":    _hls_to_hex(210, 96, 40),
        "accent":  _hls_to_hex(196, 36, 58),
        "accent2": _hls_to_hex(226, 42, 52),
    },
    "sage": {
        "canvas":  _hls_to_hex(95, 97, 58),
        "panel":   _hls_to_hex(95, 90, 55),
        "card":    _hls_to_hex(95, 94, 46),
        "tint":    _hls_to_hex(95, 96, 44),
        "accent":  _hls_to_hex(30, 44, 62),
        "accent2": _hls_to_hex(150, 34, 46),
    },
    "blush": {
        "canvas":  _hls_to_hex(335, 97, 58),
        "panel":   _hls_to_hex(335, 90, 56),
        "card":    _hls_to_hex(335, 94, 46),
        "tint":    _hls_to_hex(335, 96, 42),
        "accent":  _hls_to_hex(331, 48, 52),
        "accent2": _hls_to_hex(285, 30, 48),
    },
}

INK = {"linen": "221E1A", "mist": "1B222B", "sage": "1C241B", "blush": "261D22"}
MUTED = {"linen": "645A50", "mist": "545F6C", "sage": "566152", "blush": "66555F"}


def report():
    for name, p in PALETTES.items():
        ink, muted = INK[name], MUTED[name]
        print(f"\n== {name} ==")
        for key in ("canvas", "panel", "card", "tint"):
            h = p[key].lstrip("#")
            r, g, b = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
            hh, ll, ss = colorsys.rgb_to_hls(r, g, b)
            print(f"  {key:7s} #{p[key]}  H={round(hh*360):3d} "
                  f"L={round(ll*100):3d} S={round(ss*100):3d}")
        for key in ("accent", "accent2"):
            h = p[key].lstrip("#")
            r, g, b = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
            hh, ll, ss = colorsys.rgb_to_hls(r, g, b)
            print(f"  {key:7s} #{p[key]}  H={round(hh*360):3d} "
                  f"L={round(ll*100):3d} S={round(ss*100):3d}")

        checks = [
            ("ink on canvas", ink, p["canvas"], 7.0),
            ("ink on card", ink, p["card"], 7.0),
            ("ink on tint", ink, p["tint"], 7.0),
            ("ink on panel", ink, p["panel"], 7.0),
            ("muted on canvas", muted, p["canvas"], 4.5),
            ("muted on panel", muted, p["panel"], 4.5),
            ("accent on canvas", p["accent"], p["canvas"], 3.0),
            ("accent on panel", p["accent"], p["panel"], 3.0),
        ]
        ok = True
        for label, fg, bg, need in checks:
            r = ratio(fg, bg)
            good = r >= need
            ok &= good
            print(f"    {'OK  ' if good else 'FAIL'} {r:6.2f} (need {need}) {label}")
        jump = ratio(p["canvas"], p["panel"])
        calm = jump <= 1.75
        ok &= calm
        print(f"    {'OK  ' if calm else 'FAIL'} {jump:6.2f} canvas->panel jump "
              f"(target <= 1.75)")
        print(f"  RESULT: {'PASS' if ok else 'NEEDS WORK'}")


if __name__ == "__main__":
    report()
