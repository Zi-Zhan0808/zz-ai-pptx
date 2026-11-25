"""
palette_audit.py - measure colour discipline, not just contrast.

Contrast checks answer "can it be read". They say nothing about the failure mode
that actually makes a deck feel loud: too many hues, too saturated, too opaque,
spread evenly so nothing leads. A page can pass every contrast test and still
read as a wash of competing rectangles.

What this measures, per deck and per slide:

  hues        distinct hue families (30-degree buckets), ignoring neutrals
  sat_max     highest saturation of any filled shape (0-1)
  sat_mean    mean saturation of non-neutral fills
  alpha       lowest opacity used (a 100% opaque panel is a wall; 90% reads as
              a tint and lets the canvas breathe)
  fills       number of distinct fills on a slide
  area        share of the slide covered by non-neutral fills, i.e. how much of
              the page is shouting
  band        vertical spread of visual weight: the y-centres of the heaviest
              shapes. One tight cluster means every element is competing along
              the same line - the "ledger" look.

Thresholds are deliberately opinionated and stated here, so the report can be
argued with instead of silently obeyed.
"""
from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from pathlib import Path

from pptx import Presentation

EMU = 914400

# Opinionated limits. A deck that trips these is not illegal, it is loud.
MAX_HUES = 4            # hue families among visible patches
MAX_SAT = 0.72          # one strong accent is fine; several are a fairground
MAX_FILLS_PER_SLIDE = 4  # visible patches, not every chip and dot
MAX_AREA = 0.62         # chroma-weighted cover of the page
MIN_DOMINANCE = 0.38    # the biggest patch must own this much of the colour
MIN_CV = 0.35           # patch sizes must vary, or nothing leads
MIN_ALPHA = 0.86        # below this a panel turns into glass
PATCH_IN2 = 0.30        # smaller than this is an accent, not a surface


# ---------------------------------------------------------------- colour math
def hex_rgb(h: str):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def rgb_hsl(rgb) -> tuple:
    r, g, b = [v / 255 for v in rgb]
    mx, mn = max(r, g, b), min(r, g, b)
    l = (mx + mn) / 2
    if mx == mn:
        return 0.0, 0.0, l
    d = mx - mn
    s = d / (2 - mx - mn) if l > 0.5 else d / (mx + mn)
    if mx == r:
        h = ((g - b) / d) % 6
    elif mx == g:
        h = (b - r) / d + 2
    else:
        h = (r - g) / d + 4
    return h * 60.0, s, l


def is_neutral(s: float, l: float) -> bool:
    """Near-grey, near-white or near-black: does not count as a hue."""
    return s < 0.12 or l > 0.95 or l < 0.10


def luminance(rgb) -> float:
    c = [v / 255 for v in rgb]
    c = [(v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4) for v in c]
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]


# ---------------------------------------------------------------- extraction
def _alpha_of(shape):
    """
    Opacity actually applied to a shape's fill, as a 0-1 fraction.

    python-pptx exposes no transparency API, so this reads the alpha child of
    the fill's colour element. Absent means fully opaque, which is the usual
    case and the one worth reporting.
    """
    xml = shape._element.xml
    spPr = re.search(r"<p:spPr.*?</p:spPr>", xml, re.S)
    if not spPr:
        return 1.0
    blk = spPr.group(0)
    ln = re.search(r"<a:ln[ >].*?</a:ln>", blk, re.S)
    if ln:
        blk = blk.replace(ln.group(0), "")
    m = re.search(r'<a:alpha val="(\d+)"', blk)
    return int(m.group(1)) / 100000.0 if m else 1.0


def _fill_hex(shape, palette):
    try:
        f = shape.fill
        ftype = str(f.type)
    except Exception:
        return None
    if "GRADIENT" in ftype:
        try:
            stop = f.gradient_stops[0].color
            return _col(stop, palette)
        except Exception:
            return None
    # MSO_FILL.BACKGROUND is what <a:noFill> reports; there is no colour to read.
    if "BACKGROUND" in ftype or ftype.endswith("None"):
        return None
    try:
        return str(f.fore_color.rgb)
    except Exception:
        return _col(f.fore_color, palette)


def _col(colfmt, palette):
    try:
        return str(colfmt.rgb)
    except Exception:
        pass
    try:
        key = str(colfmt.theme_color).split(".")[-1].split(" ")[0]
        alias = {"BACKGROUND_1": "lt1", "BACKGROUND_2": "lt2",
                 "TEXT_1": "dk1", "TEXT_2": "dk2", "ACCENT_1": "accent1",
                 "ACCENT_2": "accent2", "ACCENT_3": "accent3",
                 "ACCENT_4": "accent4", "ACCENT_5": "accent5",
                 "ACCENT_6": "accent6"}
        return palette.get(alias.get(key, ""))
    except Exception:
        return None


def theme_palette(prs) -> dict:
    """Named palette slots, via the one implementation in theme_library."""
    from theme_library import theme_palette as _tp
    return _tp(prs)


def canvas_hex(prs) -> str:
    pal = theme_palette(prs)
    return pal.get("lt1", "FFFFFF")


def _bleed_canvas(prs, pal, first_slide=None):
    """
    The page colour, taken from the deck rather than from the theme.

    A builder that paints its own background draws a full-bleed rectangle; a
    deck built inside a template leaves the canvas to the master. Reading only
    the theme gets this wrong in the first case and then counts the entire page
    as "coloured fill", which makes a calm deck look flooded.
    """
    seen = Counter()
    for s in prs.slides:
        for sh in s.shapes:
            if ((sh.width or 0) >= prs.slide_width * 0.99
                    and (sh.height or 0) >= prs.slide_height * 0.99):
                h = _fill_hex(sh, pal)
                if h:
                    seen[h.upper()] += 1
    if seen and sum(seen.values()) * 2 >= len(prs.slides):
        # Two page colours are common: a deep cover, a light body. Only trust the
        # rects when they are the majority; a deck built inside a template draws
        # none at all and takes its canvas from the master, and reading the two
        # bookend covers as "the canvas" would call the whole document brand
        # green and then measure every body page as flooded.
        return seen.most_common(1)[0][0]
    return canvas_hex(prs).upper()


def collect(path, canvas=None):
    prs = Presentation(str(path))
    pal = theme_palette(prs)
    if canvas is None:
        first = prs.slides[0] if len(prs.slides) else None
        canvas = _bleed_canvas(prs, pal, first) if first else canvas_hex(prs)
    slides = []
    for i, s in enumerate(prs.slides, 1):
        fills = []
        for sh in s.shapes:
            h = _fill_hex(sh, pal)
            if not h:
                continue
            # A shape covering the whole page *is* the page, whatever its
            # colour; counting it as fill makes any deck look 100% flooded.
            if ((sh.width or 0) >= prs.slide_width * 0.99
                    and (sh.height or 0) >= prs.slide_height * 0.99):
                continue
            try:
                alpha = _alpha_of(sh)
            except Exception:
                alpha = 1.0
            w = (sh.width or 0) / EMU
            hh = (sh.height or 0) / EMU
            top = (sh.top or 0) / EMU
            fills.append({"hex": h.upper(), "alpha": round(alpha, 3),
                          "area": round(w * hh, 3), "top": round(top, 2),
                          "h": round(hh, 2), "name": sh.name})
        slides.append({"index": i, "fills": fills})
    return {"canvas": canvas.upper(), "slides": slides}


def audit(data) -> dict:
    """
    Judge colour as composition: how many voices, how loud, and who leads.

    Two filters do most of the work. A full-bleed shape is the page, not a fill.
    And a patch below PATCH_IN2 is an accent - a chip, a dot, a rule - so it
    counts towards neither the hue count nor the flood, but it also cannot be
    the thing that leads a page.

    The dominance and rhythm numbers are the ones that answer "看似塞满实则有点乱":
    eight patches of similar size, none of them larger than the rest, is a page
    with no first thing to look at - even if every one of them is tasteful.
    """
    canvas = data["canvas"]
    hue_counter, sat_vals, alphas = Counter(), [], []
    per_slide = []

    for s in data["slides"]:
        patches = []
        for f in s["fills"]:
            if f["area"] < PATCH_IN2:
                continue
            patches.append(f)
            alphas.append(f["alpha"])
            h, sat, l = rgb_hsl(hex_rgb(f["hex"]))
            if not is_neutral(sat, l):
                hue_counter[int(h // 30) * 30] += 1
                # A deep, saturated block carrying white type reads as lacquer,
                # which is a legitimate emphasis device. What reads as "pastel
                # gone wrong" is a *mid* saturated patch, so only those count
                # against the saturation budget.
                if l > 0.30:
                    sat_vals.append(sat)

        # Weight coverage by how far each patch sits from the page colour: a
        # near-white card on a white page is visually almost nothing, while a
        # saturated block of the same size is a slab. Counting raw area calls a
        # pale card grid "flooded" and lets a loud deck through.
        canvas_l = luminance(hex_rgb(canvas))

        def weight(f):
            d = abs(luminance(hex_rgb(f["hex"])) - canvas_l)
            return min(1.0, d / 0.25) * f["alpha"]

        total = sum(p["area"] for p in patches) or 1.0
        biggest = max((p["area"] for p in patches), default=0.0)
        areas = [p["area"] for p in patches]
        mean = total / len(areas) if areas else 0.0
        if len(areas) > 1 and mean:
            var = sum((a - mean) ** 2 for a in areas) / len(areas)
            cv = (var ** 0.5) / mean
        else:
            cv = 0.0
        per_slide.append({
            "index": s["index"],
            "patch_count": len(patches),
            "fills": len({p["hex"] for p in patches}),
            "coverage": round(sum(p["area"] * weight(p) for p in patches)
                              / (13.333 * 7.5), 3),
            "dominance": round(biggest / total, 3) if patches else 0.0,
            "size_cv": round(cv, 3),
        })

    issues = []
    n_hues = len(hue_counter)
    if n_hues > MAX_HUES:
        issues.append({
            "type": "too-many-hues", "value": n_hues, "limit": MAX_HUES,
            "detail": f"{n_hues} 个色相家族（每 30° 一档），上限 {MAX_HUES}；"
                      f"色相越多，读者越难判断「哪个颜色在说话」",
            "hues": sorted(hue_counter)})

    smax = max(sat_vals) if sat_vals else 0.0
    smean = sum(sat_vals) / len(sat_vals) if sat_vals else 0.0
    if smax > MAX_SAT:
        issues.append({
            "type": "saturation", "value": round(smax, 3), "limit": MAX_SAT,
            "detail": f"可见色块最高饱和度 {smax:.0%} 超过 {MAX_SAT:.0%}；"
                      f"淡彩应当靠亮度承载，而不是靠饱和度"})

    amin = min(alphas) if alphas else 1.0
    if amin < MIN_ALPHA:
        issues.append({
            "type": "transparency", "value": round(amin, 3), "limit": MIN_ALPHA,
            "detail": f"最低不透明度 {amin:.0%} 低于 {MIN_ALPHA:.0%}；"
                      f"半透明色块叠在白底上会变成灰雾，投影和打印各差一档"})

    busy = [p for p in per_slide if p["fills"] > MAX_FILLS_PER_SLIDE]
    if busy:
        worst = max(busy, key=lambda p: p["fills"])
        issues.append({
            "type": "fills-per-slide", "value": worst["fills"],
            "limit": MAX_FILLS_PER_SLIDE,
            "detail": f"第 {worst['index']} 页有 {worst['fills']} 种可见填充色，"
                      f"上限 {MAX_FILLS_PER_SLIDE}"})

    flood = [p for p in per_slide if p["coverage"] > MAX_AREA]
    if flood:
        issues.append({
            "type": "area-flood", "value": len(flood), "limit": MAX_AREA,
            "detail": f"{len(flood)} 页的可见色块覆盖超过 {MAX_AREA:.0%} 版面；"
                      f"看似塞满实则没有留白呼吸"})

    # Hierarchy: a page with several patches needs one of them to lead, and the
    # rest to differ enough in size that the eye can rank them.
    # "No focal point" is about *colour* competing, not about equal boxes. A row
    # of three cards in one colour is a grid, and equal sizes are the design; a
    # row of eight cards in five colours, none of them larger, is a collage with
    # nowhere to look first. So the test needs several distinct fills, not just
    # several patches.
    leaderless = [p for p in per_slide
                  if p["patch_count"] >= 3 and p["fills"] >= 3
                  and p["dominance"] < MIN_DOMINANCE]
    if len(leaderless) >= max(2, len(per_slide) // 3):
        issues.append({
            "type": "no-focal-point", "value": len(leaderless),
            "limit": MIN_DOMINANCE,
            "detail": f"{len(leaderless)} 页的色块面积几乎均等"
                      f"（最大一块仅占 {min(p['dominance'] for p in leaderless):.0%}），"
                      f"没有主次；应当让一块明显占主导，其余退成陪衬"})
    mono = [p for p in per_slide
            if p["patch_count"] >= 3 and p["fills"] >= 3
            and p["size_cv"] < MIN_CV]
    if len(mono) >= max(2, len(per_slide) // 3):
        issues.append({
            "type": "flat-rhythm", "value": len(mono), "limit": MIN_CV,
            "detail": f"{len(mono)} 页的色块大小过于接近（变异系数 "
                      f"{min(p['size_cv'] for p in mono):.2f}），排布像流水账；"
                      f"大小要有梯度才能读出轻重"})

    return {
        "canvas": canvas,
        "hues": sorted(hue_counter),
        "hue_families": n_hues,
        "saturation_max": round(smax, 3),
        "saturation_mean": round(smean, 3),
        "alpha_min": round(amin, 3),
        "per_slide": per_slide,
        "issues": issues,
    }


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="配色纪律体检")
    ap.add_argument("pptx")
    ap.add_argument("--out", default="audit")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)

    rep = audit(collect(a.pptx))
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    (out / "palette_audit.json").write_text(
        json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")

    if a.json:
        print(json.dumps(rep, ensure_ascii=False, indent=2))
        return 0
    print(f"画布 #{rep['canvas']}")
    print(f"色相家族 {rep['hue_families']} {rep['hues']}")
    print(f"饱和度   最高 {rep['saturation_max']:.0%}  平均 {rep['saturation_mean']:.0%}")
    print(f"最低不透明度 {rep['alpha_min']:.0%}")
    print("每页：色块数 / 色数 / 覆盖 / 主元素占比 / 大小变异")
    for p in rep["per_slide"]:
        print(f"  p{p['index']:<3} {p['patch_count']} / {p['fills']} / "
              f"{p['coverage']:.0%} / {p['dominance']:.0%} / {p['size_cv']:.2f}")
    print(f"\n问题 {len(rep['issues'])} 条")
    for i in rep["issues"]:
        print(f"  [{i['type']}] {i['detail']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
