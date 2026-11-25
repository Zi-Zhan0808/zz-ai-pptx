"""
design_check.py - Structural design lint that encodes failures already hit.

WHY THIS EXISTS
---------------
Every rule here was discovered by shipping a deck with the defect in it:

  * an agenda whose divider rules ended up on only one column ("两个意义不明的横杠")
  * a flowchart whose first step was a flat-sided PENTAGON while the rest were
    CHEVRONs ("第一个背景箭头有小 bug")
  * a title claiming "七步" over a diagram of five steps
  * decorative rules that align with nothing
  * mixed arrow shapes inside one flow

None of these are text-overflow bugs, so the existing metrics cannot see them.
They are geometry/text-consistency bugs, and the only reliable fix is to check
them mechanically every time instead of remembering to look.

Usage:
  python design_check.py deck.pptx --out audit/
  python design_check.py deck.pptx --out audit/ --json
"""
from __future__ import annotations

import argparse
import json
import re
import sys

from pptx.oxml.ns import qn
from collections import Counter
from pathlib import Path

EMU_IN = 914400

# IMPORTANT: these are the values stored in <a:prstGeom prst="...">, NOT the
# python-pptx constant names. MSO_SHAPE.PENTAGON is written as "homePlate"; a
# checker that compares against "PENTAGON" silently matches nothing and reports
# a clean deck. `_selftest_preset_names()` at the bottom of this file locks the
# mapping down.
POINTED = {
    "chevron",              # MSO_SHAPE.CHEVRON
    "homePlate",            # MSO_SHAPE.PENTAGON  (flat left, pointed right)
    "rightArrow",           # MSO_SHAPE.RIGHT_ARROW
    "notchedRightArrow",    # MSO_SHAPE.NOTCHED_RIGHT_ARROW
    "stripedRightArrow",
    "rightArrowCallout",
}
FLAT_SIDED = {
    "rect", "roundRect", "trapezoid", "parallelogram",
}
BARS = {"rightBrace", "leftBrace", "round2SameRect"}

NUMERAL_WORDS = {
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8,
    "九": 9, "十": 10, "两": 2,
}


def _inches(emu) -> float:
    return (emu or 0) / EMU_IN


def _shape_kind(shape) -> str:
    """Best-effort preset name for an autoshape ('' when not one)."""
    try:
        from pptx.oxml.ns import qn
        prst = shape._element.find(".//" + qn("a:prstGeom"))
        if prst is not None:
            return prst.get("prst", "")
    except Exception:
        pass
    return ""


def _is_line(shape) -> bool:
    try:
        return str(shape.shape_type).endswith("LINE") or _shape_kind(shape) == ""
    except Exception:
        return True


def collect_slides(pptx_path: Path):
    from pptx import Presentation
    prs = Presentation(str(pptx_path))
    out = []
    for i, slide in enumerate(prs.slides, start=1):
        shapes = []
        for sh in slide.shapes:
            item = {
                "name": sh.name,
                "kind": _shape_kind(sh),
            "name": sh.name or "",
                "type": str(sh.shape_type),
                "x": _inches(sh.left), "y": _inches(sh.top),
                "w": _inches(sh.width), "h": _inches(sh.height),
                "text": (sh.text_frame.text if sh.has_text_frame else ""),
                "connector": "CONNECTOR" in str(sh.shape_type),
            }
            shapes.append(item)
        # the background is the full-bleed rectangle drawn first
        palette = theme_colors(prs)
        bg = None
        for sh in slide.shapes:
            if (sh.width or 0) >= prs.slide_width * 0.99 and \
               (sh.height or 0) >= prs.slide_height * 0.99:
                bg = _fill_hex(sh, palette)
                if bg:
                    break
        if bg is None:
            # A deck built inside a template has no per-slide background
            # rectangle: the canvas comes from the master, so a check that only
            # looks at slide shapes sees nothing and passes everything. Read the
            # master instead.
            bg = _master_bg(slide, prs, palette)
        out.append({"index": i, "shapes": shapes, "background": bg,
                    "text": "\n".join(s["text"] for s in shapes if s["text"].strip())})
    return out


_SCHEME_ALIAS = {
    "bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
}
_THEME_ALIAS = {
    "BACKGROUND_1": "lt1", "BACKGROUND_2": "lt2",
    "TEXT_1": "dk1", "TEXT_2": "dk2",
    "ACCENT_1": "accent1", "ACCENT_2": "accent2", "ACCENT_3": "accent3",
    "ACCENT_4": "accent4", "ACCENT_5": "accent5", "ACCENT_6": "accent6",
    "HYPERLINK": "hlink", "FOLLOWED_HYPERLINK": "folHlink",
}


def theme_colors(prs) -> dict:
    """Named palette slots, via the one implementation in theme_library."""
    from theme_library import theme_palette
    pal = theme_palette(prs)
    pal.setdefault("lt1", "FFFFFF")
    pal.setdefault("dk1", "000000")
    return pal


def _color_hex(colfmt, palette):
    """One ColorFormat -> hex, following theme references if necessary."""
    try:
        return str(colfmt.rgb)
    except Exception:
        pass
    try:
        key = str(colfmt.theme_color).split(".")[-1].split(" ")[0]
        return palette.get(_THEME_ALIAS.get(key, ""))
    except Exception:
        return None


def _fill_hex(sh, palette):
    """
    Colour of a shape's fill, resolving theme references and gradients.

    Three different ways a template says "green": an srgb literal, a theme
    reference ("accent 1"), or a gradient between two of them. All three are
    common - branded covers are almost always gradients - and reading only the
    first turns the background test into a no-op precisely on the slides that
    decide whether a deck flashes.

    Every branch here can legitimately fail on some shape, and a checker that
    raises is worse than one that returns "unknown": an exception aborts the
    whole audit, so one unreadable fill hides every other finding in the deck.
    """
    try:
        f = sh.fill
        ftype = str(f.type)
    except Exception:
        return None
    if "GRADIENT" in ftype:
        try:
            stops = [_color_hex(s.color, palette) for s in f.gradient_stops]
            got = _average_hex(stops)
            if got:
                return got
        except Exception:
            pass
    if "BACKGROUND" in ftype or ftype.endswith("None"):
        return None
    try:
        return str(f.fore_color.rgb)
    except Exception:
        pass
    try:
        return _color_hex(f.fore_color, palette)
    except Exception:
        return None


def _master_bg(slide, prs, palette=None):
    """Canvas colour inherited from the layout/master (template decks)."""
    palette = palette or theme_colors(prs)
    try:
        cSld = slide.slide_layout.slide_master._element.find(qn("p:cSld"))
        bg = cSld.find(qn("p:bg")) if cSld is not None else None
        if bg is not None:
            xml = bg.xml
            m = re.search(r'<a:schemeClr val="(\w+)"', xml)
            if m:
                return palette.get(_SCHEME_ALIAS.get(m.group(1), m.group(1)))
            m = re.search(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', xml)
            if m:
                return m.group(1).upper()
    except Exception:
        pass
    return palette.get("lt1")


# ---------------------------------------------------------------- checks
def check_flow_shapes(slides):
    """A horizontal flow must use one consistent pointed shape."""
    issues = []
    for sl in slides:
        pointed = [s for s in sl["shapes"] if s["kind"] in POINTED]
        if len(pointed) < 2:
            continue
        # only shapes on a shared band form one flow
        bands: dict[int, list] = {}
        for s in pointed:
            bands.setdefault(round(s["y"], 1), []).append(s)
        for y, group in bands.items():
            if len(group) < 2:
                continue
            kinds = Counter(s["kind"] for s in group)
            if len(kinds) > 1:
                # the first step legitimately has no incoming notch; require it
                # to be the *same* visual family (pointed on the right)
                first = min(group, key=lambda s: s["x"])
                detail = ", ".join(f"{k}×{v}" for k, v in kinds.items())
                issues.append({
                    "type": "flow-shape-mixed",
                    "slide": sl["index"],
                    "detail": f"flow row mixes arrow shapes: {detail}",
                    "hint": (f"first step is {first['kind']}; a flat-sided shape "
                             f"next to chevrons reads as a broken arrow — use "
                             f"CHEVRON for every step, or PENTAGON for all"),
                })
    return issues


def check_rules(slides, content_w=11.53):
    """
    A horizontal rule must be anchored to something the reader can name.

    Allowed:
      * full-width separator (>= 85% of the content width)
      * anchor rule: sits within 0.6in below a text block, directly under it
      * repeated at the same y in every column of a multi-column list
    Anything else is a stray bar: it divides nothing and reads as leftover.
    """
    issues = []
    for sl in slides:
        # A measured span (a distance, a before/after track) is data drawn as a
        # hairline. It is not a divider, so the "unanchored rule" rule does not
        # apply to it - hence the builder names such shapes "measure-*" / "track-*".
        rules = [s_ for s_ in sl["shapes"]
                 if not str(s_.get("name", "")).startswith(("measure-", "track-"))
                 and (s_["connector"] or (s_["h"] < 0.06 and s_["w"] > 0.8))
                 and s_["w"] < content_w * 0.85]
        if not rules:
            continue
        texts = [t for t in sl["shapes"] if t["text"].strip() and not t["connector"]]
        rows: dict[float, list] = {}
        for r in rules:
            rows.setdefault(round(r["y"], 1), []).append(r)

        for y, group in rows.items():
            # (a) anchor rule: a text block ends just above it, same x
            anchored = False
            for t in texts:
                ty_end = t["y"] + t["h"]
                if 0 <= y - ty_end <= 0.62 and abs(t["x"] - group[0]["x"]) < 0.35:
                    anchored = True
                    break
            if anchored:
                continue
            # (b) repeated across columns at the same y
            if len({round(g["x"], 1) for g in group}) >= 2:
                continue
            # (c) a lone divider between list items with no column partner
            issues.append({
                "type": "unanchored-rule",
                "slide": sl["index"],
                "detail": (f"rule at y={y:.2f}in, x={group[0]['x']:.2f}in, "
                           f"width={group[0]['w']:.2f}in: not full width, not under "
                           f"a heading, and not repeated in other columns"),
                "hint": ("drop the rules, or make them full-width separators, or "
                         "repeat them in every column"),
            })
    return issues


def _cn_count(text: str) -> int | None:
    """Parse a leading count such as 七步 / 4步 / 三个."""
    m = re.search(r"([一二三四五六七八九十两|\d]+)\s*(?:步|个|条|项|类|道|部分|章节)", text)
    if not m:
        return None
    tok = m.group(1)
    if tok.isdigit():
        return int(tok)
    if tok in NUMERAL_WORDS:
        return NUMERAL_WORDS[tok]
    return None


def check_title_counts(slides):
    """
    A title that states a count must match what the slide actually shows.

    Note: python-pptx writes text boxes with <a:prstGeom prst="rect"> too, so a
    text box cannot be told apart from a rectangle by its preset. Shape *type*
    is the reliable discriminator.
    """
    issues = []
    for sl in slides:
        titles = [sh for sh in sl["shapes"]
                  if sh["text"].strip()
                  and "TEXT_BOX" in sh["type"]
                  and sh["y"] < 2.4]
        if not titles:
            continue
        # the title is the longest of the top text blocks
        title = max(titles, key=lambda t: len(t["text"]))["text"]
        claimed = _cn_count(title)
        if not claimed:
            continue
        unit = 0
        pointed = [x for x in sl["shapes"] if x["kind"] in POINTED]
        if len(pointed) >= 2:
            unit = len(pointed)
        if unit and unit != claimed:
            issues.append({
                "type": "title-count-mismatch",
                "slide": sl["index"],
                "detail": f"title claims {claimed} but the diagram has {unit} steps",
                "hint": "make the number match, or drop the number from the title",
            })
    return issues


def check_orphan_decor(slides):
    """
    Flag a decorative shape that belongs to nothing.

    A shape is fine when any of these hold:
      * a text box overlaps it (icons and number chips carry their label in a
        separate box drawn on top)
      * text sits directly beside it on the same band
      * it is part of a repeated row of near-identical shapes (a card/grid row)

    Only a genuinely isolated, unlabelled patch survives these tests.
    """
    issues = []
    for sl in slides:
        texts = [s for s in sl["shapes"] if s["text"].strip()]
        decor = [s for s in sl["shapes"]
                 if not s["text"].strip() and not s["connector"]
                 and s["kind"] in FLAT_SIDED and s["w"] * s["h"] < 4.0]
        if not decor:
            continue
        for d in decor:
            dx0, dy0 = d["x"], d["y"]
            dx1, dy1 = dx0 + d["w"], dy0 + d["h"]

            def overlaps(t):
                tx0, ty0 = t["x"], t["y"]
                tx1, ty1 = tx0 + t["w"], ty0 + t["h"]
                ox = min(dx1, tx1) - max(dx0, tx0)
                oy = min(dy1, ty1) - max(dy0, ty0)
                return ox > 0.01 and oy > 0.01

            if any(overlaps(t) for t in texts):
                continue
            beside = [t for t in texts
                      if abs(t["y"] - dy0) < 0.9
                      and not (t["x"] > dx1 + 0.15 or t["x"] + t["w"] < dx0 - 0.15)]
            if beside:
                continue
            # part of a repeated motif? (same size, same row, several copies)
            twins = [o for o in decor
                     if abs(o["w"] - d["w"]) < 0.03 and abs(o["h"] - d["h"]) < 0.03
                     and abs(o["y"] - d["y"]) < 0.06]
            if len(twins) >= 2:
                continue
            issues.append({
                "type": "orphan-decoration",
                "slide": sl["index"],
                "detail": (f"{d['kind']} at ({dx0:.2f},{dy0:.2f})in "
                           f"{d['w']:.2f}x{d['h']:.2f}in has no label, no text "
                           f"beside it, and no repeating counterpart"),
                "hint": "delete it, or attach a label",
            })
    return issues



def _lum(hexs: str) -> float:
    h = hexs.lstrip("#")
    r, g, b = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]

    def f(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _ratio(a: str, b: str) -> float:
    la, lb = _lum(a), _lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


MAX_BG_JUMP = 1.75
HEAVY_L = 0.45          # below this a page reads as "a dark field"
NEAR_BLACK_L = 0.05


def _runs(seq):
    """Collapse the page sequence into consecutive light / heavy runs."""
    runs = []
    for item in seq:
        heavy = item[2] < HEAVY_L
        if runs and runs[-1]["heavy"] == heavy:
            runs[-1]["items"].append(item)
        else:
            runs.append({"heavy": heavy, "items": [item]})
    return runs


def check_background_flash(slides):
    """
    Refuse a deck whose background alternates, not one that opens heavily.

    The injury is re-adaptation: at 17.8:1 between page backgrounds the
    audience's pupils reset on every slide change. Two pages of heavy field at
    the very start and the very end do not do that - the eye adapts twice in
    twenty minutes, which is how a branded cover has always worked. So the test
    is on *alternation*, and the shape of an acceptable sequence is

        heavy, light, light, ..., light, heavy        (bookends, allowed)
        light, light, ..., light                      (uniform, allowed)

    while heavy/light/heavy/light is not. A near-black field stays an issue
    wherever it appears, because that is the case the earlier deck actually hit
    and "it is only the cover" is how it got waved through last time.
    """
    seq = []
    for sl in slides:
        bg = sl.get("background")
        if not bg:
            continue
        seq.append((sl["index"], bg, _lum(bg)))
    if len(seq) < 2:
        return []

    runs = _runs(seq)
    bookend = len(runs) <= 3 and all(
        not r["heavy"] or i in (0, len(runs) - 1)
        for i, r in enumerate(runs))
    run_of = {}
    for i, r in enumerate(runs):
        for item in r["items"]:
            run_of[item[0]] = i

    issues = []
    for i, (idx, bg, L) in enumerate(seq):
        if L < NEAR_BLACK_L:
            issues.append({
                "type": "harsh-background",
                "slide": idx,
                "detail": f"background #{bg} is near-black",
                "hint": "a near-black field reads as a hole on a projector; "
                        "use a deeper tint of the brand hue instead",
            })
        if L > 0.98:
            issues.append({
                "type": "notice-pure-white",
                "slide": idx,
                "detail": f"background #{bg} is pure white",
                "hint": "a soft tint such as #FCF8F3 is easier on the eye; "
                        "acceptable when it is the template's own canvas",
            })
        if i == 0:
            continue
        prev_idx, prev_bg, prev_L = seq[i - 1]
        boundary = run_of[idx] != run_of[prev_idx]
        if boundary and bookend:
            # A bookend transition is chosen, not repeated. Still refuse it if
            # what it transitions to or from is a near-black hole.
            if min(L, prev_L) < NEAR_BLACK_L:
                pass
            else:
                continue
        jump = _ratio(prev_bg, bg)
        if jump > MAX_BG_JUMP:
            issues.append({
                "type": "background-flash",
                "slide": idx,
                "detail": (f"background jumps {jump:.2f}:1 from slide "
                           f"{prev_idx} (#{prev_bg} -> #{bg}); "
                           f"limit is {MAX_BG_JUMP}:1"),
                "hint": "keep one bright canvas and vary only its tint; heavy "
                        "fields belong at the ends, never alternating",
            })
    return issues


def _layout_signature(sl):
    """
    A coarse name for the shape of a page, from what is on it.

    Every text box carries prst="rect", so preset names alone classify a whole
    deck as "rect". The useful signals are structural: a table, a row of
    chevrons, several cards, or a wall of text blocks.
    """
    shapes = sl["shapes"]
    if any("TABLE" in sh["type"].upper() for sh in shapes):
        return "table"
    if sum(1 for sh in shapes if sh["kind"] in ("chevron", "homePlate")) >= 2:
        return "flow"
    if sum(1 for sh in shapes if sh["kind"] == "roundRect") >= 3:
        return "cards"
    texts = [sh for sh in shapes if sh["text"].strip()]
    if len(texts) >= 7:
        return "list"
    if len(texts) <= 3:
        return "text"
    return "split"


def check_layout_variety(slides, min_pages=8, min_kinds=3, max_share=0.65):
    """
    Refuse a deck where nearly every page has the same skeleton.

    The bar is set high on purpose. "Every slide is text-left, panel-right" is a
    real complaint, but a crude signature also calls a well-varied deck uniform,
    and a report that cries wolf gets ignored - so this fires only on decks long
    enough to show a pattern (8+ pages) that use fewer than three skeletons or
    repeat one more than two thirds of the time.
    """
    content = [s for s in slides if s["shapes"]]
    if len(content) < min_pages:
        return []
    counts = {}
    for s in content:
        k = _layout_signature(s)
        counts[k] = counts.get(k, 0) + 1
    issues = []
    if len(counts) < min_kinds:
        issues.append({
            "type": "layout-monotony", "slide": content[0]["index"],
            "detail": f"{len(content)} 页只用了 {len(counts)} 种版面"
                      f"（{sorted(counts)}），至少应有 {min_kinds} 种",
            "hint": "换用 cards / flow / table / split 等不同骨架，"
                    "不要把内容都塞进左右两栏",
        })
    top_kind, top_n = max(counts.items(), key=lambda kv: kv[1])
    if top_n / len(content) > max_share:
        issues.append({
            "type": "layout-repeat", "slide": content[0]["index"],
            "detail": f"{top_n}/{len(content)} 页是同一种版面（{top_kind}），"
                      f"占比超过 {max_share:.0%}",
            "hint": "同一骨架连续出现太多页，读者会开始跳读",
        })
    return issues


def run(pptx_path: Path) -> dict:
    slides = collect_slides(pptx_path)
    issues = []
    issues += check_flow_shapes(slides)
    issues += check_rules(slides)
    issues += check_title_counts(slides)
    issues += check_orphan_decor(slides)
    issues += check_background_flash(slides)
    issues += check_layout_variety(slides)
    # `notice-*` types are observations, not defects: a pure-white canvas is a
    # preference this skill holds for decks it designs itself, and a given
    # template is allowed to differ. They are reported so nobody is surprised,
    # but they must not fail the gate - a gate that fires on a legitimate
    # template choice gets switched off by the next person, and then it is not a
    # gate at all.
    notices = [i for i in issues if i["type"].startswith("notice-")]
    issues = [i for i in issues if not i["type"].startswith("notice-")]
    counts: dict[str, int] = {}
    for i in issues:
        counts[i["type"]] = counts.get(i["type"], 0) + 1
    return {
        "file": str(pptx_path),
        "slides": len(slides),
        "issues": issues,
        "issue_counts": counts,
        "notices": notices,
        "ok": not issues,
    }


def main():
    ap = argparse.ArgumentParser(description="Structural design lint")
    ap.add_argument("pptx")
    ap.add_argument("--out", default="audit")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    rep = run(Path(args.pptx))
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    (out / "design_check.json").write_text(
        json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.json:
        print(json.dumps(rep, ensure_ascii=False, indent=2))
    else:
        print(f"slides={rep['slides']}  issues={len(rep['issues'])}  "
              f"notices={len(rep.get('notices', []))}")
        for i in rep["issues"]:
            print(f"  [{i['type']}] slide {i['slide']}: {i['detail']}")
        for i in rep.get("notices", []):
            print(f"  ({i['type']}) slide {i['slide']}: {i['detail']}")
            if i.get("hint"):
                print(f"      -> {i['hint']}")
    return 0 if rep["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
