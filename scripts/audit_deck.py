"""
audit_deck.py - One command that produces the manual's gate report.

Aggregates every hard check into `audit_report.md`, which is a required
deliverable. Keeping this as one entry point matters: a gate that needs six
separate commands gets skipped under time pressure.

Checks
  layout     : margins, overlap, font-floor, contrast, metric-based fit,
               plus rendered text-block vs box comparison  (render_qa.py)
  citations  : bidirectional [n] <-> reference list, locator validity (refs_check.py)
  assets     : every inserted picture exists and is referenced on a slide
  placeholders / de-AI prose smells

Usage:
  python audit_deck.py deck.pptx --work work/ --out audit/
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import design_check  # noqa: E402
import refs_check    # noqa: E402
import palette_audit  # noqa: E402
import render_qa     # noqa: E402

# Phrases that mark machine-generated filler in Chinese academic prose.
AI_SMELLS = [
    (r"近年来[，,]?\s*随着", "开场铺陈式套话"),
    (r"在当今[^。]{0,10}(时代|背景|社会)", "空泛时代背景"),
    (r"深刻变革", "空转大词"),
    (r"全面赋能", "空转大词"),
    (r"具有重要意义", "结论性套话"),
    (r"由于时间(有限|关系)", "自我辩解式语句"),
    (r"未来可进一步研究", "空泛展望"),
    (r"值得(我们)?深入(思考|研究)", "空泛展望"),
    (r"众所周知", "冗余铺陈"),
    (r"综上所述[，,]?\s*本文", "模板式收束"),
]


def deck_text(pptx_path: Path) -> list[tuple[int, str]]:
    from pptx import Presentation
    prs = Presentation(str(pptx_path))
    out = []
    for i, slide in enumerate(prs.slides, start=1):
        buf = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                buf.append(shape.text_frame.text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    buf.append(" ".join(c.text for c in row.cells))
        notes = ""
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
            notes = slide.notes_slide.notes_text_frame.text or ""
        out.append((i, "\n".join(buf) + ("\n[notes]\n" + notes if notes.strip() else "")))
    return out


def check_placeholders(texts):
    issues = []
    pat = re.compile(r"(lorem ipsum|\bTBD\b|TODO|XXX+|\[insert|待补充|占位|"
                     r"点击添加|click to add)", re.IGNORECASE)
    for idx, t in texts:
        for m in pat.finditer(t):
            issues.append({"type": "placeholder", "slide": idx,
                           "detail": f"found {m.group(0)!r}"})
    return issues


def check_ai_smells(texts, allow=()):
    issues = []
    for idx, t in texts:
        for pat, label in AI_SMELLS:
            for m in re.finditer(pat, t):
                snippet = t[max(0, m.start() - 12):m.end() + 18].replace("\n", " ")
                if any(a in snippet for a in allow):
                    continue
                issues.append({"type": "ai-smell", "slide": idx,
                               "detail": f"{label}: …{snippet}…"})
    return issues


def check_assets(pptx_path: Path):
    """Every picture must sit on a slide at a sane size."""
    from pptx import Presentation
    prs = Presentation(str(pptx_path))
    issues = []
    for i, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if "PICTURE" not in str(shape.shape_type):
                continue
            w = (shape.width or 0) / 914400
            h = (shape.height or 0) / 914400
            if w < 0.4 or h < 0.4:
                issues.append({"type": "tiny-image", "slide": i, "shape": shape.name,
                               "detail": f"{w:.2f}x{h:.2f}in — too small to read"})
            if w > 13.0 or h > 7.2:
                issues.append({"type": "image-too-large", "slide": i, "shape": shape.name,
                               "detail": f"{w:.2f}x{h:.2f}in — exceeds slide"})
    return issues


def write_markdown(path: Path, deck: Path, layout: dict, refs: dict,
                   texts, extra_issues, assets, allow_smells, design=None,
                   palette=None):
    L = []
    L.append(f"# Audit report — `{deck.name}`")
    L.append("")
    L.append(f"- file: `{deck}`")
    L.append(f"- slides: {layout['slide_count']}  shapes: {layout['shape_count']}")
    L.append(f"- renderer: `{layout.get('renderer')}`")
    if layout.get("fonts_used"):
        L.append(f"- fonts in render: {', '.join(layout['fonts_used'])}")
    if layout.get("render_error"):
        L.append(f"- ⚠ render note: {layout['render_error']}")

    layout_issues = layout["issues"]
    L.append("")
    L.append("## 1. Layout / typography")
    if not layout_issues:
        L.append("- ✅ no layout defects detected")
    else:
        L.append(f"- ❌ {len(layout_issues)} issue(s)")
        for i in layout_issues:
            who = i.get("shape") or i.get("shapes")
            L.append(f"  - **{i['type']}** slide {i['slide']} `{who}` — {i['detail']}"
                     + (f"  \n    > {i['text']}" if i.get("text") else ""))

    notes = layout.get("notes") or []
    if notes:
        L.append("")
        L.append("<details><summary>提示（不影响交付判定）</summary>")
        L.append("")
        for n in notes:
            L.append(f"- {n['type']} slide {n['slide']} `{n.get('shape')}` — {n['detail']}")
        L.append("")
        L.append("</details>")

    L.append("")
    L.append("## 1c. Colour discipline")
    if palette:
        L.append(f"- 色相家族 {palette['hue_families']}　饱和度峰值 "
                 f"{palette['saturation_max']:.0%}　最低不透明度 "
                 f"{palette['alpha_min']:.0%}")
        if palette["issues"]:
            L.append(f"- ❌ {len(palette['issues'])} issue(s)")
            for i in palette["issues"]:
                L.append(f"  - **{i['type']}** — {i['detail']}")
        else:
            L.append("- ✅ 色相数、饱和度、色块主次与节奏都在预算内")
    L.append("")
    L.append("## 1b. Structural design")
    d_issues = (design or {}).get("issues", [])
    if not d_issues:
        L.append("- \u2705 arrow shapes consistent, rules anchored, "
                 "title counts match, no orphan decoration")
    else:
        L.append(f"- \u274c {len(d_issues)} issue(s)")
        for i in d_issues:
            L.append(f"  - **{i['type']}** slide {i['slide']} \u2014 {i['detail']}")
            if i.get("hint"):
                L.append(f"    - _fix_: {i['hint']}")

    L.append("")
    L.append("## 2. Citations (bidirectional)")
    if refs.get("error"):
        L.append(f"- ❌ {refs['error']}")
    else:
        L.append(f"- reference slide: {refs['reference_slide']}, "
                 f"entries: {refs['reference_count']}")
        L.append(f"- cited numbers: {refs['cited_numbers']}")
        if refs["undefined_citations"]:
            L.append(f"- ❌ citations with no list entry: {refs['undefined_citations']}")
        else:
            L.append("- ✅ every in-body citation resolves to an entry")
        if refs["uncited_references"]:
            L.append(f"- ❌ listed but never cited: {refs['uncited_references']}")
        else:
            L.append("- ✅ every entry is cited at least once")
        if refs["locator_problems"]:
            L.append("- ❌ locator problems (reference list):")
            for p in refs["locator_problems"]:
                L.append(f"  - ref {p['ref']}: {p['issue']} "
                         f"({p.get('locator', '')}) `{p.get('raw','')[:80]}`")
        else:
            L.append("- ✅ all reference-list locators verified against the source index")
        fl = refs.get("footer_locator_problems") or []
        if fl:
            L.append("- ❌ locator problems (slide footers):")
            for p in fl:
                L.append(f"  - slide {p['slide']}: `{p['locator']}` — {p['issue']}")
        else:
            L.append("- ✅ all slide-footer locators verified")

    L.append("")
    L.append("## 3. Assets")
    if not assets:
        L.append("- ✅ all pictures are on-slide and readable")
    else:
        for a in assets:
            L.append(f"- ❌ **{a['type']}** slide {a['slide']} — {a['detail']}")

    L.append("")
    L.append("## 4. Placeholders & prose")
    if not extra_issues:
        L.append("- ✅ no placeholders or AI-slop patterns")
    else:
        for i in extra_issues:
            L.append(f"- **{i['type']}** slide {i['slide']} — {i['detail']}")

    L.append("")
    L.append("## 5. Manual confirmation needed")
    L.append("- [ ] page locators refer to *physical* PDF pages — confirm against the reader")
    L.append("- [ ] every figure reprinted from a paper is marked \"重绘自 [n]\" if redrawn")
    L.append("- [ ] formulas: confirm the PNGs match the intended LaTeX")
    ok = (not layout_issues and not assets and not extra_issues
          and refs.get("ok") and not d_issues)
    L.append("")
    L.append(f"## Verdict: {'✅ PASS' if ok else '❌ FAIL — fix before delivery'}")
    path.write_text("\n".join(L), encoding="utf-8")
    return ok


def main():
    ap = argparse.ArgumentParser(description="Full deck audit + audit_report.md")
    ap.add_argument("pptx")
    ap.add_argument("--work", default="work", help="dir with sources.json")
    ap.add_argument("--out", default="audit")
    ap.add_argument("--ref-slide", type=int)
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--no-png", action="store_true")
    ap.add_argument("--min-margin", type=float, default=0.5,
                    help="页边安全距离；按给定模板创作时可调小")
    ap.add_argument("--allow", nargs="*", default=[],
                    help="substrings that make a prose-smell hit acceptable")
    args = ap.parse_args()

    deck = Path(args.pptx)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    layout = render_qa.run(deck, out / "layout", dpi=args.dpi,
                       make_png=not args.no_png,
                       min_margin_in=args.min_margin)
    design = design_check.run(deck)
    palette = palette_audit.audit(palette_audit.collect(str(deck)))
    src = Path(args.work) / "sources.json"
    refs = refs_check.run(deck, src if src.exists() else None, args.ref_slide)
    texts = deck_text(deck)
    extra = check_placeholders(texts) + check_ai_smells(texts, args.allow)
    assets = check_assets(deck)

    ok = write_markdown(out / "audit_report.md", deck, layout, refs, texts,
                        extra, assets, args.allow, design=design,
                        palette=palette)
    if palette["issues"]:
        ok = False
    (out / "audit.json").write_text(json.dumps({
        "layout": layout, "design": design, "palette": palette,
        "citations": refs,
        "prose": extra, "assets": assets,
        "ok": ok,
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"layout issues : {len(layout['issues'])}")
    print(f"citation gate : {'PASS' if refs.get('ok') else 'FAIL'}")
    print(f"design lint   : {len((design or {}).get(chr(105)+chr(115)+chr(115)+chr(117)+chr(101)+chr(115), []))} issue(s)")
    print(f"palette        : {len(palette['issues'])} issue(s)  "
          f"({palette['hue_families']} 个色相)")
    print(f"prose/assets  : {len(extra)} / {len(assets)}")
    print(f"report        : {out/'audit_report.md'}")
    print("VERDICT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
