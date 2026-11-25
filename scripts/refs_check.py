"""
refs_check.py - Bidirectional citation audit for academic decks.

The manual's S4/S6 gate has two halves and both must be enforced mechanically,
because a model reading its own deck will always believe it is consistent:

  forward : every [n] appearing in a slide resolves to a reference entry
  reverse : every reference entry is actually cited somewhere in the deck

It additionally verifies that each citation's locator (e.g. "S1 p.4") exists in
the ingested `sources.json`, which is what turns "no hallucinated references"
from a promise into a check.

Reference entries are read from a slide marked as the reference list
(`--ref-slide`) or auto-detected as the last slide containing "References" /
"参考文献". Citation markers are matched in the forms [1], [1,2], [1-3], [1–3],
and superscript-like 〔1〕.

Usage:
  python refs_check.py deck.pptx --sources work/sources.json --out audit/
  python refs_check.py deck.pptx --ref-slide 12
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

CITATION_RE = re.compile(r"[\[〔【]\s*(\d+(?:\s*[-–—~,，、]\s*\d+)*)\s*[\]〕】]")
REF_HEADING_RE = re.compile(r"^\s*(references?|参考文献|文献列表|bibliography)\s*$",
                           re.IGNORECASE)


def parse_citation_group(group: str) -> list[int]:
    """'1,3-5' -> [1,3,4,5]"""
    out: list[int] = []
    for part in re.split(r"[,，、]", group):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s*[-–—~]\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a <= b:
                out.extend(range(a, b + 1))
            else:
                out.extend(range(b, a + 1))
        elif part.isdigit():
            out.append(int(part))
    return out


@dataclass
class SlideText:
    index: int
    lines: list[str] = field(default_factory=list)
    footer_lines: list[str] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(self.lines)

    @property
    def footer_text(self) -> str:
        return "\n".join(self.footer_lines)


@dataclass
class RefEntry:
    n: int
    raw: str
    locators: list[str] = field(default_factory=list)


def _shape_text(sh) -> str:
    """
    Text of one shape, including tables.

    A citation inside a table cell is still a citation. python-pptx reports a
    table as a graphic frame with no text frame, so a reader that only looks at
    text frames silently sees nothing - and then reports a perfectly cited
    table as "listed but never cited".
    """
    if sh.has_text_frame:
        return sh.text_frame.text
    try:
        if sh.has_table:
            cells = []
            for row in sh.table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        cells.append(cell.text)
            return "\n".join(cells)
    except Exception:
        pass
    return ""


def read_slides(pptx_path: Path) -> list[SlideText]:
    from pptx import Presentation
    prs = Presentation(str(pptx_path))
    slide_h = (prs.slide_height or 0) / 914400
    # A footer band is the bottom ~1.3in of the slide. Only shapes in that band
    # can assert a page locator by way of "the source of what you are seeing";
    # a locator quoted in body text is a topic, not a claim about this slide.
    footer_top_in = max(0.0, slide_h - 1.3)
    slides: list[SlideText] = []
    for i, slide in enumerate(prs.slides, start=1):
        lines: list[str] = []
        footer: list[str] = []
        texts = []
        for shape in slide.shapes:
            t = _shape_text(shape)
            if not t.strip():
                continue
            top = shape.top if shape.top is not None else 0
            texts.append((top, t))
        for top, t in sorted(texts, key=lambda x: x[0]):
            keep = [ln for ln in t.split("\n") if ln.strip()]
            lines.extend(keep)
            if top / 914400 >= footer_top_in:
                footer.extend(keep)
        slides.append(SlideText(index=i, lines=lines, footer_lines=footer))
    return slides


def find_ref_slide(slides: list[SlideText], explicit: int | None) -> int | None:
    if explicit:
        return explicit
    # last slide that looks like a reference list: numbered entries dominate
    for s in reversed(slides):
        hits = sum(1 for ln in s.lines if re.match(r"^\s*[\[\(]?\d{1,3}[\]\.\)]\s*\S", ln))
        if hits >= 3 or any(REF_HEADING_RE.match(ln) for ln in s.lines):
            return s.index
    return None


def parse_references(slide: SlideText) -> list[RefEntry]:
    """Parse numbered references. Supports one or several entries per line."""
    entries: list[RefEntry] = []
    joined = "\n".join(slide.lines)
    # split on leading numbers: "[1]" "1." "1)" "［1］"
    # The trailing separator must stay on the same line: a lone numeral
    # (a page badge, a section number) is not an entry, and letting the
    # pattern swallow a newline made it eat the next line as that
    # numeral's body.
    pattern = re.compile(r"(?:^|\n)\s*[\[\(〔【]?(\d{1,3})[\]\)〕】\.、]?[ \t]+")
    matches = list(pattern.finditer(joined))
    for i, m in enumerate(matches):
        n = int(m.group(1))
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(joined)
        body = re.sub(r"\s+", " ", joined[start:end]).strip()
        if not body:
            continue
        locs = re.findall(r"\b([A-Z]\d+)\s*p{1,2}\.?\s*(\d+)\b", body)
        entries.append(RefEntry(n=n, raw=body,
                                locators=[f"{a} p.{b}" for a, b in locs]))
    # de-duplicate, keeping the longest body for each number
    by_n: dict[int, RefEntry] = {}
    for e in entries:
        if e.n not in by_n or len(e.raw) > len(by_n[e.n].raw):
            by_n[e.n] = e
    return [by_n[k] for k in sorted(by_n)]


def collect_citations(slides: list[SlideText], ref_slide: int | None):
    """Every [n] outside the reference slide, with the slide it appeared on."""
    used: dict[int, list[int]] = {}
    for s in slides:
        if ref_slide and s.index == ref_slide:
            continue
        for m in CITATION_RE.finditer(s.text):
            for n in parse_citation_group(m.group(1)):
                used.setdefault(n, [])
                if s.index not in used[n]:
                    used[n].append(s.index)
    return used


def load_sources(path: Path | None) -> dict:
    if not path or not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    return {s["id"]: s for s in data.get("sources", [])}


def run(pptx_path: Path, sources_path: Path | None, ref_slide: int | None):
    slides = read_slides(pptx_path)
    rs = find_ref_slide(slides, ref_slide)
    if rs is None:
        return {
            "ok": False,
            "error": "no reference slide found; pass --ref-slide N",
            "slides": len(slides),
        }
    refs = parse_references(slides[rs - 1])   # rs is 1-based
    used = collect_citations(slides, rs)
    sources = load_sources(sources_path)

    defined = {e.n for e in refs}
    cited = set(used)

    undefined = sorted(cited - defined)
    uncited = sorted(defined - cited)

    # locator validation against the ingested index
    bad_locators = []
    for e in refs:
        if not e.locators:
            bad_locators.append({"ref": e.n, "issue": "no locator (S1 p.N) present",
                                 "raw": e.raw[:110]})
            continue
        for loc in e.locators:
            sid, _, page = loc.partition(" p.")
            if sources:
                src = sources.get(sid)
                if src is None:
                    bad_locators.append({"ref": e.n, "issue": f"unknown source {sid}",
                                         "locator": loc})
                elif int(page) > src["pages"]:
                    bad_locators.append({
                        "ref": e.n,
                        "issue": f"page {page} beyond {sid} ({src['pages']} pages)",
                        "locator": loc})

    cites_per_slide = {}
    for n, sl in used.items():
        for s in sl:
            cites_per_slide[s] = cites_per_slide.get(s, 0) + 1

    # Footer locators are claims too: a slide footer reading "S1 p.99" asserts a
    # page that does not exist, even though the reference list itself is valid.
    # Only the footer band is scanned — a locator quoted inside body prose (e.g.
    # "cite S1 p.4 here") is a topic being discussed, not a claim about the slide.
    footer_locator_problems = []
    if sources:
        loc_pat = re.compile(r"\b([A-Z]\d+)\s*p{1,2}\.?\s*(\d+)\b")
        for s in slides:
            if rs and s.index == rs:
                continue
            for sid, page in loc_pat.findall(s.footer_text):
                src = sources.get(sid)
                if src is None:
                    footer_locator_problems.append({
                        "slide": s.index, "locator": f"{sid} p.{page}",
                        "issue": f"unknown source {sid}"})
                elif int(page) > src["pages"]:
                    footer_locator_problems.append({
                        "slide": s.index, "locator": f"{sid} p.{page}",
                        "issue": f"page {page} beyond {sid} ({src['pages']} pages)"})

    report = {
        "file": str(pptx_path),
        "slides": len(slides),
        "reference_slide": rs,
        "reference_count": len(refs),
        "cited_numbers": sorted(cited),
        "undefined_citations": undefined,
        "uncited_references": uncited,
        "locator_problems": bad_locators,
        "footer_locator_problems": footer_locator_problems,
        "citations_per_slide": dict(sorted(cites_per_slide.items())),
        "sources_loaded": sorted(sources.keys()),
        "ok": (not undefined and not uncited and not bad_locators
               and not footer_locator_problems),
    }
    return report


def main():
    ap = argparse.ArgumentParser(description="Bidirectional citation audit")
    ap.add_argument("pptx")
    ap.add_argument("--sources", help="sources.json produced by source_ingest.py")
    ap.add_argument("--ref-slide", type=int, help="1-based slide holding the reference list")
    ap.add_argument("--out", default="audit")
    args = ap.parse_args()

    rep = run(Path(args.pptx), Path(args.sources) if args.sources else None,
              args.ref_slide)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    (out / "refs_report.json").write_text(
        json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")

    if rep.get("error"):
        print("ERROR:", rep["error"])
        return 2
    print(f"slides={rep['slides']} reference_slide={rep['reference_slide']} "
          f"refs={rep['reference_count']}")
    print(f"cited numbers: {rep['cited_numbers']}")
    if rep["undefined_citations"]:
        print("UNDEFINED citations (in body, missing from list):",
              rep["undefined_citations"])
    if rep["uncited_references"]:
        print("UNCITED references (listed, never cited):", rep["uncited_references"])
    if rep["locator_problems"]:
        print("LOCATOR problems (reference list):")
        for p in rep["locator_problems"]:
            print("  ", p)
    if rep.get("footer_locator_problems"):
        print("LOCATOR problems (slide footers):")
        for p in rep["footer_locator_problems"]:
            print("  ", p)
    print("OK" if rep["ok"] else "FAIL")
    return 0 if rep["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
