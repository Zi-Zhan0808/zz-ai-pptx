"""
source_ingest.py - Turn the user's reference material into a citable index.

The central rule of the academic deck workflow: no claim without a locator.
This module is what makes that rule mechanical rather than aspirational.

For every input (PDF / .txt / .md) it builds a `sources.json` entry with:
  * a stable id  -> S1, S2, ...
  * the file's identity (path, sha1, page count)
  * per-page extracted text, so a citation can be pinned to a physical page
  * figure extraction, so "use the paper's own figure" is possible

It also answers the only question that matters during authoring:
  "which page of which file contains the sentence I want to cite?"

Usage:
  python source_ingest.py refs/ --out work/
  python source_ingest.py refs/ --out work/ --figures
  python source_ingest.py --find "attention sink" work/
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path

TEXT_SUFFIXES = {".txt", ".md", ".markdown"}


@dataclass
class PageRecord:
    page: int
    chars: int
    text: str = ""
    preview: str = ""


@dataclass
class Source:
    id: str
    kind: str                 # "pdf" | "text"
    path: str
    title: str = ""
    sha1: str = ""
    pages: int = 0
    year: str = ""
    authors: str = ""
    doi: str = ""
    page_records: list[PageRecord] = field(default_factory=list)
    figures: list[str] = field(default_factory=list)

    def locator(self, page: int) -> str:
        return f"{self.id} p.{page}"


def _sha1(path: Path) -> str:
    h = hashlib.sha1()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def render_page(pdf_path: Path, page: int, out_png: Path, dpi: int = 200,
                crop_rect: tuple[float, float, float, float] | None = None):
    """
    Render a whole PDF page (or a region) to PNG.

    Embedded-image extraction alone is usually not enough for papers: a figure is
    often assembled from many small images plus vector axes, so pulling the
    embedded rasters yields scattered fragments instead of Figure 2. Rendering
    the page region preserves the figure exactly as the authors laid it out, which
    is what "use the paper's own figure" actually requires.
    """
    import pymupdf

    doc = pymupdf.open(str(pdf_path))
    pg = doc[max(0, page - 1)]
    if crop_rect:
        clip = pymupdf.Rect(*crop_rect)
    else:
        clip = pg.rect
    pix = pg.get_pixmap(dpi=dpi, clip=clip)
    out_png.parent.mkdir(parents=True, exist_ok=True)
    pix.save(str(out_png))
    doc.close()
    return out_png


def find_figure_regions(pdf_path: Path, page: int) -> list[tuple[float, float, float, float]]:
    """
    Heuristically locate figure regions on a page: clusters of images/drawings
    that occupy a substantial area and are separated from body text.

    Returns candidate rects in PDF points, largest first, so the caller can pick
    the panel that visually matches the caption.
    """
    import pymupdf

    doc = pymupdf.open(str(pdf_path))
    pg = doc[max(0, page - 1)]
    rects: list = []
    for info in pg.get_image_info():
        r = info.get("bbox")
        if not r:
            continue
        x0, y0, x1, y1 = r
        if (x1 - x0) > 40 and (y1 - y0) > 40:
            rects.append(pymupdf.Rect(x0, y0, x1, y1))
    # drawing-heavy areas (vector plots, axes)
    for d in pg.get_drawings():
        r = d.get("rect")
        if r and r.width > 40 and r.height > 40:
            rects.append(r)
    doc.close()

    # merge overlapping/nearby rects into figure clusters
    merged: list = []
    for r in rects:
        placed = False
        for i, m in enumerate(merged):
            grown = pymupdf.Rect(m.x0 - 12, m.y0 - 12, m.x1 + 12, m.y1 + 12)
            if grown.intersects(r):
                merged[i] = m | r
                placed = True
                break
        if not placed:
            merged.append(pymupdf.Rect(r))
    # one more pass to fuse clusters that now touch
    changed = True
    while changed:
        changed = False
        for i in range(len(merged)):
            for j in range(i + 1, len(merged)):
                grown = pymupdf.Rect(merged[i].x0 - 18, merged[i].y0 - 18,
                                     merged[i].x1 + 18, merged[i].y1 + 18)
                if grown.intersects(merged[j]):
                    merged[i] = merged[i] | merged[j]
                    del merged[j]
                    changed = True
                    break
            if changed:
                break
    out = [(m.x0, m.y0, m.x1, m.y1) for m in merged
           if (m.x1 - m.x0) > 120 and (m.y1 - m.y0) > 100]
    out.sort(key=lambda t: (t[2] - t[0]) * (t[3] - t[1]), reverse=True)
    return out


def _guess_title_from_text(text: str) -> str:
    for line in text.splitlines():
        line = line.strip()
        if 8 <= len(line) <= 200 and not line.lower().startswith(
                ("abstract", "keywords", "doi", "arxiv")):
            return line
    return ""
    for line in text.splitlines():
        line = line.strip()
        if 8 <= len(line) <= 200 and not line.lower().startswith(
                ("abstract", "keywords", "doi", "arxiv")):
            return line
    return ""


def ingest_pdf(path: Path, sid: str, extract_figures: bool, fig_dir: Path) -> Source:
    import pymupdf

    doc = pymupdf.open(str(path))
    meta = doc.metadata or {}
    recs: list[PageRecord] = []
    figs: list[str] = []
    for pno, page in enumerate(doc, start=1):
        txt = page.get_text("text")
        recs.append(PageRecord(
            page=pno, chars=len(txt), text=txt,
            preview=re.sub(r"\s+", " ", txt.strip())[:280]))
        if extract_figures:
            # embedded raster images only; vector art is handled by page render
            for img_i, img in enumerate(page.get_images(full=True)):
                try:
                    pix = pymupdf.Pixmap(doc, img[0])
                    # Reject thumbnails, masks and spacer rectangles: a usable
                    # paper figure is at least 300px on its short side and not
                    # pathologically elongated (that is decoration, not data).
                    if min(pix.width, pix.height) < 300:
                        continue
                    if max(pix.width, pix.height) / max(1, min(pix.width, pix.height)) > 4:
                        continue
                    if pix.n - pix.alpha >= 4:
                        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                    fig_dir.mkdir(parents=True, exist_ok=True)
                    out = fig_dir / f"{sid}_p{pno}_{img_i}.png"
                    pix.save(str(out))
                    # a real figure survives PNG encoding as more than a few KB
                    if out.stat().st_size < 12_000:
                        out.unlink()
                        continue
                    figs.append(str(out))
                except Exception:
                    continue
    first_text = recs[0].text if recs else ""
    src = Source(
        id=sid, kind="pdf", path=str(path),
        title=(meta.get("title") or _guess_title_from_text(first_text)).strip(),
        sha1=_sha1(path), pages=doc.page_count,
        year=str(meta.get("creationDate", "") or "")[:4],
        authors=(meta.get("author") or "").strip(),
        page_records=recs, figures=figs,
    )
    joined = "\n".join(r.text for r in recs)
    m = re.search(r"\b(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)\b", joined)
    if m:
        src.doi = m.group(1).rstrip(".")
    doc.close()
    return src


def ingest_text(path: Path, sid: str) -> Source:
    txt = path.read_text(encoding="utf-8", errors="replace")
    # treat markdown headings as page-like anchors so locators stay stable
    chunks = re.split(r"(?m)^(#{1,3}\s+.+)$", txt)
    recs: list[PageRecord] = []
    if len(chunks) > 1:
        idx = 1
        pre = chunks[0]
        if pre.strip():
            recs.append(PageRecord(page=idx, chars=len(pre), text=pre,
                                   preview=re.sub(r"\s+", " ", pre.strip())[:280]))
            idx += 1
        for i in range(1, len(chunks), 2):
            head, body = chunks[i], chunks[i + 1] if i + 1 < len(chunks) else ""
            body_full = head + "\n" + body
            recs.append(PageRecord(page=idx, chars=len(body_full), text=body_full,
                                   preview=re.sub(r"\s+", " ", body_full.strip())[:280]))
            idx += 1
    else:
        recs.append(PageRecord(page=1, chars=len(txt), text=txt,
                               preview=re.sub(r"\s+", " ", txt.strip())[:280]))
    return Source(id=sid, kind="text", path=str(path),
                  title=path.stem, sha1=_sha1(path), pages=len(recs),
                  page_records=recs)


def build_sources(inputs: list[Path], out_dir: Path, extract_figures: bool) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    fig_root = out_dir / "figures"
    files: list[Path] = []
    for p in inputs:
        if p.is_dir():
            files.extend(sorted(q for q in p.rglob("*")
                                if q.suffix.lower() in {".pdf"} | TEXT_SUFFIXES))
        elif p.is_file():
            files.append(p)
    if not files:
        raise SystemExit(f"No ingestible files under: {[str(i) for i in inputs]}")

    sources: list[Source] = []
    for i, f in enumerate(files, start=1):
        sid = f"S{i}"
        if f.suffix.lower() == ".pdf":
            sources.append(ingest_pdf(f, sid, extract_figures, fig_root))
        elif f.suffix.lower() in TEXT_SUFFIXES:
            sources.append(ingest_text(f, sid))
        else:
            continue

    payload = {"sources": [asdict(s) for s in sources]}
    (out_dir / "sources.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = ["# Sources", ""]
    for s in sources:
        lines.append(f"## [{s.id}] {s.title or Path(s.path).name}")
        lines.append(f"- file: `{s.path}`")
        lines.append(f"- kind: {s.kind} | pages/anchors: {s.pages} | sha1: {s.sha1[:12]}")
        if s.doi:
            lines.append(f"- doi: {s.doi}")
        if s.figures:
            lines.append(f"- figures: {len(s.figures)} extracted under `{fig_root}`")
        for pr in s.page_records[:3]:
            lines.append(f"  - p.{pr.page}: {pr.preview[:160]}...")
        if s.pages > 3:
            lines.append(f"  - ... {s.pages - 3} more")
        lines.append("")
    (out_dir / "sources.md").write_text("\n".join(lines), encoding="utf-8")
    return payload


def find(query: str, work_dir: Path, limit: int = 20):
    """Locate a phrase across all ingested sources and report page locators."""
    data = json.loads((work_dir / "sources.json").read_text(encoding="utf-8"))
    q = query.lower()
    hits = []
    for s in data["sources"]:
        for pr in s["page_records"]:
            low = pr["text"].lower()
            pos = low.find(q)
            if pos >= 0:
                start = max(0, pos - 90)
                hits.append({
                    "locator": f"{s['id']} p.{pr['page']}",
                    "snippet": re.sub(r"\s+", " ", pr["text"][start:pos + len(q) + 90]),
                })
                if len(hits) >= limit:
                    return hits
    return hits


def main():
    ap = argparse.ArgumentParser(description="Ingest reference files into a citable index")
    ap.add_argument("inputs", nargs="*")
    ap.add_argument("--out", default="work")
    ap.add_argument("--figures", action="store_true", help="extract embedded figures")
    ap.add_argument("--find", help="search ingested text, print page locators")
    ap.add_argument("--figure-page", type=int,
                    help="render the largest figure region of PDF page N (1-based)")
    ap.add_argument("--figure-rect", help="x0,y0,x1,y1 in PDF points (with --figure-page)")
    ap.add_argument("--dpi", type=int, default=200, help="dpi for --figure-page")
    ap.add_argument("--list-regions", action="store_true",
                    help="with --figure-page: list candidate figure rects instead of rendering")
    args = ap.parse_args()

    out = Path(args.out)
    if args.find:
        hits = find(args.find, out)
        if not hits:
            print("no matches")
            return 1
        for h in hits:
            print(f"[{h['locator']}] ...{h['snippet']}...")
        return 0

    if args.figure_page:
        if not args.inputs:
            ap.error("--figure-page needs the PDF path as the first positional argument")
        pdf = Path(args.inputs[0])
        if args.list_regions:
            regions = find_figure_regions(pdf, args.figure_page)
            for i, r in enumerate(regions):
                print(f"region {i}: x0={r[0]:.1f} y0={r[1]:.1f} x1={r[2]:.1f} y1={r[3]:.1f} "
                      f"({r[2]-r[0]:.0f}x{r[3]-r[1]:.0f}pt)")
            if not regions:
                print("no figure regions detected")
            return 0
        rect = None
        if args.figure_rect:
            rect = tuple(float(v) for v in args.figure_rect.split(","))
        else:
            regions = find_figure_regions(pdf, args.figure_page)
            if not regions:
                print("no figure region detected; pass --figure-rect manually")
                return 1
            rect = regions[0]
            print(f"using largest region: {rect}")
        dest = out / "figures" / f"{pdf.stem[:24]}_p{args.figure_page}_fig.png"
        render_page(pdf, args.figure_page, dest, dpi=args.dpi, crop_rect=rect)
        print(f"rendered {dest}")
        return 0

    if not args.inputs:
        ap.error("provide input files/directories, or use --find / --figure-page")
    payload = build_sources([Path(p) for p in args.inputs], out, args.figures)
    for s in payload["sources"]:
        print(f"[{s['id']}] {Path(s['path']).name}  pages/anchors={s['pages']}  "
              f"figures={len(s['figures'])}  doi={s['doi'] or '-'}")
    print(f"wrote {out/'sources.json'} and {out/'sources.md'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
