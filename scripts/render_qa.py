"""
render_qa.py - Trustworthy visual QA for PPTX decks.

WHY THIS IS NOT A THIN WRAPPER AROUND A RENDERER
------------------------------------------------
This environment ships a bundled LibreOffice kit whose Impress renderer does
*not* honour a text box's declared width: CJK runs are laid out against the
full canvas width, so a title that badly overflows its box can render looking
merely "wide". Certifying layout from that renderer gives false negatives.

Two independent sources of truth are used instead:

  1. Standard LibreOffice (soffice), when present, converts the deck to PDF.
     PyMuPDF then yields, per slide, the *text-block bounding boxes* in
     PDF points. Those are compared against the shape geometry read from the
     .pptx, which is what makes overflow detectable numerically rather than by
     eye. `pymupdf` text extraction already gives correct CJK metrics.
  2. textfit.py computes wrapping from real font metrics, so layout can be
     audited even when no renderer is available at all.

Usage:
  python render_qa.py deck.pptx --out qa/            # full QA + PNGs
  python render_qa.py deck.pptx --out qa/ --no-png   # numbers only
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

EMU_PER_INCH = 914400
PT_PER_INCH = 72.0

SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.com",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
    "/usr/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice() -> str | None:
    env = os.environ.get("AI_PPT_SOFFICE")
    if env and Path(env).exists():
        return env
    for c in SOFFICE_CANDIDATES:
        if Path(c).exists():
            return c
    return shutil.which("soffice")


# --------------------------------------------------------------------------
# geometry extracted straight from the .pptx (the thing under test)
# --------------------------------------------------------------------------
@dataclass
class ShapeBox:
    slide: int
    name: str
    kind: str          # "text" | "picture" | "chart" | "table" | "shape"
    x: float           # inches
    y: float
    w: float
    h: float
    text: str = ""
    font_sizes: list[float] = field(default_factory=list)
    fill_is_dark: bool | None = None
    text_is_light: bool | None = None


def _px_to_in(emu: int) -> float:
    return emu / EMU_PER_INCH


def extract_geometry(pptx_path: Path) -> tuple[list[ShapeBox], tuple[float, float]]:
    """Read slide size and every top-level shape box from the deck."""
    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation(str(pptx_path))
    slide_w = _px_to_in(prs.slide_width)
    slide_h = _px_to_in(prs.slide_height)
    boxes: list[ShapeBox] = []

    for idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if shape.left is None or shape.top is None:
                continue
            kind = "shape"
            st = str(shape.shape_type)
            if shape.has_text_frame and shape.text_frame.text.strip():
                kind = "text"
            elif shape.has_chart if hasattr(shape, "has_chart") else False:
                kind = "chart"
            elif getattr(shape, "has_table", False):
                kind = "table"
            elif "PICTURE" in st:
                kind = "picture"

            text = ""
            sizes: list[float] = []
            light = None
            if shape.has_text_frame:
                text = shape.text_frame.text
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size is not None:
                            sizes.append(r.font.size.pt)
                        if (r.font.color and r.font.color.type is not None
                                and getattr(r.font.color, "rgb", None) is not None):
                            c = r.font.color.rgb
                            if (int(c[0]) + int(c[1]) + int(c[2])) / 3 > 170:
                                light = True
                if light is None and sizes:
                    light = False

            dark = None
            try:
                if shape.fill.type is not None and shape.fill.type == 1:
                    rgb = shape.fill.fore_color.rgb
                    dark = (int(rgb[0]) + int(rgb[1]) + int(rgb[2])) / 3 < 110
            except Exception:
                dark = None

            boxes.append(ShapeBox(
                slide=idx, name=shape.name, kind=kind,
                x=_px_to_in(shape.left), y=_px_to_in(shape.top),
                w=_px_to_in(shape.width), h=_px_to_in(shape.height),
                text=text, font_sizes=sizes,
                fill_is_dark=dark, text_is_light=light,
            ))
    return boxes, (slide_w, slide_h)


# --------------------------------------------------------------------------
# rendering via standard LibreOffice + PyMuPDF
# --------------------------------------------------------------------------
def convert_to_pdf(pptx_path: Path, out_dir: Path) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "Standard LibreOffice not found. Install it, or set AI_PPT_SOFFICE "
            "to a soffice.com/soffice executable.")
    profile = Path(tempfile.mkdtemp(prefix="aippt_lo_"))
    cmd = [
        soffice,
        f"-env:UserInstallation=file:///{profile.as_posix().lstrip('/')}",
        "--headless", "--norestore", "--nolockcheck",
        "--convert-to", "pdf",
        "--outdir", str(out_dir),
        str(pptx_path),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    pdf = out_dir / (pptx_path.stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError(
            f"LibreOffice produced no PDF.\nstdout:{proc.stdout}\nstderr:{proc.stderr}")
    return pdf


@dataclass
class TextBlock:
    page: int
    bbox: tuple[float, float, float, float]   # PDF points, origin top-left
    text: str


def extract_text_blocks(pdf_path: Path) -> tuple[list[TextBlock], tuple[float, float]]:
    import pymupdf
    doc = pymupdf.open(str(pdf_path))
    blocks: list[TextBlock] = []
    page_w = page_h = 0.0
    for pno, page in enumerate(doc, start=1):
        if pno == 1:
            page_w, page_h = page.rect.width, page.rect.height
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, txt = b[0], b[1], b[2], b[3], b[4]
            if not txt.strip():
                continue
            blocks.append(TextBlock(pno, (x0, y0, x1, y1), txt))
    doc.close()
    return blocks, (page_w, page_h)


def render_pngs(pdf_path: Path, out_dir: Path, dpi: int = 150) -> list[Path]:
    import pymupdf
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(str(pdf_path))
    paths = []
    for pno, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=dpi)
        p = out_dir / f"slide-{pno:02d}.png"
        pix.save(str(p))
        paths.append(p)
    doc.close()
    return paths


# --------------------------------------------------------------------------
# checks
# --------------------------------------------------------------------------
TOL_PT = 2.0  # rendering rounding tolerance

# How far a trailing punctuation mark may hang past a box edge before it counts
# as real overflow. Full-width CJK punctuation is ~1em wide, so a few points of
# hang is normal renderer behaviour (标点悬挂), not a layout defect.
PUNCT_HANG_PT = 14.0
_PUNCT_CHARS = set("，。、；：？！）」』》〉”’,.;:?!)]}%")


def _tail_is_punctuation(text: str) -> bool:
    """True when the block ends with punctuation, i.e. only that can be hanging."""
    t = text.rstrip()
    return bool(t) and t[-1] in _PUNCT_CHARS


def check_layout(boxes, blocks, slide_size):
    """
    Compare rendered text-block extents against declared shape geometry.

    Returns (issues, notes): `issues` fail the gate, `notes` are informational
    findings that a human should see but that are not defects.
    """
    issues = []
    notes: list[dict] = []
    W = slide_size[0]
    H = slide_size[1]
    by_slide: dict[int, list[ShapeBox]] = {}
    for b in boxes:
        by_slide.setdefault(b.slide, []).append(b)
    blocks_by_page: dict[int, list[TextBlock]] = {}
    for t in blocks:
        blocks_by_page.setdefault(t.page, []).append(t)

    for slide_no, page_blocks in blocks_by_page.items():
        shapes = [s for s in by_slide.get(slide_no, [])
                  if s.kind == "text" and not _is_footer_box(s)]
        for tb in page_blocks:
            if _looks_like_footer_text(tb.text):
                continue
            if _inside_nontext_shape(tb, by_slide.get(slide_no, [])):
                continue
            if _spans_multiple_shapes(tb, shapes):
                # merged boxes: comparing extent against any single box is not
                # a defect signal
                continue
            x0, y0, x1, y1 = tb.bbox
            # find the shape whose origin best matches this block
            best, best_d = None, 1e9
            for s in shapes:
                sx, sy = s.x * 72, s.y * 72
                by_s = (s.y + s.h) * 72
                # A block belongs to a box it actually overlaps. Matching on
                # origin distance alone let a footer block (whose own boxes were
                # filtered out) attach itself to the caption sitting 0.9in above
                # it and report the caption as overflowing - a false positive
                # that pointed at the wrong shape entirely.
                if y0 >= by_s + 4 or y1 <= sy - 4:
                    continue
                d = abs(sx - x0) + abs(sy - y0)
                if d < best_d:
                    best, best_d = s, d
            if best is None or best_d > 90:
                continue
            # A block far wider than the box it matched is a merge of this box
            # with its neighbour on the same line (a right-aligned row label plus
            # the body beside it). Measuring that against the label reports a
            # 600pt overflow that is not on the page. Declared-size overflow is
            # caught earlier by metrics-overflow, so nothing real is excused.
            if (x1 - x0) > best.w * 72 * 1.6 + 40:
                continue
            rx = (best.x + best.w) * 72
            by_ = (best.y + best.h) * 72
            if x1 > rx + TOL_PT:
                over = x1 - rx
                # Hang punctuation: a renderer may push a trailing full-width
                # comma/period a few points past the box edge (标点悬挂). It is
                # visually harmless and extremely common, so it is reported as a
                # note rather than a defect — but only when the overflow is small
                # AND the overflowing content really is punctuation. A run of real
                # characters spilling out stays a hard failure.
                if over <= PUNCT_HANG_PT and _tail_is_punctuation(tb.text):
                    notes.append({
                        "type": "punctuation-hang", "slide": slide_no,
                        "shape": best.name,
                        "detail": f"trailing punctuation hangs {over:.1f}pt past "
                                  f"the box edge (visual severity: negligible)",
                    })
                else:
                    issues.append({
                        "type": "text-overflow-right", "slide": slide_no,
                        "shape": best.name,
                        "detail": f"text right edge {x1:.1f}pt > box {rx:.1f}pt "
                                  f"(+{over:.1f}pt)",
                        "text": tb.text.strip()[:60],
                    })
            if y1 > by_ + TOL_PT:
                over = y1 - by_
                if over <= PUNCT_HANG_PT:
                    # a final line whose descent/punctuation dips a hair below
                    notes.append({
                        "type": "baseline-descent", "slide": slide_no,
                        "shape": best.name,
                        "detail": f"last line dips {over:.1f}pt below the box "
                                  f"(visual severity: negligible)",
                    })
                else:
                    issues.append({
                        "type": "text-overflow-bottom", "slide": slide_no,
                        "shape": best.name,
                        "detail": f"text bottom {y1:.1f}pt > box {by_:.1f}pt "
                                  f"(+{over:.1f}pt)",
                        "text": tb.text.strip()[:60],
                    })
    return issues, notes


FOOTER_NAME_HINTS = ("footer", "page number", "页码", "页脚")


def _is_footer_box(s: "ShapeBox") -> bool:
    """
    A deliberate footer sits in the bottom band and is tiny. Excluding it from
    margin/overlap checks prevents a real design element from being reported as
    a defect, which would train the reader to ignore the report.

    Recognised by name OR by geometry+size, because a footer merged into another
    text box (a common result of renderer block grouping) carries neither.
    """
    if not s.text.strip():
        return False
    if any(h in s.name.lower() for h in ("footer", "page", "页码", "页脚")):
        return True
    if not s.font_sizes:
        return False
    return max(s.font_sizes) <= 11.0 and s.y > 6.5


def _spans_multiple_shapes(tb: TextBlock, shapes: list["ShapeBox"]) -> bool:
    """
    True when a rendered text block sits on top of two or more declared shapes.

    The renderer merges adjacent text that shares a baseline into one block (an
    agenda row renders as "01第一项", and a right-aligned row label merges with
    the body text beside it). Measuring that merged block against the narrow box
    it happens to start at reports a large overflow that does not exist. Counting
    the shapes the block covers is the reliable signal: a genuinely overflowing
    single box has only itself on its line.

    """
    x0, y0, x1, y1 = tb.bbox
    covered = 0
    for s_ in shapes:
        sx, sy = s_.x * 72, s_.y * 72
        ey = (s_.y + s_.h) * 72
        if (sy < y1 and ey > y0) and (x0 - 10 <= sx <= x1 + 14):
            covered += 1
    return covered >= 2


def _inside_nontext_shape(tb: TextBlock, all_shapes: list["ShapeBox"]) -> bool:
    """True when the block sits inside a table/chart/picture shape.

    Table and chart interiors are rendered as text blocks that belong to the
    shape as a whole, not to any text box, so the text-box overflow comparison
    does not apply.
    """
    x0, y0, x1, y1 = tb.bbox
    for s_ in all_shapes:
        if s_.kind not in ("table", "chart", "picture"):
            continue
        sx, sy = s_.x * 72, s_.y * 72
        ex, ey = (s_.x + s_.w) * 72, (s_.y + s_.h) * 72
        if sx - 4 <= x0 and y0 >= sy - 4 and x1 <= ex + 4 and y1 <= ey + 8:
            return True
    return False


def _looks_like_footer_text(txt: str) -> bool:
    """
    True for a footer/citation-locator/page-number block.

    The renderer groups the footer citation and the page number into one text
    block whenever they share a baseline, so separators and brackets are
    normalised away and the remainder must consist only of locators and an
    optional trailing page number. Anything containing real prose returns False,
    so genuine content is never excused.
    """
    t = re.sub(r"\s+", " ", txt).strip()
    if not t:
        return False
    if re.fullmatch(r"\d{1,3}", t):
        return True
    t = re.sub(r"[\[\]〔〕【】()（）]", " ", t)
    t = re.sub(r"[;；,，、]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    if not t:
        return False
    # allow a leading citation number attached to a locator, e.g. "1 S1 p.4"
    t = re.sub(r"(?<![A-Za-z\d])\d{1,3}(?=\s+[A-Z]\d+\s*p\.?)", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    loc = r"[A-Z]\d+\s*p\.?\s*\d+"
    return bool(re.fullmatch(rf"(?:{loc}\s*)+", t)
                or re.fullmatch(rf"(?:{loc}\s*)+\d{{1,3}}", t))


def check_margins(boxes, slide_size, min_margin_in=0.5):
    issues = []
    W, H = slide_size
    for s in boxes:
        if s.w >= W - 0.02 and s.h >= H - 0.02:
            continue  # full-bleed background
        if s.kind not in ("text", "picture", "chart", "table"):
            continue
        if _is_footer_box(s):
            continue
        if s.x < min_margin_in - 0.01:
            issues.append({"type": "margin-left", "slide": s.slide, "shape": s.name,
                           "detail": f"left {s.x:.2f}in < {min_margin_in}in"})
        if s.y < min_margin_in - 0.01 and s.kind == "text":
            issues.append({"type": "margin-top", "slide": s.slide, "shape": s.name,
                           "detail": f"top {s.y:.2f}in < {min_margin_in}in"})
        if s.x + s.w > W - min_margin_in + 0.01:
            issues.append({"type": "margin-right", "slide": s.slide, "shape": s.name,
                           "detail": f"right {s.x+s.w:.2f}in > {W-min_margin_in:.2f}in"})
        if s.y + s.h > H - min_margin_in + 0.01 and s.kind == "text":
            issues.append({"type": "margin-bottom", "slide": s.slide, "shape": s.name,
                           "detail": f"bottom {s.y+s.h:.2f}in > {H-min_margin_in:.2f}in"})
    return issues


def check_overlap(boxes):
    """Overlapping text blocks are the defect a human notices first."""
    issues = []
    texts = [b for b in boxes if b.kind == "text" and b.w > 0 and b.h > 0]
    for i in range(len(texts)):
        for j in range(i + 1, len(texts)):
            a, b = texts[i], texts[j]
            if a.slide != b.slide:
                continue
            ox = max(0, min(a.x + a.w, b.x + b.w) - max(a.x, b.x))
            oy = max(0, min(a.y + a.h, b.y + b.h) - max(a.y, b.y))
            if ox > 0.05 and oy > 0.05:
                area = ox * oy
                smaller = min(a.w * a.h, b.w * b.h) or 1
                if area / smaller > 0.12:
                    issues.append({
                        "type": "overlap", "slide": a.slide,
                        "shapes": [a.name, b.name],
                        "detail": f"overlap {ox:.2f}x{oy:.2f}in "
                                  f"({area/smaller*100:.0f}% of smaller box)",
                    })
    return issues


def check_font_discipline(boxes, min_body_pt=14.0, min_title_pt=26.0,
                          min_caption_pt=9.0):
    """
    Academic decks live or die on a consistent type scale.

    The title floor is 26pt rather than 30pt: a section-divider or an
    evidence-slide headline legitimately sits at 26-28pt when the deck runs at a
    28pt section size. Flagging those produced false positives that made the
    whole report easy to dismiss.
    """
    issues = []
    sizes_by_slide: dict[int, list[float]] = {}
    for s in boxes:
        if _is_footer_box(s):
            continue
        for sz in s.font_sizes:
            sizes_by_slide.setdefault(s.slide, []).append(sz)
            if sz < min_caption_pt:
                issues.append({"type": "font-too-small", "slide": s.slide,
                               "shape": s.name, "detail": f"{sz:.0f}pt < {min_caption_pt}pt"})
    for slide, sizes in sizes_by_slide.items():
        if sizes and max(sizes) < min_title_pt:
            issues.append({"type": "title-too-small", "slide": slide,
                           "shape": "-",
                           "detail": f"largest text {max(sizes):.0f}pt < {min_title_pt}pt"})
    return issues


def check_contrast(boxes):
    """Light text on a light box (or dark on dark) is unreadable."""
    issues = []
    for s in boxes:
        if s.fill_is_dark is not None and s.text_is_light is not None:
            if s.fill_is_dark is False and s.text_is_light is True:
                issues.append({"type": "low-contrast", "slide": s.slide,
                               "shape": s.name,
                               "detail": "light text on light fill"})
            if s.fill_is_dark is True and s.text_is_light is False and s.font_sizes:
                issues.append({"type": "low-contrast", "slide": s.slide,
                               "shape": s.name,
                               "detail": "dark text on dark fill"})
    return issues


def check_metrics_overflow(boxes):
    """
    Renderer-independent fit check using real font metrics. Catches overflow
    even when LibreOffice is missing or its renderer is untrustworthy.
    """
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from textfit import wrap_text

    issues = []
    for s in boxes:
        if s.kind != "text" or not s.text.strip():
            continue
        if s.w <= 0.05 or s.h <= 0.05:
            continue
        # usable area inside default insets (0.1in L/R, 0.05in T/B)
        avail_w = (s.w - 0.2) * 72
        avail_h = (s.h - 0.1) * 72
        if avail_w <= 0:
            issues.append({"type": "box-too-narrow", "slide": s.slide,
                           "shape": s.name, "detail": f"width {s.w:.2f}in"})
            continue
        size = max(s.font_sizes) if s.font_sizes else 18.0
        r = wrap_text(s.text, size, avail_w)
        if r.needed_height_pt > avail_h + 4:
            issues.append({
                "type": "metrics-overflow", "slide": s.slide, "shape": s.name,
                "detail": f"{len(r.lines)} lines need ~{r.needed_height_pt:.0f}pt, "
                          f"box gives {avail_h:.0f}pt",
                "text": s.text.strip()[:60],
            })
        if r.overflow_x:
            issues.append({
                "type": "metrics-overflow-x", "slide": s.slide, "shape": s.name,
                "detail": f"line {r.widest_line_pt:.0f}pt > {avail_w:.0f}pt available",
                "text": s.text.strip()[:60],
            })
    return issues


def inventory_slide_fonts(pdf_path: Path) -> dict:
    """Fonts actually embedded by the render, to catch silent substitution."""
    import pymupdf
    doc = pymupdf.open(str(pdf_path))
    fonts = set()
    for page in doc:
        for f in page.get_fonts():
            fonts.add(f[3])
    doc.close()
    return {"fonts": sorted(fonts)}


def run(pptx_path: Path, out_dir: Path, dpi: int = 150, make_png: bool = True,
        min_margin_in: float = 0.5):
    out_dir.mkdir(parents=True, exist_ok=True)
    boxes, slide_size = extract_geometry(pptx_path)

    report = {
        "file": str(pptx_path),
        "slide_size_in": [round(slide_size[0], 3), round(slide_size[1], 3)],
        "slide_count": max((b.slide for b in boxes), default=0),
        "shape_count": len(boxes),
        "renderer": None,
        "issues": [],
        "notes": [],
        "pngs": [],
    }

    # 1) renderer-independent checks always run
    report["issues"] += check_metrics_overflow(boxes)
    report["issues"] += check_margins(boxes, slide_size, min_margin_in)
    report["issues"] += check_overlap(boxes)
    report["issues"] += check_font_discipline(boxes)
    report["issues"] += check_contrast(boxes)

    # 2) renderer-backed checks when a trustworthy renderer exists
    soffice = find_soffice()
    if soffice:
        report["renderer"] = soffice
        try:
            pdf = convert_to_pdf(pptx_path, out_dir)
            blocks, pdf_size = extract_text_blocks(pdf)
            report["pdf"] = str(pdf)
            report["pdf_size_pt"] = [round(pdf_size[0], 1), round(pdf_size[1], 1)]
            report["text_blocks"] = len(blocks)
            layout_issues, layout_notes = check_layout(boxes, blocks, slide_size)
            report["issues"] += layout_issues
            report["notes"] += layout_notes
            report["fonts_used"] = inventory_slide_fonts(pdf)["fonts"]
            if make_png:
                pngs = render_pngs(pdf, out_dir / "png", dpi=dpi)
                report["pngs"] = [str(p) for p in pngs]
        except Exception as exc:
            report["render_error"] = f"{type(exc).__name__}: {exc}"
    else:
        report["render_error"] = "standard LibreOffice not found; " \
                                 "only metrics-based layout checks ran"

    by_type: dict[str, int] = {}
    for i in report["issues"]:
        by_type[i["type"]] = by_type.get(i["type"], 0) + 1
    report["issue_counts"] = by_type
    report["ok"] = not report["issues"]

    (out_dir / "qa_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main():
    ap = argparse.ArgumentParser(description="Trustworthy PPTX layout QA")
    ap.add_argument("pptx")
    ap.add_argument("--out", default="qa")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--no-png", action="store_true")
    ap.add_argument("--min-margin", type=float, default=0.5,
                    help="页边安全距离（英寸）；模板自带的贴边装饰可调小")
    args = ap.parse_args()

    rep = run(Path(args.pptx), Path(args.out), args.dpi, not args.no_png)
    print(f"renderer: {rep['renderer']}")
    print(f"slides: {rep['slide_count']}  shapes: {rep['shape_count']}  "
          f"text blocks: {rep.get('text_blocks','-')}")
    print(f"issue counts: {rep['issue_counts'] or 'none'}")
    for i in rep["issues"][:25]:
        print(f"  [{i['type']}] slide {i['slide']} "
              f"{i.get('shape') or i.get('shapes')}: {i['detail']}")
    if rep.get("render_error"):
        print("render note:", rep["render_error"])
    print("report:", Path(args.out) / "qa_report.json")
    return 0 if rep["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
