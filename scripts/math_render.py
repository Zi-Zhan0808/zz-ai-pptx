"""
math_render.py - Render LaTeX to a transparent PNG for insertion into PPTX.

WHY IMAGES AND NOT OMML
-----------------------
Native OMML equations are the nicer authoring format, but every OOXML equation
written by a non-Microsoft producer is a compatibility bet: WPS and older
PowerPoint builds frequently drop or mangle them, and the deck then silently
loses its mathematics. A rendered PNG is verifiable in QA, identical
everywhere, and is what the manual's `math_mode=image` default describes.

A formula that is a picture is still "verifiable" as long as the PNG is
generated from the source LaTeX that is kept next to it, which is why every
render also appends to `<out>/math_manifest.json`.

Usage:
  python math_render.py --latex "E = mc^2" --out work/math --name emc2
  python math_render.py --file formulas.txt --out work/math
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

DEFAULT_DPI = 600


def _mathtext_to_png(latex: str, out_png: Path, fontsize: int = 28,
                     color: str = "000000", dpi: int = DEFAULT_DPI) -> tuple[int, int]:
    """Render with matplotlib mathtext (no LaTeX installation required)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    # matplotlib mathtext understands a large subset of LaTeX math; normalise the
    # few tokens it does not, rather than asking the user to rewrite formulas.
    tex = latex.strip().strip("$")
    tex = tex.replace("\\displaystyle", "")
    tex = re.sub(r"\\text\{([^}]*)\}", r"\\mathrm{\1}", tex)

    # Render on a deliberately oversized canvas, then crop to the actual ink and
    # re-pad. mathtext's reported bbox omits radical tails and italic overhang,
    # so sizing from it (or from bbox_inches="tight") clips real glyphs.
    fig = plt.figure(figsize=(14, 3), dpi=dpi)
    fig.patch.set_alpha(0.0)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()
    ax.text(0.01, 0.5, f"${tex}$", fontsize=fontsize, color="#" + color,
            va="center", ha="left")
    out_png.parent.mkdir(parents=True, exist_ok=True)
    raw = out_png.with_suffix(".raw.png")
    fig.savefig(raw, dpi=dpi, transparent=True)
    plt.close(fig)

    # Crop to ink, then pad symmetrically so nothing touches an edge.
    from PIL import Image, ImageChops
    with Image.open(raw) as im:
        im = im.convert("RGBA")
    bbox = im.getbbox()   # alpha-aware bounds of non-transparent pixels
    if bbox:
        im = im.crop(bbox)
    pad_px = max(10, int(fontsize * dpi / 72 * 0.18))
    canvas = Image.new("RGBA", (im.width + pad_px * 2, im.height + pad_px * 2),
                       (0, 0, 0, 0))
    canvas.paste(im, (pad_px, pad_px), im)
    canvas.save(out_png)
    raw.unlink(missing_ok=True)
    return canvas.size


def render(latex: str, out_dir: Path, name: str, fontsize: int = 28,
           color: str = "000000", dpi: int = DEFAULT_DPI,
           numbered: str | None = None) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    png = out_dir / f"{name}.png"
    w, h = _mathtext_to_png(latex, png, fontsize, color, dpi)

    if numbered:
        # append the equation number in parentheses, academic style
        from PIL import Image, ImageDraw, ImageFont
        base = Image.open(png).convert("RGBA")
        font = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", max(20, int(h * 0.30)))
        label = f"({numbered})"
        tw = font.getbbox(label)[2] + 8
        canvas = Image.new("RGBA", (base.width + tw + 16, base.height), (0, 0, 0, 0))
        canvas.paste(base, (0, 0), base)
        d = ImageDraw.Draw(canvas)
        d.text((base.width + 12, base.height // 2 - font.getbbox(label)[3] // 2),
               label, font=font, fill="#" + color)
        canvas.save(png)
        w, h = canvas.size

    record = {
        "name": name, "latex": latex, "png": str(png),
        "width_px": w, "height_px": h, "dpi": dpi,
        "size_in": [round(w / dpi, 3), round(h / dpi, 3)],
        "numbered": numbered, "color": color,
    }

    manifest = out_dir / "math_manifest.json"
    data = json.loads(manifest.read_text(encoding="utf-8")) if manifest.exists() else []
    data = [d for d in data if d.get("name") != name]
    data.append(record)
    manifest.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                        encoding="utf-8")
    return record


def main():
    ap = argparse.ArgumentParser(description="LaTeX -> transparent PNG for PPTX")
    ap.add_argument("--latex", help="LaTeX math expression")
    ap.add_argument("--file", help="file with one expression per line")
    ap.add_argument("--out", default="math")
    ap.add_argument("--name", help="output basename (default: auto)")
    ap.add_argument("--fontsize", type=int, default=28)
    ap.add_argument("--color", default="000000")
    ap.add_argument("--dpi", type=int, default=DEFAULT_DPI)
    ap.add_argument("--number", help="equation number, e.g. 3.1")
    args = ap.parse_args()

    out = Path(args.out)
    if args.latex:
        rec = render(args.latex, out, args.name or "formula", args.fontsize,
                     args.color, args.dpi, args.number)
        print(f"{rec['png']}  {rec['width_px']}x{rec['height_px']}px  "
              f"{rec['size_in'][0]}x{rec['size_in'][1]}in @{rec['dpi']}dpi")
        return 0
    if args.file:
        lines = [l.strip() for l in Path(args.file).read_text(encoding="utf-8").splitlines()
                 if l.strip() and not l.strip().startswith("#")]
        for i, l in enumerate(lines, start=1):
            rec = render(l, out, f"formula{i}", args.fontsize, args.color, args.dpi)
            print(f"{rec['name']}: {rec['size_in'][0]}x{rec['size_in'][1]}in")
        return 0
    ap.error("give --latex or --file")


if __name__ == "__main__":
    sys.exit(main())
