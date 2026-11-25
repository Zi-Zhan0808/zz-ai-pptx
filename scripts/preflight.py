"""
preflight.py - Report what this machine can actually verify.

Run before building a deck. It answers the only question that matters when the
workflow claims to "check" a deck: which checks are real here, and which are
degraded? Without this, a missing renderer is silently indistinguishable from a
clean deck.

Exit code is 0 when every capability is available, 1 when any check is degraded.
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.com",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]

FONT_CANDIDATES = [
    (r"C:\Windows\Fonts\msyh.ttc", "Microsoft YaHei"),
    (r"C:\Windows\Fonts\msyhbd.ttc", "Microsoft YaHei Bold"),
    (r"C:\Windows\Fonts\simsun.ttc", "SimSun"),
    (r"C:\Windows\Fonts\simhei.ttf", "SimHei"),
    (r"C:\Windows\Fonts\arial.ttf", "Arial"),
    (r"C:\Windows\Fonts\calibri.ttf", "Calibri"),
    # 等线 is what many Chinese institutional templates use; template mode
    # renders in it, so its availability is worth reporting too.
    (r"C:\Windows\Fonts\Deng.ttf", "DengXian / 等线"),
    (r"C:\Windows\Fonts\Dengb.ttf", "DengXian Bold / 等线 Bold"),
    ("/System/Library/Fonts/PingFang.ttc", "PingFang SC"),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "Noto Sans CJK"),
]


def check_modules() -> dict:
    out = {}
    for mod, purpose in [
        ("pptx", "build and read .pptx"),
        ("pymupdf", "PDF text, page render, figure crop"),
        ("PIL", "image sizing, formula compositing, contact sheets"),
        ("matplotlib", "LaTeX → PNG formulas"),
        ("lxml", "OOXML inspection"),
    ]:
        try:
            __import__(mod)
            out[mod] = {"ok": True, "purpose": purpose}
        except Exception as exc:
            out[mod] = {"ok": False, "purpose": purpose, "error": str(exc)}
    return out


def find_soffice() -> str | None:
    import os
    env = os.environ.get("AI_PPT_SOFFICE")
    if env and Path(env).exists():
        return env
    for c in SOFFICE_CANDIDATES:
        if Path(c).exists():
            return c
    return shutil.which("soffice")


def check_fonts() -> dict:
    found = {}
    for path, name in FONT_CANDIDATES:
        if Path(path).exists():
            found[name] = path
    cjk = [n for n in found if n in
           ("Microsoft YaHei", "SimSun", "SimHei", "PingFang SC", "Noto Sans CJK")]
    return {"found": found, "cjk_available": bool(cjk), "cjk_faces": cjk}


def main():
    report = {
        "modules": check_modules(),
        "soffice": find_soffice(),
        "fonts": check_fonts(),
    }

    sys.path.insert(0, str(HERE))
    try:
        from textfit import wrap_text
        r = wrap_text("\u4e2d\u6587\u6392\u7248\u5ea6\u91cf", 18, 200)
        report["textfit"] = {"ok": bool(r.lines), "probe_lines": len(r.lines)}
    except Exception as exc:
        report["textfit"] = {"ok": False, "error": str(exc)}

    degraded = []
    mods = report["modules"]
    if not mods["pptx"]["ok"]:
        degraded.append("BLOCKING: python-pptx missing — cannot build decks")
    if not mods["pymupdf"]["ok"]:
        degraded.append("PDF ingest and page rendering unavailable")
    if not mods["PIL"]["ok"]:
        degraded.append("image sizing unavailable")
    if not mods["matplotlib"]["ok"]:
        degraded.append("formula rendering unavailable — use supplied images or Unicode")
    if not report["soffice"]:
        degraded.append(
            "no standard LibreOffice — visual QA DEGRADED to metrics only; "
            "do not claim visual verification")
    if not report["fonts"]["cjk_available"]:
        degraded.append("no CJK font found — Chinese text cannot be measured or rendered")
    report["degraded"] = degraded

    print("== capability report ==")
    for mod, info in mods.items():
        print(f"  {'OK  ' if info['ok'] else 'MISS'} {mod:12s} {info['purpose']}")
    print(f"  {'OK  ' if report['soffice'] else 'MISS'} soffice      "
          f"{report['soffice'] or 'not found'}")
    print(f"  {'OK  ' if report['textfit']['ok'] else 'MISS'} textfit      "
          f"metrics-based text fitting")
    print(f"  fonts: {', '.join(report['fonts']['found']) or 'none'}")
    if degraded:
        print("\n== degraded ==")
        for d in degraded:
            print("  !", d)
    else:
        print("\n== all checks available ==")

    out = Path("preflight.json")
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nwrote {out}")
    return 0 if not degraded else 1


if __name__ == "__main__":
    sys.exit(main())
