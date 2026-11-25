"""
e2e.py - End-to-end regression for the zz-ai-pptx skill.

Runs the documented pipeline exactly as SKILL.md prescribes, on a real paper, and
asserts that the gates behave: good input passes, bad input is caught.

This is the test that matters, because the skill's entire value is that its
checks are not decorative.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent      # .../zz-ai-pptx/scripts
SKILL = HERE.parent                          # .../zz-ai-pptx
SCRIPTS = HERE
WORK = Path("e2e_work")


def sh(cmd: list[str], expect_fail=False):
    r = subprocess.run([sys.executable, *cmd], capture_output=True, text=True,
                       encoding="utf-8", errors="replace",
                       env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"})
    out = (r.stdout or "") + (r.stderr or "")
    if expect_fail and r.returncode == 0:
        raise AssertionError(f"expected failure but succeeded: {cmd}\n{out}")
    if not expect_fail and r.returncode != 0:
        raise AssertionError(f"command failed: {cmd}\n{out}")
    return out


def main():
    pdf = sys.argv[1] if len(sys.argv) > 1 else None
    if not pdf or not Path(pdf).exists():
        print("usage: python e2e.py <paper.pdf>")
        return 2

    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir()

    print("[1/7] preflight")
    print(sh([str(SCRIPTS / "preflight.py")]).strip().splitlines()[-1])

    print("[2/7] ingest")
    out = sh([str(SCRIPTS / "source_ingest.py"), pdf, "--out", str(WORK)])
    print("   ", out.strip().splitlines()[0])

    print("[3/7] locate a claim")
    hits = sh([str(SCRIPTS / "source_ingest.py"), "--find", "m6A", "--out", str(WORK)])
    first = hits.strip().splitlines()[0]
    locator = first[first.find("[") + 1: first.find("]")]
    print("   locator:", locator)

    print("[4/7] figure region")
    regs = sh([str(SCRIPTS / "source_ingest.py"), pdf, "--out", str(WORK),
               "--figure-page", "4", "--list-regions"])
    print("   ", regs.strip().splitlines()[0])
    sh([str(SCRIPTS / "source_ingest.py"), pdf, "--out", str(WORK),
        "--figure-page", "4", "--dpi", "220"])
    figs = list((WORK / "figures").glob("*_fig.png"))
    assert figs, "figure render produced nothing"

    print("[5/7] formula")
    sh([str(SCRIPTS / "math_render.py"), "--latex",
        r"\mathrm{softmax}\left(\frac{QK^{T}}{\sqrt{d_k}}\right)V",
        "--out", str(WORK / "math"), "--name", "eq1", "--number", "1.1"])

    print("[6/7] build deck")
    build = f'''
import sys
sys.path.insert(0, r"{SCRIPTS}")
from build_deck import Deck, Theme
d = Deck(Theme.navy())
d.title_slide("METTL3 \\u4e0d\\u662f\\u5206\\u5316\\u6807\\u5fd7\\uff0c\\u800c\\u662f\\u8eab\\u4efd\\u7ef4\\u6301\\u7684\\u5fc5\\u8981\\u6761\\u4ef6",
              subtitle="\\u4ece\\u8868\\u8fbe\\u4e0b\\u964d\\u5230\\u8868\\u578b\\u584c\\u9677",
              meta="\\u7ec4\\u4f1a \\u00b7 20 \\u5206\\u949f")
d.stats("\\u5206\\u5316\\u542f\\u52a8 48 \\u5c0f\\u65f6\\u5185\\uff0cMETTL3 \\u4e0e m6A \\u540c\\u6b65\\u4e0b\\u964d",
        [("3.2\\u00d7", "METTL3 \\u4e0b\\u8c03\\u500d\\u6570"),
         ("48 h", "\\u8eab\\u4efd\\u57fa\\u56e0\\u4e0b\\u964d\\u65f6\\u957f"),
         ("11", "\\u91cd\\u53e0 m6A \\u4f4d\\u70b9\\u6570")],
        kicker="\\u8bc1\\u636e", cite="{locator}")
d.figure("\\u56de\\u8865\\u540e\\u8eab\\u4efd\\u57fa\\u56e0\\u90e8\\u5206\\u6062\\u590d", r"{figs[0]}",
         kicker="\\u56fe\\u8bc1", caption="\\u539f\\u56fe\\u81ea [1] Fig.2\\uff0c\\u672a\\u91cd\\u7ed8\\u3002",
         cite="[1] {locator}",
         side_note=["\\u9762\\u677f a\\uff1a\\u514d\\u75ab\\u8367\\u5149", "\\u9762\\u677f c\\uff1a\\u8868\\u8fbe\\u91cf", "\\u9762\\u677f h\\uff1a\\u5dee\\u5f02\\u57fa\\u56e0"])
d.math("\\u6ce8\\u610f\\u529b\\u5f52\\u4e00\\u5316\\u51b3\\u5b9a\\u4e86\\u5c3a\\u5ea6\\u654f\\u611f\\u6027",
       r"{WORK / 'math' / 'eq1.png'}", eq_no="(1.1)",
       explain=["\\u5206\\u6bcd\\u7684\\u6839\\u53f7\\u907f\\u514d\\u5185\\u79ef\\u968f\\u7ef4\\u5ea6\\u7206\\u70b8\\u3002"], cite="[1] {locator}")
d.references(["Wen J, et al. Nat Commun. 2025. doi:10.1038/s41467-025-67484-z. S1 p.4"])
d.save(r"{WORK / 'final.pptx'}")
print("built")
'''
    (WORK / "build.py").write_text(build, encoding="utf-8")
    print("   ", sh([str(WORK / "build.py")]).strip())

    print("[7/7] audit — must PASS")
    sh([str(SCRIPTS / "audit_deck.py"), str(WORK / "final.pptx"),
        "--work", str(WORK), "--out", str(WORK / "audit"), "--dpi", "130"])
    rep = (WORK / "audit" / "audit_report.md").read_text(encoding="utf-8")
    assert "✅ PASS" in rep, "good deck did not pass the gate:\n" + rep
    pngs = sorted((WORK / "audit" / "layout" / "png").glob("slide-*.png"))
    assert len(pngs) == 5, f"expected 5 rendered slides, got {len(pngs)}"

    def run_audit(deck: Path, out: Path):
        return subprocess.run(
            [sys.executable, str(SCRIPTS / "audit_deck.py"), str(deck),
             "--work", str(WORK), "--out", str(out), "--no-png"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"})

    def expect_fail(label, mutated_build, outdir):
        (WORK / "build_bad.py").write_text(mutated_build, encoding="utf-8")
        sh([str(WORK / "build_bad.py")])
        deck = Path(WORK / "final_bad.pptx")
        assert deck.exists(), f"{label}: negative build produced no deck"
        r = run_audit(deck, outdir)
        assert r.returncode != 0, (
            f"{label}: audit passed a deck that should fail\n"
            + (r.stdout or "") + (r.stderr or ""))
        print(f"    correctly rejected: {label}")
        deck.unlink()

    print("\n[neg] reference list with an out-of-range page")
    # The reference entry is where a fabricated page number does its damage:
    # a locator on the reference slide is the citation's proof of existence.
    expect_fail("reference locator S1 p.99",
                build.replace("S1 p.4\"]", "S1 p.99\"]")
                     .replace("final.pptx", "final_bad.pptx"),
                WORK / "audit_bad1")

    print("\n[neg] slide footer citing a page that does not exist")
    expect_fail("footer locator S1 p.99",
                build.replace(locator, "S1 p.99")
                     .replace("final.pptx", "final_bad.pptx"),
                WORK / "audit_bad2")

    print("\n[neg] builder must refuse overflowing text")
    overflow = f'''
import sys
sys.path.insert(0, r"{SCRIPTS}")
from build_deck import Deck, Theme, FitError
d = Deck(Theme.navy())
try:
    d.bullets("x", ["\\u8fd9\\u662f\\u4e00\\u6bb5\\u975e\\u5e38\\u957f\\u7684\\u6587\\u5b57" * 150])
    print("NOERROR")
except FitError:
    print("REFUSED")
'''
    (WORK / "overflow.py").write_text(overflow, encoding="utf-8")
    r = sh([str(WORK / "overflow.py")]).strip()
    assert r == "REFUSED", f"builder did not refuse overflow: {r}"

    print("\n===== E2E OK =====")
    print("deliverables:")
    for f in ["final.pptx", "sources.md", "audit/audit_report.md"]:
        p = WORK / f
        print(f"  {'OK ' if p.exists() else 'MISS'} {p}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
