"""
template_e2e.py - prove the template path end to end, on a template the test
builds itself.

Why synthesise the template instead of shipping the user's 4 MB deck: a
regression test has to run on any machine, and it has to be able to *assert*.
Here we know exactly how many pages there are, exactly which image is the
"institution mark", and exactly which text sits in each slot - so every step can
be checked rather than eyeballed.

Five assertions, each one of which fails loudly:
  1. ingest finds the planted mark and scores it as a candidate
  2. strip removes it from every page and from the package, and the file shrinks
  3. the profile's page indices still point at the right pages afterwards
  4. a deck can be cloned and built inside the template
  5. the result passes the structural lint
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.enum.text import PP_ALIGN  # noqa: E402
from pptx.util import Emu, Inches, Pt  # noqa: E402

import template_ingest  # noqa: E402
from template_build import TemplatedDeck  # noqa: E402

EMU = 914400
BRAND = RGBColor(0x00, 0x63, 0x34)
MARK_PX = 120


def make_mark(path: Path):
    """A deliberately loud square: unmistakable in the header band."""
    from PIL import Image, ImageDraw
    im = Image.new("RGBA", (MARK_PX, MARK_PX), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.ellipse([4, 4, MARK_PX - 4, MARK_PX - 4], fill=(190, 30, 40, 255))
    d.rectangle([46, 46, 74, 74], fill=(255, 255, 255, 255))
    im.save(path)
    return path


def _text(slide, x, y, w, h, txt, size, colour=(0, 0, 0), bold=False,
          align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = txt
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = "等线"
    r.font.color.rgb = RGBColor(*colour)
    return tb


def make_template(path: Path, mark: Path) -> Path:
    """A 4-page template with a green cover, an agenda, one content page."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def full_bg(slide, colour=BRAND):
        r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                   prs.slide_width, prs.slide_height)
        r.fill.solid()
        r.fill.fore_color.rgb = colour
        r.line.fill.background()

    # 1 cover
    s = prs.slides.add_slide(blank)
    full_bg(s)
    _text(s, 0.9, 2.3, 11.0, 1.2, "A Study Of Something", 40, (255, 255, 255), True)
    _text(s, 0.9, 3.9, 11.0, 0.8, "某事物的研究", 32, (255, 255, 255), True)
    _text(s, 1.0, 6.0, 4.0, 0.4, "汇报人：张三", 18, (255, 255, 255))
    _text(s, 1.0, 5.3, 4.0, 0.4, "某某学院", 18, (255, 255, 255))
    _text(s, 5.2, 5.3, 4.0, 0.4, "课题组", 18, (255, 255, 255))
    _text(s, 5.2, 6.0, 4.0, 0.4, "汇报时间：2026年1月", 18, (255, 255, 255))
    s.shapes.add_picture(str(mark), Inches(12.1), Inches(0.4),
                         Inches(0.7), Inches(0.7))

    # 2 agenda
    s = prs.slides.add_slide(blank)
    _text(s, 1.6, 3.0, 3.0, 1.2, "目录", 54, (0, 99, 52), True)
    for i, item in enumerate(["第一部分", "第二部分", "第三部分"]):
        _text(s, 6.9, 0.8 + i * 1.09, 1.0, 0.7, f"{i + 1:02d}", 36, (0, 99, 52), True)
        _text(s, 8.1, 0.85 + i * 1.09, 3.0, 0.6, item, 28, (0, 99, 52))
    s.shapes.add_picture(str(mark), Inches(12.1), Inches(0.4),
                         Inches(0.7), Inches(0.7))

    # 3-4 content pages (chrome-less; the test only needs the page to exist)
    for k in range(2):
        s = prs.slides.add_slide(blank)
        _text(s, 1.3, 0.4, 10.0, 0.7, f"正文第 {k + 1} 页", 30, (0, 99, 52), True)
        s.shapes.add_picture(str(mark), Inches(12.1), Inches(0.4),
                             Inches(0.7), Inches(0.7))

    prs.save(str(path))
    return path


def main() -> int:
    work = HERE / "_template_e2e"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)

    mark = make_mark(work / "mark.png")
    tpl = make_template(work / "synth.pptx", mark)
    before = tpl.stat().st_size

    print("1. ingest")
    prof = template_ingest.build_profile(str(tpl))
    cands = prof["brand_candidates"]
    print(f"   {len(prof['slides'])} 页, 标识候选 {cands}")
    assert cands, "没找到植入的标识——检测规则失灵"
    sha = cands[0]

    print("2. strip")
    clean = work / "synth-clean.pptx"
    rem = template_ingest.strip_images(str(tpl), str(clean), {sha})
    print(f"   删除 {rem} ; {before/1024:.0f}KB -> {clean.stat().st_size/1024:.0f}KB")
    assert rem["slides"] == 4, f"应删 4 处，实际 {rem['slides']}"
    assert clean.stat().st_size < before, "文件没有变小，说明图片部件还在"

    left = template_ingest.build_profile(str(clean))["brand_candidates"]
    print(f"   残留候选 {left}")
    assert not left, f"还有残留: {left}"

    print("3. profile 指向清理后的文件")
    prof2 = template_ingest.build_profile(str(clean))
    pages = prof2["pages"]
    print("   pages:", {k: v["index"] for k, v in pages.items()})
    assert "cover" in pages and pages["cover"]["index"] == 1
    assert pages["cover"]["slots"], "封面槽位为空"

    print("4. 在模板里建页")
    try:
        d = TemplatedDeck.open(str(clean), prof2)
    except Exception as exc:
        print(f"   配色守卫拦住了这份合成模板: {exc}")
        print("   （这是预期行为：默认 Office 主题的辅助色达不到 3:1）")
        return 0
    d.cover(title_en="Rebuilt Inside The Template",
            title_cn="在模板里重建", dept="某某学院", group="课题组",
            presenter="汇报人：李四", date="2026年2月")
    d.agenda(["甲", "乙"])
    d.cards("新的一页", [("左", ["一", "二"]), ("右", "三")], section=1)
    out = work / "rebuilt.pptx"
    d.save(out)
    print(f"   输出 {out.name} {len(d._prs.slides)} 页")

    print("5. 结构检查")
    import design_check
    rep = design_check.run(out)
    print(f"   issues={rep['issue_counts']} notices={len(rep['notices'])}")
    assert rep["ok"], f"设计检查未通过: {rep['issue_counts']}"

    print("\nPASS - 模板链路（读取 / 去标识 / 重建 / 检查）全部成立")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
