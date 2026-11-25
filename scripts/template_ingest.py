"""
template_ingest.py - read a deck somebody else designed, so the builder can
reuse it instead of imitating it.

Why this exists
---------------
Asking a model to "make it look like this template" produces an imitation:
roughly similar colours, roughly similar chrome, and a hundred small mistakes
in type, spacing and ornament. Reading the file produces the thing itself.

The distinction that matters for authoring
------------------------------------------
A .pptx carries three independent layers, and they are the only three ways a
template can hand anything to a new deck:

  theme    colours and fonts, by name  -> re-skin any layout we draw ourselves
  layout   ornamental chrome the designer drew once (title bar, number badge,
           rules, footer band) -> inherited automatically by every slide we
           add on that layout, and still editable by the user afterwards
  slide    finished pages -> cloned verbatim, for signature pages (cover,
           agenda, closing) whose composition is not worth re-deriving

`ingest` reports all three, plus the one thing that is *always* wrong to
inherit: the original author's institution marks.

Brand marks
-----------
A logo is not identifiable by name - the shape is called "Picture 19". It is
identifiable by repetition and placement, so we score every image on evidence
(repeats across slides, lives in a layout, sits in the header band, extreme
aspect for a wordmark) and report the score rather than a verdict. Removal is
a separate, explicit step: a figure that legitimately appears twice must not be
silently deleted, so the caller confirms.

Usage
-----
    # 1. read it
    python template_ingest.py 模板.pptx --out work/template

    # 2. remove the author's institution marks
    python template_ingest.py 模板.pptx --out work/template \
        --strip 9fda18b8,42a4a3d7,d347a1f8,0da48213

Both steps write work/template/template_profile.json, which build_deck.Deck
loads via `Deck(theme, template=...)`.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

EMU = 914400

# An institution mark lives in the header band, in practice. Content figures
# start below it, because the chrome (title, rules) occupies the top.
BRAND_BAND_IN = 1.35
WORDMARK_ASPECT = 2.5

_PIC = 13
_GROUP = 6


def sha8(blob: bytes) -> str:
    return hashlib.sha1(blob).hexdigest()[:8]


def _in(sh) -> tuple:
    """(left, top, width, height) in inches, tolerant of missing geometry."""
    def conv(v):
        return round(v / EMU, 3) if v is not None else 0.0
    return (conv(sh.left), conv(sh.top), conv(sh.width), conv(sh.height))


def _iter_shapes(shapes):
    """Yield every shape, descending into groups."""
    for sh in shapes:
        yield sh
        if sh.shape_type == _GROUP:
            yield from _iter_shapes(sh.shapes)


# --------------------------------------------------------------------------
# theme
# --------------------------------------------------------------------------
def read_theme(prs) -> dict:
    """
    Colours and fonts as the file declares them.

    We read the XML rather than the python-pptx font API because the API
    resolves *effective* fonts run by run; what a re-skin needs is the theme's
    named slots (accent1..6, dk2, lt2), which only exist in the XML.
    """
    master = prs.slide_masters[0]
    theme_part = master.part.part_related_by(RT.THEME)
    xml = theme_part.blob.decode("utf-8")

    scheme = re.search(r"<a:clrScheme.*?</a:clrScheme>", xml, re.S)
    colors = {}
    if scheme:
        for m in re.finditer(
            r"<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>\s*"
            r'<a:(srgbClr|sysClr)\s+(?:val|lastClr)="([0-9A-Fa-f]{6})"',
            scheme.group(0),
        ):
            colors[m.group(1)] = m.group(3).upper()
    # bg1/tx1 are the scheme aliases PowerPoint actually paints with.
    colors.setdefault("dk1", "000000")
    colors.setdefault("lt1", "FFFFFF")

    fonts = {}
    fs = re.search(r"<a:fontScheme.*?</a:fontScheme>", xml, re.S)
    if fs:
        for kind in ("majorFont", "minorFont"):
            blk = re.search(rf"<a:{kind}>.*?</a:{kind}>", fs.group(0), re.S)
            if not blk:
                continue
            latin = re.search(r'<a:latin typeface="([^"]*)"', blk.group(0))
            eas = [t for t in re.findall(r'<a:ea typeface="([^"]*)"', blk.group(0)) if t]
            fonts[kind] = {
                "latin": latin.group(1) if latin else "",
                "ea": eas[0] if eas else "",
            }
    return {"colors": colors, "fonts": fonts}


# --------------------------------------------------------------------------
# layouts (the reusable chrome)
# --------------------------------------------------------------------------
def read_layouts(prs) -> list:
    """
    One entry per layout, with the geometry a builder needs to stay off the
    chrome: where the title sits and where content may start.
    """
    out = []
    for idx, lay in enumerate(prs.slide_layouts):
        chrome = []
        title_box = None
        for sh in lay.shapes:
            l, t, w, h = _in(sh)
            is_ph = sh.is_placeholder
            ph_type = None
            if is_ph:
                try:
                    ph_type = str(sh.placeholder_format.type)
                except Exception:
                    ph_type = None
            entry = {
                "name": sh.name,
                "id": sh.shape_id,
                "kind": str(sh.shape_type).split(" ")[0],
                "ph": ph_type,
                "box": [l, t, w, h],
            }
            if is_ph and ph_type and "TITLE" in ph_type and title_box is None:
                title_box = entry["box"]
            if not is_ph:
                chrome.append(entry)

        # Header chrome = anything ending in the band. Content must clear it.
        top_bottom = max(
            (e["box"][1] + e["box"][3] for e in chrome
             if e["box"][1] + e["box"][3] <= BRAND_BAND_IN + 0.9),
            default=0.0,
        )
        if title_box:
            top_bottom = max(top_bottom, title_box[1] + title_box[3])
        content_top = round(top_bottom + 0.16, 2) if top_bottom else 0.9

        out.append({
            "index": idx,
            "name": lay.name,
            "title_box": title_box,
            "content_top": content_top,
            "chrome": chrome,
            "slide_count": sum(1 for s in prs.slides
                               if s.slide_layout.name == lay.name),
        })
    return out


# --------------------------------------------------------------------------
# slides (cloneable finished pages)
# --------------------------------------------------------------------------
def read_slides(prs) -> list:
    """
    A catalogue, not a dump: enough for someone to choose which page to clone
    without opening the file. Texts are kept short on purpose.
    """
    out = []
    for i, s in enumerate(prs.slides, 1):
        texts, pics = [], 0
        for sh in s.shapes:
            if sh.shape_type == _PIC:
                pics += 1
            if sh.has_text_frame:
                t = " ".join(sh.text_frame.text.split())
                if t:
                    texts.append(t)
        texts.sort(key=len, reverse=True)
        big = texts[0][:70] if texts else ""
        out.append({
            "n": i,
            "layout": s.slide_layout.name,
            "shape_count": len(s.shapes),
            "pictures": pics,
            "title_guess": big,
            "text_chars": sum(len(t) for t in texts),
        })
    return out


# --------------------------------------------------------------------------
# signature pages
# --------------------------------------------------------------------------
# A deck built in somebody's template needs three composed pages that are not
# worth re-deriving: the cover, the agenda, the closing. They are found by the
# words people put on them, then their text boxes are named by what they hold,
# so template_build can refill them without knowing this file.
_ROLE_WORDS = {
    "closing": ("敬请", "谢谢", "批评指正", "thank you", "thanks"),
    "agenda": ("目录", "contents"),
}
# Pages that explain the template rather than use it. Inheriting one of these
# as "the cover" is the most likely way a template-based deck goes wrong.
_SKIP_WORDS = ("教程", "网盘", "一键换", "替换/删除", "使用说明", "模板说明",
               "instruction", "tutorial")
_LABEL_RULES = [
    ("presenter", r"^\s*汇报人|^\s*报告人|^\s*答辩人|^\s*speaker|presented by"),
    ("date", r"^\s*汇报时间|^\s*报告时间|^\s*日期|^\s*date"),
    ("dept", r"学院|大学|研究所|研究院|institute|universit"),
    ("group", r"课题组|实验室|课题组|group|lab\b"),
    ("heading", r"^\s*目录|^\s*contents"),
    ("kicker", r"^\s*literature\s+report|^\s*journal\s+club|报告人$"),
]


def _all_texts(slide) -> list:
    out = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            t = " ".join(sh.text_frame.text.split())
            if t:
                out.append((sh.name, t, _in(sh)))
    return out


def _full_bleed(slide, canvas) -> bool:
    """
    True when a shape covers essentially the whole page.

    The hero variants of a cover (a green field, a dark photograph) are the
    ones worth cloning; the same template's lighter variants are usually
    decoration comps. Geometry, not guesswork.
    """
    cw, chh = canvas
    for sh in slide.shapes:
        l, t, w, h = _in(sh)
        if (w >= cw * 0.92 and h >= chh * 0.92 and l <= 0.2 and t <= 0.2):
            return True
    return False


def guess_pages(prs) -> dict:
    """
    Propose the three composed pages, with a starter slot map for each.

    The map is a proposal, and it is wrong for some templates - which is why
    every entry carries the box's original text beside its guessed role. A
    caller reads three lines and corrects what is off, instead of hunting for
    "which of the nine text boxes is the date".

    Two rules are role-specific on purpose. An agenda has no title, only a
    heading and a list, so nothing else is mapped - guessing one of its six
    entries as "the title" is how the previous version got it wrong. A closing
    page has one big sentence, so its longest text is mapped to `text`.
    """
    canvas = (prs.slide_width / EMU, prs.slide_height / EMU)
    found = {"cover": [], "agenda": [], "closing": []}

    for i, s in enumerate(prs.slides, 1):
        texts = _all_texts(s)
        joined = " ".join(t for _, t, _ in texts).lower()
        if not joined.strip():
            continue
        if any(w in joined for w in _SKIP_WORDS):
            continue
        role = None
        for r, words in _ROLE_WORDS.items():
            if any(w.lower() in joined for w in words):
                role = r
                break
        if role is None:
            labelled = sum(1 for _, t, _ in texts
                           if any(re.search(p, t, re.I)
                                  for _, p in _LABEL_RULES[:4]))
            latin = [t for _, t, _ in texts if re.match(r"^[\x20-\x7e]{8,}$", t)]
            if labelled >= 1 or latin:
                role = "cover"
        if role:
            found[role].append((i, _full_bleed(s, canvas)))

    pages = {}
    for role, hits in found.items():
        if not hits:
            continue
        # Prefer the hero variant: a full-bleed field reads as "the" cover.
        hero = [h for h in hits if h[1]]
        idx = (hero or hits)[0][0]
        s = prs.slides[idx - 1]
        texts = _all_texts(s)
        slots, review = {}, []

        def take(key, name, was):
            if key not in slots:
                slots[key] = name
                review.append({"role": key, "shape": name, "was": was[:40]})

        if role == "agenda":
            for name, text, _b in texts:
                if re.search(_LABEL_RULES[4][1], text, re.I):
                    take("heading", name, text)
        elif role == "closing":
            for name, text, _b in texts:
                hit = next((lab for lab, pat in _LABEL_RULES
                            if pat and re.search(pat, text, re.I)
                            and lab != "heading"), None)
                if hit:
                    take(hit, name, text)
            # The sentence itself: the biggest type left on the page. Length is
            # a poor proxy here - an English kicker is often longer than the
            # Chinese closing line it sits above.
            rest = [(n, t) for n, t, _b in texts if n not in slots.values()]
            if rest:
                def weight(item):
                    sh = next(x for x in s.shapes if x.name == item[0])
                    size = None
                    if sh.has_text_frame:
                        for p in sh.text_frame.paragraphs:
                            for r in p.runs:
                                if r.font.size:
                                    size = max(size or 0, r.font.size.pt)
                    h = _in(sh)[3]
                    return (size or 0, h, len(item[1]))
                pick = max(rest, key=weight)
                take("text", pick[0], pick[1])
        else:  # cover
            for name, text, _b in texts:
                hit = next((lab for lab, pat in _LABEL_RULES
                            if re.search(pat, text, re.I)), None)
                if hit and hit != "heading":
                    take(hit, name, text)
                    continue
                if re.search(r"[\u4e00-\u9fff]", text) and len(text) >= 6:
                    take("title_cn", name, text)
                elif re.match(r"^[\x20-\x7e]{8,}$", text):
                    take("title_en", name, text)
        pages[role] = {"index": idx, "slots": slots, "widen": {},
                       "review": review}
    return pages


# --------------------------------------------------------------------------
# brand marks
# --------------------------------------------------------------------------
def find_brand_marks(prs) -> list:
    """
    Score every image on how much it looks like an institution mark.

    The score is deliberately additive and legible, because the caller is going
    to override it sometimes and needs to know why a picture scored what it did.

      +3  used by a slide layout or the master (design furniture, not content)
      +2  on two or more slides
      +3  on four or more slides
      +2  entirely inside the header band
      +2  aspect >= 2.5 (a wordmark is a strip; a figure is not)
      -4  big and roughly 4:3 or 16:9 (a figure, near-certainly)
      +2  shares exact pixel dimensions with a stronger mark (same asset family)

    Anything at 4 or above is a candidate. Nothing is deleted here.
    """
    n_slides = len(prs.slides)
    occ = {}

    def note(part_kind, slide_no, sh):
        if sh.shape_type != _PIC:
            return
        try:
            blob = sh.image.blob
        except Exception:
            return
        k = sha8(blob)
        rec = occ.setdefault(k, {
            "sha8": k, "px": None, "slides": set(), "layouts": set(),
            "boxes": [], "blob": blob,
        })
        try:
            rec["px"] = list(sh.image.size)
        except Exception:
            pass
        if part_kind == "slide":
            rec["slides"].add(slide_no)
        else:
            rec["layouts"].add(part_kind)
        rec["boxes"].append((part_kind, _in(sh)))

    for i, s in enumerate(prs.slides, 1):
        for sh in _iter_shapes(s.shapes):
            note("slide", i, sh)
    for idx, lay in enumerate(prs.slide_layouts):
        for sh in _iter_shapes(lay.shapes):
            note(f"layout{idx}", None, sh)
    for sh in _iter_shapes(prs.slide_masters[0].shapes):
        note("master", None, sh)

    scored = []
    for k, rec in occ.items():
        score, why = 0, []
        if rec["layouts"]:
            score += 3
            why.append("in layout/master")
        ns = len(rec["slides"])
        if ns >= 4:
            score += 3
            why.append(f"on {ns} slides")
        elif ns >= 2:
            score += 2
            why.append(f"on {ns} slides")

        boxes = [b for _, b in rec["boxes"]]
        if boxes and all(b[1] + b[3] <= BRAND_BAND_IN for b in boxes):
            score += 2
            why.append("header band")

        aspect = 0.0
        if rec["px"] and rec["px"][1]:
            aspect = rec["px"][0] / rec["px"][1]
        if aspect >= WORDMARK_ASPECT:
            score += 2
            why.append(f"wordmark {aspect:.1f}:1")

        biggest = max((b[2] * b[3] for b in boxes), default=0)
        if biggest >= 3.5 * 2.5 and 0.6 <= aspect <= 1.8:
            score -= 4
            why.append("large figure")

        scored.append({
            "sha8": k, "px": rec["px"], "score": score, "evidence": why,
            "slides": sorted(rec["slides"]),
            "layouts": sorted(rec["layouts"]),
            "boxes": [[b_[0], list(b_[1])] for b_ in rec["boxes"]][:6],
            "aspect": round(aspect, 2),
        })

    # Same-asset families: two crops of one seal share pixel dimensions.
    # Only among small clusters, or the 70 copied 96x96 icons would drag each
    # other over the line.
    families = {}
    for s in scored:
        if s["px"]:
            families.setdefault(tuple(s["px"]), []).append(s)
    for dims, members in families.items():
        if len(members) > 4:
            continue
        if any(m["score"] >= 4 for m in members):
            for m in members:
                if m["score"] < 4:
                    m["score"] += 2
                    m["evidence"].append("shares size with a stronger mark")

    scored.sort(key=lambda s: (-s["score"], -len(s["slides"])))
    return scored


# --------------------------------------------------------------------------
# removal
# --------------------------------------------------------------------------
def _drop_orphan_image_rels(part) -> int:
    """
    Delete image relationships the part's XML no longer mentions.

    Removing the <p:pic> element is only half the job: the relationship from
    the slide to the image part survives, so the image stays reachable, stays
    in the saved package, and the file keeps its 450 KB of someone else's
    stationery. Only image rels are touched - the relationship to the slide
    layout is implicit in the XML and must not be disturbed.
    """
    xml = part.blob.decode("utf-8", "ignore") if hasattr(part, "blob") else ""
    if not xml:
        xml = part._element.xml
    n = 0
    for rId, rel in list(part.rels.items()):
        if rel.reltype != RT.IMAGE:
            continue
        if f'r:embed="{rId}"' in xml or f'r:link="{rId}"' in xml or f'r:id="{rId}"' in xml:
            continue
        part.rels.pop(rId)
        n += 1
    return n


def drop_slides(prs, indices) -> int:
    """
    Remove whole pages by 1-based number.

    Templates ship pages that are not part of the deck: a "how to swap the
    logo" instruction page, a colour-swatch page. After the logo is gone, an
    instruction page telling the user to swap the logo is worse than absent, so
    removal belongs to the same step.
    """
    drop = set(int(i) for i in indices)
    keep = [s for i, s in enumerate(prs.slides, 1) if i not in drop]
    sldIdLst = prs.slides._sldIdLst
    removed = 0
    for sldId in list(sldIdLst):
        slide = None
        for s in prs.slides:
            if prs.part.related_part(sldId.rId) is s.part:
                slide = s
                break
        if slide is None:
            continue
        if any(slide.part is s.part for s in keep):
            continue
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)
        removed += 1
    return removed


def strip_images(src: str, dst: str, shas: set, drop_pages=()) -> dict:
    """
    Delete every picture whose image has one of `shas`, in slides, layouts and
    masters alike, then drop the media parts nothing references any more.

    Removal is by element, not by hiding: a hidden picture still ships in the
    file and still becomes visible the moment someone toggles it.
    """
    prs = Presentation(src)
    removed = {"slides": 0, "layouts": 0, "masters": 0, "rels": 0, "pages": 0}

    def purge(shapes) -> int:
        n = 0
        for sh in list(shapes):
            if sh.shape_type == _GROUP:
                n += purge(sh.shapes)
            if sh.shape_type == _PIC:
                try:
                    if sha8(sh.image.blob) in shas:
                        sh._element.getparent().remove(sh._element)
                        n += 1
                except Exception:
                    pass
        return n

    for s in prs.slides:
        hit = purge(s.shapes)
        if hit:
            removed["slides"] += hit
            removed["rels"] += _drop_orphan_image_rels(s.part)
    for lay in prs.slide_layouts:
        hit = purge(lay.shapes)
        if hit:
            removed["layouts"] += hit
            removed["rels"] += _drop_orphan_image_rels(lay.part)
    for m in prs.slide_masters:
        hit = purge(m.shapes)
        if hit:
            removed["masters"] += hit
            removed["rels"] += _drop_orphan_image_rels(m.part)

    if drop_pages:
        removed["pages"] = drop_slides(prs, drop_pages)

    prs.save(dst)
    dropped = _prune_media(dst)
    removed["media_parts"] = dropped
    return removed


def _prune_media(path: str) -> int:
    """
    Rewrite the package without media parts no .rels points at.

    A removed picture leaves its image behind: python-pptx keeps the part,
    PowerPoint keeps the bytes, and the file stays 5 MB of other people's
    decoration. Parts are only reachable through relationships, so an
    unreferenced media part is unreachable by construction.
    """
    tmp = path + ".tmp"
    dropped = 0
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        referenced = set()
        for n in names:
            if not n.endswith(".rels"):
                continue
            base = n.rsplit("/_rels/", 1)[0]
            for tgt in re.findall(r'Target="([^"]+)"', zin.read(n).decode("utf-8", "ignore")):
                if tgt.startswith(("http:", "https:", "file:", "mailto:")):
                    continue
                # relative to the part that owns the .rels file
                stack = (base or "").split("/")
                for seg in tgt.lstrip("/").split("/"):
                    if seg == "..":
                        if stack:
                            stack.pop()
                    elif seg not in ("", "."):
                        stack.append(seg)
                referenced.add("/".join(stack))

        keep, gone = [], []
        for n in names:
            if n.startswith("ppt/media/") and n not in referenced:
                gone.append(n)
            else:
                keep.append(n)

        if not gone:
            return 0

        ct_name = "[Content_Types].xml"
        ct = zin.read(ct_name).decode("utf-8")
        for g in gone:
            ct = re.sub(r'<Override[^>]*PartName="/%s"[^>]*/>' % re.escape(g), "", ct)
        dropped = len(gone)

        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in keep:
                data = ct.encode("utf-8") if n == ct_name else zin.read(n)
                zout.writestr(n, data)
    shutil.move(tmp, path)
    return dropped


# --------------------------------------------------------------------------
# profile
# --------------------------------------------------------------------------
def find_badge_color(prs) -> str | None:
    """
    The colour the template paints its page numeral in.

    This is not a guess: the numeral sits in the roundel the layout draws, so its
    colour has to be the one that reads on *that* disc. On a template with a
    green disc the numeral is white; hard-coding the brand colour there produces
    dark green on dark green, which is exactly the kind of "close enough"
    substitution that makes an inherited design look broken.
    """
    from theme_library import theme_palette
    palette = theme_palette(prs)
    alias = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}

    for s in prs.slides:
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            if not re.fullmatch(r"\d{1,2}", sh.text_frame.text.strip()):
                continue
            left, top = _in(sh)[0], _in(sh)[1]
            if left > 2.0 or top > 1.3:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    try:
                        return str(r.font.color.rgb).upper()
                    except Exception:
                        pass
                    try:
                        key = str(r.font.color.theme_color).split(".")[-1]
                        key = key.split(" ")[0]
                        name = {"BACKGROUND_1": "lt1", "TEXT_1": "dk1",
                                "ACCENT_1": "accent1", "ACCENT_2": "accent2",
                                "ACCENT_3": "accent3"}.get(key)
                        if name and name in palette:
                            return palette[name]
                    except Exception:
                        pass
            m = re.search(r'<a:schemeClr val="(\w+)"', sh._element.xml)
            if m:
                key = alias.get(m.group(1), m.group(1))
                if key in palette:
                    return palette[key]
    return None


def build_profile(path: str) -> dict:
    prs = Presentation(path)
    theme = read_theme(prs)
    layouts = read_layouts(prs)
    slides = read_slides(prs)
    marks = find_brand_marks(prs)
    return {
        "source": Path(path).name,
        "canvas": {
            "w": round(prs.slide_width / EMU, 3),
            "h": round(prs.slide_height / EMU, 3),
            "ratio": round(prs.slide_width / prs.slide_height, 3),
        },
        "theme": theme,
        "layouts": layouts,
        "slides": slides,
        "pages": guess_pages(prs),
        "badge_color": find_badge_color(prs),
        "brand_marks": marks,
        "brand_candidates": [m["sha8"] for m in marks if m["score"] >= 4],
    }


def write_markdown(prof: dict, png=None) -> str:
    L = []
    A = L.append
    A(f"# 模板画像 · {prof['source']}")
    A("")
    c = prof["canvas"]
    A(f"画布 {c['w']} × {c['h']} in（比例 {c['ratio']}），共 {len(prof['slides'])} 页、"
      f"{len(prof['layouts'])} 个版式。")
    A("")
    A("## 主题色")
    A("")
    A("| 槽位 | 色值 | 用途 |")
    A("|---|---|---|")
    use = {
        "dk1": "正文/标题墨色", "lt1": "画布", "dk2": "深色次调",
        "lt2": "浅色次调", "accent1": "主色", "accent2": "主色次档",
        "accent3": "主色三档", "accent4": "辅助色", "accent5": "辅助色",
        "accent6": "辅助色", "hlink": "超链接", "folHlink": "已访问链接",
    }
    for k in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
              "accent4", "accent5", "accent6", "hlink", "folHlink"):
        if k in prof["theme"]["colors"]:
            A(f"| `{k}` | `#{prof['theme']['colors'][k]}` | {use.get(k, '')} |")
    A("")
    f = prof["theme"]["fonts"]
    if f:
        A("字体：" + "；".join(
            f"{k} 西文 `{v.get('latin') or '—'}` 中文 `{v.get('ea') or '—'}`"
            for k, v in f.items()))
    A("")
    A("## 版式（可复用的装饰层）")
    A("")
    A("| # | 名称 | 使用页数 | 标题框 | 内容可起始于 | 装饰元素 |")
    A("|---|---|---|---|---|---|")
    for lay in prof["layouts"]:
        tb = lay["title_box"]
        tbs = f"({tb[0]},{tb[1]}) {tb[2]}×{tb[3]}" if tb else "—"
        A(f"| {lay['index']} | {lay['name']} | {lay['slide_count']} | {tbs} | "
          f"{lay['content_top']} in | {len(lay['chrome'])} 个 |")
    A("")
    A("## 页面清单（可克隆的成品页）")
    A("")
    A("| 页 | 版式 | 形状 | 图 | 主要文字 |")
    A("|---|---|---|---|---|")
    for s in prof["slides"]:
        A(f"| {s['n']} | {s['layout']} | {s['shape_count']} | {s['pictures']} | "
          f"{s['title_guess'] or '（无文字）'} |")
    A("")
    A("## 机构标识（需要用户确认是否删除）")
    A("")
    A("| 图 | 像素 | 得分 | 出现页 | 证据 |")
    A("|---|---|---|---|---|")
    for m in prof["brand_marks"][:12]:
        if m["score"] <= 0:
            continue
        px = f"{m['px'][0]}×{m['px'][1]}" if m["px"] else "—"
        A(f"| `{m['sha8']}` | {px} | {m['score']} | "
          f"{','.join(map(str, m['slides'])) or '—'} | {'；'.join(m['evidence'])} |")
    A("")
    A("得分 ≥ 4 的建议当作机构标识处理。**删除前请确认**：一个合法重复出现的"
      "插图（例如对照图用两次）也会得分。")
    A("")
    pages = prof.get("pages") or {}
    if pages:
        A("## 签名页（可整页克隆）")
        A("")
        A("自动识别的槽位是**提案**，请对照「原本写的是」核对后再用：")
        A("")
        for role, e in pages.items():
            A(f"### {role} — 第 {e['index']} 页")
            A("")
            A("| 槽位 | 形状名 | 原本写的是 |")
            A("|---|---|---|")
            for row in e.get("review", []):
                A(f"| `{row['role']}` | `{row['shape']}` | {row['was']} |")
            if e.get("widen"):
                A("")
                A("需加宽的窄框：" + "，".join(
                    f"`{k}` → {v}in" for k, v in e["widen"].items()))
            A("")
    return "\n".join(L)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="读取一个给定模板，输出可复用的画像")
    ap.add_argument("template")
    ap.add_argument("--out", default="template_profile",
                    help="输出目录（写 profile.json / profile.md / clean.pptx）")
    ap.add_argument("--strip", default="",
                    help="要删除的图片 sha8，逗号分隔；给 all 则删除所有候选")
    ap.add_argument("--strip-all-candidates", action="store_true",
                    help="删除所有得分 >= 4 的候选")
    ap.add_argument("--drop-pages", default="",
                    help="要整页删除的页码，逗号分隔（如模板自带的使用说明页）")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args(argv)

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)

    prof = build_profile(a.template)
    (out / "template_profile.json").write_text(
        json.dumps(prof, ensure_ascii=False, indent=1), encoding="utf-8")
    (out / "template_profile.md").write_text(write_markdown(prof), encoding="utf-8")

    if not a.quiet:
        print(f"画布 {prof['canvas']['w']}×{prof['canvas']['h']} in，"
              f"{len(prof['slides'])} 页，{len(prof['layouts'])} 版式")
        print("主题色:", ", ".join(f"{k}=#{v}" for k, v in
                                  prof["theme"]["colors"].items()))
        print("标识候选:", prof["brand_candidates"] or "（无）")

    shas = set()
    if a.strip_all_candidates:
        shas |= set(prof["brand_candidates"])
    if a.strip and a.strip.lower() == "all":
        shas |= set(prof["brand_candidates"])
    elif a.strip:
        shas |= {s.strip().lower() for s in a.strip.split(",") if s.strip()}

    if shas or a.drop_pages:
        clean = out / (Path(a.template).stem + "-clean.pptx")
        drop = [int(x) for x in a.drop_pages.split(",") if x.strip()]
        rem = strip_images(a.template, str(clean), shas, drop)

        # Re-read the cleaned file, because the profile has to describe the file
        # the builder will actually open. Dropping a page renumbers every page
        # after it, so a `pages.cover.index` of 5 taken from the original would
        # silently point at a different slide.
        cleaned = build_profile(str(clean))
        cleaned["stripped"] = {"shas": sorted(shas), "dropped_pages": drop,
                               **rem, "output": clean.name,
                               "source_original": Path(a.template).name}
        prof = cleaned
        (out / "template_profile.json").write_text(
            json.dumps(prof, ensure_ascii=False, indent=1), encoding="utf-8")
        (out / "template_profile.md").write_text(
            write_markdown(prof), encoding="utf-8")
        if not a.quiet:
            print(f"已删除 幻灯片 {rem['slides']} 处、版式 {rem['layouts']} 处、"
                  f"母版 {rem['masters']} 处，顺带断开 {rem['rels']} 条图片关系；"
                  f"清理无引用媒体 {rem['media_parts']} 个")
            if rem["pages"]:
                print(f"整页删除 {rem['pages']} 页，剩余 {len(prof['slides'])} 页")
            before = Path(a.template).stat().st_size
            after = clean.stat().st_size
            print(f"体积 {before/1048576:.1f} MB -> {after/1048576:.1f} MB")
            print("输出:", clean)
        if not (rem["slides"] or rem["layouts"] or rem["masters"]):
            print("警告: 没有删到任何形状，检查 sha8 是否正确", file=sys.stderr)
            return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
