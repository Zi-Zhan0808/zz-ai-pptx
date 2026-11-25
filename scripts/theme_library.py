"""
theme_library.py - a colour system with a budget, not a palette list.

The failure this replaces
-------------------------
"淡彩多色配色" as a prompt produces what any list of pretty colours produces: a
dozen mid-saturated fills, all fully opaque, all roughly equally loud. Measured
on such a deck: 8 hue families, peak saturation 77%, nine distinct fills on one
page, coloured area totalling more than the page itself. Nothing is wrong with
any single colour - the deck is wrong because every element claims the same
amount of attention, so the reader has nowhere to look first.

The fixing idea
---------------
Colour should get *lighter as it gets bigger*, and darker as it gets smaller.
A panel covering a third of the page must be a tint you barely notice; the small
thing you want noticed must be the only dark saturated mark on the slide. That
inverts the usual habit of picking a bright colour and filling shapes with it.

So every theme here is a ladder over a single hue, and the levels are ranked:

    canvas   L .972   the page
    wash     L .940   large bands, table stripes
    panel    L .900   card fills
    veil     L .845   one emphasis block per page, at most
    accent   L .375   rules, numerals, chips - small, dark, saturated
    ink      L .135   body text
    muted    L .455   captions

Area budget, enforced by scripts/palette_audit.py:

    surfaces (canvas..veil)   90% of coloured area, saturation <= .34
    accent                     <= 8% of coloured area, saturation .62
    a second hue               allowed, <= 6%, and never as a surface

Modes
-----
    mono      one hue, seven levels                      (quiet, academic)
    duotone   one hue for surfaces + one for accent      (the safe default)
    trio      surfaces + two accents                     (only for lively occasions)

`pick_for()` chooses a family from the topic. It is deterministic given a seed,
so the same brief reproduces the same deck, but different briefs land on
different hues - which is the point: a reader should not be able to tell that
every deck came out of the same tool.
"""
from __future__ import annotations

import colorsys
import hashlib
from dataclasses import dataclass

# ---------------------------------------------------------------- hue families
# The mood prose is for humans reading this file; `KEYWORDS` is what actually
# matches. Prose cannot be matched reliably - a topic says "史料", the prose says
# "历史", and the lookup silently returns the neutral default.
FAMILIES = [
    ("clay",      18,  "人文 · 历史 · 教育 · 田野 · 考古 · 社会学 · 政策"),
    ("vermilion", 8,   "警示 · 风险 · 冲突 · 批评 · 争议 · 局限"),
    ("amber",     36,  "能源 · 材料 · 制造 · 工程 · 交通 · 机械 · 热"),
    ("ochre",     48,  "地理 · 土壤 · 林业 · 农业 · 野外"),
    ("olive",     72,  "植物 · 生态 · 环境 · 可持续 · 碳 · 气候"),
    ("moss",      96,  "生物 · 细胞 · 分子 · 医学 · 遗传 · 蛋白"),
    ("pine",      146, "健康 · 临床 · 医院 · 护理 · 药学 · 公共卫生"),
    ("celadon",   172, "数据 · 统计 · 方法 · 测量 · 实验 · 信号"),
    ("teal",      192, "计算 · 算法 · 网络 · 系统 · 软件 · 信息"),
    ("indigo",    214, "物理 · 数学 · 理论 · 模型 · 推理 · 逻辑"),
    ("violet",    252, "认知 · 心理 · 语言 · 学习 · 神经"),
    ("wisteria",  282, "艺术 · 设计 · 媒体 · 传播 · 文化"),
    ("rose",      330, "社会 · 性别 · 访谈 · 质性 · 关怀"),
    ("graphite",  210, "综述 · 总论 · 元分析 · 通用"),
]

KEYWORDS = {
    "clay":      ("历史", "史料", "考古", "文献学", "地方志", "人文", "哲学",
                  "文学", "教育", "课程", "教学", "田野", "民俗", "档案"),
    "vermilion": ("风险", "危机", "冲突", "争议", "批判", "缺陷", "漏洞",
                  "失败", "攻击", "安全", "伦理"),
    "amber":     ("能源", "电池", "材料", "制造", "工程", "机械", "热管理",
                  "燃烧", "光伏", "风能", "储氢", "焊接"),
    "ochre":     ("林业", "土壤", "地理", "农业", "森林", "草原", "流域",
                  "水土", "荒漠", "作物", "育种", "农田"),
    "olive":     ("生态", "碳汇", "碳排放", "碳中和", "气候", "植被", "物种",
                  "多样性", "环境", "污染", "遥感反演"),
    "moss":      ("细胞", "蛋白", "基因", "分子", "遗传", "酶", "代谢",
                  "自噬", "凋亡", "rna", "dna", "免疫", "肿瘤", "抗体"),
    "pine":      ("临床", "患者", "疾病", "医院", "护理", "药", "治疗",
                  "诊断", "队列", "随访", "卒中", "疫", "公共卫生"),
    "celadon":   ("数据", "统计", "测量", "实验", "指标", "评估", "问卷",
                  "误差", "偏差", "显著性", "回归", "抽样"),
    "teal":      ("算法", "模型压缩", "网络", "系统", "软件", "计算机",
                  "深度学习", "神经网络", "检索", "分布式", "芯片", "架构"),
    "indigo":    ("物理", "数学", "理论", "推理", "逻辑", "证明", "量子",
                  "力学", "方程", "优化理论"),
    "violet":    ("认知", "心理", "语言", "学习", "神经", "记忆", "注意",
                  "行为", "脑"),
    "wisteria":  ("设计", "艺术", "媒体", "传播", "文化", "视觉", "交互",
                  "影像", "叙事", "美学"),
    "rose":      ("性别", "访谈", "质性", "关怀", "老龄化", "社区", "身份",
                  "情感", "群体"),
    "graphite":  ("综述", "总论", "元分析", "方法学", "导论"),
}

NO_SURFACE = {"vermilion"}

# Levels are stated as *perceived* luminance, not HSL lightness. HSL lightness is
# not luminance: at L=.375 a yellow is far brighter than a blue, so pinning L
# makes half the families fail contrast while the other half looks washed out.
# Solving for luminance keeps every family equally readable and equally calm.
LEVELS = {
    #  name      S      target relative luminance
    # A canvas at luminance .955 is #FAFAF8 - technically tinted, visually white,
    # which is the "always the same pale page" problem in another costume. .900
    # still clears 9.5:1 against the ink while carrying a hue you can name.
    "canvas": (0.38, 0.900),
    "wash":   (0.40, 0.820),
    "panel":  (0.42, 0.730),
    "veil":   (0.45, 0.575),
    # Cards sit *above* the canvas in brightness, not below it. A card that is
    # darker than the page reads as a hole; one that is a shade lighter reads as
    # a sheet lying on the page, which is what a card is.
    "card":   (0.30, 0.962),
    "accent": (0.62, 0.140),
    "ink":    (0.26, 0.050),
    "muted":  (0.18, 0.155),
}

# Families whose whole point is to be quiet: graphite is the "no opinion" option,
# so its saturation is scaled down rather than deleted - a hint of blue keeps it
# from looking like a printer error.
SAT_SCALE = {"graphite": 0.22}


def _hex(h: float, s: float, l: float) -> str:
    r, g, b = colorsys.hls_to_rgb((h % 360) / 360.0, l, s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def _lum(hexstr: str) -> float:
    h = hexstr.lstrip("#")
    c = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    c = [(v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4) for v in c]
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]


def contrast(a: str, b: str) -> float:
    la, lb = _lum(a), _lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def _solve(hue: float, sat: float, target: float) -> str:
    """
    The HSL lightness at which this hue reaches `target` relative luminance.

    Bisection, because luminance is monotone in lightness for a fixed hue and
    saturation - so 40 iterations land on the answer and no closed form is
    needed. Falls back to the closest end if the target is unreachable at this
    saturation (a very pale hue cannot be driven to luminance .05).
    """
    lo, hi = 0.0, 1.0
    best = _hex(hue, sat, 0.5)
    for _ in range(40):
        mid = (lo + hi) / 2
        cand = _hex(hue, sat, mid)
        got = _lum(cand)
        best = cand
        if abs(got - target) < 0.0015:
            return cand
        if got < target:
            lo = mid
        else:
            hi = mid
    return best


def _level(hue: float, level: str, shift: float = 0.0,
           sat_scale: float = 1.0) -> str:
    s, target = LEVELS[level]
    return _solve(hue + shift, s * sat_scale, target)


@dataclass
class Spec:
    """A resolved colour system: every level named, ready to hand to Theme."""
    name: str
    mode: str
    hue: float
    family: str
    canvas: str = ""
    wash: str = ""
    panel: str = ""
    veil: str = ""
    card: str = ""
    accent: str = ""
    ink: str = ""
    muted: str = ""
    accent2: str = ""     # second accent hue (trio only)
    accent3: str = ""     # third accent hue (trio only)
    on_accent: str = "FFFFFF"

    def __post_init__(self):
        if not self.canvas:
            k = SAT_SCALE.get(self.family, 1.0)
            self.canvas = _level(self.hue, "canvas", sat_scale=k)
            self.wash = _level(self.hue, "wash", sat_scale=k)
            self.panel = _level(self.hue, "panel", sat_scale=k)
            self.veil = _level(self.hue, "veil", sat_scale=k)
            self.card = _level(self.hue, "card", sat_scale=k)
            self.accent = _level(self.hue, "accent")
            self.ink = _level(self.hue, "ink", sat_scale=k)
            self.muted = _level(self.hue, "muted", sat_scale=k)
        # accents sit a deliberate distance from the dominant hue: a neighbour
        # (rotating 20-30 deg) reads as the same family, a complement reads as a
        # second voice. Two voices is the most a single page can carry.
        if self.mode in ("duotone", "trio") and not self.accent2:
            self.accent2 = _level(self.hue + 148, "accent")
        if self.mode == "trio" and not self.accent3:
            self.accent3 = _level(self.hue + 96, "accent")
        if self.mode == "mono":
            self.accent2 = ""

    @property
    def accent_hues(self) -> int:
        return {"mono": 0, "duotone": 1, "trio": 2}[self.mode]

    def verify(self) -> list:
        """The pairings that must hold, checked rather than assumed."""
        bad = []
        for label, fg, bg, need in (
            ("正文在画布上", self.ink, self.canvas, 7.0),
            ("注释在画布上", self.muted, self.canvas, 4.5),
            ("强调在画布上", self.accent, self.canvas, 4.5),
            ("强调在面板上", self.accent, self.panel, 4.0),
            ("正文在面板上", self.ink, self.panel, 6.0),
            ("正文在卡片上", self.ink, self.card, 7.0),
            ("反白在强调底上", self.on_accent, self.accent, 3.0),
        ):
            got = contrast(fg, bg)
            if got < need:
                bad.append(f"{label} {got:.1f}:1 < {need}:1")
        # the ladder must climb: each surface level visibly lighter than the next
        if _lum(self.card) <= _lum(self.canvas):
            bad.append(f"卡片 {self.card} 不比画布 {self.canvas} 更亮，会看成凹陷")
        order = [self.canvas, self.wash, self.panel, self.veil]
        for i in range(len(order) - 1):
            if _lum(order[i]) <= _lum(order[i + 1]):
                bad.append(f"阶梯不单调: {order[i]} 不亮于 {order[i + 1]}")
        return bad


def theme_palette(prs) -> dict:
    """
    The deck's named colour slots, read from theme XML.

    `<a:sysClr val="window" lastClr="FFFFFF"/>` puts the real colour on
    `lastClr`, not `val`, so a pattern that expects the hex immediately after the
    element name silently loses lt1/dk1 - and every default derived from them
    turns white on white. Three modules used to carry their own copy of this
    lookup, all three with that bug. One implementation cannot drift from itself.
    """
    import re as _re
    out = {}
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        xml = prs.slide_masters[0].part.part_related_by(RT.THEME).blob.decode(
            "utf-8", "ignore")
    except Exception:
        return out
    for name in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6", "hlink", "folHlink"):
        blk = _re.search(rf"<a:{name}>(.*?)</a:{name}>", xml, _re.S)
        if not blk:
            continue
        m = _re.search(r'(?:lastClr|val)="([0-9A-Fa-f]{6})"', blk.group(1))
        if m:
            out[name] = m.group(1).upper()
    return out


def build(family: str = "teal", mode: str = "duotone",
          font: str = "Microsoft YaHei") -> Spec:
    """A named family at a named mode. Unknown names raise rather than guess."""
    table = {f[0]: f for f in FAMILIES}
    if family not in table:
        raise ValueError(f"没有这个色族 {family!r}；可选 {sorted(table)}")
    if mode not in ("mono", "duotone", "trio"):
        raise ValueError(f"没有这个模式 {mode!r}；可选 mono / duotone / trio")
    if mode != "mono" and family in NO_SURFACE:
        # A warning colour used as a page surface stops being a warning.
        family = "clay"
    _, hue, _ = table[family]
    spec = Spec(name=f"{family}-{mode}", mode=mode, hue=hue, family=family)
    bad = spec.verify()
    if bad:
        raise ValueError("色阶校验失败：" + "；".join(bad))
    return spec


def pick_for(topic: str = "", *, mode: str = "duotone",
             seed: str | None = None, prefer: str | None = None) -> Spec:
    """
    Choose a family from the topic, without always choosing the same one.

    Scoring counts keyword hits in the topic; ties are broken by a hash of the
    seed so two decks with equally plausible topics still differ. A topic with no
    hits gets graphite, which is sober rather than decorative - silence is better
    than a random costume.
    """
    if prefer:
        return build(prefer, mode)
    text = (topic or "").lower()
    scored = []
    for name, hue, _mood in FAMILIES:
        hits = sum(1 for w in KEYWORDS.get(name, ()) if w.lower() in text)
        scored.append((hits, name))
    best = max(h for h, _ in scored)
    if best == 0:
        return build("graphite", mode)
    tied = sorted(n for h, n in scored if h == best)
    pick = hashlib.sha1((seed or topic or "x").encode("utf-8")).hexdigest()
    return build(tied[int(pick[:8], 16) % len(tied)], mode)


def all_families() -> list:
    return [{"family": n, "hue": h, "mood": m} for n, h, m in FAMILIES]
