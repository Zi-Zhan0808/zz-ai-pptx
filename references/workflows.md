# Worked patterns by occasion

Four common academic situations. Each names the argument the deck must make, a
slide skeleton, and the failure mode that specific occasion invites.

All of them obey the same constraints: claim-style titles, a locator on every
content slide, the paper's own figures, and a judgement slide that is actually
your judgement.

---

## 1. 组会汇报 (lab meeting, 10–20 min)

**Argument the deck must make:** *what is new, whether it holds, and what the
group should do about it.* The audience knows the field. Do not teach the
background.

**Skeleton (9–12 slides)**

1. Title — the one finding, as a sentence
2. 问题与已知 — where the current understanding stops (2–3 slides max)
3. 方法要点 — only what changes how the result should be read
4. 主要结果 — 1 figure per slide, panel-by-panel pointing
5. 结果与文献的冲突/一致 — name the conflict explicitly
6. **我的判断** — which explanation you find stronger and why
7. 局限与下一步 — one concrete experiment
8. 讨论问题 — a question the room can actually answer
9. 参考文献

**Beware:** presenting the paper's structure instead of your assessment. A lab
meeting wants your read of the paper, not a translation of it.

---

## 2. 论文精讲 (single-paper deep read, 30–45 min)

**Argument:** *why this paper's central claim is (or is not) supported by its own
evidence.* Requires reading beyond the abstract.

**Skeleton (15–25 slides)**

1. Title — the central claim as the authors state it
2. 研究的真问题 — what was unknown before
3. 实验设计 — the causal structure: what is manipulated, what is controlled
4. 关键图 — 3–5 figures, each with "what it shows" and "what it does **not** show"
5. 替代解释 — at least one alternative explanation and whether it was excluded
6. 统计与稳健性 — effect sizes, n, whether the claim survives the reported test
7. 我的评价 — what holds, what is overclaimed
8. 可延伸的方向 — specific, testable
9. 参考文献

**Beware:** the "全部图表都讲一遍" trap. Choosing which figures to skip is the
expertise being demonstrated. Five well-chosen figures beat fifteen narrated ones.

---

## 3. 综述汇报 (multi-paper synthesis, 20–30 min)

**Argument:** *what the field collectively shows, and where the disagreements
are.* A synthesis is not a list of papers.

**Skeleton (12–18 slides)**

1. Title — the state of the question, as a claim
2. 检索与选取 — how many papers, which criteria (this is a locator slide)
3. 三条主要证据线 — organise by mechanism/claim, **not** by paper
4. 一致之处 — what independent studies converge on
5. 冲突之处 — where they disagree, and a candidate reason (system, dose,
   endpoint, species)
6. 证据强度分层 — which claims rest on n=1, which on independent replication
7. 我的判断 — what the field can and cannot yet conclude
8. 缺口 — the experiment nobody has done
9. 参考文献 — every paper cited, numbered

**Beware:** "paper-by-paper recap" disguised as a review. If slide N is "Paper X
showed…", you are summarising, not synthesising. Reorganise by claim and let one
slide carry three papers that agree.

Multi-source citation: `[1,3-5]` is fine and `refs_check.py` expands ranges.

---

## 4. 开题 / 答辩预演 (proposal or defence rehearsal)

**Argument:** *the question matters, the plan can answer it, and the committee's
likely objections are already addressed.*

**Skeleton (12–18 slides)**

1. Title — the question, precisely
2. 研究缺口 — what is unknown and why it matters for a decision (not "for science")
3. 假设 — falsifiable, one sentence
4. 技术路线 — a flow the committee can critique
5. 关键预实验 — evidence you can already answer it
6. 可行性与风险 — each risk with a fallback
7. 预期贡献 — what changes if it works
8. 参考文献

**Beware:** a literature review that eats 80% of the slot. The committee is
judging the plan. Two slides of gap analysis, then the plan.

For a defence rehearsal, add an **anticipation slide** near the end: the two
hardest questions you expect, and your answer's first line. Do not put the full
answers on the slide — you want to say them.

---

## Cross-cutting: the judgement slide

Every pattern includes one. It is the slide that cannot be generated from the
source material, and therefore the only one that proves a person read it.

What makes it work:

- **A direction**: 「我倾向于 X 而不是 Y，因为 Z」
- **A specific conflict**: 「[2] 报告了相反趋势，但其体系是过表达，可能解释差异」
- **A falsifiable next step**: 「若回补后身份基因部分恢复，机制成立；若不恢复，则 m6A 只是伴随现象」

What makes it fail: 「有待进一步研究」. That is not a judgement, it is a shrug.

## Time budgeting

| Slot | Content slides | Per slide |
|---|---|---|
| 10 min | 5–7 | ~1.5 min |
| 20 min | 10–13 | ~1.5 min |
| 30 min | 15–20 | ~1.5 min |
| 45 min | 22–30 | ~1.5 min |

Leave ~20% of the slot for questions. If the outline exceeds the budget, cut
slides — do not speed up. A deck read at 30 seconds per slide conveys nothing.
