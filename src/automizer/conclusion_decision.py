# src/automizer/conclusion_decision.py
from __future__ import annotations

from typing import List, Tuple, Optional

from src.automizer.conclusion_types import ConclusionTemplate, KeywordRule
from src.automizer.conclusion_consts import STATUS_EMPTY, STATUS_NEED_APPROVE, STATUS_APPROVED


def group_by_template(
    candidates: List[Tuple[ConclusionTemplate, KeywordRule]]
):
    groups: dict[str, list[KeywordRule]] = {}
    order: list[ConclusionTemplate] = []
    seen: set[str] = set()

    for tmpl, rule in candidates:
        name = tmpl.name
        groups.setdefault(name, []).append(rule)
        if name not in seen:
            order.append(tmpl)
            seen.add(name)
    return groups, order

def compute_scores(
    candidates: list[tuple[ConclusionTemplate, KeywordRule]]
) -> dict[str, int]:
    """
    Повертає:
        {template_name: сумарний score}
    """
    scores: dict[str, int] = {}

    for tmpl, rule in candidates:
        scores.setdefault(tmpl.name, 0)
        scores[tmpl.name] += int(getattr(rule, "probability", 0))

    return scores

def decide_status_by_candidates(
    candidates: list[tuple[ConclusionTemplate, KeywordRule]]
) -> tuple[str, str]:
    """
    Повертає:
        (status, leader_template_name)
    """

    if not candidates:
        return STATUS_EMPTY, ""

    scores = compute_scores(candidates)

    ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)

    leader_name, leader_score = ranked[0]
    runner_score = ranked[1][1] if len(ranked) > 1 else 0

    # Умова авто-апруву
    if leader_score >= 2 and (
        len(ranked) == 1 or leader_score >= runner_score + 2
    ):
        return STATUS_APPROVED, leader_name

    return STATUS_NEED_APPROVE, leader_name


def choose_primary(
    candidates: list[tuple[ConclusionTemplate, KeywordRule]],
    leader_name: str,
):
    """
    Повертає:
        (template, matched_word, multi)
    """

    leader_candidates = [
        (tmpl, rule)
        for tmpl, rule in candidates
        if tmpl.name == leader_name
    ]

    if not leader_candidates:
        return None, "", False

    tmpl = leader_candidates[0][0]

    # слово з p=1 якщо є
    rule = next(
        (r for _, r in leader_candidates if r.probability == 1),
        leader_candidates[0][1],
    )

    multi = len(set(t.name for t, _ in candidates)) > 1

    return tmpl, rule.word, multi



from dataclasses import dataclass

@dataclass(frozen=True)
class DecisionExplain:
    status: str
    leader_name: str
    leader_score: int
    runner_up_score: int
    autoapprove: bool
    reason: str
    scores: dict[str, int]


def compute_scores(
    candidates: list[tuple[ConclusionTemplate, KeywordRule]]
) -> dict[str, int]:
    scores: dict[str, int] = {}
    for tmpl, rule in candidates:
        scores[tmpl.name] = scores.get(tmpl.name, 0) + int(getattr(rule, "probability", 0))
    return scores


def explain_decision(
    candidates: list[tuple[ConclusionTemplate, KeywordRule]]
) -> DecisionExplain:
    if not candidates:
        return DecisionExplain(
            status=STATUS_EMPTY,
            leader_name="",
            leader_score=0,
            runner_up_score=0,
            autoapprove=False,
            reason="Збігів не знайдено (0 кандидатів).",
            scores={},
        )

    scores = compute_scores(candidates)
    ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)

    leader_name, leader_score = ranked[0]
    runner_up_score = ranked[1][1] if len(ranked) > 1 else 0

    # умова авто-апруву:
    autoapprove = leader_score >= 2 and (len(ranked) == 1 or leader_score >= runner_up_score + 2)

    if autoapprove:
        if len(ranked) == 1:
            reason = f'Єдина категорія "{leader_name}" має score={leader_score} (>=2).'
        else:
            reason = f'Лідер "{leader_name}" має score={leader_score} і випереджає конкурентів на ≥2 (runner_up={runner_up_score}).'
        status = STATUS_APPROVED
    else:
        reason = f'Недостатня перевага/сила сигналу: leader={leader_score}, runner_up={runner_up_score}.'
        status = STATUS_NEED_APPROVE

    return DecisionExplain(
        status=status,
        leader_name=leader_name,
        leader_score=leader_score,
        runner_up_score=runner_up_score,
        autoapprove=autoapprove,
        reason=reason,
        scores=dict(ranked),
    )
