# src/automizer/analysis_engine.py
from __future__ import annotations
from dataclasses import dataclass
from typing import Any, List, Tuple

import pandas as pd

from src.automizer.conclusions import (
    autopick_for_row,
    # find_template_candidates,
    # ConclusionTemplate,
)

from src.automizer.conclusion_types import ConclusionTemplate
from src.automizer.conclusion_rules import find_template_candidates
from src.automizer.conclusion_decision import explain_decision
from src.automizer.conclusion_render import render_with_place_unit


@dataclass(frozen=True)
class RowAnalysisResult:
    status: str
    matched_template: str
    matched_word: str
    multi: bool
    notes: str
    candidates: List[Tuple[ConclusionTemplate, Any]]


def analyze_row(
    *,
    row: pd.Series,
    reference_df: pd.DataFrame,
    templates: list[ConclusionTemplate],
) -> RowAnalysisResult:
    """
    Єдина точка входу для аналізу одного перехоплення.
    Поки що просто делегує в стару логіку.
    """

    status, m_tmpl, m_word, multi, notes = autopick_for_row(
        row=row,
        reference_df=reference_df,
        templates=templates,
    )

    text = str(row.get("р\\обмін", "") or "")
    candidates = find_template_candidates(text, templates)

    return RowAnalysisResult(
        status=status,
        matched_template=m_tmpl,
        matched_word=m_word,
        multi=multi,
        notes=notes,
        candidates=candidates,
    )
    
    
@dataclass(frozen=True)
class TestAnalysisResult:
    status: str
    leader_template: str
    leader_score: int
    runner_up_score: int
    notes: str
    explain_text: str


def analyze_text(
    *,
    text: str,
    freq_value: Any,
    reference_df: pd.DataFrame,
    templates: list[ConclusionTemplate],
) -> TestAnalysisResult:
    candidates = find_template_candidates(text or "", templates)
    exp = explain_decision(candidates)

    notes = ""
    if exp.leader_name:
        tmpl = next((t for t in templates if t.name == exp.leader_name), None)
        if tmpl:
            notes = render_with_place_unit(tmpl.description, freq_value, reference_df)

    # красивий текст пояснення для UI
    lines = []
    lines.append(f"Статус: {exp.status}")
    if exp.leader_name:
        lines.append(f'Лідер: "{exp.leader_name}" (score={exp.leader_score}, runner_up={exp.runner_up_score})')
    lines.append(f"Причина: {exp.reason}")

    if exp.scores:
        lines.append("Scores:")
        for name, sc in exp.scores.items():
            lines.append(f"  - {name}: {sc}")

    if candidates:
        lines.append("Спрацювали слова:")
        # згрупуємо для зручності
        by_t: dict[str, list[str]] = {}
        for t, r in candidates:
            by_t.setdefault(t.name, []).append(f"{r.word} (p={getattr(r,'probability',0)})")
        for tname, items in by_t.items():
            lines.append(f'  {tname}: ' + "; ".join(items))

    return TestAnalysisResult(
        status=exp.status,
        leader_template=exp.leader_name,
        leader_score=exp.leader_score,
        runner_up_score=exp.runner_up_score,
        notes=notes,
        explain_text="\n".join(lines),
    )
