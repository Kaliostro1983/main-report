# src/automizer/conclusions.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Tuple

import json
import logging
import pandas as pd

from src.automizer.freqdict import lookup_place_unit, COL_FREQ

from src.automizer.conclusion_rules import find_template_candidates
from src.automizer.conclusion_decision import decide_status_by_candidates, choose_primary, group_by_template
from src.automizer.conclusion_render import render_with_place_unit

from src.automizer.conclusion_types import ConclusionTemplate, KeywordRule

from src.automizer.conclusion_consts import (
    STATUS_EMPTY, STATUS_NEED_APPROVE, STATUS_APPROVED,
    COL_STATUS, COL_NOTES, COL_MATCHED_TEMPLATE, COL_MATCHED_WORD, COL_MULTI_MATCH,
    TEXT_COL,
)
from src.automizer.conclusion_types import ConclusionTemplate, KeywordRule

logger = logging.getLogger(__name__)


def _norm_text(s: Any) -> str:
    return str(s).strip().lower()


def load_conclusions(conclusions_path: str | Path) -> list[ConclusionTemplate]:
    """
    Підтримує новий формат:
      {"keywords": [{"word":"...", "probability":0|1}, ...]}
    і зворотну сумісність:
      {"keywords": ["рядок1", "рядок2", ...]} -> probability=0
    """
    p = Path(conclusions_path)
    if not p.exists():
        raise FileNotFoundError(f"conclusions.json not found: {p}")

    raw = json.loads(p.read_text(encoding="utf-8"))
    items = raw.get("conclusions", [])
    result: list[ConclusionTemplate] = []

    for item in items:
        shortcut = str(item.get("shortcut") or item.get("shorcut") or "").strip()
        description = str(item.get("description", "")).strip()
        name = str(item.get("name", "")).strip()

        kw_list: list[KeywordRule] = []
        raw_kws = item.get("keywords", [])
        for k in raw_kws:
            if isinstance(k, dict):
                word = str(k.get("word", "")).strip()
                prob = int(k.get("probability", 0))
                exceptions = k.get("exceptions") or []
                exceptions = [str(x).strip().lower() for x in exceptions]
            else:
                # зворотна сумісність для старих JSON
                word = str(k).strip()
                prob = 0
            if not word:
                continue
            prob = 1 if prob == 1 else 0
            kw_list.append(
                KeywordRule(
                    word=word,
                    probability=prob,
                    exceptions=exceptions,
                )
            )

        result.append(
            ConclusionTemplate(
                name=name,
                description=description,
                keywords=kw_list,
                shortcut=shortcut,
            )
        )
    return result


# conclusions.py
import re







def autopick_for_row(
    row: pd.Series,
    reference_df: pd.DataFrame,
    templates: list[ConclusionTemplate],
) -> tuple[str, str, str, bool, str]:
    text = str(row.get(TEXT_COL, "") or "")
    freq_val = row.get(COL_FREQ)

    candidates = find_template_candidates(text, templates)
    status, leader_name = decide_status_by_candidates(candidates)

    matched_template = ""
    matched_word = ""
    notes = ""

    first_tmpl, chosen_word, multi_templates = choose_primary(
        candidates,
        leader_name,
    )

    if first_tmpl:
        matched_template = first_tmpl.name
        matched_word = chosen_word
        notes = render_with_place_unit(first_tmpl.description, freq_val, reference_df)

    return status, matched_template, matched_word, multi_templates, notes




def ensure_df_service_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Гарантуємо наявність службових колонок."""
    if COL_STATUS not in df.columns:
        df[COL_STATUS] = STATUS_EMPTY
    if COL_NOTES not in df.columns:
        df[COL_NOTES] = ""
    if COL_MATCHED_TEMPLATE not in df.columns:
        df[COL_MATCHED_TEMPLATE] = ""
    if COL_MATCHED_WORD not in df.columns:
        df[COL_MATCHED_WORD] = ""
    if COL_MULTI_MATCH not in df.columns:
        df[COL_MULTI_MATCH] = False
    return df


def apply_autopick_to_df(
    df: pd.DataFrame,
    reference_df: pd.DataFrame,
    templates: list[ConclusionTemplate],
    *,
    skip_approved: bool = True,
) -> pd.DataFrame:
    """
    Виконує автопідбір для всього датафрейму.
    Якщо skip_approved=True — рядки зі статусом 'approved' не перераховуються.
    """
    df = ensure_df_service_columns(df)

    for idx, row in df.iterrows():
        if skip_approved and df.at[idx, COL_STATUS] == STATUS_APPROVED:
            continue

        status, m_tmpl, m_word, multi, notes = autopick_for_row(row, reference_df, templates)

        df.at[idx, COL_STATUS] = status
        df.at[idx, COL_MATCHED_TEMPLATE] = m_tmpl
        df.at[idx, COL_MATCHED_WORD] = m_word
        df.at[idx, COL_MULTI_MATCH] = multi

        if status != STATUS_EMPTY:
            df.at[idx, COL_NOTES] = notes
        else:
            df.at[idx, COL_NOTES] = df.at[idx, COL_NOTES] or ""

        # logger.info('row=%s status=%s tmpl="%s" word="%s" multi=%s', idx, status, m_tmpl, m_word, multi)

    return df
