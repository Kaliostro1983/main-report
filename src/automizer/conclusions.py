# src/automizer/conclusions.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Tuple

import json
import logging
import pandas as pd

from src.automizer.freqdict import lookup_place_unit, COL_FREQ

logger = logging.getLogger(__name__)

# --------- Статуси висновків ---------
STATUS_EMPTY: str = "empty"
STATUS_NEED_APPROVE: str = "need_approve"
STATUS_APPROVED: str = "approved"

# --------- Колонки службові ---------
COL_STATUS = "__status"
COL_NOTES = "примітки"
COL_MATCHED_TEMPLATE = "__matched_template"
COL_MATCHED_WORD = "__matched_word"
COL_MULTI_MATCH = "__multi_match"

TEXT_COL = "р\\обмін"


@dataclass
class KeywordRule:
    word: str
    probability: int  # 0 або 1


@dataclass
class ConclusionTemplate:
    name: str
    description: str
    keywords: list[KeywordRule]
    shortcut: str = ""


def _norm_text(s: Any) -> str:
    return str(s).strip().lower()


def _render_with_place_unit(description: str, freq_value: Any, reference_df: pd.DataFrame) -> str:
    place, unit = lookup_place_unit(freq_value, reference_df)
    return description.replace("{PLACE}", place).replace("{UNIT}", unit)


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
            else:
                # зворотна сумісність для старих JSON
                word = str(k).strip()
                prob = 0
            if not word:
                continue
            prob = 1 if prob == 1 else 0
            kw_list.append(KeywordRule(word=word, probability=prob))

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

def _prefix_regex(word: str) -> re.Pattern:
    # Префіксний збіг на початку слова; для дуже коротких префіксів вимагаємо повний збіг
    w_raw = (word or "").strip().lower().replace("ё", "е")
    w = re.escape(w_raw)
    if len(w_raw) < 3:
        pattern = rf"\b{w}\b"
    else:
        pattern = rf"\b{w}[\w\-]*"
    return re.compile(pattern, flags=re.IGNORECASE | re.UNICODE)

def find_template_candidates(text: str, templates: list[ConclusionTemplate]):
    """
    Повертає список [(tmpl, rule), ...] у тому ж порядку, як у JSON.
    На відміну від старої версії, НЕ зупиняється на першому збігу в межах шаблону:
    якщо у шаблона спрацювало кілька слів-тригерів, додаємо їх усі.
    """
    text_norm = (text or "")
    matches = []
    for tmpl in templates:
        for rule in tmpl.keywords:  # rule має поля .word, .probability
            pat = _prefix_regex(rule.word)
            if pat.search(text_norm):
                matches.append((tmpl, rule))
                # ВАЖЛИВО: без break — збираємо всі спрацьовані слова одного шаблону
    return matches




# NEW helper: згрупувати збіги за шаблоном
# ↓ ДОДАЙ це поруч із іншими хелперами
def _group_by_template(candidates: list[tuple[ConclusionTemplate, KeywordRule]]):
    """
    Групує збіги за НАЗВОЮ шаблону.
    Повертає:
      groups: dict[name -> list[KeywordRule]]
      order:  list[ConclusionTemplate] у порядку першої появи в candidates
    """
    groups: dict[str, list[KeywordRule]] = {}
    order: list[ConclusionTemplate] = []
    seen: set[str] = set()

    for tmpl, rule in candidates:
        name = tmpl.name
        if name not in groups:
            groups[name] = []
        groups[name].append(rule)
        if name not in seen:
            order.append(tmpl)
            seen.add(name)
    return groups, order



def decide_status_by_candidates(candidates: list[tuple[ConclusionTemplate, KeywordRule]]) -> str:
    """
    Логіка статусу:
      - 0 збігів -> empty
      - рівно 1 тип висновку:
          * якщо збіглося ≥2 слів цього типу -> approved
          * інакше 1 слово -> за p цього слова (1 -> approved, 0 -> need_approve)
      - >1 типів висновків -> need_approve
    """
    if len(candidates) == 0:
        return STATUS_EMPTY

    groups, _ = _group_by_template(candidates)
    if len(groups) == 1:
        only_rules = next(iter(groups.values()))
        if len(only_rules) >= 2:
            return STATUS_APPROVED
        rule = only_rules[0]
        return STATUS_APPROVED if getattr(rule, "probability", 0) == 1 else STATUS_NEED_APPROVE

    return STATUS_NEED_APPROVE




def autopick_for_row(
    row: pd.Series,
    reference_df: pd.DataFrame,
    templates: list[ConclusionTemplate],
) -> tuple[str, str, str, bool, str]:
    text = str(row.get(TEXT_COL, "") or "")
    freq_val = row.get(COL_FREQ)

    candidates = find_template_candidates(text, templates)
    status = decide_status_by_candidates(candidates)

    matched_template = ""
    matched_word = ""
    notes = ""

    groups, order = _group_by_template(candidates)
    multi_templates = len(groups) > 1

    if candidates:
        first_tmpl = order[0]  # перший за порядком появи
        rules_for_first = groups.get(first_tmpl.name, [])

        # взяти слово з p=1, якщо є; інакше перше
        rule_p1 = next((r for r in rules_for_first if getattr(r, "probability", 0) == 1), None)
        chosen_rule = rule_p1 or (rules_for_first[0] if rules_for_first else None)

        matched_template = first_tmpl.name
        matched_word = chosen_rule.word if chosen_rule else ""
        notes = _render_with_place_unit(first_tmpl.description, freq_val, reference_df)

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
