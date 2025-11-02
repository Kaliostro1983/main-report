# src/automizer/conclusions.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json
from typing import List, Optional, Any

import pandas as pd

import logging

logger = logging.getLogger(__name__)

COL_PLACE = "Зона функціонування"
COL_UNIT = "Хто"
COL_FREQ = "Частота"


@dataclass
class ConclusionTemplate:
    name: str
    description: str
    keywords: list[str]
    shortcut: str  # у файлі поле назване "shorcut", обробимо нижче


def load_conclusions(path: Path | str) -> list[ConclusionTemplate]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"conclusions.json not found: {p}")

    raw = json.loads(p.read_text(encoding="utf-8"))
    items = raw.get("conclusions", [])
    result: list[ConclusionTemplate] = []
    for item in items:
        shortcut = item.get("shortcut") or item.get("shorcut") or ""
        tmpl = ConclusionTemplate(
            name=item.get("name", "").strip(),
            description=item.get("description", "").strip(),
            keywords=[k.strip() for k in item.get("keywords", [])],
            shortcut=str(shortcut).strip(),
        )
        result.append(tmpl)
    return result


def _norm_text(s: str) -> str:
    return (
        s.lower()
        .replace("ї", "і")
        .replace(",", " ")
        .replace(".", " ")
        .replace("!", " ")
        .replace("?", " ")
    )


def _find_place_unit_by_freq(freq_value: Any, reference_df: pd.DataFrame) -> tuple[str, str]:
    """
    Пошук по reference_df (аркуш "freq") за значенням колонки "Частота".
    Якщо не знайшли — повертаємо "невизначено".
    """
    if freq_value is None:
        return "невизначено", "невизначено"

    # df у тебе сирий з екселя: там може бути і float, і str
    target = str(freq_value).strip()
    if not target:
        return "невизначено", "невизначено"

    # приводимо обидві сторони до str
    tmp = reference_df
    if COL_FREQ in tmp.columns:
        mask = tmp[COL_FREQ].astype(str).str.strip() == target
        matches = tmp[mask]
        if len(matches) > 0:
            row = matches.iloc[0]
            place = str(row.get(COL_PLACE, "")).strip() or "невизначено"
            unit = str(row.get(COL_UNIT, "")).strip() or "невизначено"
            return place, unit

    return "невизначено", "невизначено"


def render_template(
    tmpl: ConclusionTemplate,
    freq_value: Any,
    reference_df: pd.DataFrame,
) -> str:
    place, unit = _find_place_unit_by_freq(freq_value, reference_df)
    text = tmpl.description.replace("{PLACE}", place).replace("{UNIT}", unit)
    return text


def try_autopick_conclusion(
    intercept_text: str,
    freq_value: Any,
    reference_df: pd.DataFrame,
    templates: list[ConclusionTemplate],
) -> Optional[str]:
    """
    Проходимося по шаблонах і при першому збігу keywords у тексті повертаємо сформований висновок.
    """
    if not intercept_text:
        return None

    norm_intercept = _norm_text(intercept_text)
    for tmpl in templates:
        for kw in tmpl.keywords:
            if not kw:
                continue
            if _norm_text(kw) in norm_intercept:
                # ЛОГУЄМО ЗНАЙДЕНИЙ ЗБІГ
                logger.info(
                    'в повідомлення "%s" знайдено слово "%s" висновку "%s"',
                    intercept_text,
                    kw,
                    tmpl.name,
                )
                return render_template(tmpl, freq_value, reference_df)
    return None
