# src/automizer/freqdict.py
from __future__ import annotations

from typing import Any, Tuple
import pandas as pd

COL_FREQ = "Частота"
COL_PLACE = "Зона функціонування"
COL_UNIT = "Хто"


def lookup_place_unit(freq_value: Any, reference_df: pd.DataFrame) -> Tuple[str, str]:
    if freq_value is None:
        return "невизначено", "невизначено"

    target = str(freq_value).strip()
    if not target:
        return "невизначено", "невизначено"

    if COL_FREQ not in reference_df.columns:
        return "невизначено", "невизначено"

    mask = reference_df[COL_FREQ].astype(str).str.strip() == target
    matches = reference_df[mask]
    if len(matches) == 0:
        return "невизначено", "невизначено"

    row = matches.iloc[0]
    place = str(row.get(COL_PLACE, "")).strip() or "невизначено"
    unit = str(row.get(COL_UNIT, "")).strip() or "невизначено"
    return place, unit
