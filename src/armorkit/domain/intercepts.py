# src/armorkit/domain/intercepts.py
from __future__ import annotations
from typing import Optional
import pandas as pd

from pathlib import Path
from src.armorkit.domain.reference import get_network_name_by_freq, read_reference_sheet

from src.armorkit.domain.schema import message_columns
from src.armorkit.domain.freqnorm import freq4_str


def _resolve_comment_col(df: pd.DataFrame, comment_col: Optional[str]) -> str:
    """Визначає колонку з коментарем."""
    if comment_col:
        return comment_col
    # спробуємо взяти зі схеми (повертає (msg_col, comment_col))
    try:
        _, ccol = message_columns(df)
        return ccol
    except Exception:
        pass
    # простий фолбек
    for name in ("Коментар", "коментар"):
        if name in df.columns:
            return name
    raise KeyError("Не знайдено колонку з коментарем")


def filter_with_comments(df: pd.DataFrame, comment_col: Optional[str] = None) -> pd.DataFrame:
    """
    Повертає лише рядки, де поле коментаря непорожнє.
    """
    ccol = _resolve_comment_col(df, comment_col)
    ser = df[ccol].astype(str).str.strip()
    out = df[ser.ne("") & df[ccol].notna()].copy()
    return out


def network_is_empty(df: pd.DataFrame, freq4: str,
                     freq_col: str = "Частота",
                     comment_col: Optional[str] = None) -> bool:
    """
    True, якщо для частоти немає жодного перехоплення з коментарем.
    """
    # нормалізуємо частоту у фреймі до ###.####
    tmp = df.copy()
    tmp["__f4"] = tmp[freq_col].map(freq4_str)
    part = tmp[tmp["__f4"] == freq4]
    with_comments = filter_with_comments(part, comment_col=comment_col)
    return with_comments.empty


def resolve_network_title(freq4: str, reference_df, ref_xlsx_path: str | Path) -> str:
    name = get_network_name_by_freq(freq4, reference_df)
    if name and str(name).strip() != "—":
        return name
    meta = read_reference_sheet(freq4, ref_xlsx_path)
    return meta.get("Назва") or meta.get("Призначення") or "—"