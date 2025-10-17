# src/reportgen/normalize_freq.py
from __future__ import annotations
from typing import Optional
import math
import logging
import pandas as pd

log = logging.getLogger(__name__)

MASK_PREFIXES = ("100", "200", "300")
FREQ_NOT_FOUND = "111.1111"
COL_TEXT = "р\\обмін"

def _to_float_safe(x) -> Optional[float]:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    s = str(x).strip().replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def _format_mask3(x) -> Optional[str]:
    v = _to_float_safe(x)
    return f"{v:.3f}" if v is not None else None

def is_real_freq(value) -> bool:
    if value is None:
        return False
    s = str(value).strip()
    if not s:
        return False
    return not s.startswith(MASK_PREFIXES)

def _match_in_columns_exact(ref_df: pd.DataFrame, col_names: list[str], target: str) -> pd.DataFrame:
    if not target:
        return ref_df.iloc[0:0]
    mask = False
    for c in col_names:
        if c in ref_df.columns:
            mask = mask | (ref_df[c].astype(str).str.strip() == target)
    return ref_df[mask]

def get_true_freq_by_mask(mask_like, ref_df: pd.DataFrame) -> str:
    mask3 = _format_mask3(mask_like)
    if mask3 is None:
        log.warning("WARN: Маска %r некоректна -> %s", mask_like, FREQ_NOT_FOUND)
        return FREQ_NOT_FOUND

    tmp = ref_df.copy()
    for col in ("Маска_3", "Маска_Ш"):
        if col in tmp.columns:
            tmp[col] = tmp[col].apply(_format_mask3)

    matches = _match_in_columns_exact(tmp, ["Маска_3", "Маска_Ш"], mask3)

    if len(matches) == 0:
        log.warning("WARN: Маска %s не знайдена у Маска_3/Маска_Ш -> %s", mask3, FREQ_NOT_FOUND)
        return FREQ_NOT_FOUND
    if len(matches) > 1:
        log.warning("WARN: Маска %s має декілька збігів (%d). Узято перший.", mask3, len(matches))

    row = matches.iloc[0]
    if "Частота" not in row.index or pd.isna(row["Частота"]):
        log.warning("WARN: У збігу для маски %s відсутня 'Частота' -> %s", mask3, FREQ_NOT_FOUND)
        return FREQ_NOT_FOUND

    return str(row["Частота"]).strip()

def _first_nonempty_line(text: str) -> str:
    if text is None or (isinstance(text, float) and math.isnan(text)):
        return ""
    for line in str(text).splitlines():
        st = line.strip()
        if st:
            return st
    return ""

def get_true_freq_by_text(text, ref_df: pd.DataFrame) -> str:
    line = _first_nonempty_line(text)
    if not line:
        log.warning("WARN: Порожній текст для пошуку маски за 'р\\обмін' -> %s", FREQ_NOT_FOUND)
        return FREQ_NOT_FOUND

    matches = _match_in_columns_exact(ref_df, ["Маска_А", "Маска_Акв"], line)
    if len(matches) == 0:
        log.warning("WARN: Маска за текстом '%s' не знайдена у Маска_А/Маска_Акв -> %s", line, FREQ_NOT_FOUND)
        return FREQ_NOT_FOUND
    if len(matches) > 1:
        log.warning("WARN: Текстова маска '%s' має декілька збігів (%d). Узято перший.", line, len(matches))

    row = matches.iloc[0]
    if "Частота" not in row.index or pd.isna(row["Частота"]):
        log.warning("WARN: Для текстової маски '%s' відсутня 'Частота' -> %s", line, FREQ_NOT_FOUND)
        return FREQ_NOT_FOUND

    return str(row["Частота"]).strip()

def normalize_frequency_column(intercepts_df: pd.DataFrame, ref_df: pd.DataFrame) -> pd.DataFrame:
    if "Частота" not in intercepts_df.columns:
        raise KeyError("У перехопленнях відсутня колонка 'Частота'")
    if COL_TEXT not in intercepts_df.columns:
        log.warning("WARN: Відсутня колонка '%s' — пошук за текстом буде обмежений.", COL_TEXT)

    # ВАЖЛИВО: дозволяємо писати '111.1111' як str
    intercepts_df["Частота"] = intercepts_df["Частота"].astype("object")

    for i in range(len(intercepts_df)):
        raw = intercepts_df.at[i, "Частота"] if "Частота" in intercepts_df.columns else None
        raw_str = None if (raw is None or (isinstance(raw, float) and math.isnan(raw))) else str(raw).strip()

        if raw_str:
            if is_real_freq(raw_str):
                continue
            else:
                true_f = get_true_freq_by_mask(raw_str, ref_df)
                intercepts_df.at[i, "Частота"] = true_f
        else:
            text_val = intercepts_df.at[i, COL_TEXT] if COL_TEXT in intercepts_df.columns else None
            true_f = get_true_freq_by_text(text_val, ref_df)
            intercepts_df.at[i, "Частота"] = true_f

    return intercepts_df
