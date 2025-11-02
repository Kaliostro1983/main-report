# src/reportgen/normalize_freq.py
from __future__ import annotations

import math
from typing import Optional
import re
import logging

import pandas as pd


FREQ_NOT_FOUND = "111.1111"
COL_TEXT = r"р\обмін"
TEXT_MASK_COLUMNS = ["Маска_А", "Маска_Акв"]
NUM_MASK_COLUMNS = ["Маска_3", "Маска_Ш"]
MASK_PREFIXES = ("100", "200", "300")


logger = logging.getLogger(__name__)


# те саме ім'я, що в твоєму xlsx
COL_TEXT = r"р\обмін"

# дата/час на початку перехоплення — те, що ми маємо ігнорувати
DATE_LINE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4},\s*\d{2}:\d{2}:\d{2}$")


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
    """
    Частота, а не маска: не починається з 100/200/300 і не порожня.
    """
    if value is None:
        return False
    s = str(value).strip()
    if not s:
        return False
    return not s.startswith(MASK_PREFIXES)


def get_true_freq_by_mask(mask_like, ref_df: pd.DataFrame) -> str:
    """
    Пошук по числовим/ш-колонкам (Маска_3, Маска_Ш).
    Використовується, коли у нас у полі 'Частота' насправді 300.3250 і т.п.
    """
    # print(f"DEBUG: Шукаємо частоту по масці {mask_like!r}")
    mask3 = _format_mask3(mask_like)
    if mask3 is None:
        print(f"WARNING: Маска {mask_like!r} некоректна -> {FREQ_NOT_FOUND}")
        return FREQ_NOT_FOUND

    tmp = ref_df.copy()
    # нормалізуємо числові колонки до 3 знаків
    for col in ("Маска_3", "Маска_Ш"):
        if col in tmp.columns:
            tmp[col] = tmp[col].apply(_format_mask3)

    matches = tmp[
        (tmp["Маска_3"].astype(str).str.strip() == mask3)
        | (tmp["Маска_Ш"].astype(str).str.strip() == mask3)
    ]

    if len(matches) == 0:
        print(f"WARNING: Маска {mask3} не знайдена у Маска_3/Маска_Ш -> {FREQ_NOT_FOUND}")
        return FREQ_NOT_FOUND
    if len(matches) > 1:
        print(f"WARNING: Маска {mask3} має декілька збігів ({len(matches)}). Узято перший.")

    row = matches.iloc[0]
    freq_val = row.get("Частота")
    if freq_val is None or (isinstance(freq_val, float) and math.isnan(freq_val)):
        print(f"WARNING: У збігу для маски {mask3} відсутня 'Частота' -> {FREQ_NOT_FOUND}")
        return FREQ_NOT_FOUND

    # print(f"DEBUG: Знайдено частоту {freq_val} для маски {mask3}")
    return str(freq_val).strip()


def get_true_freq_by_text(text, ref_df: pd.DataFrame) -> str:
    """
    Шукаємо по колонках 'Маска_А' і 'Маска_Акв'
    """
    if text is None or (isinstance(text, float) and math.isnan(text)):
        print(f"WARNING: Порожній текст для пошуку маски -> {FREQ_NOT_FOUND}")
        return FREQ_NOT_FOUND

    # беремо перший непорожній рядок
    line = ""
    for ln in str(text).splitlines():
        st = ln.strip()
        if st:
            line = st
            break

    if not line:
        print(f"WARNING: Порожня перша строка у текстовій масці -> {FREQ_NOT_FOUND}")
        return FREQ_NOT_FOUND

    # шукаємо по обох колонках
    mask = False
    for col in TEXT_MASK_COLUMNS:
        if col in ref_df.columns:
            mask = mask | (ref_df[col].astype(str).str.strip() == line)

    matches = ref_df[mask]
    if len(matches) == 0:
        print(f"WARNING: Маска за текстом '{line}' не знайдена -> {FREQ_NOT_FOUND}")
        return FREQ_NOT_FOUND
    if len(matches) > 1:
        print(f"WARNING: Текстова маска '{line}' має декілька збігів ({len(matches)}). Узято перший.")

    row = matches.iloc[0]
    freq_val = row.get("Частота")
    if freq_val is None or (isinstance(freq_val, float) and math.isnan(freq_val)):
        print(f"WARNING: Для текстової маски '{line}' відсутня 'Частота' -> {FREQ_NOT_FOUND}")
        return FREQ_NOT_FOUND
    
    print(f"DEBUG: Знайдено частоту {freq_val} для текстової маски '{line}'")

    return str(freq_val).strip()


def normalize_frequency_column(intercepts_df: pd.DataFrame, ref_df: pd.DataFrame, masks_df: pd.DataFrame = None) -> pd.DataFrame:
    """
    Те, що в тебе крутилося у word_report.py:
    - якщо в 'Частота' вже нормальна частота — лишаємо;
    - якщо там маска 100/200/300 — шукаємо по Маска_3/Маска_Ш;
    - інакше — пробуємо витягнути з тексту 'р\обмін'.
    """
    if "Частота" not in intercepts_df.columns:
        raise KeyError("У перехопленнях відсутня колонка 'Частота'")
    
    intercepts_df["Частота"] = intercepts_df["Частота"].astype("object")

    total = len(intercepts_df)
    for i in range(total):
        raw_freq = intercepts_df.at[i, "Частота"]

        # 1. якщо це адекватна частота — нічого не робимо
        if is_real_freq(raw_freq):
            continue

        # 2. якщо це схоже на маску 100/200/300 — шукаємо по масках
        if raw_freq is not None and str(raw_freq).strip().startswith(MASK_PREFIXES):
            true_f = get_true_freq_by_mask(raw_freq, ref_df)
            true_f = str(true_f)
            intercepts_df.at[i, "Частота"] = true_f
            continue

        # 3. інакше — шукаємо по тексту перехоплення
        text_val = intercepts_df.at[i, COL_TEXT] if COL_TEXT in intercepts_df.columns else None
        if masks_df is not None:
            # якщо є маски, то шукаємо спочатку в них
            print(f"DEBUG: Шукаємо частоту по тексту у masks_df для рядка {i}")
            true_f = get_true_freq_by_text(text_val, masks_df)
        else:
            true_f = get_true_freq_by_text(text_val, ref_df)
        
        true_f = str(true_f)
        intercepts_df.at[i, "Частота"] = true_f

    return intercepts_df
