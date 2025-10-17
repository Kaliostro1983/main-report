from collections import Counter
import logging
from typing import Optional
import pandas as pd
from typing import Dict, Any

log = logging.getLogger(__name__)  # => 'armorkit.domain.reference'

# -----------------------
# Довідник: назва мережі
# -----------------------
_NET_NAME_CANDIDATES = [
    "Назва радіомережі", "Назва мережі", "Радіомережа",
    "Назва", "Мережа", "Опис", "Призначення"
]


def get_network_name_by_freq(freq4: str, ref_df: pd.DataFrame) -> str:
    if "Частота" not in ref_df.columns:
        return "—"
    df = ref_df.copy()
    def _f4(x):
        try: return f"{float(str(x).replace(',', '.')):.4f}"
        except: return None
    df["__f4"] = df["Частота"].map(_f4)
    m = df[df["__f4"] == freq4]
    if m.empty: return "—"
    row = m.iloc[0]
    for c in _NET_NAME_CANDIDATES:
        if c in row.index and pd.notna(row[c]) and str(row[c]).strip():
            return str(row[c]).strip()
    return "—"


# -----------------------
# Довідник: повна «Хто» для групи
# -----------------------
def full_tag_for_group(freq_list, ref_df: pd.DataFrame, fallback_short: str) -> str:
    if not freq_list:
        return fallback_short
    tmp = ref_df.copy()
    def _f4(x):
        try: return f"{float(str(x).replace(',', '.')):.4f}"
        except: return None
    tmp["__f4"] = tmp["Частота"].map(_f4)
    vals = []
    if "Хто" not in tmp.columns:
        return fallback_short
    for f in freq_list:
        m = tmp[tmp["__f4"] == f]
        if not m.empty and pd.notna(m.iloc[0]["Хто"]):
            vals.append(str(m.iloc[0]["Хто"]).strip())
    if not vals:
        return fallback_short
    return Counter(vals).most_common(1)[0][0]


# -----------------------
# Еталонні вкладки: 'Призначення', 'Склад кореспондентів'
# -----------------------
def read_reference_sheet(freq4: str, ref_xlsx_path: str) -> dict:
    """
    Повертає {'Призначення': str|None, 'Склад кореспондентів': str|None}
    з аркуша, названого freq4 (наприклад '145.9500').
    Очікується таблиця з колонками 'Категорія'/'Значення' (регістро-незалежно).
    """
    result = {"Призначення": None, "Склад кореспондентів": None}
    try:
        xls = pd.ExcelFile(ref_xlsx_path, engine="openpyxl")
    except Exception as e:
        log.warning("Не вдалося відкрити довідник '%s': %s", ref_xlsx_path, e)
        return result

    if freq4 not in xls.sheet_names:
        log.info("Еталонний аркуш для %s не знайдено у файлі довідника.", freq4)
        return result

    try:
        df = pd.read_excel(xls, sheet_name=freq4)
    except Exception as e:
        log.warning("Не вдалося зчитати аркуш '%s': %s", freq4, e)
        return result

    if df.empty:
        log.info("Аркуш еталонки '%s' порожній.", freq4)
        return result

    # знайдемо колонки категорії/значення
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    cat_col = next((c for c, n in cols_lower.items() if "категор" in n), None)
    val_col = next((c for c, n in cols_lower.items() if "значен" in n), None)
    if not cat_col or not val_col:
        # fallback: інколи перша колонка - "№", тому беремо 2-гу та 3-тю
        try:
            cat_col, val_col = df.columns[1], df.columns[2]
        except Exception:
            log.info("На аркуші '%s' не знайшов пар колонок 'Категорія'/'Значення'.", freq4)
            return result

    def norm(s: str) -> str:
        return str(s).strip().lower()

    wanted = {"призначення": "Призначення",
              "склад кореспондентів": "Склад кореспондентів"}

    found = []
    for _, row in df.iterrows():
        key = norm(row.get(cat_col, ""))
        if key in wanted:
            val = row.get(val_col, None)
            if pd.notna(val) and str(val).strip():
                result[wanted[key]] = str(val).strip()
                found.append(wanted[key])

    if found:
        log.info("Еталонка %s: витягнуто поля %s.", freq4, ", ".join(found))
    else:
        log.info("Для %s еталонні поля не знайдені на аркуші.", freq4)

    return result
