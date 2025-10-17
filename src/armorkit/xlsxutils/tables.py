from pathlib import Path
import pandas as pd

def load_sheet_df(freq4: str, xlsx_path: Path) -> pd.DataFrame | None:
    """
    Читає вкладку Excel з назвою == freq4 (4 знаки після крапки).
    Очікує колонки: '№', 'Категорія', 'Значення'.
    Повертає DataFrame або None (якщо вкладки немає/структура неочікувана).
    """
    try:
        df = pd.read_excel(xlsx_path, sheet_name=freq4, dtype=str)
    except Exception:
        return None

    df.columns = [str(c).strip() for c in df.columns]

    # підбираємо фактичні імена потрібних колонок (case-insensitive)
    def _find(colname: str) -> str | None:
        low = colname.lower()
        for c in df.columns:
            if str(c).strip().lower() == low:
                return c
        return None

    c_n = _find("№")
    c_cat = _find("Категорія")
    c_val = _find("Значення")
    if not all([c_n, c_cat, c_val]):
        return None

    df = df[[c_n, c_cat, c_val]].rename(columns={c_n: "№", c_cat: "Категорія", c_val: "Значення"})

    # приберемо пусті рядки
    df["Категорія"] = df["Категорія"].fillna("").map(str).str.strip()
    df["Значення"] = df["Значення"].fillna("").map(str).str.strip()
    df = df[~((df["Категорія"] == "") & (df["Значення"] == ""))]

    # акуратне сортування за № (якщо числовий)
    with pd.option_context("mode.chained_assignment", None):
        try:
            df["__n"] = pd.to_numeric(df["№"], errors="coerce")
            df = df.sort_values(["__n", "№"]).drop(columns=["__n"])
        except Exception:
            pass

    return df.reset_index(drop=True)