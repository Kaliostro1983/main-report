# src/xlsxutils/writer.py
from __future__ import annotations

from pathlib import Path
import pandas as pd


def _strip_service_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c in df.columns if not str(c).startswith("__")]
    return df[cols].copy()


def save_df_over_original(df: pd.DataFrame, original_path: str | Path) -> Path:
    """
    1. Прибираємо службові колонки (__approved і т.п.)
    2. Пишемо у тимчасовий файл
    3. Замінюємо оригінал
    """
    original_path = Path(original_path)
    to_save = _strip_service_columns(df)

    tmp_path = original_path.with_suffix(".new.xlsx")
    # пишемо без індексу, як у твоїх вхідних
    to_save.to_excel(tmp_path, index=False)

    # видаляємо старий і підміняємо
    if original_path.exists():
        original_path.unlink()

    tmp_path.rename(original_path)
    return original_path
