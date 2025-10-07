# src/reportgen/export_xlsx.py
from __future__ import annotations
from pathlib import Path
import pandas as pd
from src.reportgen.io_utils import safe_save_xlsx

def save_df_xlsx(df: pd.DataFrame, path: str) -> str:
    """
    Зберігає DataFrame у XLSX. Якщо файл відкритий у Excel — автоматично
    збереже під новою назвою '<name>__opened.xlsx' і поверне фактичний шлях.
    """
    def _write(p: Path):
        with pd.ExcelWriter(p, engine="openpyxl") as xw:
            df.to_excel(xw, index=False)

    saved = safe_save_xlsx(_write, path)
    return str(saved.resolve())
