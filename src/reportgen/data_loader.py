from __future__ import annotations
from pathlib import Path
from typing import List, Dict
import pandas as pd

def load_tables(paths: List[str]) -> Dict[str, pd.DataFrame]:
    """Завантажує CSV/Excel файли за шляхами, повертає dict ім'я->DataFrame."""
    result: Dict[str, pd.DataFrame] = {}
    for p in paths:
        path = Path(p)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        if path.suffix.lower() in {".csv"}:
            df = pd.read_csv(path)
        elif path.suffix.lower() in {".xlsx", ".xls"}:
            df = pd.read_excel(path)
        else:
            raise ValueError(f"Unsupported file type: {path.suffix}")
        # Додаємо ім'я джерела для аудиту
        df = df.copy()
        df["__source__"] = path.name
        result[path.stem] = df
    return result

def combine_tables(tables: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Об'єднує всі таблиці вертикально, вирівнюючи колонки."""
    if not tables:
        raise ValueError("No tables to combine")
    combined = pd.concat(tables.values(), axis=0, ignore_index=True, sort=False)
    return combined
