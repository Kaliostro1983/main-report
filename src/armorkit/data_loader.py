# src/reportgen/data_loader.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List
from typing import Optional

import pandas as pd

from .settings import load_config, Config
from src.armorkit.io_utils import read_excel, find_latest


__all__ = [
    "LoadedInputs",
    "load_inputs",
    "load_reference",
    "load_latest_report_path",
    "load_report",
    "load_tables",
    "combine_tables",
]


@dataclass
class LoadedInputs:
    cfg_path: str
    freq_path: str
    report_path: str
    reference_df: pd.DataFrame
    intercepts_df: pd.DataFrame
    masks_df: Optional[pd.DataFrame] = None


def load_reference(freq_path: str | Path) -> pd.DataFrame:
    """
    Зчитує довідник частот (XLSX) – головний аркуш.
    ВАЖЛИВО: тут використовується read_excel з armorkit, без sheet_name.
    """
    path = Path(freq_path)
    df = read_excel(path)
    if df.empty:
        raise ValueError(f"Reference (frequencies) file is empty: {path}")
    return df


def load_masks(freq_path: str | Path) -> Optional[pd.DataFrame]:
    path = Path(freq_path)
    try:
        df = pd.read_excel(
            path,
            sheet_name="masks",
            engine="openpyxl",
            dtype={"Частота": "object", "Маска_А": "object"},  # ← важливо
        )
    except Exception as ex:
        print(f"DEBUG: load_masks FAILED: {ex}")
        return None

    if df is None or df.empty:
        print("DEBUG: load_masks -> empty dataframe")
        return None

    # Перший великий трейс
    # print(f"DEBUG: load_masks -> shape={df.shape}")
    # print(f"DEBUG: load_masks -> columns={list(df.columns)}")

    # ПОКАЗАТИ всі рядки, які він реально прочитав
    # for idx, row in df.iterrows():
    #     print(f"DEBUG: masks[{idx}] = Ч:{repr(row.get('Частота'))} | М:{repr(row.get('Маска_А'))}")

    # маленький захист — тільки якщо реально є колонка "Маска"
    if "Маска_А" not in df.columns or "Частота" not in df.columns:
        print("DEBUG: load_masks -> required columns not found")
        return None

    # print(f"DEBUG: load_masks -> loaded {len(df)} rows")
    return df


def load_latest_report_path(reports_dir: str | Path, mask: str = "report_*.xlsx") -> Path:
    """
    Повертає шлях до найсвіжішого звіту в каталозі.
    """
    return find_latest(str(reports_dir), mask)


def load_report(report_path: str | Path) -> pd.DataFrame:
    """
    Зчитує свіжий XLSX зі перехопленнями.
    """
    path = Path(report_path)
    df = read_excel(path)
    if df.empty:
        raise ValueError(f"Intercepts report is empty: {path}")
    return df


def load_inputs(config_path: str = "config.yml") -> LoadedInputs:
    """
    Головна точка входу: читаємо конфіг, довідник і свіжий звіт.
    (без masks, без додаткових аркушів)
    """
    cfg: Config = load_config(config_path)

    freq_path = cfg.paths.freq_file
    latest_report = load_latest_report_path(cfg.paths.reports_dir, cfg.paths.report_mask)

    reference_df = load_reference(freq_path)
    intercepts_df = load_report(latest_report)
    masks_df = load_masks(freq_path)
    print(f"DEBUG: Loaded masks_df with {len(masks_df) if masks_df is not None else 0} rows")

    return LoadedInputs(
        cfg_path=config_path,
        freq_path=freq_path,
        report_path=str(latest_report),
        reference_df=reference_df,
        intercepts_df=intercepts_df,
        masks_df=masks_df
    )


def load_tables(paths: List[str]) -> Dict[str, pd.DataFrame]:
    """
    Утиліта: завантажує кілька таблиць і повертає їх як dict.
    """
    result: Dict[str, pd.DataFrame] = {}
    for p in paths:
        path = Path(p)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        if path.suffix.lower() in {".csv"}:
            df = pd.read_csv(path)
        elif path.suffix.lower() in {".xlsx", ".xls"}:
            df = pd.read_excel(
                path,
                engine="openpyxl" if path.suffix.lower() == ".xlsx" else None,
            )
        else:
            raise ValueError(f"Unsupported file type: {path.suffix}")

        df = df.copy()
        df["__source__"] = path.name
        result[path.stem] = df
    return result


def combine_tables(tables: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Об’єднує всі таблиці вертикально.
    """
    if not tables:
        raise ValueError("No tables to combine")
    combined = pd.concat(tables.values(), axis=0, ignore_index=True, sort=False)
    return combined
