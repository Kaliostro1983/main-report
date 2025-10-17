# src/reportgen/data_loader.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

import pandas as pd

from ..reportgen.settings import load_config, Config
from ..reportgen.io_utils import read_excel, find_latest



__all__ = [
    "LoadedInputs",
    "load_inputs",
    "load_reference",
    "load_latest_report_path",
    "load_report",
    "load_tables",
    "combine_tables",
]


# =========================
# Публічні структури
# =========================

@dataclass
class LoadedInputs:
    """Зібрані вхідні дані першого кроку конвеєра."""
    cfg_path: str
    freq_path: str
    report_path: str
    reference_df: pd.DataFrame
    intercepts_df: pd.DataFrame


def _read_excel_first_sheet(path: str) -> pd.DataFrame:
    # головний аркуш довідника/репорту (не еталонні вкладки)
    return pd.read_excel(path)

def load_inputs(config_path: str = "config.yml") -> LoadedInputs:
    cfg: Config = load_config(config_path)

    freq_path = cfg.paths.freq_file
    latest_report = load_latest_report_path(cfg.paths.reports_dir, cfg.paths.report_mask)

    reference_df = _read_excel_first_sheet(freq_path)
    intercepts_df = _read_excel_first_sheet(latest_report)

    return LoadedInputs(
        cfg_path=config_path,
        freq_path=freq_path,
        report_path=latest_report,
        reference_df=reference_df,
        intercepts_df=intercepts_df,
    )
    
    
# =========================
# Основні функції читання
# =========================

def load_reference(freq_path: str | Path) -> pd.DataFrame:
    """
    Зчитує довідник радіомереж (XLSX).
    Повертає raw DataFrame без нормалізації/мапінгу колонок.
    """
    path = Path(freq_path)
    df = read_excel(path)  # engine='openpyxl' всередині io_utils.read_excel
    if df.empty:
        raise ValueError(f"Reference (frequencies) file is empty: {path}")
    return df


def load_latest_report_path(reports_dir: str | Path, mask: str = "report_*.xlsx") -> Path:
    """
    Повертає шлях до найсвіжішого XLSX-звіту за маскою в указаній директорії.
    """
    return find_latest(str(reports_dir), mask)


def load_report(report_path: str | Path) -> pd.DataFrame:
    """
    Зчитує файл перехоплень (XLSX).
    Повертає raw DataFrame без нормалізації/мапінгу колонок.
    """
    path = Path(report_path)
    df = read_excel(path)
    if df.empty:
        raise ValueError(f"Intercepts report is empty: {path}")
    return df


def load_inputs(config_path: str = "config.yml") -> LoadedInputs:
    """
    Комплексне зчитування двох джерел:
      1) Довідник частот (cfg.paths.freq_file)
      2) Найсвіжіший звіт перехоплень у каталозі (cfg.paths.reports_dir + cfg.paths.report_mask)

    Повертає LoadedInputs з raw DataFrame'ами.
    """
    cfg = load_config(config_path)
    freq_path = Path(cfg.paths.freq_file)
    latest_report = load_latest_report_path(cfg.paths.reports_dir, cfg.paths.report_mask)

    reference_df = load_reference(freq_path)
    intercepts_df = load_report(latest_report)

    return LoadedInputs(
        cfg_path=str(Path(config_path).resolve()),
        freq_path=str(freq_path.resolve()),
        report_path=str(latest_report.resolve()),
        reference_df=reference_df,
        intercepts_df=intercepts_df,
    )


# =========================
# Додаткові універсальні утиліти
# (можуть знадобитись для пакетного читання)
# =========================

def load_tables(paths: List[str]) -> Dict[str, pd.DataFrame]:
    """
    Завантажує CSV/Excel файли за шляхами, повертає dict ім'я->DataFrame.
    Для CSV використовується pd.read_csv, для XLSX/XLS — pd.read_excel.
    """
    result: Dict[str, pd.DataFrame] = {}
    for p in paths:
        path = Path(p)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        if path.suffix.lower() in {".csv"}:
            df = pd.read_csv(path)
        elif path.suffix.lower() in {".xlsx", ".xls"}:
            df = pd.read_excel(path, engine="openpyxl" if path.suffix.lower() == ".xlsx" else None)
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
