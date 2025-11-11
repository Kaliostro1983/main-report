# src/automizer/service.py
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict
import logging
import pandas as pd

from src.armorkit.data_loader import load_inputs
from src.armorkit.normalize_freq import normalize_frequency_column
from src.automizer.conclusions import (
    load_conclusions,
    apply_autopick_to_df,
    ensure_df_service_columns,
)

logger = logging.getLogger(__name__)


def initialize_data(
    *,
    config_path: str | Path,
    conclusions_path: str | Path,
    skip_approved_on_reload: bool = True,
) -> Dict[str, Any]:
    """
    1) Завантажуємо дані через load_inputs.
    2) Нормалізуємо частоти в перехопленнях (маски -> реальні частоти).
    3) Гарантуємо службові колонки.
    4) Завантажуємо шаблони та виконуємо автопідбір.
    5) Повертаємо все необхідне для UI.
    """
    li = load_inputs(str(config_path))

    # 1. Отримуємо датафрейми з loader-а
    intercepts_df: pd.DataFrame = li.intercepts_df.copy()
    reference_df: pd.DataFrame = li.reference_df.copy()
    masks_df: pd.DataFrame = li.masks_df.copy()

    # 2. Нормалізуємо назви колонок (зайві пробіли, різні регістри не чіпаємо)
    intercepts_df.rename(columns=lambda c: str(c).strip(), inplace=True)
    reference_df.rename(columns=lambda c: str(c).strip(), inplace=True)
    masks_df.rename(columns=lambda c: str(c).strip(), inplace=True)

    # 3. Нормалізація частот (порядок аргументів як у робочому прикладі)
    normalize_frequency_column(intercepts_df, reference_df, masks_df)

    # 4. Службові колонки для подальшої роботи UI
    intercepts_df = ensure_df_service_columns(intercepts_df)

    # 5. Завантажуємо шаблони і виконуємо автопідбір
    templates = load_conclusions(conclusions_path)
    intercepts_df = apply_autopick_to_df(
        df=intercepts_df,
        reference_df=reference_df,
        templates=templates,
        skip_approved=skip_approved_on_reload,
    )

    # 6. Пакуємо результат
    return {
        "df": intercepts_df,
        "ref_df": reference_df,
        "templates": templates,
        "report_path": li.report_path,
        "cfg_path": li.cfg_path,
    }
