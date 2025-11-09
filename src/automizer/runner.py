# src/automizer/runner.py
from __future__ import annotations

from pathlib import Path

from src.armorkit.data_loader import load_inputs  # у тебе вже є
from src.armorkit.normalize_freq import normalize_frequency_column
from src.automizer.ui import AutomizerApp

import logging


def main(config_path: str = "config.yml") -> None:
    
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    
    # 1. зчитали все як у word_report
    li = load_inputs(config_path)

    # 2. нормалізували частоти (вже готова функція)
    normalize_frequency_column(li.intercepts_df, li.reference_df, li.masks_df)

    # 3. запустили інтерфейс
    app = AutomizerApp(
        intercepts_df=li.intercepts_df,
        reference_df=li.reference_df,
        report_path=li.report_path,
        config_path=li.cfg_path,
        conclusions_path=Path("src/automizer/conclusions.json"),
    )
    app.run()


if __name__ == "__main__":
    main()
