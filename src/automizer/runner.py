# src/automizer/runner.py
from __future__ import annotations

from pathlib import Path
import logging

from src.automizer.service import initialize_data
from src.automizer.ui import AutomizerApp


def main(config_path: str = "config.yml") -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s:%(name)s:%(message)s")

    conclusions_path = Path("src/automizer/conclusions.json")

    # Єдина точка входу ініціалізації
    init = initialize_data(
        config_path=config_path,
        conclusions_path=conclusions_path,
        skip_approved_on_reload=False,  # перший старт: порахувати ВСІ
    )

    app = AutomizerApp(
        intercepts_df=init["df"],
        reference_df=init["ref_df"],
        report_path=init["report_path"],
        config_path=init["cfg_path"],
        conclusions_path=conclusions_path,
        templates=init["templates"],
    )
    app.run()


if __name__ == "__main__":
    main()
