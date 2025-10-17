# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import logging
import pandas as pd

from src.armorkit.data_loader import load_inputs
from .report import build_docx
from src.armorkit.domain.freqnorm import freq4_str
from src.armorkit.xlsxutils.tables import load_sheet_df

# --------- логування ---------
log = logging.getLogger("eralonky")
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


# --------- хелпери ---------
def _next_free(path: Path) -> Path:
    if not path.exists():
        return path
    i = 2
    while True:
        cand = path.with_name(f"{path.stem} ({i}){path.suffix}")
        if not cand.exists():
            return cand
        i += 1


# def _freq4_str(x) -> str | None:
#     """Привести частоту до рядка з 4 знаками після крапки."""
#     try:
#         return f"{float(str(x).replace(',', '.')):.4f}"
#     except Exception:
#         return None





# --------- основний сценарій ---------
def run() -> Path:
    # 1) завантаження
    li = load_inputs("config.yml")  # очікуємо: li.reference_df, li.freq_path
    ref: pd.DataFrame = li.reference_df.copy()

    if not {"Статус", "Частота"}.issubset(ref.columns):
        raise KeyError("У довіднику бракує колонок 'Статус' та/або 'Частота'.")

    # 2) відібрати частоти зі статусом "Спостерігається"
    ref["freq4"] = ref["Частота"].map(freq4_str)
    mask = ref["Статус"].astype(str).str.strip().str.lower() == "спостерігається"
    freqs = (
        ref.loc[mask & ref["freq4"].notna(), "freq4"]
        .drop_duplicates()
        .sort_values(key=lambda s: pd.to_numeric(s.str.replace(",", "."), errors="coerce"))
        .tolist()
    )
    if not freqs:
        log.warning("Не знайдено жодної частоти зі статусом 'Спостерігається' у довіднику.")

    # 3) читання еталонок з вкладок
    xlsx_path = Path(li.freq_path)
    sections: list[dict] = []
    missing: list[str] = []

    for f in freqs:
        sheet_df = load_sheet_df(f, xlsx_path)
        if sheet_df is None or sheet_df.empty:
            log.warning(f"[ERALONKY] Не знайдено еталонку для частоти {f} (вкладка відсутня/порожня).")
            missing.append(f)
            continue

        # перетворення у "№. Категорія: Значення"
        lines: list[str] = []
        for _, row in sheet_df.iterrows():
            n = str(row.get("№", "")).strip()
            cat = str(row.get("Категорія", "")).strip()
            val = str(row.get("Значення", "")).strip()
            if not (cat or val):
                continue
            prefix = f"{n}. " if n else ""
            lines.append(f"{prefix}{cat}: {val}" if cat else f"{prefix}{val}")

        if not lines:
            log.warning(f"[ERALONKY] На вкладці {f} немає придатних рядків для еталонки.")
            missing.append(f)
            continue

        sections.append({"freq": f, "lines": lines})

    # 4) побудова документа
    out_dir = Path("build"); out_dir.mkdir(parents=True, exist_ok=True)
    today = datetime.now().strftime("%d.%m.%Y")
    out_path = _next_free(out_dir / f"Форма 1.5.3 (63 ОМБр {today}).docx")
    build_docx(sections, out_path)

    # 5) зведення по відсутніх еталонках
    if missing:
        print("\n[SUMMARY] Не знайдено еталонки для частот:")
        for x in missing:
            print(f"  - {x}")
    else:
        print("\n[SUMMARY] Еталонки знайдено для всіх відібраних частот.")

    print(f"[OK] Saved: {out_path}")
    return out_path


if __name__ == "__main__":
    run()
