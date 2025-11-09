# src/stillalive/generate_stillalive_report.py

from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

from src.armorkit.data_loader import load_config, load_inputs
from src.armorkit.normalize_freq import normalize_frequency_column
from src.armorkit.dates import parse_period_from_filename, format_for_filename


def _get_observed_frequencies(reference_df: pd.DataFrame) -> pd.DataFrame:
    """
    Вибирає частоти, які перебувають на спостереженні (Статус == 'Спостерігається').
    Повертає датафрейм з колонками 'Частота', 'Маска_3'.
    """
    observed = reference_df[reference_df["Статус"] == "Спостерігається"].copy()
    observed = observed[["Частота", "Маска_3"]].drop_duplicates(subset=["Частота"], keep="first")
    observed["Частота"] = observed["Частота"].astype(str)
    return observed


def _prepare_daily_counts(intercepts_df: pd.DataFrame,
                          observed_freqs_df: pd.DataFrame) -> pd.DataFrame:
    """
    Формує підсумковий датафрейм у wide-форматі:
    - "Частота"
    - "Маска"
    - далі по одній колонці на кожен день періоду (формат "dd.mm")
    Значення = кількість перехоплень цієї частоти у цю дату.
    ВАЖЛИВО: включаємо всі частоти зі статусом "Спостерігається",
    навіть якщо перехоплень по них 0.
    """

    work = intercepts_df.copy()
    work["Частота"] = work["Частота"].astype(str)

    if "Дата" not in work.columns:
        raise ValueError("У перехопленнях немає колонки 'Дата', не можу побудувати добову активність.")

    # Нормалізація дати
    work["_dt"] = pd.to_datetime(work["Дата"], errors="coerce")
    work = work.dropna(subset=["_dt"])
    work["_day_label"] = work["_dt"].dt.strftime("%d.%m")

    # Лічильник перехоплень по (Частота, день)
    counts = (
        work.groupby(["Частота", "_day_label"])
        .size()
        .reset_index(name="count")
    )

    # Всі унікальні дні періоду
    all_days = (
        counts["_day_label"]
        .drop_duplicates()
        .sort_values(key=lambda s: s.apply(lambda x: (int(x.split(".")[1]), int(x.split(".")[0]))))
        .tolist()
    )

    # Повний список частот зі статусом "Спостерігається"
    freqs_df = observed_freqs_df.copy()
    freqs_df["Частота"] = freqs_df["Частота"].astype(str)
    freqs_df = freqs_df.drop_duplicates(subset=["Частота"], keep="first")

    # Повна комбінація (частота × день)
    freq_list = freqs_df["Частота"].tolist()
    cartesian = (
        pd.MultiIndex.from_product(
            [freq_list, all_days],
            names=["Частота", "_day_label"]
        ).to_frame(index=False)
    )

    cartesian = cartesian.merge(counts, on=["Частота", "_day_label"], how="left")
    cartesian["count"] = cartesian["count"].fillna(0).astype(int)

    # Wide-формат
    pivot = cartesian.pivot_table(
        index="Частота",
        columns="_day_label",
        values="count",
        fill_value=0,
        aggfunc="sum"
    ).reset_index()

    # Додаємо маску
    mask_map = freqs_df.set_index("Частота")["Маска_3"].to_dict()
    pivot.insert(1, "Маска", pivot["Частота"].map(mask_map).fillna(""))

    final_cols = ["Частота", "Маска"] + all_days
    pivot = pivot[final_cols]

    return pivot


def _export_to_xlsx(df: pd.DataFrame, period_start: str, period_end: str, output_dir: Path) -> Path:
    """
    Записує df у Excel з форматуванням:
    - Заголовки колонок жирним.
    - Колонки "Частота" і "Маска" — жирним у всіх рядках.
    - Комірки дат: 0 -> червоний фон, >0 -> зелений фон.
    """

    start_s = format_for_filename(period_start)
    end_s = format_for_filename(period_end)
    filename = f"Активність радіомереж ({start_s} - {end_s}).xlsx"
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / filename

    wb = Workbook()
    ws = wb.active
    ws.title = "still_alive"

    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    n_rows = ws.max_row
    n_cols = ws.max_column

    bold_font = Font(bold=True)
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    # Заголовки
    for c in range(1, n_cols + 1):
        ws.cell(row=1, column=c).font = bold_font

    # Колонки "Частота" і "Маска"
    for r in range(1, n_rows + 1):
        ws.cell(row=r, column=1).font = bold_font
        ws.cell(row=r, column=2).font = bold_font

    # Підсвітка значень
    for r in range(2, n_rows + 1):
        for c in range(3, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            try:
                num = int(cell.value)
                if num == 0:
                    cell.fill = red_fill
                else:
                    cell.fill = green_fill
            except Exception:
                continue

    wb.save(out_path)
    return out_path


def build_stillalive_report(config_path: str) -> Path:
    """
    Основний пайплайн:
    1. Завантажує конфіг і дані.
    2. Визначає період звіту з назви файлу.
    3. Нормалізує частоти у перехопленнях.
    4. Вибирає частоти зі статусом "Спостерігається".
    5. Формує підсумковий датафрейм по днях.
    6. Зберігає Excel у build/.
    """

    cfg = load_config(config_path)
    li = load_inputs(config_path)

    period_start, period_end = parse_period_from_filename(li.report_path)
    normalize_frequency_column(li.intercepts_df, li.reference_df, li.masks_df)

    observed_freqs_df = _get_observed_frequencies(li.reference_df)
    summary_df = _prepare_daily_counts(li.intercepts_df, observed_freqs_df)

    out_dir = Path(getattr(cfg.paths, "output_dir", "build"))
    out_path = _export_to_xlsx(summary_df, period_start, period_end, out_dir)

    return out_path
