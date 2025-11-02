# src/freqexport/generate_lists.py

from pathlib import Path
import pandas as pd

from src.armorkit.data_loader import load_config, load_inputs


def _pick_mask_or_freq(df: pd.DataFrame) -> pd.Series:
    """
    Для кожного рядка повертає значення, яке піде у файл:
    - якщо Маска_3 не порожня -> Маска_3
    - інакше -> Частота
    Результат у вигляді рядка (str).
    """
    mask_series = df.get("Маска_3", pd.Series([""] * len(df))).fillna("").astype(str).str.strip()
    freq_series = df.get("Частота", pd.Series([""] * len(df))).fillna("").astype(str).str.strip()

    out = mask_series.copy()
    out[out == ""] = freq_series[out == ""]
    return out


def _extract_list(reference_df: pd.DataFrame, flt: pd.Series) -> list:
    """
    Приймає reference_df та булеву маску flt (які рядки залишити).
    1. Вибирає ці рядки.
    2. Формує колонку "маска/частота" за правилом _pick_mask_or_freq.
    3. Видаляє дублі.
    4. Повертає список рядків.
    """
    sub = reference_df[flt].copy()
    if sub.empty:
        return []

    values = _pick_mask_or_freq(sub)
    unique_values = values.drop_duplicates(keep="first").tolist()
    return unique_values


def _write_txt(lines: list, path: Path):
    """
    Створює директорію (якщо треба) і пише кожне значення в окремий рядок.
    Навіть якщо список порожній — створюємо (порожній) файл, щоб пайплайн не ламався.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        for line in lines:
            f.write(str(line).strip() + "\n")


def build_freq_lists(config_path: str) -> dict:
    """
    Основний раннер.
    1. Завантажує довідник частот (reference_df) через load_inputs().
    2. Будує три множини: green / yellow / blue.
    3. Записує їх у build/green.txt, build/yellow.txt, build/blue.txt.
    4. Повертає словник з шляхами.
    """

    cfg = load_config(config_path)
    li = load_inputs(config_path)

    ref_df = li.reference_df.copy()

    # Нормалізуємо критичні поля (strip) — але не lower(), бо ти дав точні значення
    for col in ["Слухає", "Дешифрування", "Статус", "Маска_3", "Частота"]:
        if col in ref_df.columns:
            ref_df[col] = ref_df[col].fillna("").astype(str).str.strip()
        else:
            # якщо колонки немає — додаємо порожню, щоб уникнути KeyError нижче
            ref_df[col] = ""

    # GREEN:
    green_mask = (
        (ref_df["Слухає"] == "Шаєн_63") &
        (ref_df["Дешифрування"] == "так") &
        (ref_df["Статус"] == "Спостерігається")
    )
    green_list = _extract_list(ref_df, green_mask)

    # YELLOW:
    yellow_mask = (
        (ref_df["Слухає"] != "Шаєн_63") &
        (ref_df["Дешифрування"] == "так") &
        (ref_df["Статус"] == "Спостерігається")
    )
    yellow_list = _extract_list(ref_df, yellow_mask)

    # BLUE:
    # Нове правило: Дешифрування == "ні"
    blue_mask = (ref_df["Дешифрування"] == "ні")
    blue_list = _extract_list(ref_df, blue_mask)

    # Куди писати
    out_dir = Path(getattr(cfg.paths, "output_dir", "build"))
    green_path = out_dir / "green.txt"
    yellow_path = out_dir / "yellow.txt"
    blue_path = out_dir / "blue.txt"

    _write_txt(green_list, green_path)
    _write_txt(yellow_list, yellow_path)
    _write_txt(blue_list, blue_path)

    return {
        "green": green_path,
        "yellow": yellow_path,
        "blue": blue_path,
    }


if __name__ == "__main__":
    # Простий запуск без main.py: python src/freqexport/generate_lists.py config.yml
    # Але краще запускати через main.py --mode freq-lists --config config.yml
    import sys
    if len(sys.argv) < 2:
        print("Використання: python generate_lists.py <config.yml>")
        sys.exit(1)
    cfg_path = sys.argv[1]
    res = build_freq_lists(cfg_path)
    print("Записані файли:")
    for k, v in res.items():
        print(f" {k}: {v}")
