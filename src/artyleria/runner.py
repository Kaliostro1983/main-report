# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import pandas as pd

# 1) вхідні дані беремо з armorkit (єдиний шар)
from src.armorkit.config import load_config
from src.armorkit.data_loader import load_inputs
# 2) нормалізація частоти — та сама, що у попередніх звітах
from src.armorkit.normalize_freq import normalize_frequency_column

from .report import build_docx


def _to4(x) -> str | None:
    try:
        return f"{float(str(x).replace(',', '.')):.4f}"
    except Exception:
        return None


def _image_for(freq4: str) -> Path | None:
    """
    Картинки шукаємо в локальній теці images (поруч із runner.py / report.py).
    Ім’я файлу = частота з 4 знаками після крапки (наприклад 136.5600.png).
    """
    folder = Path(__file__).parent / "images"
    try:
        filename = f"{float(str(freq4).replace(',', '.')):.4f}.png"
    except Exception:
        filename = f"{freq4}.png"
    path = folder / filename
    return path if path.exists() else None


def _next_free(path: Path) -> Path:
    if not path.exists():
        return path
    i = 2
    while True:
        cand = path.with_name(f"{path.stem} ({i}){path.suffix}")
        if not cand.exists():
            return cand
        i += 1


def run(config_path: str = "config.yml") -> Path:
    # ---- 1) Завантаження (довідник + свіжа вибірка перехоплень) ----
    cfg = load_config(config_path)
    li = load_inputs(config_path)
    
    loaded = load_inputs("config.yml")
    ref_df: pd.DataFrame = getattr(loaded, "reference_df")
    inter_df: pd.DataFrame = (
        getattr(loaded, "report_df", None)
        if getattr(loaded, "report_df", None) is not None
        else getattr(loaded, "intercepts_df")
    )

    # ---- 2) Нормалізація частот у перехопленнях ----
    inter_df = normalize_frequency_column(inter_df, ref_df, li.masks_df)

    # ---- 3) Перелік артмереж із довідника ----
    need_cols = {"Частота", "Теги"}
    miss = need_cols - set(ref_df.columns)
    if miss:
        raise KeyError(f"У довіднику відсутні колонки: {miss}")

    art_ref = ref_df.copy()
    art_ref["__is_arta"] = art_ref["Теги"].astype(str).str.contains("Арта", case=False, na=False)
    art_ref["freq4"] = art_ref["Частота"].map(_to4)
    art_ref = art_ref[art_ref["__is_arta"] & art_ref["freq4"].notna()]

    names = (
        art_ref.set_index("freq4")["Радіомережа"]
        if "Радіомережа" in art_ref.columns
        else pd.Series(dtype=str)
    )

    # ---- 4) Вибрати тільки перехоплення артмереж + сортування ----
    df = inter_df.copy()
    if "Частота" not in df.columns:
        raise KeyError("У перехопленнях немає колонки 'Частота' після нормалізації.")

    df["freq4"] = df["Частота"].map(_to4)
    df = df[df["freq4"].isin(art_ref["freq4"])]

    if not {"Дата", "Час"}.issubset(df.columns):
        raise KeyError("Очікуються колонки 'Дата' та 'Час' у перехопленнях.")
    df = df.sort_values(by=["Дата", "Час"], ascending=True)

    # ---- 5) Сформувати групи для рендера ----
    groups = []
    for freq4, g in df.groupby("freq4", sort=True):
        # Назва р/м
        name = str(names.get(freq4, "") or "").strip()
        if not name:
            print(f"[WARN] Відсутня назва радіомережі у довіднику для частоти {freq4}")
            name = "НВ підрозділу"

        # знайти фактичну колонку з текстом і уніфікувати як 'text'
        text_col = None
        for c in g.columns:
            if str(c).strip().lower().replace("\\", "/") == "р/обмін":
                text_col = c
                break

        base_cols = [c for c in ["Дата", "Час"] if c in g.columns]
        tmp = g[base_cols].copy()
        tmp["text"] = g[text_col].astype(str) if text_col else ""

        inters = tmp.to_dict("records")
        if not inters:
            continue  # блок створюємо лише якщо є перехоплення

        groups.append(
            {
                "freq": freq4,
                "name": name,
                "image": _image_for(freq4),
                "intercepts": inters,  # у кожному записі: 'Дата','Час','text'
            }
        )

    # ---- 6) Запис DOCX ----
    out_dir = Path("build")
    out_dir.mkdir(parents=True, exist_ok=True)
    today = datetime.now().strftime("%d.%m.%Y")
    out_path = _next_free(out_dir / f"Звіт з артилерії {today}.docx")

    build_docx(groups, out_path)
    print(f"[OK] Report saved to: {out_path}")
    return out_path


if __name__ == "__main__":
    run()
