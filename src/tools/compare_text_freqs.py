# src/tools/compare_text_freqs.py
from __future__ import annotations

import logging
import traceback

from src.armorkit.data_loader import load_inputs
from src.armorkit.normalize_freq import get_true_freq_by_text2

COL_TEXT = r"р\обмін"


def is_empty_freq(val) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and val != val:  # NaN
        return True
    return str(val).strip() == ""


def main(config_path: str = "config.yml", limit: int = 200) -> None:
    print("[check-new] start")

    try:
        li = load_inputs(config_path)
    except Exception as e:
        print("[check-new] ERROR load_inputs:", e)
        traceback.print_exc()
        return

    df = li.intercepts_df
    masks_df = li.masks_df

    print(f"[check-new] rows in intercepts: {len(df)}")
    print(f"[check-new] rows in masks_df: {'NONE' if masks_df is None else len(masks_df)}")

    tested = 0
    not_found = 0
    found = 0

    for idx, row in df.iterrows():
        if tested >= limit:
            break

        text_val = row.get(COL_TEXT, "")
        freq_val = row.get("Частота", None)

        # ми тестуємо САМЕ ті рядки, де частота порожня,
        # бо функція так і задумана
        if not is_empty_freq(freq_val):
            continue

        tested += 1

        print("---------------")
        print(f"ROW {idx}")
        print("TEXT FIRST LINES:")
        # покажемо перші 3 рядки тексту, щоб ти бачив, що вона реально бере
        for j, ln in enumerate(str(text_val).splitlines()):
            if j >= 3:
                break
            print(f"  {j+1}: {ln!r}")

        resolved = get_true_freq_by_text2(text_val, masks_df, current_freq=freq_val)
        print(f"RESOLVED: {resolved}")

        if resolved == "111.1111":
            not_found += 1
        else:
            found += 1

    print("========== SUMMARY ==========")
    print(f"tested rows     : {tested}")
    print(f"resolved (found): {found}")
    print(f"not found       : {not_found}")
    print("[check-new] done")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    main()
