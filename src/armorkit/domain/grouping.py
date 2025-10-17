import pandas as pd
from typing import Dict, List

from src.armorkit.domain.freqnorm import freq4_str


def unique_freq_counts(df: pd.DataFrame) -> Dict[str, int]:
    """Підрахунок кількості перехоплень по кожній частоті (###.####)."""
    f4 = df["Частота"].map(freq4_str)
    return f4.value_counts().sort_index().to_dict()


def group_by_tag(freqs: List[str], ref_df: pd.DataFrame, order: List[str]) -> Dict[str, List[str]]:
    """
    Групує частоти за колонкою 'Хто' у довіднику. Ті, що не увійшли в список order — у 'Інші радіомережі'.
    """
    tag_map = {}
    for f in freqs:
        # шукаємо у довіднику запис по частоті (точній)
        row = ref_df.loc[ref_df["Частота"].map(freq4_str) == f]
        tag = row["Хто"].iloc[0] if not row.empty and "Хто" in row.columns else None
        tag_map[f] = tag

    groups: Dict[str, List[str]] = {k: [] for k in order}
    groups["Інші радіомережі"] = []
    for f in freqs:
        tag = tag_map.get(f)
        if tag in order:
            groups[tag].append(f)
        else:
            groups["Інші радіомережі"].append(f)
    # прибираємо порожні групи
    return {k: v for k, v in groups.items() if v}