# src/reportgen/grouping.py
from __future__ import annotations
from collections import OrderedDict, Counter
from typing import Iterable, Dict, List, Tuple
import re
import pandas as pd
import logging

from src.reportgen.normalize_freq import FREQ_NOT_FOUND

log = logging.getLogger(__name__)

REF_FREQ_COL = "Частота"
REF_TAG_COL  = "Хто"

def _to_float(x):
    try:
        return float(str(x).replace(",", "."))
    except Exception:
        return None

def _freq4_str(x) -> str | None:
    v = _to_float(x)
    return f"{v:.4f}" if v is not None else None

def _numeric_sort_key(x: str) -> float:
    v = _to_float(x)
    return v if v is not None else float("inf")

def _normalize_tag(tag: str | None, cfg_grouping: dict | None) -> str | None:
    if tag is None or (isinstance(tag, float) and pd.isna(tag)):
        return None
    s = str(tag).strip()
    rules = (cfg_grouping or {}).get("tag_normalization", {})
    explicit = (rules.get("map") or {})
    if s in explicit:
        return explicit[s]
    for pat in (rules.get("patterns") or []):
        m = re.search(pat.get("match", ""), s)
        if m:
            return re.sub(pat.get("match", ""), pat.get("to", "\\g<0>"), s)
    return s

def unique_frequencies_with_counts(intercepts_df: pd.DataFrame) -> Tuple[List[str], Counter]:
    if "Частота" not in intercepts_df.columns:
        raise KeyError("У перехопленнях немає колонки 'Частота'")
    ser = intercepts_df["Частота"].map(_freq4_str).dropna()
    ser = ser[ser != FREQ_NOT_FOUND]  # прибираємо службовий маркер
    counts = Counter(ser)
    freqs = sorted(counts.keys(), key=_numeric_sort_key)
    return freqs, counts

def tag_for_frequency(freq: str, ref_df: pd.DataFrame, cfg_grouping: dict | None) -> str | None:
    if REF_FREQ_COL not in ref_df.columns:
        raise KeyError(f"У довіднику немає колонки '{REF_FREQ_COL}'")
    if REF_TAG_COL not in ref_df.columns:
        raise KeyError(f"У довіднику немає колонки '{REF_TAG_COL}'")

    ref = ref_df.copy()
    ref["_freq4"] = ref[REF_FREQ_COL].map(_freq4_str)

    m = ref[ref["_freq4"] == _freq4_str(freq)]
    if m.empty:
        return None
    if len(m) > 1:
        log.warning("WARN: Частота %s має кілька рядків у довіднику. Узято перший.", freq)

    tag_raw = m.iloc[0][REF_TAG_COL]
    return _normalize_tag(tag_raw, cfg_grouping)

def group_frequencies_by_tag(
    freqs: Iterable[str],
    ref_df: pd.DataFrame,
    allowed_tags: List[str],
    other_bucket: str,
    cfg_grouping: dict | None,
) -> "OrderedDict[str, List[str]]":
    buckets: Dict[str, List[str]] = {tag: [] for tag in allowed_tags}
    buckets[other_bucket] = []

    for f in freqs:
        tag = tag_for_frequency(f, ref_df, cfg_grouping)
        if tag in buckets:
            buckets[tag].append(f)
        else:
            buckets[other_bucket].append(f)

    for k in buckets:
        buckets[k] = sorted(set(buckets[k]), key=_numeric_sort_key)

    ordered = OrderedDict((tag, buckets.get(tag, [])) for tag in allowed_tags)
    ordered[other_bucket] = buckets.get(other_bucket, [])
    return ordered
