# src/reportgen/enrich.py
import pandas as pd

def map_reference(ref_df: pd.DataFrame, cols_map: dict) -> pd.DataFrame:
    c = cols_map["reference"]
    ref = ref_df.rename(columns={v: k for k, v in c.items() if v in ref_df.columns}).copy()
    if "frequency" in ref.columns:
        ref["frequency"] = pd.to_numeric(ref["frequency"], errors="coerce")
        ref["frequency"] = ref["frequency"].map(lambda v: f"{v:.4f}" if pd.notna(v) else v)
    return ref

def join_by_frequency(intercepts: pd.DataFrame, ref: pd.DataFrame) -> pd.DataFrame:
    return intercepts.merge(ref, how="left", on=["frequency"], suffixes=("", "_ref"))


def callsigns_set(df: pd.DataFrame) -> list[str]:
    if "callsign_norm" not in df.columns:
        return []
    # легка дедуплікація схожих помилок (МОЛОКО vs МАЛАКО)
    uniq = {}
    for s in df["callsign_norm"].dropna().unique():
        key = s.replace("А", "А").replace("О", "О")  # місце для більш розумної нормалізації кир/лат
        uniq.setdefault(key, s)
    return sorted(uniq.values())

def activity_window(df: pd.DataFrame) -> str:
    if "datetime" not in df.columns or df["datetime"].dropna().empty:
        return "—"
    first = df["datetime"].min()
    last  = df["datetime"].max()
    if first.date() == last.date():
        return f"{first:%d.%m.%Y} {first:%H:%M:%S} – {last:%H:%M:%S}"
    return f"{first:%d.%m.%Y %H:%M:%S} – {last:%d.%m.%Y %H:%M:%S}"

def group_bucket(tag: str | None, allowed: list[str], other_name: str) -> str:
    if tag in (allowed or []):
        return tag
    return other_name
