# src/reportgen/normalize.py
import pandas as pd
import re

def _callsign_norm(s: str) -> str:
    if pd.isna(s): return s
    s = str(s).strip().upper()
    s = re.sub(r"\s+", "-", s)     # пробіли → дефіси
    s = re.sub(r"-{2,}", "-", s)
    return s

def normalize_intercepts(df: pd.DataFrame, cols_map: dict) -> pd.DataFrame:
    c = cols_map["intercepts"]
    df = df.rename(columns={v: k for k, v in c.items() if v in df.columns}).copy()

    # datetime з окремих date+time
    if "date" in df.columns and "time" in df.columns:
        df["datetime"] = pd.to_datetime(
            df["date"].astype(str).str.strip() + " " + df["time"].astype(str).str.strip(),
            errors="coerce", dayfirst=True
        )

    # частота у форматі ###.#### (4 знаки після крапки)
    if "frequency" in df.columns:
        df["frequency"] = pd.to_numeric(df["frequency"], errors="coerce")
        df["frequency"] = df["frequency"].map(lambda v: f"{v:.4f}" if pd.notna(v) else v)

    # позивні
    if "from_callsign" in df.columns:
        df["from_callsign_norm"] = df["from_callsign"].apply(_callsign_norm)
    if "to_callsign" in df.columns:
        df["to_callsign_norm"] = df["to_callsign"].apply(_callsign_norm)

    # сортування за часом
    if "datetime" in df.columns:
        df = df.sort_values(["datetime"]).reset_index(drop=True)

    return df
