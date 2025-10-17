import re
import pandas as pd
from src.armorkit.domain.freqnorm import freq4_str

def normalize_callsign(s: str) -> str:
    # верхній регістр, пробіли -> дефіс, прибираємо повторні дефіси
    x = str(s).strip().upper().replace(" ", "-")
    while "--" in x:
        x = x.replace("--", "-")
    return x


def extract_callsigns_for_freq(df: pd.DataFrame, freq4: str, aliases: dict[str, str] | None) -> list[str]:
    """
    Збирає унікальні позивні з колонок 'Хто' та 'Кому' для заданої частоти.
    Розділювач у клітинці — кома або крапка з комою. Ігнорує 'НВ'.
    Застосовує aliases (cfg.callsign_aliases), якщо задані.
    """
    alias = aliases or {}
    work = df.copy()
    work["__f4"] = work["Частота"].map(freq4_str)
    part = work[work["__f4"] == freq4]

    calls: set[str] = set()
    for col in ("Хто", "Кому"):
        if col not in part.columns:
            continue
        for raw in part[col].dropna().astype(str):
            # розбиваємо за , або ; 
            for token in re.split(r"[;,]", raw):
                t = normalize_callsign(token)
                if not t or t == "НВ":
                    continue
                t = alias.get(t, t)  # виправлення опечаток
                calls.add(t)

    return sorted(calls)


def build_callsign_str_for_freq(df: pd.DataFrame, freq4: str) -> str:
    """
    Повертає рядок позивних для заданої частоти:
    - фільтр по 'Частота' (формат ###.####)
    - беремо колонки 'хто' та 'кому'
    - ділимо тільки за комою
    - UPPERCASE, пробіли -> дефіс
    - прибираємо 'НВ'
    - унікалізуємо
    """
    # 1) фільтр по частоті
    tmp = df.copy()
    tmp["__f4"] = tmp["Частота"].map(freq4_str)
    part = tmp[tmp["__f4"] == freq4]

    # 2) збираємо токени з 'хто' і 'кому'
    items: list[str] = []
    for col in ("хто", "кому"):
        if col not in part.columns:
            continue  # (мінімальний захист)
        for v in part[col].dropna().astype(str):
            v = v.strip()
            if "," in v:
                items.extend([t.strip() for t in v.split(",") if t.strip()])
            else:
                items.append(v)

    # 3) нормалізація + множина
    norm = []
    for t in items:
        t = t.upper().replace(" ", "-")
        if t and t != "НВ":
            norm.append(t)

    uniq = sorted(set(norm))
    # 5) у рядок
    return ", ".join(uniq) if uniq else "—"


def normalize_callsign(s: str) -> str:
    x = str(s).strip().upper()
    x = x.replace(" ", "-")
    for ch in ['"', "'", "«", "»", "[", "]", "(", ")", "–", "—"]:
        x = x.replace(ch, "")
    while "--" in x:
        x = x.replace("--", "-")
    return x