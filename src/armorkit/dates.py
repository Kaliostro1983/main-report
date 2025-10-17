from datetime import datetime
import re
import pandas as pd
import pathlib
from pathlib import Path


# -----------------------
# Парсинг періоду з назв
# -----------------------
_PERIOD_RE = re.compile(
    r"report_(\d{4}-\d{2}-\d{2}T\d{2}-\d{2})_(\d{4}-\d{2}-\d{2}T\d{2}-\d{2})",
    re.IGNORECASE,
)


def parse_period_from_filename(path: str) -> tuple[str, str]:
    m = _PERIOD_RE.search(Path(path).name)
    if not m:
        return "", ""
    def fmt(s: str) -> str:
        dt = datetime.strptime(s, "%Y-%m-%dT%H-%M")
        return dt.strftime("%d.%m.%Y %H:%M")
    return fmt(m.group(1)), fmt(m.group(2))


def combine_date_time(date_val, time_val) -> pd.Timestamp | None:
    if pd.isna(date_val) and pd.isna(time_val):
        return None
    try:
        t = str(time_val).strip() if pd.notna(time_val) else "00:00"
        if len(t) == 5 and t[2] == ":":
            t = t + ":00"
        s = f"{str(date_val).strip()} {t}".strip()
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return None
    
    
def format_for_filename(dt_obj) -> str:
        """
        Приймає або datetime, або рядок дати/часу і повертає
        формат для імені файлу: 'ДД.ММ.РРРР_ГГ-ММ'.
        Підтримує кілька поширених форматів рядка.
        """
        if isinstance(dt_obj, datetime):
            return dt_obj.strftime("%d.%m.%Y_%H-%M")

        if isinstance(dt_obj, str):
            for fmt in ("%Y-%m-%dT%H-%M",     # 2025-10-02T12-00   (з назви report_*.xlsx)
                        "%d.%m.%Y %H:%M",     # 02.10.2025 12:00
                        "%Y-%m-%d %H:%M:%S",  # 2025-10-02 12:00:00
                        ):
                try:
                    dt = datetime.strptime(dt_obj.strip(), fmt)
                    return dt.strftime("%d.%m.%Y_%H-%M")
                except ValueError:
                    continue
            # якщо не розпізнали — повернемо як є (краще так, ніж падати)
            return dt_obj

        # на всяк випадок — явно конвертуємо
        return str(dt_obj)
    
    
def build_report_filename(start, end, prefix="Звіт РЕР", sufix='docx') -> str:
    return f"{prefix} ({format_for_filename(start)} - {format_for_filename(end)}).{sufix}"