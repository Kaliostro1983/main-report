import pandas as pd
from pathlib import Path
import re
from datetime import datetime
from typing import Callable

def next_available_path(p: Path, reason_suffix: str = "") -> Path:
    """
    Якщо p зайнятий/недоступний, повертає новий шлях:
    <name>__<suffix|timestamp>.<ext>
    """
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = reason_suffix.strip("_") or stamp
    candidate = p.with_name(f"{p.stem}__{suffix}{p.suffix}")
    if not candidate.exists():
        return candidate
    # рідкісний випадок колізії -> лічильник
    i = 2
    while candidate.exists():
        candidate = p.with_name(f"{p.stem}__{suffix}_{i}{p.suffix}")
        i += 1
    return candidate

def safe_save_docx(doc, path: str | Path) -> Path:
    """
    Пробує зберегти DOCX. Якщо файл відкритий користувачем (PermissionError),
    зберігає під новою назвою і повертає новий шлях.
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    try:
        doc.save(p)
        return p
    except PermissionError:
        alt = next_available_path(p, reason_suffix="opened")
        print(f"[INFO] Файл зайнятий: {p.name}. Збережено як: {alt.name}")
        doc.save(alt)
        return alt

def safe_save_xlsx(writer_fn: Callable[[Path], None], path: str | Path) -> Path:
    """
    Узагальнений сейвер для XLSX (pandas). Приймає функцію, яка пише у файл.
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    try:
        writer_fn(p)
        return p
    except PermissionError:
        alt = next_available_path(p, reason_suffix="opened")
        print(f"[INFO] Файл зайнятий: {p.name}. Збережено як: {alt.name}")
        writer_fn(alt)
        return alt