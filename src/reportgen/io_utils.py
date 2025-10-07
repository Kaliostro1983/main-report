from pathlib import Path
import pandas as pd

from pathlib import Path
from datetime import datetime
import re
from typing import Callable

_TS_RE = re.compile(
    r"report_(\d{4}-\d{2}-\d{2}T\d{2}-\d{2})_(\d{4}-\d{2}-\d{2}T\d{2}-\d{2})",
    re.IGNORECASE,
)


def find_latest(path_dir: str, mask: str) -> Path:
    base = Path(path_dir)
    files = list(base.glob(mask))
    if not files:
        raise FileNotFoundError(f"No files match: {base}\\{mask}")
    return max(files, key=lambda p: p.stat().st_ctime)

def read_excel(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")
    # engine='openpyxl' для .xlsx
    return pd.read_excel(path, engine="openpyxl")

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


def peleng_path(beamshots_dir: str, freq: str) -> str | None:
    d = Path(beamshots_dir)
    candidates = [
        d / f"{freq}.png",                           # 4 знаки (нормалізована частота, типу 408.3150)
        d / f"{float(freq):.3f}.png" if freq else None,  # запасний варіант з 3 знаками (408.315)
    ]
    for p in candidates:
        if p and p.exists():
            return str(p)
    return None


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
