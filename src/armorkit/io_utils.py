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

