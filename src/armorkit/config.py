from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

@dataclass
class PathsCfg:
    freq_file: str
    reports_dir: str
    output_dir: str
    beamshots_dir: str | None = None
    report_mask: str = "report_*.xlsx"

@dataclass
class Config:
    paths: PathsCfg
    grouping: Dict[str, List[str]] = None
    callsign_aliases: Dict[str, str] = None

def load_config(path: str | Path) -> Config: ...
