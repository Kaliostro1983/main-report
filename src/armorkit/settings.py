# src/reportgen/settings.py
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional
import yaml

@dataclass
class PathsCfg:
    freq_file: str
    reports_dir: str
    output_dir: str = "build"
    beamshots_dir: Optional[str] = None
    report_mask: str = "report_*.xlsx"   # <-- ДОДАЛИ

@dataclass
class Config:
    paths: PathsCfg
    grouping: Dict[str, Any] | None = None
    callsign_aliases: Dict[str, str] | None = None

def _as_dict(d: Dict[str, Any] | None) -> Dict[str, Any]:
    return d if isinstance(d, dict) else {}

def load_config(path: str = "config.yml") -> Config:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Config not found: {p}")
    with p.open("r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    paths_raw = _as_dict(raw.get("paths"))
    grouping = _as_dict(raw.get("grouping"))
    callsign_aliases = _as_dict(raw.get("callsign_aliases"))

    freq_file = paths_raw.get("freq_file")
    reports_dir = paths_raw.get("reports_dir")
    if not freq_file or not reports_dir:
        raise ValueError("Config.paths must include 'freq_file' and 'reports_dir'.")

    paths = PathsCfg(
        freq_file=freq_file,
        reports_dir=reports_dir,
        output_dir=paths_raw.get("output_dir", "build"),
        beamshots_dir=paths_raw.get("beamshots_dir"),
        report_mask=paths_raw.get("report_mask", "report_*.xlsx"),  # <-- ДОДАЛИ
    )
    
    return Config(
        paths=paths,
        grouping=grouping,
        callsign_aliases=callsign_aliases,   # <-- Виправлено (було call_sign_aliases)
    )