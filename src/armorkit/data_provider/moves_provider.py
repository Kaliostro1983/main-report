from __future__ import annotations

from pathlib import Path
from typing import Any, Optional
import pandas as pd


def _build_gsheet_csv_url(spreadsheet_id: str, gid: Optional[str] = None) -> str:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv"
    if gid:
        url += f"&gid={gid}"
    return url


def _moves_cfg_from_cfg_object(cfg: Any) -> Optional[Any]:
    if cfg is None:
        return None
    if isinstance(cfg, dict):
        return cfg.get("moves")
    if hasattr(cfg, "moves"):
        return getattr(cfg, "moves")
    return None


def _value(obj: Any, key: str, default: Any = None) -> Any:
    if obj is None:
        return default
    if isinstance(obj, dict):
        return obj.get(key, default)
    if hasattr(obj, key):
        v = getattr(obj, key)
        return default if v is None else v
    return default


def _moves_cfg_from_yaml(config_path: str) -> dict:
    try:
        import yaml
    except Exception as e:
        raise RuntimeError("Потрібен пакет PyYAML: pip install pyyaml") from e

    with open(config_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if not isinstance(data, dict):
        return {}
    moves = data.get("moves") or {}
    return moves if isinstance(moves, dict) else {}


def load_moves_df(cfg: Any, config_path: str, default_xlsx_path: Path) -> pd.DataFrame:
    moves_cfg = _moves_cfg_from_cfg_object(cfg)

    # Якщо Config не містить moves — беремо з YAML напряму
    if not moves_cfg:
        moves_cfg = _moves_cfg_from_yaml(config_path)

    source = str(_value(moves_cfg, "source", "xlsx")).strip().lower()

    if source == "gsheet_csv":
        spreadsheet_id = str(_value(moves_cfg, "spreadsheet_id", "")).strip()
        gid = str(_value(moves_cfg, "gid", "")).strip() or None

        if not spreadsheet_id:
            raise ValueError("moves.spreadsheet_id не задано для moves.source=gsheet_csv")

        url = _build_gsheet_csv_url(spreadsheet_id, gid=gid)

        try:
            return pd.read_csv(url)
        except Exception as e:
            raise RuntimeError(
                "Не вдалось прочитати Google Sheet як CSV. "
                "Перевір доступ: 'будь-хто з посиланням (Viewer)' та правильність spreadsheet_id/gid."
            ) from e

    # дефолт: xlsx
    xlsx_path = _value(moves_cfg, "xlsx_path", None)
    moves_path = Path(xlsx_path) if xlsx_path else default_xlsx_path

    if not moves_path.exists():
        raise FileNotFoundError(f"Файл з переміщеннями не знайдено: {moves_path}")

    return pd.read_excel(moves_path)