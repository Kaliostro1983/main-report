# src/automizer/conclusion_render.py
from __future__ import annotations

from typing import Any
import pandas as pd

from src.automizer.freqdict import lookup_place_unit


def render_with_place_unit(description: str, freq_value: Any, reference_df: pd.DataFrame) -> str:
    place, unit = lookup_place_unit(freq_value, reference_df)
    return description.replace("{PLACE}", place).replace("{UNIT}", unit)
