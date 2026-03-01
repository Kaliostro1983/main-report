# src/automizer/conclusion_types.py
from __future__ import annotations
from dataclasses import dataclass

@dataclass
class KeywordRule:
    word: str
    probability: int
    exceptions: list[str] | None = None

@dataclass
class ConclusionTemplate:
    name: str
    description: str
    keywords: list[KeywordRule]
    shortcut: str = ""