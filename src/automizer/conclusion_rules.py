# src/automizer/conclusion_rules.py
from __future__ import annotations

import re
from typing import Any, List, Tuple

from src.automizer.conclusion_types import ConclusionTemplate, KeywordRule


def _prefix_regex(word: str) -> re.Pattern:
    # Повністю копіюємо поведінку, щоб не змінювати результати
    w_raw = (word or "").strip().lower().replace("ё", "е")
    w = re.escape(w_raw)
    if len(w_raw) < 3:
        pattern = rf"\b{w}\b"
    else:
        pattern = rf"\b{w}[\w\-]*"
    return re.compile(pattern, flags=re.IGNORECASE | re.UNICODE)


def find_template_candidates(
    text: str,
    templates: List[ConclusionTemplate],
) -> List[Tuple[ConclusionTemplate, KeywordRule]]:
    """
    Повертає список [(tmpl, rule), ...] у тому ж порядку, як у JSON.
    Не зупиняється на першому збігу в межах шаблону.
    """
    text_norm = (text or "")
    matches: List[Tuple[ConclusionTemplate, KeywordRule]] = []

    for tmpl in templates:
        for rule in tmpl.keywords:
            pat = _prefix_regex(rule.word)
            if pat.search(text_norm):

                # 🔥 перевірка винятків
                if rule.exceptions:
                    blocked = False
                    for exc in rule.exceptions:
                        exc_pat = _prefix_regex(exc)
                        if exc_pat.search(text_norm):
                            blocked = True
                            break

                    if blocked:
                        continue

                matches.append((tmpl, rule))
    return matches
