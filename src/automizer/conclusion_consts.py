# src/automizer/conclusion_consts.py
from __future__ import annotations

# --------- Статуси висновків ---------
STATUS_EMPTY: str = "empty"
STATUS_NEED_APPROVE: str = "need_approve"
STATUS_APPROVED: str = "approved"

# --------- Колонки службові ---------
COL_STATUS = "__status"
COL_NOTES = "примітки"
COL_MATCHED_TEMPLATE = "__matched_template"
COL_MATCHED_WORD = "__matched_word"
COL_MULTI_MATCH = "__multi_match"

# Назва текстової колонки
TEXT_COL = "р\\обмін"