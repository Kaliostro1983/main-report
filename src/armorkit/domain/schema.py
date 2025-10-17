import pandas as pd

COL_FREQ="Частота", 
COL_DATE="Дата", 
COL_TIME="Час", 
COL_WHO="хто", 
COL_TO="кому"

def message_columns(df: pd.DataFrame) -> tuple[str | None, str | None]:
    """
    Визначає назви колонок для тексту перехоплення та коментаря.
    Повертає (msg_col, cmt_col). Якщо не знайдено — None.
    """
    msg_candidates = ["р\\обмін", "р/обмін", "Перехоплення"]
    cmt_candidates = ["примітки", "Коментар", "коментар"]
    msg_col = next((c for c in msg_candidates if c in df.columns), None)
    cmt_col = next((c for c in cmt_candidates if c in df.columns), None)
    return msg_col, cmt_col