from __future__ import annotations
from pathlib import Path
from typing import List
import pandas as pd
from .data_loader import load_tables, combine_tables
from .export.word_report import export_docx
from .pdf_report import export_pdf

def run_pipeline(inputs: List[str], out_dir: str = "build") -> dict:
    """Основний конвеєр: завантаження -> об'єднання -> експорт DOCX та PDF."""
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    tables = load_tables(inputs)
    df = combine_tables(tables)

    docx_path = Path(out_dir) / "report.docx"
    pdf_path = Path(out_dir) / "report.pdf"

    export_docx(df, str(docx_path))
    export_pdf(df, str(pdf_path))

    return {"rows": len(df), "docx": str(docx_path), "pdf": str(pdf_path)}
