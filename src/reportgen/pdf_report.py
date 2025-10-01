from __future__ import annotations
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import simpleSplit
from datetime import datetime
from pathlib import Path
import pandas as pd

def export_pdf(df: pd.DataFrame, out_path: str, title: str = "Аналітичний звіт") -> str:
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(out_path, pagesize=A4)
    width, height = A4

    # Заголовок
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2, height - 20*mm, title)
    c.setFont("Helvetica", 10)
    c.drawRightString(width - 15*mm, height - 27*mm, f"Згенеровано: {datetime.now():%Y-%m-%d %H:%M:%S}")

    # Резюме
    y = height - 40*mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(15*mm, y, "Резюме")
    y -= 8*mm
    c.setFont("Helvetica", 10)
    lines = [
        f"Кількість рядків: {len(df)}",
        f"Кількість колонок: {len(df.columns)}",
        f"Джерел (файлів): {df['__source__'].nunique() if '__source__' in df.columns else '—'}",
    ]
    for line in lines:
        c.drawString(20*mm, y, line); y -= 6*mm

    # Таблична вибірка (текстово, перші 15 рядків)
    y -= 4*mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(15*mm, y, "Вибірка даних (top 15)")
    y -= 8*mm
    c.setFont("Helvetica", 8)

    cols = list(df.columns)
    sample = df.head(15)
    col_line = " | ".join(str(cn) for cn in cols)
    for wrapped in simpleSplit(col_line, "Helvetica", 8, width - 30*mm):
        c.drawString(15*mm, y, wrapped); y -= 5*mm
    c.line(15*mm, y, width-15*mm, y); y -= 5*mm

    for _, row in sample.iterrows():
        text = " | ".join(str(row.get(cn, "")) for cn in cols)
        wrapped_lines = simpleSplit(text, "Helvetica", 8, width - 30*mm)
        for wline in wrapped_lines:
            if y < 20*mm:
                c.showPage(); y = height - 20*mm
                c.setFont("Helvetica", 8)
            c.drawString(15*mm, y, wline); y -= 5*mm

    c.showPage()
    c.save()
    return out_path
