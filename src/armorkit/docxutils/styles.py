import logging
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -----------------------
# Базові стилі/утиліти
# -----------------------
def set_base_styles(doc: Document):
    style = doc.styles['Normal']
    style.font.size = Pt(12)

def add_title(doc: Document, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    return p