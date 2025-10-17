import logging
from docx import Document   
from docx.shared import Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT, WD_ALIGN_VERTICAL


def set_col_widths(table, factors):
    total = sum(factors)
    total_inches = 6.5
    widths = [Inches(total_inches * f / total) for f in factors]
    for i, w in enumerate(widths):
        for row in table.rows:
            row.cells[i].width = w

def center_cell(cell):
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def vcenter(cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def set_row_min_height(row, cm: float = 0.9):
    row.height = Cm(cm)
    row.height_rule = WD_ROW_HEIGHT.AT_LEAST