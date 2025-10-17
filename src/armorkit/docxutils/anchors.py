from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_internal_link(cell, text: str, anchor: str):
    """
    Додає внутрішній лінк на закладку 'anchor' всередині осередка таблиці.
    """
    p = cell.paragraphs[0]
    # очистимо існуючий текст
    for r in list(p.runs):
        r._element.getparent().remove(r._element)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), anchor)
    # стиль гіперлінка без синього/підкреслення: створимо run і налаштуємо вручну
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    # прибираємо підкреслення й колір
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'none')
    color = OxmlElement('w:color'); color.set(qn('w:val'), '000000')
    rPr.append(u); rPr.append(color)
    t = OxmlElement('w:t'); t.text = text
    new_run.append(rPr); new_run.append(t)
    hyperlink.append(new_run)
    p._p.append(hyperlink)
    
    
    # -----------------------
# Внутрішні гіперпосилання / закладки
# -----------------------
def bookmark(paragraph, name: str):
    """Додає закладку перед абзацом."""
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), '0')
    start.set(qn('w:name'), name)

    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), '0')

    p = paragraph._p
    p.insert(0, start)
    p.append(end)