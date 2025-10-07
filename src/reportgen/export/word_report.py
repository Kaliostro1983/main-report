# src/reportgen/export/word_report.py
from __future__ import annotations

from collections import OrderedDict, Counter
from pathlib import Path
from datetime import datetime
import re
import pandas as pd

from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from src.reportgen.data_loader import load_inputs
from src.reportgen.normalize_freq import normalize_frequency_column, FREQ_NOT_FOUND
from src.reportgen.grouping import (
    unique_frequencies_with_counts,
    group_frequencies_by_tag,
)
from src.reportgen.settings import load_config

# зверху файла, поруч з іншими імпортами
from src.reportgen.io_utils import safe_save_docx

import logging
log = logging.getLogger(__name__)

# -----------------------
# Парсинг періоду з назв
# -----------------------
_PERIOD_RE = re.compile(
    r"report_(\d{4}-\d{2}-\d{2}T\d{2}-\d{2})_(\d{4}-\d{2}-\d{2}T\d{2}-\d{2})",
    re.IGNORECASE,
)

def _normalize_callsign(s: str) -> str:
    # верхній регістр, пробіли -> дефіс, прибираємо повторні дефіси
    x = str(s).strip().upper().replace(" ", "-")
    while "--" in x:
        x = x.replace("--", "-")
    return x

def _network_is_empty(intercepts_df: pd.DataFrame, freq4: str) -> bool:
    """
    True  -> для цієї частоти НІ ОДНОГО перехоплення з непорожнім коментарем.
    False -> є хоча б один рядок із коментарем (секцію можна публікувати).
    """
    # визначаємо назву колонки коментаря
    _, cmt_col = _message_columns(intercepts_df)
    if not cmt_col:
        return True  # немає поля коментаря => вважаємо порожньою

    df = intercepts_df.copy()
    df["__f4"] = df["Частота"].map(_freq4_str)
    part = df[df["__f4"] == freq4]

    if part.empty:
        return True

    # фільтр НЕпорожніх коментарів (без 'nan'/'None')
    cm = part[cmt_col].astype(str).fillna("").str.strip()
    cm = cm.replace({"nan": "", "None": "", "NONE": ""})
    part = part[cm.ne("")]
    return part.empty


def _build_callsign_set_for_freq(df: pd.DataFrame, freq4: str, cfg_grouping: dict | None) -> list[str]:
    """
    Повертає відсортовану унікальну множину позивних для частоти:
    колонки 'Хто' та 'Кому', розділювач — кома. Виключає 'НВ'.
    Застосовує callsign_aliases з конфіга (якщо є).
    """
    alias = ((cfg_grouping or {}).get("callsign_aliases") or
             ({}))  # сумісність: якщо зберігаєш aliases не в grouping — перенеси туди
    # допускаємо також верхнього рівня cfg (якщо ти зберігаєш там)
    try:
        from src.reportgen.settings import load_config
        # нічого не робимо — просто залишив сумісність, якщо зручно тримати aliases не тут
    except Exception:
        pass

    s = set()
    # фільтр по частоті
    tmp = df.copy()
    tmp["__f4"] = tmp["Частота"].map(_freq4_str)
    part = tmp[tmp["__f4"] == freq4]

    for col in ("хто", "кому"):
        if col not in part.columns:
            continue
        for v in part[col].dropna().astype(str):
            for token in v.split(","):
                t = _normalize_callsign(token)
                if not t or t == "НВ":
                    continue
                # застосувати alias, якщо є
                t = alias.get(t, t)
                s.add(t)

    return sorted(s)


def build_callsign_str_for_freq(df: pd.DataFrame, freq4: str) -> str:
    """
    Повертає рядок позивних для заданої частоти:
    - фільтр по 'Частота' (формат ###.####)
    - беремо колонки 'хто' та 'кому'
    - ділимо тільки за комою
    - UPPERCASE, пробіли -> дефіс
    - прибираємо 'НВ'
    - унікалізуємо
    """
    # 1) фільтр по частоті
    tmp = df.copy()
    tmp["__f4"] = tmp["Частота"].map(_freq4_str)
    part = tmp[tmp["__f4"] == freq4]

    # 2) збираємо токени з 'хто' і 'кому'
    items: list[str] = []
    for col in ("хто", "кому"):
        if col not in part.columns:
            continue  # (мінімальний захист)
        for v in part[col].dropna().astype(str):
            v = v.strip()
            if "," in v:
                items.extend([t.strip() for t in v.split(",") if t.strip()])
            else:
                items.append(v)

    # 3) нормалізація + множина
    norm = []
    for t in items:
        t = t.upper().replace(" ", "-")
        if t and t != "НВ":
            norm.append(t)

    uniq = sorted(set(norm))
    # 5) у рядок
    return ", ".join(uniq) if uniq else "—"



def _parse_period_from_filename(path: str) -> tuple[str, str]:
    m = _PERIOD_RE.search(Path(path).name)
    if not m:
        return "", ""
    def fmt(s: str) -> str:
        dt = datetime.strptime(s, "%Y-%m-%dT%H-%M")
        return dt.strftime("%d.%m.%Y %H:%M")
    return fmt(m.group(1)), fmt(m.group(2))

def _to_datetime(date_val, time_val) -> pd.Timestamp | None:
    if pd.isna(date_val) and pd.isna(time_val):
        return None
    try:
        t = str(time_val).strip() if pd.notna(time_val) else "00:00"
        if len(t) == 5 and t[2] == ":":
            t = t + ":00"
        s = f"{str(date_val).strip()} {t}".strip()
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return None

def _freq4_str(x) -> str | None:
    try:
        v = float(str(x).replace(",", "."))
        return f"{v:.4f}"
    except Exception:
        return None
    
# у word_report.py біля хелперів

def _normalize_callsign(s: str) -> str:
    x = str(s).strip().upper()
    x = x.replace(" ", "-")
    for ch in ['"', "'", "«", "»", "[", "]", "(", ")", "–", "—"]:
        x = x.replace(ch, "")
    while "--" in x:
        x = x.replace("--", "-")
    return x

def _extract_callsigns_for_freq(df: pd.DataFrame, freq4: str, aliases: dict[str, str] | None) -> list[str]:
    """
    Збирає унікальні позивні з колонок 'Хто' та 'Кому' для заданої частоти.
    Розділювач у клітинці — кома або крапка з комою. Ігнорує 'НВ'.
    Застосовує aliases (cfg.callsign_aliases), якщо задані.
    """
    alias = aliases or {}
    work = df.copy()
    work["__f4"] = work["Частота"].map(_freq4_str)
    part = work[work["__f4"] == freq4]

    calls: set[str] = set()
    for col in ("Хто", "Кому"):
        if col not in part.columns:
            continue
        for raw in part[col].dropna().astype(str):
            # розбиваємо за , або ; 
            for token in re.split(r"[;,]", raw):
                t = _normalize_callsign(token)
                if not t or t == "НВ":
                    continue
                t = alias.get(t, t)  # виправлення опечаток
                calls.add(t)

    return sorted(calls)


# -----------------------
# Базові стилі/утиліти
# -----------------------
def _set_base_styles(doc: Document):
    style = doc.styles['Normal']
    style.font.size = Pt(12)

def _add_title(doc: Document, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    return p

def _set_col_widths(table, factors):
    total = sum(factors)
    total_inches = 6.5
    widths = [Inches(total_inches * f / total) for f in factors]
    for i, w in enumerate(widths):
        for row in table.rows:
            row.cells[i].width = w

def _center_cell(cell):
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def _vcenter(cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def _set_row_min_height(row, cm: float = 0.9):
    row.height = Cm(cm)
    row.height_rule = WD_ROW_HEIGHT.AT_LEAST

# -----------------------
# Довідник: назва мережі
# -----------------------
_NET_NAME_CANDIDATES = [
    "Назва радіомережі", "Назва мережі", "Радіомережа",
    "Назва", "Мережа", "Опис", "Призначення"
]
def _get_network_name_by_freq(freq4: str, ref_df: pd.DataFrame) -> str:
    if "Частота" not in ref_df.columns:
        return "—"
    df = ref_df.copy()
    def _f4(x):
        try: return f"{float(str(x).replace(',', '.')):.4f}"
        except: return None
    df["__f4"] = df["Частота"].map(_f4)
    m = df[df["__f4"] == freq4]
    if m.empty: return "—"
    row = m.iloc[0]
    for c in _NET_NAME_CANDIDATES:
        if c in row.index and pd.notna(row[c]) and str(row[c]).strip():
            return str(row[c]).strip()
    return "—"

# -----------------------
# Довідник: повна «Хто» для групи
# -----------------------
def _full_tag_for_group(freq_list, ref_df: pd.DataFrame, fallback_short: str) -> str:
    if not freq_list:
        return fallback_short
    tmp = ref_df.copy()
    def _f4(x):
        try: return f"{float(str(x).replace(',', '.')):.4f}"
        except: return None
    tmp["__f4"] = tmp["Частота"].map(_f4)
    vals = []
    if "Хто" not in tmp.columns:
        return fallback_short
    for f in freq_list:
        m = tmp[tmp["__f4"] == f]
        if not m.empty and pd.notna(m.iloc[0]["Хто"]):
            vals.append(str(m.iloc[0]["Хто"]).strip())
    if not vals:
        return fallback_short
    return Counter(vals).most_common(1)[0][0]

# -----------------------
# Еталонні вкладки: 'Призначення', 'Склад кореспондентів'
# -----------------------
def _read_reference_sheet(freq4: str, ref_xlsx_path: str) -> dict:
    """
    Повертає {'Призначення': str|None, 'Склад кореспондентів': str|None}
    з аркуша, названого freq4 (наприклад '145.9500').
    Очікується таблиця з колонками 'Категорія'/'Значення' (регістро-незалежно).
    """
    result = {"Призначення": None, "Склад кореспондентів": None}
    try:
        xls = pd.ExcelFile(ref_xlsx_path, engine="openpyxl")
    except Exception as e:
        log.warning("Не вдалося відкрити довідник '%s': %s", ref_xlsx_path, e)
        return result

    if freq4 not in xls.sheet_names:
        log.info("Еталонний аркуш для %s не знайдено у файлі довідника.", freq4)
        return result

    try:
        df = pd.read_excel(xls, sheet_name=freq4)
    except Exception as e:
        log.warning("Не вдалося зчитати аркуш '%s': %s", freq4, e)
        return result

    if df.empty:
        log.info("Аркуш еталонки '%s' порожній.", freq4)
        return result

    # знайдемо колонки категорії/значення
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    cat_col = next((c for c, n in cols_lower.items() if "категор" in n), None)
    val_col = next((c for c, n in cols_lower.items() if "значен" in n), None)
    if not cat_col or not val_col:
        # fallback: інколи перша колонка - "№", тому беремо 2-гу та 3-тю
        try:
            cat_col, val_col = df.columns[1], df.columns[2]
        except Exception:
            log.info("На аркуші '%s' не знайшов пар колонок 'Категорія'/'Значення'.", freq4)
            return result

    def norm(s: str) -> str:
        return str(s).strip().lower()

    wanted = {"призначення": "Призначення",
              "склад кореспондентів": "Склад кореспондентів"}

    found = []
    for _, row in df.iterrows():
        key = norm(row.get(cat_col, ""))
        if key in wanted:
            val = row.get(val_col, None)
            if pd.notna(val) and str(val).strip():
                result[wanted[key]] = str(val).strip()
                found.append(wanted[key])

    if found:
        log.info("Еталонка %s: витягнуто поля %s.", freq4, ", ".join(found))
    else:
        log.info("Для %s еталонні поля не знайдені на аркуші.", freq4)

    return result

# -----------------------
# Позивні (Хто/Кому)
# -----------------------
_CALLS_SPLIT_RE = re.compile(r"[;,]")  # кома або крапка з комою

def _normalize_callsign(token: str) -> str:
    s = token.strip().upper()
    if not s: return ""
    # пробіли -> дефіс
    s = re.sub(r"\s+", "-", s)
    # кілька дефісів -> один
    s = re.sub(r"-{2,}", "-", s)
    # дозволені символи: літера (лат/кирил), цифра, дефіс
    s = re.sub(r"[^A-ZА-ЯІЇЄҐ0-9-]", "", s)
    return s

def _extract_callsigns_for_freq(df: pd.DataFrame, freq4: str, cfg: dict | None) -> list[str]:
    """
    Множина позивних для частоти з колонок 'Хто' та 'Кому'.
    - розділювачі: , ;
    - нормалізація (CAPS, пробіли→дефіс, очищення)
    - вилучаємо 'НВ'
    - застосовуємо aliases з конфігу (якщо є)
    """
    if "Частота" not in df.columns:
        return []
    df2 = df.copy()
    df2["__f4"] = df2["Частота"].map(_freq4_str)
    part = df2[df2["__f4"] == freq4]

    out = []
    for col in ["Хто", "Кому"]:
        if col not in part.columns:
            continue
        for val in part[col].fillna(""):
            for raw_tok in _CALLS_SPLIT_RE.split(str(val)):
                tok = _normalize_callsign(raw_tok)
                if not tok:
                    continue
                # прибираємо НВ
                if tok == "НВ":
                    continue
                out.append(tok)

    # aliases з конфігу
    aliases = ((cfg or {}).get("callsign_aliases") or {})
    out = [aliases.get(x, x) for x in out]

    # унікальність + сортування
    uniq = sorted(set(out))
    return uniq

# -----------------------
# Вставка зображення пеленгів (заглушка)
# -----------------------
def insert_bearing_image(doc: Document, freq4: str):
    """
    Заглушка. Залишаємо порожньою — реалізуємо пізніше.
    """
    return

# -----------------------
# Внутрішні гіперпосилання / закладки
# -----------------------
def _bookmark(paragraph, name: str):
    """Додає закладку перед абзацом."""
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), '0')
    start.set(qn('w:name'), name)

    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), '0')

    p = paragraph._p
    p.insert(0, start)
    p.append(end)
    
def _message_columns(df: pd.DataFrame) -> tuple[str | None, str | None]:
    """
    Визначає назви колонок для тексту перехоплення та коментаря.
    Повертає (msg_col, cmt_col). Якщо не знайдено — None.
    """
    msg_candidates = ["р\\обмін", "р/обмін", "Перехоплення"]
    cmt_candidates = ["примітки", "Коментар", "коментар"]
    msg_col = next((c for c in msg_candidates if c in df.columns), None)
    cmt_col = next((c for c in cmt_candidates if c in df.columns), None)
    return msg_col, cmt_col




def _append_executor_block(doc: Document) -> None:
    """Додає примітку і рядок виконавця в кінці звіту (без розриву сторінки)."""
    note = (
        "* Добування розвідувальної інформації про противника здійснюється підрозділами "
        "РЕР, розгорнутими в смугах відповідальності ОТУ та надається для первинної обробки "
        "у відповідні підрозділи РЕР центрів (відділів) розвідки ОТУ, де інформація "
        "узагальнюється та надається короткий опис подій.\n\n"
        "Матеріали радіоперехоплень в узагальненому вигляді від підрозділів РЕР ЦР (ВР) ОТУ у "
        "визначений час надаються черговому розвідки ОКП ОСУВ “Хортиця”. Інформація, яка "
        "потребує невідкладного доведення до начальника центру розвідки ОКП ОСУВ "
        "“Хортиця”, надається черговому розвідки ОКП ОСУВ “Хортиця” негайно по тлф з "
        "подальшим документальним підтвердженням."
    )
    p1 = doc.add_paragraph(note)
    for r in p1.runs:
        r.italic = True
        r.font.size = Pt(12)

    doc.add_paragraph()  # порожній рядок

    p2 = doc.add_paragraph(
        "Командир взводу РЕР _________СТЕПУРА Андрій Іванович__________________"
    )
    for r in p2.runs:
        r.font.size = Pt(12)



def _render_frequency_section(doc: Document, freq4: str, count: int, li, cfg) -> None:
    # Якір
    title_p = doc.add_paragraph()
    anchor = f"freq-{freq4.replace('.', '_')}"
    _bookmark(title_p, anchor)

    # Заголовок
    net_name = _get_network_name_by_freq(freq4, li.reference_df)
    run = title_p.add_run(f"[{freq4}] - {net_name} - ({count})")
    run.bold = True
    run.font.size = Pt(12)

    # Еталонка
    ref_sheet = _read_reference_sheet(freq4, li.freq_path)
    purpose = ref_sheet.get("Призначення") or "—"

    # Вузли: з головного листа; якщо порожньо — зі "Склад кореспондентів" еталонки
    nodes = "—"
    if "Частота" in li.reference_df.columns:
        df_ref = li.reference_df.copy()
        df_ref["__f4"] = df_ref["Частота"].map(_freq4_str)
        m = df_ref[df_ref["__f4"] == freq4]
        if not m.empty:
            for col in ["Вузли зв’язку", "Вузли зв'язку", "Вузли", "Вузли звʹязку"]:
                if col in m.columns and pd.notna(m.iloc[0][col]) and str(m.iloc[0][col]).strip():
                    nodes = str(m.iloc[0][col]).strip()
                    break
    if nodes == "—":
        nodes = ref_sheet.get("Склад кореспондентів") or "—"

    # Позивні (беремо aliases з cfg.callsign_aliases)
    calls_text = build_callsign_str_for_freq(li.intercepts_df, freq4)

    # ТРИ абзаци з готовими значеннями
    doc.add_paragraph(f"Призначення радіомережі: {purpose}")
    doc.add_paragraph(f"Вузли зв’язку: {nodes}")
    doc.add_paragraph(f"Список позивних: {calls_text}")

    # Далі — як було: таблиця з 2 колонок тільки для перехоплень з коментарем
    doc.add_paragraph("Найважливіші перехоплення з коментарями:").runs[0].bold = True

    msg_col, cmt_col = _message_columns(li.intercepts_df)
    df = li.intercepts_df.copy()
    df["__f4"] = df["Частота"].map(_freq4_str)
    part = df[df["__f4"] == freq4].copy()

    if cmt_col:
        cm = part[cmt_col].astype(str).fillna("").str.strip().replace({"nan": "", "None": "", "NONE": ""})
        part = part[cm.ne("")]
    else:
        part = part.iloc[0:0]

    if part.empty:
        # якщо з якихось причин сюди дійшли без записів — просто не друкуємо пусту таблицю
        log.info("Секція %s: відсутні перехоплення з коментарями (таблиця пропущена).", freq4)
        doc.add_page_break()
        return

    if all(c in part.columns for c in ("Дата", "Час")):
        part["__dt"] = [_to_datetime(d, t) for d, t in zip(part["Дата"], part["Час"])]
        part = part.sort_values("__dt", kind="stable")

    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Перехоплення"; hdr[1].text = "Коментар"
    for c in hdr: _center_cell(c); _vcenter(c)
    _set_row_min_height(t.rows[0], cm=0.9); _set_col_widths(t, [3.6, 2.4])

    for _, row in part.iterrows():
        msg = str(row[msg_col]).strip() if msg_col and pd.notna(row.get(msg_col)) else ""
        cmt = str(row[cmt_col]).strip() if cmt_col and pd.notna(row.get(cmt_col)) else ""
        cells = t.add_row().cells
        cells[0].text = msg
        cells[1].text = cmt
        _set_row_min_height(t.rows[-1], cm=0.9)

    insert_bearing_image(doc, freq4)
    doc.add_page_break()



def _add_internal_link(cell, text: str, anchor: str):
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
# Перша сторінка (ОГЛЯД) — БЕЗ ЗМІН ВІД ТВОЄЇ ОСТАННЬОЇ ВЕРСІЇ,
# але замість plain-text частоти додаємо КЛІКАЛЬНЕ посилання.
# -----------------------
def _render_overview_page(doc: Document, cfg, li, groups, counts, period_start, period_end):
    _set_base_styles(doc)

    _add_title(doc, "Донесення")

    psub = doc.add_paragraph("за результатами ведення радіоелектронної розвідки\n"
                             "у зоні відповідальності тактичної групи “Кремінна”")
    psub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in psub.runs:
        r.bold = True; r.font.size = Pt(12)

    pper = doc.add_paragraph()
    pper.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = pper.add_run(f"(з {period_start.split(' ')[1]} {period_start.split(' ')[0]} "
                       f"по {period_end.split(' ')[1]} {period_end.split(' ')[0]} року)")
    run.font.size = Pt(12)

    doc.add_paragraph()
    total_intercepts = int(len(li.intercepts_df))
    pinfo = doc.add_paragraph()
    pinfo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = pinfo.add_run(f"За поточний період отримано {total_intercepts} перехоплень.")
    rr.bold = True; rr.font.size = Pt(12)

    doc.add_paragraph()
    title_tbl = doc.add_paragraph("Активність радіомереж:")
    title_tbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_tbl.runs[0].bold = True; title_tbl.runs[0].font.size = Pt(12)
    doc.add_paragraph()

       # Таблиця: № | Частота | Радіомережа | Перехоплення
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "№"
    hdr[1].text = "Частота"
    hdr[2].text = "Радіомережа"
    hdr[3].text = "Перехоплення"

    # заголовки по центру (горизонтально і вертикально) + висота
    for c in hdr:
        _center_cell(c)
        _vcenter(c)
    _set_row_min_height(t.rows[0], cm=0.9)

    # Ширини колонок:
    # 1 — в 4 рази менша; 2 — на 35% менша; 3 — вдвічі більша; 4 — в 4 рази менша.
    # Пропорції: [0.25, 0.65, 2.0, 0.25]
    _set_col_widths(t, [0.25, 0.65, 2.0, 0.25])

    row_counter = 1
    for short_tag, flist in groups.items():
        # повний напис «Хто» (за довідником)
        full_tag = _full_tag_for_group(flist, li.reference_df, short_tag)

        # рядок-заголовок групи (злиті комірки, жирним, по центру)
        r = t.add_row()
        c = r.cells
        c[0].merge(c[1]).merge(c[2]).merge(c[3])
        c[0].text = f"Радіомережі {full_tag}"
        # стилізація комірки заголовка групи
        p = c[0].paragraphs[0]
        if p.runs:
            p.runs[0].bold = True
        _center_cell(c[0])
        _vcenter(c[0])
        _set_row_min_height(t.rows[-1], cm=0.9)

        if not flist:
            # порожня група
            r = t.add_row().cells
            r[2].text = "Не виявлено"
            if r[2].paragraphs[0].runs:
                r[2].paragraphs[0].runs[0].italic = True
            # центруємо 1,2,4 колонки
            for idx in (0, 1, 3):
                _center_cell(r[idx])
                _vcenter(r[idx])
            _set_row_min_height(t.rows[-1], cm=0.9)
            continue

        # рядки з даними по частотах
        for f in flist:
            r = t.add_row().cells
            # №
            r[0].text = str(row_counter)
            _center_cell(r[0]); _vcenter(r[0])

            # Частота як клікабельний лінк на закладку розділу частоти
            if not _network_is_empty(li.intercepts_df, f):
                anchor = f"freq-{f.replace('.', '_')}"
                _add_internal_link(r[1], f, anchor)
            else:
                r[1].text = f  # просто текст без гіперпосилання
            _center_cell(r[1]); _vcenter(r[1])

            # Назва мережі (з довідника)
            r[2].text = _get_network_name_by_freq(f, li.reference_df)

            # Кількість перехоплень
            r[3].text = str(counts.get(f, 0))
            _center_cell(r[3]); _vcenter(r[3])

            _set_row_min_height(t.rows[-1], cm=0.9)
            row_counter += 1
            
    # doc.add_page_break()

            
# --- ПУБЛІЧНИЙ API: згенерувати DOCX-чернетку ---
__all__ = ["build_draft_docx"]

def build_draft_docx(config_path: str = "config.yml") -> str:
    
    import logging
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    """
    Генерує DOCX:
      - перша сторінка (огляд з клікабельними частотами)
      - розділи по кожній частоті (page break між ними)
    Повертає абсолютний шлях до збереженого файлу.
    """
    cfg = load_config(config_path)
    li = load_inputs(config_path)

    # нормалізуємо «Частота» в перехопленнях
    normalize_frequency_column(li.intercepts_df, li.reference_df)

    # групи та лічильники
    freqs, counts = unique_frequencies_with_counts(li.intercepts_df)
    allowed = (cfg.grouping or {}).get("allowed_tags", [])
    other   = (cfg.grouping or {}).get("other_bucket", "Інші радіомережі")
    groups = group_frequencies_by_tag(freqs, li.reference_df, allowed, other, cfg.grouping)

    # період з назви файла репорту
    period_start, period_end = _parse_period_from_filename(li.report_path)

    # створюємо документ
    doc = Document()

    # 1) Перша сторінка-огляд
    _render_overview_page(doc, cfg, li, groups, counts, period_start, period_end)

    # 2) Детальні розділи по частотах
  # 1) Зібрати частоти, які МАЮТЬ коментовані перехоплення (зберігаємо порядок груп)
    pub_freqs: list[tuple[str, str]] = []  # (short_tag, freq4)
    for short_tag, flist in groups.items():
        for f in flist:
            if not _network_is_empty(li.intercepts_df, f):
                pub_freqs.append((short_tag, f))
            else:
                log.info("Секцію %s пропущено — немає перехоплень із коментарями.", f)

    # 2) Рендер секцій з розривом сторінки МІЖ ними
    for idx, (short_tag, f) in enumerate(pub_freqs, start=1):
        _render_frequency_section(doc, f, counts.get(f, 0), li, cfg)
        if idx < len(pub_freqs):
            doc.add_page_break()
                
    # після останньої секції — примітка + підпис (без page break)
    _append_executor_block(doc)

    # збереження


    def _fmt_fname_dt(dt_obj) -> str:
        """
        Приймає або datetime, або рядок дати/часу і повертає
        формат для імені файлу: 'ДД.ММ.РРРР_ГГ-ММ'.
        Підтримує кілька поширених форматів рядка.
        """
        if isinstance(dt_obj, datetime):
            return dt_obj.strftime("%d.%m.%Y_%H-%M")

        if isinstance(dt_obj, str):
            for fmt in ("%Y-%m-%dT%H-%M",     # 2025-10-02T12-00   (з назви report_*.xlsx)
                        "%d.%m.%Y %H:%M",     # 02.10.2025 12:00
                        "%Y-%m-%d %H:%M:%S",  # 2025-10-02 12:00:00
                        ):
                try:
                    dt = datetime.strptime(dt_obj.strip(), fmt)
                    return dt.strftime("%d.%m.%Y_%H-%M")
                except ValueError:
                    continue
            # якщо не розпізнали — повернемо як є (краще так, ніж падати)
            return dt_obj

        # на всяк випадок — явно конвертуємо
        return str(dt_obj)


    # ...
    start_s = _fmt_fname_dt(period_start)
    end_s   = _fmt_fname_dt(period_end)
    file_name = f"Звіт РЕР ({start_s} - {end_s}).docx"


    out_dir = Path(getattr(cfg.paths, "output_dir", "build"))
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / file_name

    saved = safe_save_docx(doc, out_path)  # як і раніше використовуємо безпечне збереження
    return str(saved.resolve())



