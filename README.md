# Report Project (Python)

Заготовка для проєкту, який обробляє дані з кількох таблиць (CSV/Excel) і формує звіт у форматах DOCX та PDF.

## Структура
```
report-project/
├─ src/reportgen/
│  ├─ __init__.py
│  ├─ data_loader.py
│  ├─ word_report.py
│  ├─ pdf_report.py
│  └─ report_pipeline.py
├─ data/input/
│  ├─ sample1.csv
│  └─ sample2.csv
├─ tests/
│  └─ test_sanity.py
├─ main.py
├─ requirements.txt
└─ .gitignore
```

## Швидкий старт
```bash
python -m venv .venv
# Windows PowerShell:
.\.venv\Scripts\Activate.ps1
# або CMD:
.\.venv\Scripts\activate.bat
# або Linux/Mac:
# source .venv/bin/activate

pip install --upgrade pip
pip install -r requirements.txt

python main.py --inputs data/input/*.csv --out-dir build
```

DOCX зберігається як `build/report.docx`, PDF — `build/report.pdf`.

## Примітки
- DOCX генерується через `python-docx`.
- PDF генерується напряму через `reportlab` (не потребує MS Word).
- В майбутньому можна додати конвертацію DOCX→PDF через `docx2pdf` на Windows з MS Word.
