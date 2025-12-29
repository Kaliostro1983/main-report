# src/pelengreport/runner.py
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import sys
import logging

# імпорти відносно пакета
from .parser import parse_whatsapp_text
from .report import build_docx

def _repo_root() -> Path:
    # .../src/pelengreport/runner.py -> .../src -> .../
    return Path(__file__).resolve().parents[2]

def _next_free_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 2
    while True:
        cand = path.with_name(f"{stem} ({i}){suffix}")
        if not cand.exists():
            return cand
        i += 1

def _resolve_input_path(arg: str | None) -> Path:
    root = _repo_root()

    # якщо передано шлях аргументом — поважаємо
    if arg:
        p = Path(arg)
        if not p.is_absolute():
            # спробувати відносно CWD
            cand = Path.cwd() / p
            if cand.exists():
                return cand
            # спробувати відносно кореня репо
            cand = root / p
            if cand.exists():
                return cand
        if p.exists():
            return p
        raise FileNotFoundError(f"Не знайдено вхідний файл: {p}")

    # автопошук у стандартних каталогах
    candidates_dirs = [
        root / "pelengreport" / "data",            # якщо data лежить поруч з коренем
        root / "src" / "pelengreport" / "data",    # ✅ твій випадок
        Path.cwd(),                                 # поточна тека
    ]

    txt_files: list[Path] = []
    for d in candidates_dirs:
        if d.exists():
            txt_files += list(d.glob("*.txt")) + list(d.glob("*.TXT"))

    if not txt_files:
        raise FileNotFoundError(
            "Не знайдено *.txt у жодному з каталогів:\n"
            + "\n".join(f" - {d}" for d in candidates_dirs)
            + "\n\nПередай шлях явно, напр.:\n"
            "  python -m src.pelengreport.runner src\\pelengreport\\data\\peleng.txt"
        )

    txt_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return txt_files[0]

def run(input_txt: Path, out_dir: Path | None = None) -> Path:
    root = _repo_root()
    out_dir = Path(out_dir or (root / "build"))
    out_dir.mkdir(parents=True, exist_ok=True)

    today = datetime.now().strftime("%d.%m.%Y")
    out_path = _next_free_path(out_dir / f"форма_1.2.15 {today}.docx")

    print(f"[i] Using input: {input_txt}")
    with open(input_txt, "r", encoding="utf-8-sig") as f:   # <- utf-8-sig на випадок BOM з WhatsApp
        lines = f.readlines()
        
    records = list(parse_whatsapp_text(lines))  # ← парсимо WhatsApp у список записів
    logging.warning(
            f"Знайдено {len(records)} записів у вхідному файлі."
        )
    build_docx(records, out_path)               # ← генеруємо DOCX «як у старій версії»


    records = list(parse_whatsapp_text(lines))
    build_docx(records, out_path)
    
    print(f"[OK] Report saved to: {out_path}")
    return out_path

if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    inp = _resolve_input_path(arg)
    run(inp)
