from __future__ import annotations
import argparse
import hashlib
import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple

from .db import ensure_db, bulk_insert, DEFAULT_DB_PATH
from .parser import iter_blocks, ParseError

CHAT_ID_DEFAULT = "Ocheret"

@dataclass
class ImportStats:
    total: int
    inserted: int
    skipped: int
    warnings: List[str]

def _row_id(chat_id: str, date: str, time_s: str, freq: float, who: str, komu: str, body_full: str) -> Tuple[str, str]:
    body_hash = hashlib.sha1(body_full.encode("utf-8")).hexdigest()
    rid = hashlib.sha1(f"{chat_id}|{date}|{time_s}|{freq:.6f}|{who}|{komu}|{body_hash}".encode("utf-8")).hexdigest()
    return rid, body_hash

def import_whatsapp_blocks(path: Path, chat_id: str = CHAT_ID_DEFAULT, db_path: Path = DEFAULT_DB_PATH) -> ImportStats:
    text = path.read_text(encoding="utf-8", errors="ignore")
    lines = text.splitlines()
    total = inserted = skipped = 0
    warnings: List[str] = []

    with ensure_db(db_path) as con:
        to_insert = []
        now_utc = int(time.time())
        i = 0
        while i < len(lines):
            try:
                for blk in iter_blocks(lines[i:]):
                    # Щоб не зациклити, потрібно знайти, скільки рядків реально спожито в підблоці.
                    # Простий прийом: реконструюємо сигнатуру початку наступного блоку.
                    # Але тут зробимо інакше: одразу зупинимо зовнішній while і перейдемо на for по всьому файлу.
                    pass
            except Exception:
                break
            break  # захисний break, реальний імпорт нижче

        # Простіше: один прохід по всьому файлу генератором
        for blk in iter_blocks(lines):
            total += 1
            try:
                rid, src_body_hash = _row_id(chat_id, blk.date, blk.time, blk.freq_mhz, blk.who, blk.komu, blk.body_full)
                to_insert.append((
                    rid, chat_id, blk.date, blk.time, blk.ts_utc, blk.freq_mhz, blk.radionet,
                    blk.who, blk.komu, blk.body_full, now_utc, blk.src_hash
                ))
            except ParseError as e:
                skipped += 1
                warnings.append(str(e))
            except Exception as e:
                skipped += 1
                warnings.append(f"Unexpected: {e}")

        if to_insert:
            before = con.total_changes
            bulk_insert(con, to_insert)
            con.commit()
            inserted = con.total_changes - before

    return ImportStats(total=total, inserted=inserted, skipped=total - inserted, warnings=warnings)

def main():
    p = argparse.ArgumentParser(description="Import WhatsApp export into SQLite (test_intercepts)")
    p.add_argument("--file", required=True, help="path to exported .txt")
    p.add_argument("--chat", default=CHAT_ID_DEFAULT, help="chat_id (default: Ocheret)")
    p.add_argument("--db", default=str(DEFAULT_DB_PATH), help="SQLite path (default: src/data/data.db)")
    args = p.parse_args()

    stats = import_whatsapp_blocks(Path(args.file), chat_id=args.chat, db_path=Path(args.db))
    print(f"TOTAL: {stats.total} | INSERTED: {stats.inserted} | SKIPPED: {stats.skipped}")
    if stats.warnings:
        print("WARNINGS:")
        for w in stats.warnings[:20]:
            print(" -", w)
        if len(stats.warnings) > 20:
            print(f" ... and {len(stats.warnings)-20} more")

if __name__ == "__main__":
    main()
