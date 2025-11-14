from __future__ import annotations
import sqlite3
from pathlib import Path
from typing import Iterable

# Шлях до БД за ТЗ: src/data/data.db
DEFAULT_DB_PATH = (Path(__file__).resolve().parents[1] / "data" / "data.db")

DDL = """
PRAGMA journal_mode=WAL;
CREATE TABLE IF NOT EXISTS test_intercepts (
  id               TEXT      PRIMARY KEY,
  chat_id          TEXT      NOT NULL,
  date             TEXT      NOT NULL,   -- 'YYYY-MM-DD'
  time             TEXT      NOT NULL,   -- 'HH:MM:SS'
  ts_utc           INTEGER   NOT NULL,   -- epoch seconds
  freq_mhz         REAL      NOT NULL,
  radionet         TEXT      NOT NULL,
  who              TEXT      NOT NULL,
  komu             TEXT      NOT NULL,
  body_full        TEXT      NOT NULL,
  ingested_at_utc  INTEGER   NOT NULL,
  src_hash         TEXT      NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_ti_ts        ON test_intercepts(ts_utc);
CREATE INDEX IF NOT EXISTS idx_ti_freq      ON test_intercepts(freq_mhz);
CREATE INDEX IF NOT EXISTS idx_ti_who_komu  ON test_intercepts(who, komu);
"""

def ensure_db(db_path: Path = DEFAULT_DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(db_path))
    con.execute("PRAGMA foreign_keys=ON;")
    con.executescript(DDL)
    return con

def bulk_insert(con: sqlite3.Connection, rows: Iterable[tuple]):
    con.executemany(
        """INSERT OR IGNORE INTO test_intercepts
           (id, chat_id, date, time, ts_utc, freq_mhz, radionet, who, komu,
            body_full, ingested_at_utc, src_hash)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        rows,
    )
