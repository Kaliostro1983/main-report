from __future__ import annotations
import re
import hashlib
from dataclasses import dataclass
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from typing import List, Iterator, Tuple

KYIV_TZ = ZoneInfo("Europe/Kyiv")

DATE_LINE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}, \d{2}:\d{2}:\d{2}$")

@dataclass
class Block:
    date: str         # 'YYYY-MM-DD'
    time: str         # 'HH:MM:SS'
    ts_utc: int       # epoch seconds
    freq_mhz: float
    radionet: str
    who: str
    komu: str
    body_full: str
    src_hash: str     # sha1(raw block text)

class ParseError(Exception):
    pass

def _parse_datetime(datestr: str) -> Tuple[str, str, int]:
    # datestr: 'DD.MM.YYYY, HH:MM:SS'
    dt = datetime.strptime(datestr, "%d.%m.%Y, %H:%M:%S").replace(tzinfo=KYIV_TZ)
    ts_utc = int(dt.astimezone(timezone.utc).timestamp())
    return dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M:%S"), ts_utc

def _is_new_block_line(line: str) -> bool:
    return bool(DATE_LINE_RE.match(line.strip()))

def iter_blocks(lines: List[str]) -> Iterator[Block]:
    i, n = 0, len(lines)
    while i < n:
        line = lines[i].rstrip("\n")
        if not _is_new_block_line(line):
            i += 1
            continue

        raw_parts = [line]  # для src_hash
        # 1) дата+час
        date_str, time_str, ts_utc = _parse_datetime(line.strip())
        i += 1
        if i >= n: break

        # 2) частота
        freq_line = lines[i].rstrip("\n").strip()
        raw_parts.append(freq_line)
        i += 1
        try:
            freq_mhz = float(freq_line.replace(" ", ""))
        except ValueError as e:
            raise ParseError(f"Bad freq line: {freq_line}") from e
        if not (0.001 <= freq_mhz <= 4000.0):
            raise ParseError(f"Out-of-range freq: {freq_mhz}")

        # 3) радіомережа
        if i >= n: break
        radionet = lines[i].rstrip("\n").strip()
        raw_parts.append(radionet)
        i += 1

        # 4) хто
        if i >= n: break
        who = lines[i].rstrip("\n").strip()
        raw_parts.append(who)
        i += 1
        if not who:
            raise ParseError("Empty 'who'")

        # 5) кому
        if i >= n: break
        komu = lines[i].rstrip("\n").strip()
        raw_parts.append(komu)
        i += 1
        if not komu:
            raise ParseError("Empty 'komu'")

        # 6) додаткові адресати (ігноруємо) — до першого порожнього рядка
        while i < n:
            peek = lines[i].rstrip("\n")
            if peek.strip() == "":
                raw_parts.append(peek)
                i += 1
                break
            # якщо одразу почався текст перехоплення з тире — теж вважаємо кінець адресатів
            if peek.startswith("—") or peek.startswith("- "):
                # не споживаємо порожній рядок; одразу підемо в body
                break
            # якщо випадково натрапили на новий блок — це кінець поточного (тіла немає)
            if _is_new_block_line(peek.strip()):
                break
            # інакше — це ще один адресат; ігноруємо, просто споживаємо рядок
            raw_parts.append(peek)
            i += 1

        # 7) тіло перехоплення до наступної дати-часу або EOF
        body_lines: List[str] = []
        while i < n:
            peek = lines[i].rstrip("\n")
            if _is_new_block_line(peek.strip()):
                # не споживаємо; вийдемо, щоб головний цикл побачив новий блок
                break
            body_lines.append(peek)
            raw_parts.append(peek)
            i += 1

        body_full = "\n".join(body_lines).rstrip("\n")
        # src_hash з цілого сирого блоку
        src_hash = hashlib.sha1(("\n".join(raw_parts)).encode("utf-8")).hexdigest()

        # Невеликі обрізання за ТЗ
        radionet = radionet.strip()
        if len(radionet) > 500:
            radionet = radionet[:500]
        who = who.strip()
        komu = komu.strip()
        if len(body_full) > 50000:
            body_full = body_full[:50000]

        yield Block(
            date=date_str,
            time=time_str,
            ts_utc=ts_utc,
            freq_mhz=freq_mhz,
            radionet=radionet,
            who=who,
            komu=komu,
            body_full=body_full,
            src_hash=src_hash,
        )
