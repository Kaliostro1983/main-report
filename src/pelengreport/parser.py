# src/pelengreport/parser.py
import re

HDR = re.compile(
    r"Пеленг\s+РЕР_63:\s*(?P<val>\d+(?:\.\d+)?)\s*/\s*(?P<date>\d{2}\.\d{2}\.\d{4})\s+(?P<time>\d{1,2}[:.]\d{2})",
    re.IGNORECASE,
)
SPACE_RE = re.compile(r"\s+")
MGRS_LAST_TWO_5 = re.compile(r"^\d{5}$")

def norm_time(t: str) -> str:
    t = t.strip().replace(".", ":")
    if len(t) == 4 and t[1] == ":":
        t = "0" + t
    if len(t) == 5:
        t = t + ":00"
    return t

def sanitize_mgrs(line: str) -> str:
    s = SPACE_RE.sub(" ", (line or "").strip())
    parts = s.split(" ")
    if len(parts) < 4:
        raise ValueError("Недостатньо токенів")
    t0, t1, d1, d2 = parts[0].upper(), parts[1].upper(), parts[-2], parts[-1]
    if not (MGRS_LAST_TWO_5.match(d1) and MGRS_LAST_TWO_5.match(d2)):
        raise ValueError("Цифрові блоки мають бути по 5 цифр")
    return f"{t0} {t1} {d1} {d2}"

def parse_whatsapp_text(lines: list[str]):
    """Один заголовок → 1..N MGRS; кожна координата окремим записом."""
    def looks_like_header(s: str) -> bool: return bool(HDR.search(s))
    def looks_like_mgrs(s: str) -> bool:
        try:
            sanitize_mgrs(s); return True
        except Exception:
            return False

    i, n = 0, len(lines)
    while i < n:
        m = HDR.search(lines[i])
        if not m:
            i += 1
            continue

        freq_or_mask = m.group("val")
        date_s = m.group("date")
        time_s = norm_time(m.group("time"))
        dt_s = f"{date_s} {time_s}"

        # опис на наступному рядку (+ можливе склеювання ще одного)
        i += 1
        if i >= n: break
        desc = lines[i].strip()
        j = i + 1
        if j < n and not looks_like_header(lines[j]) and not looks_like_mgrs(lines[j]):
            desc = SPACE_RE.sub(" ", (desc + " " + lines[j].strip())).strip()
            i = j

        # збираємо 1..N MGRS
        coords = []
        k = i + 1
        while k < n and not looks_like_header(lines[k]):
            s = lines[k].strip()
            if not s:
                k += 1
                continue
            try:
                coords.append(sanitize_mgrs(s))
            except Exception:
                break
            k += 1

        for mgrs in coords:
            yield {"freq_or_mask": freq_or_mask, "unit_desc": desc, "dt": dt_s, "mgrs": mgrs}
        i = k

__all__ = ["parse_whatsapp_text", "sanitize_mgrs", "norm_time"]
