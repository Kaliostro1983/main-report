def freq4_str(x) -> str | None:
    try:
        v = float(str(x).replace(",", "."))
        return f"{v:.4f}"
    except Exception:
        return None