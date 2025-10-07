from pathlib import Path
import pandas as pd
from .settings import load_config
from .io_utils import find_latest, read_excel
from .normalize import normalize_intercepts
from .enrich import map_reference, join_by_frequency, callsigns_set, activity_window, group_bucket

def run_pipeline_with_config(config_path: str = "config.yml") -> dict:
    cfg = load_config(config_path)
    paths = cfg.paths
    cols = {"reference": cfg.columns.reference, "intercepts": cfg.columns.intercepts}

    # 1) Довідник
    ref_raw = read_excel(paths.freq_file)
    ref = map_reference(ref_raw, cols)

    # 2) Звіт перехоплень (останній)
    latest = find_latest(paths.reports_dir, paths.report_mask)
    rpt_raw = read_excel(latest)
    intr = normalize_intercepts(rpt_raw, cols)

    # 3) Збагачення
    df = join_by_frequency(intr, ref)

    # 4) Агрегати по мережах
    out = []
    other_bucket = (cfg.grouping or {}).get("other_bucket", "Інші радіомережі")
    allowed = (cfg.grouping or {}).get("allowed_tags", [])
    for freq, g in df.groupby("frequency", dropna=False):
        calls = callsigns_set(g)
        window = activity_window(g)
        tag = g["tags"].dropna().iloc[0] if "tags" in g.columns and not g["tags"].dropna().empty else None
        bucket = group_bucket(tag, allowed, other_bucket)
        out.append({
            "frequency": freq,
            "name": g["name"].dropna().iloc[0] if "name" in g.columns and not g["name"].dropna().empty else "",
            "purpose": g["purpose"].dropna().iloc[0] if "purpose" in g.columns and not g["purpose"].dropna().empty else "",
            "nodes": g["nodes"].dropna().iloc[0] if "nodes" in g.columns and not g["nodes"].dropna().empty else "",
            "bucket": bucket,
            "callsigns": calls,
            "activity": window,
            "messages": g[["datetime","message","comment"]].sort_values("datetime", na_position="last").to_dict("records")
        })

    # 5) Експорт (зараз — як було; далі під твою верстку DOCX/PDF)
    out_dir = Path(paths.output_dir); out_dir.mkdir(parents=True, exist_ok=True)
    # TODO: виклик word/pdf експортерів

    return {"latest_report": str(latest), "items": len(out), "output_dir": str(out_dir)}
