from __future__ import annotations
import argparse
import logging
from pathlib import Path  # ⚡
from glob import glob
from src.activefrequencies.report import build_active_frequencies_docx
from src.armorkit.data_loader import load_inputs
from src.armorkit.normalize_freq import normalize_frequency_column
# from src.reportgen.export_xlsx import save_df_xlsx

from src.reportgen.grouping import unique_frequencies_with_counts, group_frequencies_by_tag
from src.armorkit.settings import load_config
from src.reportgen.export.word_report import build_draft_docx

def parse_args():
    ap = argparse.ArgumentParser(description="Report generator")
    
    ap.add_argument("--inputs", nargs="+", required=False, default=["data/input/*.csv"],
                    help="Список шляхів до CSV/XLSX (можна з масками)")
    
    ap.add_argument("--out-dir", default="build", help="Куди зберегти звіти")
      
    return ap.parse_args()
    

def expand_inputs(patterns):
    files = []
    for p in patterns:
        files.extend(glob(p))
    return sorted(set(files))

def main():
    ap = argparse.ArgumentParser(description="Report generator")
    ap.add_argument("--config", default="config.yml", help="Шлях до YAML-конфіга")
    ap.add_argument(
        "--mode",
        choices=["read", "normalize", "freq-groups", "draft-docx", "run", "active-freqs", "peleng-gui", "artyleria-report", "eralonky", "enemies", "simple-report", "still-alive", "move_enemies", "freq-lists", "automizer"],
        default="read",
        help="read=зчитати; normalize=нормалізувати 'Частота' і зберегти XLSX; "
             "freq-groups=вивести групи частот; draft-docx=згенерувати DOCX-чернетку; run=повний конвеєр; active-freqs=звіт 'Активні мережі'",
    )
    ap.add_argument("--log-level", default="INFO", help="DEBUG, INFO, WARNING, ERROR")
    args = ap.parse_args()

    logging.basicConfig(level=getattr(logging, args.log_level.upper(), logging.INFO),
                        format="%(levelname)s: %(message)s")

    if args.mode == "read":
        li = load_inputs(args.config)
        print("CONFIG :", li.cfg_path)
        print("FREQ   :", li.freq_path, "| shape:", li.reference_df.shape)
        print("REPORT :", li.report_path, "| shape:", li.intercepts_df.shape)
        print("Reference columns:", list(li.reference_df.columns)[:12])
        print("Intercepts columns:", list(li.intercepts_df.columns)[:12])
        return

    # if args.mode == "normalize":
    #     li = load_inputs(args.config)
    #     normalize_frequency_column(li.intercepts_df, li.reference_df)
    #     out_path = "build/normalized_intercepts.xlsx"
    #     save_df_xlsx(li.intercepts_df, out_path)
    #     print(f"OK: normalized intercepts saved to {out_path}")
    #     return

    if args.mode == "freq-groups":
        li = load_inputs(args.config)
        cfg = load_config(args.config)
        normalize_frequency_column(li.intercepts_df, li.reference_df)
        freqs, counts = unique_frequencies_with_counts(li.intercepts_df)
        allowed = (cfg.grouping or {}).get("allowed_tags", [])
        other   = (cfg.grouping or {}).get("other_bucket", "Інші радіомережі")
        groups = group_frequencies_by_tag(freqs, li.reference_df, allowed, other, cfg.grouping)
        print("\n=== ГРУПИ РАДІОМЕРЕЖ ===")
        for bucket, items in groups.items():
            print(f"\n[{bucket}]  ({len(items)})")
            for f in items:
                print(f"  - {f} ({counts.get(f,0)})")
        return

    if args.mode == "draft-docx":
        path = build_draft_docx(args.config)
        print(f"OK: DOCX збережено → {path}")
        return

    if args.mode == "run":
        print("Full pipeline will be implemented next.")
        return
    
    elif args.mode == "active-freqs":
        path = build_active_frequencies_docx(args.config)
        print(f"OK: DOCX збережено → {path}")
        
        
    elif args.mode == "peleng-gui":
        # ВАЖЛИВО: назва пакета має бути без дефіса!
        # Тобто директорія модуля повинна називатися src/peleng_gen/, не src/peleng-gen/
        from src.pelenggen.gui import main as peleng_gui_main
        # Якщо GUI сам читає config.yml — достатньо просто викликати:
        peleng_gui_main()
        return
    
    elif args.mode == "artyleria-report":
        from src.artyleria.runner import run as arty_run
        arty_run()
        return
    
    elif args.mode == "eralonky":
        from src.etalonky.runner import run as eralonky_run
        eralonky_run()
        return
    
    elif args.mode == "enemies":
        from src.enemies.generate_enemies_report import main as enemies_main
        enemies_main()
        return
    
    elif args.mode == "simple-report":
        from src.simplereport.generate_simple_report import build_simple_report_docx
        path = build_simple_report_docx(args.config)
        print(f"OK: DOCX збережено → {path}")
        return
    
    elif args.mode == "still-alive":
        from src.stillalive.generate_stillalive_report import build_stillalive_report
        path = build_stillalive_report(args.config)
        print(f"[OK] StillAlive збережено: {path}")
        
    elif args.mode == "automizer":
        # runner може приймати config або ні — зробимо сумісно
        # на початку файлу поруч з іншими імпортами
        from src.automizer.runner import main as automizer_main

        try:
            automizer_main(args.config)
        except TypeError:
            automizer_main()
        return
    
    elif args.mode == "move_enemies":
        from src.movecallsigns.who_move import create_move_report
        path = create_move_report(args.config)
        print(f"OK: DOCX збережено → {path}")
        return
    
    
    elif args.mode == "freq-lists":
        from src.freqexport.generate_lists import build_freq_lists
        paths = build_freq_lists(args.config)
        print("[OK] Збережено списки частот:")
        for k, v in paths.items():
            print(f"  {k}: {v}")






if __name__ == "__main__":
    main()
