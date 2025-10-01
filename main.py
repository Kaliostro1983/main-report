from __future__ import annotations
import argparse
from glob import glob
from src.reportgen.report_pipeline import run_pipeline

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
    args = parse_args()
    files = expand_inputs(args.inputs)
    if not files:
        raise SystemExit("Не знайдено жодного вхідного файлу (перевірте --inputs)")
    result = run_pipeline(files, out_dir=args.out_dir)
    print(f"OK: {result}")

if __name__ == "__main__":
    main()
