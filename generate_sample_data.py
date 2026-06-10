"""
Sample Data Generator
=====================
Creates realistic dummy input files for the transaction report parser:

1. sample_data/ALPHA_<date>.raw   - pipe-delimited report (Network Alpha)
2. sample_data/BETA_<yymmdd>.txt  - fixed-width report (Network Beta)
3. sample_data/terminal_lookup.xlsx - terminal prefix -> provider mapping

All data is randomly generated. No real transaction data is used.

Usage:
    python generate_sample_data.py [--days 3] [--rows 200]
"""

import argparse
import random
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

SAMPLE_DIR = Path(__file__).parent / "sample_data"

# Fictional payment providers and their 4-digit terminal prefixes
PROVIDERS = {
    "PT MAJU PAYMENT": ["1001", "1002", "1003"],
    "WARUNG DIGITAL": ["2201", "2202"],
    "AGEN PINTAR": ["3305", "3306", "3307"],
    "KIOSK NUSANTARA": ["4410"],
    "TOKO SEJAHTERA": ["5512", "5513"],
}

# A couple of prefixes intentionally NOT in the lookup file,
# so the parser's "unmatched" handling can be demonstrated.
UNKNOWN_PREFIXES = ["9901", "9902"]

RESPONSE_CODES = ["00", "00", "00", "00", "51", "51", "05", "91"]  # weighted


def random_terminal_id() -> str:
    all_prefixes = [p for plist in PROVIDERS.values() for p in plist]
    prefix = random.choice(all_prefixes * 9 + UNKNOWN_PREFIXES)  # ~5% unknown
    return prefix + f"{random.randint(0, 9999):04d}"


def random_card() -> str:
    return f"4{random.randint(10**14, 10**15 - 1)}"[:8] + "********"


def write_alpha_file(date: datetime, rows: int) -> Path:
    """Pipe-delimited format: one header record (H) then data records (D)."""
    path = SAMPLE_DIR / f"ALPHA_{date:%Y%m%d}.raw"
    lines = ["H|ALPHA-DAILY|" + date.strftime("%Y%m%d")]
    for _ in range(rows):
        t = f"{random.randint(6, 22):02d}{random.randint(0, 59):02d}{random.randint(0, 59):02d}"
        lines.append("|".join([
            "D",
            "ALPHA-DAILY",
            date.strftime("%Y%m%d"),                 # report date
            t,                                        # time HHMMSS
            random_terminal_id(),                     # terminal id (8 chars)
            random_card(),                            # masked card
            str(random.choice([20, 50, 100, 200, 500]) * 1000),  # amount
            random.choice(RESPONSE_CODES),            # response code
            f"{random.randint(1, 999999):06d}",       # trace number
            str(random.choice([0, 1500, 2500])),      # network fee
        ]))
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def write_beta_file(date: datetime, rows: int) -> Path:
    """
    Fixed-width format. Layout (1-based positions):
        1-6    date (YYMMDD)
        7-12   time (HHMMSS)
        13-20  terminal id
        21-36  masked card (16)
        37-49  amount, right-aligned (13)
        50-51  response code
        52-57  trace number
    """
    path = SAMPLE_DIR / f"BETA_{date:%y%m%d}.txt"
    lines = ["RPT-BETA DAILY TRANSACTION FILE"]  # header line, skipped by parser
    for _ in range(rows):
        t = f"{random.randint(6, 22):02d}{random.randint(0, 59):02d}{random.randint(0, 59):02d}"
        amount = random.choice([20, 50, 100, 150, 300]) * 1000
        lines.append(
            date.strftime("%y%m%d")
            + t
            + random_terminal_id()
            + random_card().ljust(16)
            + str(amount).rjust(13)
            + random.choice(RESPONSE_CODES)
            + f"{random.randint(1, 999999):06d}"
        )
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def write_lookup_file() -> Path:
    path = SAMPLE_DIR / "terminal_lookup.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Lookup"
    ws.append(["PROVIDER", "TERMINAL PREFIX"])
    for provider, prefixes in PROVIDERS.items():
        for i, prefix in enumerate(prefixes):
            # Provider name only on the first row of each group,
            # to simulate a merged-cell style source file
            ws.append([provider if i == 0 else None, prefix])
    wb.save(path)
    return path


def main() -> None:
    ap = argparse.ArgumentParser(description="Generate dummy transaction files")
    ap.add_argument("--days", type=int, default=3, help="number of report days")
    ap.add_argument("--rows", type=int, default=200, help="rows per file")
    args = ap.parse_args()

    SAMPLE_DIR.mkdir(exist_ok=True)
    random.seed(42)  # reproducible output

    print(f"Generating {args.days} day(s) of sample data in {SAMPLE_DIR}/ ...")
    base = datetime.now() - timedelta(days=args.days)
    for d in range(args.days):
        date = base + timedelta(days=d + 1)
        print(f"  {write_alpha_file(date, args.rows).name}")
        print(f"  {write_beta_file(date, args.rows).name}")
    print(f"  {write_lookup_file().name}")
    print("Done.")


if __name__ == "__main__":
    main()
