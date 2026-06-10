#!/usr/bin/env python3
"""
Test BatchProcessor against the existing 14-17 May end product file.
Compares calculated counts with the known Final tab values, then
creates a test end product file to verify Excel output.

Run with:  python test_batch.py
"""

import os
import shutil
import pandas as pd
from datetime import datetime

from BatchProcessor import calculate_counts, create_end_product, create_batch_folder, SOURCES

KNOWN_FILE = (
    r"C:\Users\CheltonGraham\OneDrive - Amidel (Pty) Ltd"
    r"\Documents\Sales\Sales Auto Hub\Scraping and Reports"
    r"\ICT & RFQ\End Products\May"
    r"\RFQ_and_ICT_Checker_(14 - 17 May).xlsx"
)
TEST_OUTPUT = r"C:\Users\CheltonGraham\Desktop\eTenderScraping\data\TEST_end_product.xlsx"


def load_known_final(path):
    """Load the Final tab from an existing end product file as a dict of {source: (NNT, ICT, RFQ)}."""
    df = pd.read_excel(path, sheet_name="Final", header=None)
    expected = {}
    for _, row in df.iterrows():
        source = row.iloc[0]
        if isinstance(source, str) and source not in ("Source", "Report Date"):
            try:
                nnt = int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
                ict = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
                rfq = int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
                expected[source] = (nnt, ict, rfq)
            except (ValueError, TypeError):
                pass
    return expected


def main():
    print("=" * 60)
    print("BatchProcessor Test")
    print("=" * 60)
    print()

    # ── Step 1: Load existing tender data ────────────────────────────
    print(f"Loading tender data from:\n  {KNOWN_FILE}\n")
    df = pd.read_excel(KNOWN_FILE, sheet_name="Tender Data")
    print(f"  Rows loaded: {len(df)}")
    print()

    # ── Step 2: Calculate counts ──────────────────────────────────────
    print("Calculating counts...")
    counts = calculate_counts(df)

    # ── Step 3: Load known-good Final tab values ──────────────────────
    expected = load_known_final(KNOWN_FILE)

    # ── Step 4: Compare ───────────────────────────────────────────────
    print()
    print(f"{'Source':<25} {'Expected':>20}   {'Got':>20}   {'Match'}")
    print("-" * 75)

    all_pass = True
    for source, _ in SOURCES:
        if source not in expected:
            continue
        exp_nnt, exp_ict, exp_rfq = expected[source]
        got = counts.get(source, {"NNT": 0, "ICT": 0, "RFQ": 0})
        got_nnt, got_ict, got_rfq = got["NNT"], got["ICT"], got["RFQ"]

        match = (exp_nnt == got_nnt and exp_ict == got_ict and exp_rfq == got_rfq)
        if not match:
            all_pass = False

        exp_str = f"NNT={exp_nnt} ICT={exp_ict} RFQ={exp_rfq}"
        got_str = f"NNT={got_nnt} ICT={got_ict} RFQ={got_rfq}"
        flag    = "OK" if match else "FAIL <<<"
        print(f"{source:<25} {exp_str:>20}   {got_str:>20}   {flag}")

    print()
    if all_pass:
        print("All counts match.")
    else:
        print("Some counts do not match — review FAIL rows above.")

    # ── Step 5: Create test end product file ──────────────────────────
    print()
    print("Creating test end product file...")
    batch_folder = create_batch_folder("2026-05-14", "2026-05-17", "M")

    path = create_end_product(
        df,
        date_from="2026-05-14",
        date_to="2026-05-17",
        batch_type="M",
        report_date=datetime(2026, 5, 17),
        batch_folder=batch_folder,
    )
    # Copy to fixed test output path for easy inspection
    os.makedirs(os.path.dirname(TEST_OUTPUT), exist_ok=True)
    shutil.copy(path, TEST_OUTPUT)
    print(f"  Written to: {TEST_OUTPUT}")
    print("  Open it and verify the Final tab shows the same numbers as above.")

    print()
    print("Test complete.")
    print("=" * 60)


if __name__ == "__main__":
    main()
