#!/usr/bin/env python3
"""One-off: restore the full 283-tender Tender Data sheet in the (M) 16-19 July
eTenders batch, then regenerate the equation counts. Uses the Raw Data sheet
(which already has all 283) as the source of truth.
"""
import io, sys, os, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from BatchProcessor import (SOURCES, calculate_counts, create_end_product,
                             update_equation_file, TENDER_COLUMNS, BATCH_COLUMNS)

BATCH = os.path.join(HERE, 'data', 'etenders.gov.za', '(M) 16-19 July 2026')
EP = os.path.join(BATCH, 'end product', 'RFQ_and_ICT_Checker_(16 - 19 July).xlsx')

def _backup(p):
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    dst = f"{os.path.splitext(p)[0]}_pre_unfilter_backup_{ts}.xlsx"
    shutil.copy2(p, dst)
    print(f"  backup: {os.path.basename(dst)}")

# Read the full 283-row Raw Data
raw = pd.read_excel(EP, sheet_name='Raw Data')
print(f"Raw Data has {len(raw)} rows")

# Ensure DUPLICATED column exists (0 = single source, since this is eTenders-only)
if 'DUPLICATED' not in raw.columns:
    raw['DUPLICATED'] = 0

_backup(EP)

# Rebuild the end product using Raw Data as the "kept" set. Same rows go into
# Tender Data (full 283) and Raw Data (mirror).
date_from = '2026-07-16'
date_to = '2026-07-19'
report_date = datetime.strptime(date_to, '%Y-%m-%d')
batch_type = 'M'

# Convert timestamps back to string 'YYYY/MM/DD' format so the write helper
# formats them as dates
for col in ('PUBLICATION_DATE', 'CLOSING_DATE', 'BRIEFING_DATE', 'REPORT_DATE'):
    if col in raw.columns:
        raw[col] = raw[col].apply(lambda v: v.strftime('%Y/%m/%d') if hasattr(v, 'strftime') else (v if pd.notna(v) else ''))

new_path = create_end_product(raw, date_from, date_to, batch_type, report_date, BATCH,
                              raw_df=raw)
print(f"Rebuilt end product: {new_path}")

# Regenerate equation counts
counts = calculate_counts(raw)
update_equation_file(counts, batch_type, report_date, BATCH)
print("Equation updated with counts based on full 283 rows")

# Quick summary
watchlist_total = sum(v['NNT'] for k, v in counts.items() if k != 'eTenders (All)')
print(f"\nSummary:")
print(f"  eTenders (All) [total captured]: {len(raw)}")
print(f"  TOTAL WATCHLIST (sum of entity rows): {watchlist_total}")
print(f"  Gap (organs not on watchlist): {len(raw) - watchlist_total}")
