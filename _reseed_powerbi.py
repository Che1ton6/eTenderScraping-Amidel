"""One-time script: rebuilds eTender_PowerBI_Data.xlsx in daily-breakdown format from all batch folders."""
import os
import re
import sys
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(__file__))
from BatchProcessor import SOURCES, calculate_counts

DATA_DIR      = os.path.join(os.path.dirname(__file__), "data")
POWER_BI_FILE = r"C:\eTenderData\eTender_PowerBI_Data.xlsx"
PBI_COLUMNS   = ["BATCH_LABEL", "BATCH_TYPE", "REPORT_DATE", "SOURCE", "NNT", "ICT", "RFQ"]
HDR_FONT      = Font(bold=True, color="FFFFFF")
HDR_FILL      = PatternFill("solid", fgColor="1F4E79")


def parse_folder(name):
    m = re.match(r"\(([MT])\)\s+(\d+)-(\d+)\s+(\w+)\s+(\d{4})", name)
    if not m:
        return None
    btype, d1, d2, month_name, year = m.groups()
    end   = datetime.strptime(f"{d2} {month_name} {year}", "%d %B %Y")
    start = end.replace(day=int(d1))
    return btype, start, end


all_rows = []

for folder_name in sorted(os.listdir(DATA_DIR)):
    folder_path = os.path.join(DATA_DIR, folder_name)
    if not os.path.isdir(folder_path):
        continue
    parsed = parse_folder(folder_name)
    if not parsed:
        print(f"Skipping: {folder_name}")
        continue

    btype, start, end = parsed
    if start.month == end.month:
        batch_label = f"({btype}) {start.day}-{end.day} {end.strftime('%b %y')}"
    else:
        batch_label = f"({btype}) {start.strftime('%d %b')}-{end.strftime('%d %b %y')}"

    batches_dir = os.path.join(folder_path, "batches")
    current = start
    while current <= end:
        day_str  = current.strftime("%Y_%m_%d")
        day_file = os.path.join(batches_dir, f"tenders_{day_str}.xlsx")
        if os.path.exists(day_file):
            day_df     = pd.read_excel(day_file)
            day_counts = calculate_counts(day_df)
        else:
            day_counts = {src: {"NNT": 0, "ICT": 0, "RFQ": 0} for src, _ in SOURCES}

        for source, _ in SOURCES:
            data = day_counts.get(source, {"NNT": 0, "ICT": 0, "RFQ": 0})
            all_rows.append({
                "BATCH_LABEL": batch_label,
                "BATCH_TYPE":  btype,
                "REPORT_DATE": current,
                "SOURCE":      source,
                "NNT":         data["NNT"],
                "ICT":         data["ICT"],
                "RFQ":         data["RFQ"],
            })
        current += timedelta(days=1)

    print(f"  {batch_label}: {(end - start).days + 1} days")

combined = pd.DataFrame(all_rows, columns=PBI_COLUMNS)
combined["REPORT_DATE"] = pd.to_datetime(combined["REPORT_DATE"])
combined.sort_values("REPORT_DATE", ascending=False, inplace=True)
combined.reset_index(drop=True, inplace=True)

wb = Workbook()
ws = wb.active
ws.title = "Tender Data"
for col_idx, col_name in enumerate(PBI_COLUMNS, 1):
    cell = ws.cell(row=1, column=col_idx, value=col_name)
    cell.font = HDR_FONT
    cell.fill = HDR_FILL
for row_idx, row in enumerate(combined.itertuples(index=False), 2):
    for col_idx, col_name in enumerate(PBI_COLUMNS, 1):
        value = getattr(row, col_name)
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        if col_name == "REPORT_DATE" and hasattr(value, "strftime"):
            cell.number_format = "YYYY/MM/DD"

last_col = get_column_letter(len(PBI_COLUMNS))
last_row = len(combined) + 1
tbl = Table(displayName="TenderData", ref=f"A1:{last_col}{last_row}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
wb.save(POWER_BI_FILE)
print(f"\nDone — {len(combined)} rows written to {POWER_BI_FILE}")
