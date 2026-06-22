#!/usr/bin/env python3
"""
Parts 2 & 3: Creates the local batch folder structure, saves per-day Excel
files, builds the end product file (Tender Data + Final tabs), and appends
the new batch column to the equation file.
"""

import re
import os
import logging
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side

EQUATION_FILE = (
    r"C:\Users\CheltonGraham\OneDrive - Amidel (Pty) Ltd"
    r"\Documents\Sales\Sales Auto Hub\Scraping and Reports"
    r"\ICT & RFQ\Old Method\Bhekis conditional Product\RFQ_and_ICT_Equation.xlsx"
)

POWER_BI_FILE = os.path.join("data", "etenders.gov.za", "eTender_PowerBI_Data.xlsx")

# Exact department matching patterns taken from the Final tab Excel formulas.
# (label, dept_pattern) — None means "all rows" (eTenders All).
SOURCES = [
    ("eTenders (All)",        None),
    ("Eskom",                 "ESKOM"),
    ("SITA",                  "State Information Technology Agency"),
    ("Transnet",              "Transnet*"),
    ("JPC",                   "*Joburg Property*"),
    ("Matatiele LM",          "Matatiele Local Municipality"),
    ("Ntabankulu LM",         "Ntabankulu Local Municipality"),
    ("Umzimvubu LM",          "Umzimvubu Local Municipality"),
    ("Winnie Mandela LM",     "Winnie Madikizela*"),
    ("Mnquma LM",             "Mnquma Local Municipality"),
    ("Great Kei LM",          "Great Kei Local Municipality"),
    ("Amahlathi LM",          "Amahlathi Local Municipality"),
    ("Raymond Mhlaba LM",     "Raymond Mhlaba Local Municipality"),
    ("JOSHCO",                "*Johannesburg Social Housing*"),
    ("GPL",                   "Gauteng Provincial Legislature"),
    ("RAF",                   "Road Accident Fund"),
    ("MM Trading Company",    "*Metropolitan Trading*"),
    ("Nelson Mandela Bay MM", "*Nelson Mandela Bay*"),
    ("Buffalo City MM",       "*Buffalo City*"),
    ("EC DPW",                "*Eastern Cape*Public Works*"),
    ("DEL",                   "*Labour*"),
    ("SIU",                   "Special Investigation Unit*"),
    ("GDoH",                  "*Gauteng*Health*"),
    ("DIRCO",                 "*International Relation*"),
    ("DOJ",                   "*Justice*Constitutional*"),
    ("CP JHB",                "City Power*Johannesburg*"),
    ("W JHB",                 "*Johannesburg Water*"),
]

ICT_CATEGORIES = ["Information service activities", "Information and communication"]
RFQ_TYPE       = "Request for Quotation"

TENDER_COLUMNS = [
    "REPORT_DATE", "RECORD_ID", "TENDER_ID", "PUBLICATION_DATE",
    "CLOSING_DATE", "CLOSING_TIME", "TENDER_TYPE", "TENDER_DESCRIPTION",
    "TENDER_SOURCE", "DEPARTMENT", "PROVINCE",
    "ESUBMISSION", "CATEGORY", "IS_THERE_A_BRIEFING_SESSION",
    "BRIEFING_DATE", "COMPULSORY_BRIEFING", "BRIEFING_SESSION_VENUE", "LINK", "SOE",
    "COST_OF_SALES_ESTIMATE", "CAPABILITY_AVAILABLE", "CAPABILITY_GROUP", "REQUIREMENTS",
]
# End product and daily batch files exclude briefing fields (Tender Summary only)
BATCH_COLUMNS = [c for c in TENDER_COLUMNS if c != "BRIEFING_DATE"]
DATE_COLUMNS = {"REPORT_DATE", "PUBLICATION_DATE", "CLOSING_DATE", "BRIEFING_DATE"}

HDR_FONT = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F4E79")


# ── Folder structure ──────────────────────────────────────────────────────────

def create_batch_folder(date_from: str, date_to: str, batch_type: str,
                        root_dir: str = None) -> str:
    """
    Create and return the root batch folder.
    root_dir defaults to data/etenders.gov.za; pass data/All_Tenders for watchlist runs.
    """
    if root_dir is None:
        root_dir = os.path.join("data", "etenders.gov.za")

    start = datetime.strptime(date_from, "%Y-%m-%d")
    end   = datetime.strptime(date_to,   "%Y-%m-%d")

    if start.month == end.month:
        label = f"({batch_type}) {start.day}-{end.day} {end.strftime('%B %Y')}"
    else:
        label = (f"({batch_type}) {start.strftime('%d %b')}"
                 f"-{end.strftime('%d %b %Y')}")

    import shutil
    root = os.path.join(root_dir, label)
    if os.path.exists(root):
        shutil.rmtree(root)
        logging.info(f"Existing batch folder removed for re-scrape: {root}")
    os.makedirs(os.path.join(root, "batches"),          exist_ok=True)
    os.makedirs(os.path.join(root, "end product"),      exist_ok=True)
    os.makedirs(os.path.join(root, "Display Equation"), exist_ok=True)
    logging.info(f"Batch folder ready: {root}")
    return root


# ── Excel helpers ─────────────────────────────────────────────────────────────

def write_tender_rows(ws, df: pd.DataFrame, columns: list = None) -> None:
    """Write headers + data rows to a Tender Data worksheet."""
    if columns is None:
        columns = TENDER_COLUMNS
    for col_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL

    ordered = df.reindex(columns=columns)
    for row_idx, row in enumerate(ordered.itertuples(index=False), 2):
        for col_idx, (col_name, value) in enumerate(zip(columns, row), 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if col_name in DATE_COLUMNS and isinstance(value, str) and value:
                try:
                    cell.value = datetime.strptime(value, "%Y/%m/%d")
                    cell.number_format = "YYYY/MM/DD"
                except ValueError:
                    pass
            if col_name == "LINK" and isinstance(value, str) and value.startswith("http"):
                cell.hyperlink = value
                cell.style = "Hyperlink"


def _write_final_sheet(ws, batch_type: str, report_date: datetime) -> None:
    """Write the Final worksheet with the same formulas as the existing files."""
    for col_idx, heading in enumerate(
        ["Source", "Number of New Tenders", "ICT Tenders", "RFQs"], 1
    ):
        cell = ws.cell(row=1, column=col_idx, value=heading)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL

    ws["A2"] = "Report Date"
    ws["B2"] = f"({batch_type}) {report_date.strftime('%d/%m/%Y')}"

    ws["A3"] = "eTenders (All)"
    ws["B3"] = "=COUNTA('Tender Data'!J:J)-1"
    ws["C3"] = (
        "=COUNTIF('Tender Data'!M:M,\"Information service activities\")"
        "+COUNTIF('Tender Data'!M:M,\"Information and communication\")"
    )
    ws["D3"] = "=COUNTIF('Tender Data'!G:G,\"Request for Quotation\")"

    for offset, (label, pattern) in enumerate(SOURCES[1:]):
        r = 4 + offset
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2,
                value=f"=COUNTIF('Tender Data'!J:J,\"{pattern}\")")
        ws.cell(row=r, column=3, value=(
            f"=COUNTIFS('Tender Data'!M:M,\"Information service activities\","
            f"'Tender Data'!J:J,\"{pattern}\")"
            f"+COUNTIFS('Tender Data'!M:M,\"Information and communication\","
            f"'Tender Data'!J:J,\"{pattern}\")"
        ))
        ws.cell(row=r, column=4, value=(
            f"=COUNTIFS('Tender Data'!J:J,\"{pattern}\","
            f"'Tender Data'!G:G,\"Request for Quotation\")"
        ))

    last_dept = 3 + len(SOURCES)
    ws.cell(row=last_dept, column=2, value=f"=SUM(B4:B{last_dept - 1})")
    ws.cell(row=last_dept, column=3, value=f"=SUM(C4:C{last_dept - 1})")
    ws.cell(row=last_dept, column=4, value=f"=SUM(D4:D{last_dept - 1})")


# ── Per-day file ──────────────────────────────────────────────────────────────

def save_daily_file(tender_data: list, day: str, batch_folder: str) -> str:
    """
    Save one day's tender data to {batch_folder}/batches/tenders_YYYY_MM_DD.xlsx.
    If tender_data is empty the filename gets a ' (No Tenders)' suffix so the
    day is still accounted for in the batches folder.
    Returns the saved file path.
    """
    date_str  = day.replace("-", "_")
    suffix    = " (No Tenders)" if not tender_data else ""
    filepath  = os.path.join(batch_folder, "batches", f"tenders_{date_str}{suffix}.xlsx")

    # Assign record IDs (first scraped = highest number)
    records = []
    total = len(tender_data)
    for idx, tender in enumerate(tender_data, 1):
        t = tender.copy()
        t["RECORD_ID"] = total - idx + 1
        records.append(t)

    df = pd.DataFrame(records)

    wb = Workbook()
    ws = wb.active
    ws.title = "Tender Data"
    write_tender_rows(ws, df, BATCH_COLUMNS)
    wb.save(filepath)

    logging.info(f"Daily file saved: {filepath} ({total} tenders)")
    return filepath


# ── End product file ──────────────────────────────────────────────────────────

def create_end_product(df: pd.DataFrame, date_from: str, date_to: str,
                       batch_type: str, report_date: datetime,
                       batch_folder: str,
                       raw_df: pd.DataFrame = None) -> str:
    """
    Create the end product Excel file inside {batch_folder}/end product/.
    Sheets: Tender Data (deduplicated), Final (formulas), Raw Data (all scraped, optional).
    Returns the saved file path.
    """
    start = datetime.strptime(date_from, "%Y-%m-%d")
    end   = datetime.strptime(date_to,   "%Y-%m-%d")

    if start.month == end.month:
        label = f"{start.day} - {end.day} {end.strftime('%B')}"
    else:
        label = f"{start.strftime('%d %b')} - {end.strftime('%d %b')}"

    filename = f"RFQ_and_ICT_Checker_({label}).xlsx"
    filepath = os.path.join(batch_folder, "end product", filename)

    # Drop rows with no real publication date (proxy-dated watchlist tenders)
    df = df[df["PUBLICATION_DATE"].notna() & (df["PUBLICATION_DATE"].astype(str).str.strip() != "")]
    if df.empty:
        logging.warning("create_end_product: all rows removed by pub-date filter")
        return filepath

    # Deduplicate across all sources before final output
    df = pd.DataFrame(deduplicate_tenders(df.to_dict("records")))

    # Assign record IDs across the full combined dataset
    records = []
    total = len(df)
    for idx, row in enumerate(df.itertuples(index=False), 1):
        t = {col: getattr(row, col, None) for col in TENDER_COLUMNS if col != "RECORD_ID"}
        t["RECORD_ID"] = total - idx + 1
        records.append(t)

    df_out = pd.DataFrame(records)

    wb = Workbook()
    wb.remove(wb.active)
    write_tender_rows(wb.create_sheet("Tender Data"), df_out, BATCH_COLUMNS)
    _write_final_sheet(wb.create_sheet("Final"), batch_type, report_date)

    if raw_df is not None and not raw_df.empty:
        raw_records = []
        raw_total = len(raw_df)
        for idx, row in enumerate(raw_df.itertuples(index=False), 1):
            t = {col: getattr(row, col, None) for col in TENDER_COLUMNS if col != "RECORD_ID"}
            t["RECORD_ID"] = raw_total - idx + 1
            raw_records.append(t)
        write_tender_rows(wb.create_sheet("Raw Data"), pd.DataFrame(raw_records), BATCH_COLUMNS)
        logging.info(f"Raw Data sheet written: {raw_total} rows ({raw_total - total} duplicates)")

    wb.save(filepath)
    logging.info(f"End product created: {filepath}")
    return filepath


# ── Deduplication ────────────────────────────────────────────────────────────

def deduplicate_tenders(tenders: list) -> list:
    """
    Remove duplicate tender entries by TENDER_ID (first occurrence wins).
    Records with a blank/None TENDER_ID are always kept as-is.
    """
    seen: set = set()
    out: list = []
    dupes = 0
    for t in tenders:
        tid = str(t.get("TENDER_ID") or "").strip()
        if tid and tid in seen:
            dupes += 1
            logging.info(f"Duplicate tender skipped: {tid}")
            continue
        out.append(t)
        if tid:
            seen.add(tid)
    if dupes:
        logging.info(f"Deduplication removed {dupes} duplicate tender(s)")
    return out


# ── Counting ──────────────────────────────────────────────────────────────────

def _xl_match(value: str, pattern: str) -> bool:
    if not isinstance(value, str):
        return False
    regex = re.escape(pattern).replace(r"\*", ".*")
    return bool(re.fullmatch(regex, value, re.IGNORECASE))


def calculate_counts(df: pd.DataFrame) -> dict:
    dept = df["DEPARTMENT"].fillna("")
    cat  = df["CATEGORY"].fillna("")
    typ  = df["TENDER_TYPE"].fillna("")

    is_ict = cat.str.lower().isin([c.lower() for c in ICT_CATEGORIES])
    is_rfq = typ.str.lower() == RFQ_TYPE.lower()

    result = {}
    for label, pattern in SOURCES:
        if pattern is None:
            mask = pd.Series([True] * len(df), index=df.index)
        else:
            mask = dept.apply(lambda v, p=pattern: _xl_match(v, p))
        result[label] = {
            "NNT": int(mask.sum()),
            "ICT": int((mask & is_ict).sum()),
            "RFQ": int((mask & is_rfq).sum()),
        }
    return result


# ── Equation file ─────────────────────────────────────────────────────────────

MAX_EQUATION_BATCHES = 6


def update_equation_file(counts: dict, batch_type: str, report_date: datetime,
                         batch_folder: str) -> None:
    """
    Insert a new NNT/ICT/RFQ column block at column B (newest-first), shift
    older batches right, trim to MAX_EQUATION_BATCHES, then copy to the
    Display Equation folder inside batch_folder.
    """
    import shutil

    wb = load_workbook(EQUATION_FILE)
    ws = wb["Sheet1"]

    batch_label = f"({batch_type}) {report_date.strftime('%d/%m/%y')}"

    # Remove any existing column for this batch label (handles re-runs)
    col = 2
    while ws.cell(row=1, column=col).value is not None:
        if ws.cell(row=1, column=col).value == batch_label:
            ws.delete_cols(col, 3)
            break
        col += 3

    # Insert 3 blank columns at position B, pushing existing data right
    ws.insert_cols(2, 3)

    # Write new batch in columns B, C, D (indices 2, 3, 4)
    ws.cell(row=1, column=2, value=batch_label)
    ws.cell(row=2, column=2, value="NNT")
    ws.cell(row=2, column=3, value="ICT")
    ws.cell(row=2, column=4, value="RFQ")

    for row_offset, (source, _) in enumerate(SOURCES):
        data = counts.get(source, {"NNT": 0, "ICT": 0, "RFQ": 0})
        ws.cell(row=3 + row_offset, column=1, value=source)
        ws.cell(row=3 + row_offset, column=2, value=data["NNT"])
        ws.cell(row=3 + row_offset, column=3, value=data["ICT"])
        ws.cell(row=3 + row_offset, column=4, value=data["RFQ"])

    # Count batches and trim oldest if over the limit
    num_batches = 0
    col = 2
    while ws.cell(row=1, column=col).value is not None:
        num_batches += 1
        col += 3

    while num_batches > MAX_EQUATION_BATCHES:
        oldest_col = 2 + (num_batches - 1) * 3
        ws.delete_cols(oldest_col, 3)
        num_batches -= 1

    # Apply thin borders to the full data range
    _thin = Side(style="thin")
    _border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
    total_rows = 2 + len(SOURCES)
    total_cols = 1 + num_batches * 3
    for r in range(1, total_rows + 1):
        for c in range(1, total_cols + 1):
            ws.cell(row=r, column=c).border = _border

    wb.save(EQUATION_FILE)
    logging.info(f"Equation file updated: {batch_label}")

    # Copy updated equation file to Display Equation folder
    dest = os.path.join(batch_folder, "Display Equation",
                        os.path.basename(EQUATION_FILE))
    shutil.copy2(EQUATION_FILE, dest)
    logging.info(f"Display Equation copy saved: {dest}")


# ── Power BI export ───────────────────────────────────────────────────────────

PBI_COLUMNS = ["BATCH_LABEL", "BATCH_TYPE", "REPORT_DATE", "SOURCE", "NNT", "ICT", "RFQ"]


def update_power_bi_export(batch_folder: str, date_from: str, date_to: str, batch_type: str) -> None:
    """
    Append (or replace) this batch's rows in the running Power BI data file.
    One row per SOURCE per day — reads each daily batch file from batch_folder/batches/.
    Existing rows for this batch_label are replaced on re-runs.
    """
    start = datetime.strptime(date_from, "%Y-%m-%d")
    end   = datetime.strptime(date_to,   "%Y-%m-%d")

    if start.month == end.month:
        batch_label = f"({batch_type}) {start.day}-{end.day} {end.strftime('%b %y')}"
    else:
        batch_label = f"({batch_type}) {start.strftime('%d %b')}-{end.strftime('%d %b %y')}"

    # Build new rows — one per source per day
    new_rows = []
    batches_dir = os.path.join(batch_folder, "batches")

    current = start
    while current <= end:
        day_str  = current.strftime("%Y_%m_%d")
        day_file = os.path.join(batches_dir, f"tenders_{day_str}.xlsx")

        if os.path.exists(day_file):
            day_df     = pd.read_excel(day_file)
            day_counts = calculate_counts(day_df)
        else:
            day_counts = {source: {"NNT": 0, "ICT": 0, "RFQ": 0} for source, _ in SOURCES}

        for source, _ in SOURCES:
            data = day_counts.get(source, {"NNT": 0, "ICT": 0, "RFQ": 0})
            new_rows.append({
                "BATCH_LABEL": batch_label,
                "BATCH_TYPE":  batch_type,
                "REPORT_DATE": current,
                "SOURCE":      source,
                "NNT":         data["NNT"],
                "ICT":         data["ICT"],
                "RFQ":         data["RFQ"],
            })

        current += timedelta(days=1)

    new_df = pd.DataFrame(new_rows, columns=PBI_COLUMNS)

    # Load existing data and drop any prior rows for this batch (idempotent)
    if os.path.exists(POWER_BI_FILE):
        try:
            existing = pd.read_excel(POWER_BI_FILE)
            existing = existing[existing["BATCH_LABEL"] != batch_label]
            combined = pd.concat([existing, new_df], ignore_index=True)
        except Exception:
            combined = new_df
    else:
        combined = new_df

    # Sort newest batch first
    combined["REPORT_DATE"] = pd.to_datetime(combined["REPORT_DATE"])
    combined.sort_values("REPORT_DATE", ascending=False, inplace=True)
    combined.reset_index(drop=True, inplace=True)

    # Trim to 6 most recent batches (same rolling window as equation file)
    batch_dates = (
        combined.groupby("BATCH_LABEL")["REPORT_DATE"]
        .max()
        .sort_values(ascending=False)
    )
    keep_labels = batch_dates.head(MAX_EQUATION_BATCHES).index.tolist()
    combined = combined[combined["BATCH_LABEL"].isin(keep_labels)]

    # Write with styled header and named Excel Table for Power BI compatibility
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter

    os.makedirs(os.path.dirname(POWER_BI_FILE), exist_ok=True)
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
    logging.info(f"Power BI export updated: {batch_label} ({len(new_rows)} rows) -> {POWER_BI_FILE}")
