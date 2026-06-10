#!/usr/bin/env python3
"""
Tender Summary
Creates a single Excel file — Tender Summary.xlsx — inside the batch folder.
One tab per tracked source (from SOURCES) that has >= 1 tender of any type.
Within each tab, RFQs sort to the top, then all others by closing date.
"""

import os
import re
import logging
import pandas as pd
from openpyxl import Workbook

from BatchProcessor import SOURCES, RFQ_TYPE, write_tender_rows


def _xl_match(value: str, pattern: str) -> bool:
    if not isinstance(value, str):
        return False
    regex = re.escape(pattern).replace(r"\*", ".*")
    return bool(re.fullmatch(regex, value, re.IGNORECASE))


def create_tender_summary(df: pd.DataFrame, batch_folder: str) -> int:
    """
    Creates Tender Summary.xlsx in batch_folder.
    Each tab = one SOURCES entry that has >= 1 tender (any type).
    Within each tab, RFQs sort to the top, then all others by closing date.
    Returns the number of tabs written (0 if file was not created).
    """
    if df is None or len(df) == 0:
        logging.info("No tenders — skipping Tender Summary")
        return 0

    dept_series = df["DEPARTMENT"].fillna("").astype(str)
    is_rfq = df["TENDER_TYPE"].fillna("").str.lower() == RFQ_TYPE.lower()

    wb = Workbook()
    wb.remove(wb.active)  # remove the blank default sheet
    tabs_created = 0

    for label, pattern in SOURCES:
        if pattern is None:
            continue  # skip "eTenders (All)"

        dept_mask = dept_series.apply(lambda v, p=pattern: _xl_match(v, p))
        source_df = df[dept_mask].copy()

        if len(source_df) == 0:
            continue

        # RFQs first, then all others sorted by closing date
        source_df["_is_rfq"] = is_rfq[source_df.index]
        source_df = source_df.sort_values(
            ["_is_rfq", "CLOSING_DATE"], ascending=[False, True]
        ).drop(columns=["_is_rfq"])

        # Assign sequential record IDs for this tab
        source_df = source_df.reset_index(drop=True)
        source_df["RECORD_ID"] = range(1, len(source_df) + 1)

        safe_name = label[:31]  # Excel sheet name limit
        ws = wb.create_sheet(safe_name)
        write_tender_rows(ws, source_df)

        logging.info(f"Tender Summary tab: {label} ({len(source_df)} tender(s))")
        tabs_created += 1

    if tabs_created == 0:
        logging.info("No tenders matched any tracked source — Tender Summary not created")
        return 0

    filepath = os.path.join(batch_folder, "Tender Summary.xlsx")
    wb.save(filepath)
    logging.info(f"Tender Summary saved: {filepath} ({tabs_created} tab(s))")
    return tabs_created
