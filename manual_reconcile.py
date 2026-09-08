"""
Merge auto_tenders.xlsx + manual_tenders.xlsx → master_tenders.xlsx.

Auto is the scraper's cumulative output. Manual is Thando's file from
SharePoint Manual_Scrapes/. Master is the derived, deduped union — the
only file Power BI reads.

Dedupe key: normalized TENDER_ID.
Conflict resolution: manual wins (INGESTION_METHOD=BOTH), else source stands.
"""
from __future__ import annotations

import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

from master_schema import TEMPLATE_COLUMN_ORDER, apply_template_schema, read_template_schema

log = logging.getLogger("etender.merge")

MASTER_SHEET = "TENDER REPORT"
_DDMMYYYY = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
DATE_COLUMNS = ("REPORT_DATE", "PUBLICATION_DATE", "CLOSING_DATE", "BRIEFING_DATE")

def _normalize_tid(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v).strip()).upper()

def _clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _normalize_dates(df: pd.DataFrame) -> pd.DataFrame:
    for col in DATE_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(_iso_date)
    return df

def _iso_date(v) -> Optional[str]:
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return None
    m = _DDMMYYYY.match(str(v).strip())
    if m:
        d, mo, y = m.groups()
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    try:
        return pd.to_datetime(v, errors="coerce").strftime("%Y-%m-%d")
    except Exception:
        return str(v)

def _align(df: pd.DataFrame) -> pd.DataFrame:
    df = _clean_headers(df)
    df = _normalize_dates(df)
    for col in TEMPLATE_COLUMN_ORDER:
        if col not in df.columns:
            df[col] = None
    df = df[[c for c in TEMPLATE_COLUMN_ORDER if c in df.columns]]
    return df

def merge_to_master(auto_path: str, manual_path: Optional[str], master_out: str) -> dict:
    """
    Build master_tenders.xlsx from auto + manual sources.
    Returns stats dict for the run summary.
    """
    auto_p = Path(auto_path)
    if not auto_p.exists():
        raise FileNotFoundError(f"auto_tenders not found: {auto_path}")

    auto_raw = pd.read_excel(auto_p, sheet_name=MASTER_SHEET, dtype=str)
    auto = _align(read_template_schema(auto_raw))
    if "INGESTION_METHOD" not in auto.columns or auto["INGESTION_METHOD"].isna().all():
        auto["INGESTION_METHOD"] = "AUTOMATIC"
    auto["INGESTION_METHOD"] = auto["INGESTION_METHOD"].fillna("AUTOMATIC")
    auto["_tid"] = auto["TENDER_ID"].apply(_normalize_tid)
    log.info("merge: auto=%d rows", len(auto))

    manual_new_count = 0
    flipped_count = 0
    if manual_path and Path(manual_path).exists():
        manual_raw = pd.read_excel(manual_path, dtype=str)
        manual = _align(manual_raw)
        manual["_tid"] = manual["TENDER_ID"].apply(_normalize_tid)
        manual = manual[manual["_tid"] != ""]
        log.info("merge: manual=%d rows", len(manual))

        auto_tids = set(auto.loc[auto["_tid"] != "", "_tid"])
        flip_mask = auto["_tid"].isin(set(manual["_tid"])) & (auto["INGESTION_METHOD"] == "AUTOMATIC")
        flipped_count = int(flip_mask.sum())
        auto.loc[flip_mask, "INGESTION_METHOD"] = "BOTH"

        new_rows = manual[~manual["_tid"].isin(auto_tids)].copy()
        new_rows["INGESTION_METHOD"] = "MANUAL"
        manual_new_count = len(new_rows)

        combined = pd.concat([auto, new_rows], ignore_index=True)
    else:
        log.info("merge: no manual file (skipped)")
        combined = auto

    combined = combined.drop(columns=["_tid"], errors="ignore")
    combined["REPORT_DATE"] = pd.to_datetime(combined["REPORT_DATE"], errors="coerce")
    combined = combined.sort_values("REPORT_DATE", ascending=False).reset_index(drop=True)
    combined["RECORD_ID"] = range(1, len(combined) + 1)

    for col in TEMPLATE_COLUMN_ORDER:
        if col not in combined.columns:
            combined[col] = None
    combined = combined[TEMPLATE_COLUMN_ORDER]
    combined = apply_template_schema(combined)

    out = Path(master_out)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix(".xlsx.tmp")
    combined.to_excel(tmp, sheet_name=MASTER_SHEET, index=False)
    shutil.move(str(tmp), str(out))

    log.info(
        "merge: wrote %s (auto=%d, manual_new=%d, flipped=%d, total=%d)",
        out.name, len(auto), manual_new_count, flipped_count, len(combined),
    )
    return {
        "auto_rows": len(auto),
        "manual_new_rows": manual_new_count,
        "flipped_to_both": flipped_count,
        "master_total_rows": len(combined),
    }
