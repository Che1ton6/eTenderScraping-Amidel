"""
Data integrity audit for auto_tenders.xlsx / master_tenders.xlsx on SharePoint.

Downloads both files fresh from SharePoint and checks for exactly the class
of corruption this project has hit before: garbage values in ESUBMISSION
(e.g. province names from a column-shift bug), duplicate TENDER_IDs, missing
schema columns, and batches dated in the future (a scrape for a date range
that hasn't fully happened yet — e.g. a "Thu-Sun" batch checked on the Thu).
"""
from __future__ import annotations

import logging
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from master_schema import TEMPLATE_COLUMN_ORDER

log = logging.getLogger("etender.audit")

_VALID_ESUB = {"YES", "NO", "UNVERIFIED"}


def _audit_one(path: str, label: str) -> dict:
    p = Path(path)
    if not p.exists():
        return {"file": label, "ok": False, "issues": [f"{label} was not downloaded / does not exist"]}

    df = pd.read_excel(p, sheet_name="TENDER REPORT", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    issues: list[str] = []
    total = len(df)

    # 1. Schema completeness
    missing_cols = [c for c in TEMPLATE_COLUMN_ORDER if c not in df.columns]
    if missing_cols:
        issues.append(f"Missing expected columns: {missing_cols}")

    # 2. ESUBMISSION garbage check — the exact bug class hit before (province
    # names from a column-shift bug, or anything else that isn't a real answer).
    esub_summary: dict[str, int] = {}
    if "ESUBMISSION" in df.columns:
        for v in df["ESUBMISSION"]:
            key = "" if pd.isna(v) else str(v).strip()
            esub_summary[key or "(blank)"] = esub_summary.get(key or "(blank)", 0) + 1
        garbage = [v for v in df["ESUBMISSION"].dropna().unique()
                   if str(v).strip().upper() not in _VALID_ESUB]
        if garbage:
            issues.append(
                f"{len(garbage)} invalid ESUBMISSION value(s) found "
                f"(should only be Yes/No/Unverified/blank): {sorted(str(g) for g in garbage)[:10]}"
            )

    # 3. Duplicate TENDER_IDs
    dup_count = 0
    if "TENDER_ID" in df.columns:
        norm = df["TENDER_ID"].astype(str).str.strip().str.upper()
        dup_count = int(norm[norm != ""].duplicated().sum())
        if dup_count:
            issues.append(f"{dup_count} duplicate TENDER_ID(s) found")

    # 4. Future-dated batches — a scrape for a date range that hasn't fully
    # happened yet (the exact "17-20 September" bug hit earlier).
    future_rows = 0
    if "REPORT_DATE" in df.columns:
        rd = pd.to_datetime(df["REPORT_DATE"], errors="coerce")
        today = pd.Timestamp(date.today())
        future_rows = int((rd > today).sum())
        if future_rows:
            issues.append(
                f"{future_rows} row(s) dated after today ({date.today()}) — "
                f"likely a premature/incomplete batch"
            )

    return {
        "file": label,
        "ok": len(issues) == 0,
        "total_rows": total,
        "esubmission_distribution": esub_summary,
        "duplicate_tender_ids": dup_count,
        "future_dated_rows": future_rows,
        "issues": issues,
    }


def run_audit(auto_path: str, master_path: str) -> dict:
    auto_result = _audit_one(auto_path, "auto_tenders.xlsx")
    master_result = _audit_one(master_path, "master_tenders.xlsx")
    overall_ok = auto_result["ok"] and master_result["ok"]
    return {
        "ok": overall_ok,
        "checked_at": datetime.utcnow().isoformat() + "Z",
        "auto_tenders": auto_result,
        "master_tenders": master_result,
    }
