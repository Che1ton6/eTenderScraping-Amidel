"""
Master tenders write-time schema transform.

Internal Python schema uses UPPER_SNAKE identifiers (COST_OF_SALES_ESTIMATE,
DUPLICATED, etc.) which every scraper writes to. The Excel file that Bheki
sees must match the DAILYTENDER_REPORT_3.0 template's TENDER_REPORT (2)
column layout + a few extras (INGESTION_METHOD, REPORT_YEAR/MONTH/DAY).

apply_template_schema() bridges the two: takes an internal-schema DataFrame,
returns a template-schema DataFrame ready to write to Excel.
"""
from __future__ import annotations
import pandas as pd


WATCHLIST_DEPARTMENTS = [
    "All State Owened Entities (See pdf document shared) ",
    "Eskom ",
    "SITA/State Information Technology Agency",
    "Trasnet ",
    "South African Reserve Bank (SARB)",
    "Gauteng Department of Health ",
    "Department of International Relation & Coperation (DIRCO)",
    "Mbashe Local Municipality ",
    "Ngqushwa Local\xa0\xa0Municipality",
    "Matatiele Local Municipality",
    "Ntabankulu Local Municipality",
    "Umzimvubu Local Municipality",
    "Winnie Madikizela-Mandela Local Municipality",
    "Mnquma Local Municipality",
    "Great Kei Local Municipality",
    "Amahlathi Local Municipality",
    "Raymond Mhlaba Local Municipality",
    "City Power Johannesburg (SOC) Ltd",
    "Joburg Market /Johannesburg Fresh Produce Market (Pty) Ltd",
    "Johannesburg City Parks and Zoo NPC",
    "Johannesburg City Theatres (Pty) Ltd",
    "Johannesburg Development Agency (Pty) Ltd",
    "City of Joburg Property Company (SOC) Ltd",
    "Johannesburg Roads Agency (Pty) Ltd",
    "Johannesburg Social Housing Company (SOC) Ltd",
    "Johannesburg Tourism Company (SOC) Ltd",
    "Johannesburg Water (SOC) Ltd",
    "Johannesburg Metropolitan Bus Services (SOC) Ltd",
    "Pikitup Johannesburg (Pty) Ltd",
    "Metropolitan Trading Company (Pty) Ltd",
    "Nelson Mandela Bay Metropolitan Municipality",
    "Buffalo City Metropolitan Municipality",
    "Eastern Cape Department of Public Works and Infrastructure",
    " National Department of Labour",
    " Special Investigations \xa0Unit (SIU)",
    " Department of Science and Technology\xa0",
    "Road Accident Fund (RAF)",
]

# Final Excel header order. First 26 = template TENDER_REPORT (2); then 4 extras.
TEMPLATE_COLUMN_ORDER = [
    "REPORT_DATE",
    "RECORD_ID",
    "TENDER_ID",
    "PUBLICATION_DATE",
    "CLOSING_DATE",
    "CLOSING_TIME",
    "TENDER_TYPE",
    "TENDER_SOURCE",
    "DEPARTMENT",
    "PROVINCE",
    "ESUBMISSION",
    "CATEGORY",
    "TENDER_DESCRIPTION",
    "IS_THERE_A_BRIEFING_SESSION",
    "BRIEFING_DATE",
    "COMPULSORY_BRIEFING",
    "BRIEFING_SESSION_VENUE",
    "LINK",
    "SOE",
    "Cost of Sales Estimate",
    "CAPABILITY_AVAILABLE",
    "CAPABILITY_GROUP",
    "REQUIREMENTS",
    "Watch_List_1",
    "Watch_List_Final",
    "Watch_List_Final_Original",
    # extras Bheki asked to keep beyond the template:
    "REPORT_YEAR",
    "REPORT_MONTH",
    "REPORT_DAY",
    "INGESTION_METHOD",
]

# Internal -> template header rename applied at write time
_INTERNAL_TO_TEMPLATE = {
    "COST_OF_SALES_ESTIMATE": "Cost of Sales Estimate",
}

# Reverse — applied at read time so template-schema files can be loaded back
# into the internal pipeline without silently dropping renamed columns.
_TEMPLATE_TO_INTERNAL = {v: k for k, v in _INTERNAL_TO_TEMPLATE.items()}


def read_template_schema(df: pd.DataFrame) -> pd.DataFrame:
    """
    Inverse of apply_template_schema for the round trip.

    Renames template headers back to internal identifiers. Derived columns
    (Watch_List_*, REPORT_YEAR/MONTH/DAY) are left in place — internal
    pipelines that don't expect them will just ignore or overwrite them.
    DUPLICATED is re-added as 0 so downstream dedup logic keeps working.
    """
    out = df.rename(columns=_TEMPLATE_TO_INTERNAL)
    if "DUPLICATED" not in out.columns:
        out["DUPLICATED"] = 0
    return out


def _yes_no(is_watch: bool) -> str:
    return "Yes" if is_watch else "NO"


def _upper_yes(v) -> bool:
    if v is None:
        return False
    return str(v).strip().upper() == "YES"


def apply_template_schema(df: pd.DataFrame) -> pd.DataFrame:
    """
    Transform an internal-schema DataFrame into the Excel template layout.

    - Renames COST_OF_SALES_ESTIMATE -> "Cost of Sales Estimate"
    - Adds Watch_List_1 (mirrors watchlist logic: SOE=YES OR DEPARTMENT in list)
    - Adds Watch_List_Final (same logic, upper-cased YES/NO)
    - Adds Watch_List_Final_Original (blank, template parity)
    - Adds REPORT_YEAR / REPORT_MONTH / REPORT_DAY from REPORT_DATE
    - Drops DUPLICATED
    - Reorders to TEMPLATE_COLUMN_ORDER
    """
    out = df.copy()

    if "COST_OF_SALES_ESTIMATE" in out.columns:
        out = out.rename(columns=_INTERNAL_TO_TEMPLATE)

    watch_set = set(WATCHLIST_DEPARTMENTS)
    soe_series = out.get("SOE", pd.Series([None] * len(out)))
    dept_series = out.get("DEPARTMENT", pd.Series([None] * len(out)))
    is_watch = [
        _upper_yes(s) or (d in watch_set)
        for s, d in zip(soe_series, dept_series)
    ]
    out["Watch_List_1"] = ["Yes" if w else "NO" for w in is_watch]
    out["Watch_List_Final"] = ["YES" if w else "NO" for w in is_watch]
    out["Watch_List_Final_Original"] = None

    rd = pd.to_datetime(out.get("REPORT_DATE"), errors="coerce")
    out["REPORT_YEAR"] = rd.dt.year
    out["REPORT_MONTH"] = rd.dt.month
    out["REPORT_DAY"] = rd.dt.day

    if "DUPLICATED" in out.columns:
        out = out.drop(columns=["DUPLICATED"])

    for col in TEMPLATE_COLUMN_ORDER:
        if col not in out.columns:
            out[col] = None

    return out[TEMPLATE_COLUMN_ORDER]
