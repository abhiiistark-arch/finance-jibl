"""Map validated PDF rows to the fixed Excel schema."""

from __future__ import annotations

import pandas as pd

from app.utils.cleaner import clean_cell_string, format_excel_numeric, format_risk_end_date, normalize_header

OUTPUT_COLUMNS = [
    "Sr No",
    "Month",
    "Quarter",
    "Invoice Date",
    "Invoice Number",
    "Insurer Name",
    "Policy No",
    "Endorsement No",
    "Policy details (New/Renewal)",
    "Insured Name",
    "Product Code",
    "Product Name",
    "Policy Issue date",
    "Risk Start Date",
    "Risk End Date",
    "Product Type",
    "Basic/OD Premium",
    "TP Premium",
    "Total Premium",
    "Basic/OD Rate (%)",
    "Basic/ OD Brokerage Amount",
    "TP Brokerage (%)",
    "TP Brokerage Amount",
    "Total Brokerage",
    "State",
    "Group/ Non Group",
]

# PDF source keys (canonical names after fuzzy match resolution)
FIELD_POLICY_OR_ENDT = "Policy/ Endt number"
FIELD_INSURED_NAME = "Insured Name"
FIELD_POLICY_EXPIRY = "Policy Expiry Date"
FIELD_ELG_PREMIUM = "ELG Premium Amount"
FIELD_COMMISSION = "Commission Amount"

# Drop rows whose only extracted content is spacer/junk — avoids gaps in Excel.
_SUBSTANCE_KEYS = frozenset(
    {
        "Policy No",
        "Insured Name",
        "Risk End Date",
        "Basic/OD Premium",
        "Basic/ OD Brokerage Amount",
    }
)


def _row_has_extracted_content(mapped: dict[str, object]) -> bool:
    return any(str(mapped.get(k, "") or "").strip() for k in _SUBSTANCE_KEYS)


def rows_to_output_dataframe(rows: list[dict[str, object]]) -> pd.DataFrame:
    """Build output DataFrame with exact column order; Sr No incremented from 1."""
    if not rows:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    out_rows: list[dict[str, object]] = []
    for row in rows:
        mapped = map_row(row)
        if not _row_has_extracted_content(mapped):
            continue
        out_rows.append(mapped)

    for i, mapped in enumerate(out_rows, start=1):
        mapped["Sr No"] = i

    return pd.DataFrame(out_rows, columns=OUTPUT_COLUMNS)


def map_row(src: dict[str, object]) -> dict[str, object]:
    """
    Only these fields carry data from PDF; all others empty string.
    """
    dept_heading_norm = normalize_header("Department Name")
    policy = clean_cell_string(src.get(FIELD_POLICY_OR_ENDT, ""))
    if normalize_header(policy) == dept_heading_norm:
        policy = ""
    insured = clean_cell_string(src.get(FIELD_INSURED_NAME, ""))
    if normalize_header(insured) == dept_heading_norm:
        insured = ""
    expiry_fmt = format_risk_end_date(src.get(FIELD_POLICY_EXPIRY, ""))
    elg_prem = format_excel_numeric(src.get(FIELD_ELG_PREMIUM, ""))
    comm_amt = format_excel_numeric(src.get(FIELD_COMMISSION, ""))

    base: dict[str, object] = {c: "" for c in OUTPUT_COLUMNS}
    base["Policy No"] = policy
    base["Insured Name"] = insured
    base["Risk End Date"] = expiry_fmt
    base["Basic/OD Premium"] = elg_prem
    base["Basic/ OD Brokerage Amount"] = comm_amt
    return base
