"""
Batch-process every PDF in `input/`, writing one Excel workbook per PDF into `output/`.
Run from repo root: `python process_all_input_pdfs.py`
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

import pandas as pd

from app.app import configure_logging
from app.extractor.mapper import rows_to_output_dataframe
from app.extractor.pdf_parser import extract_valid_table_rows

BASE = Path(__file__).resolve().parent
INPUT_DIR = BASE / "input"
OUTPUT_DIR = BASE / "output"


def main() -> int:
    configure_logging()
    log = logging.getLogger("batch")

    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    pdfs = sorted(INPUT_DIR.glob("*.pdf"))
    if not pdfs:
        log.warning("No PDF files under %s", INPUT_DIR)
        return 0

    summaries: list[str] = []

    for pdf_path in pdfs:
        parsed = extract_valid_table_rows(pdf_path, source_name=str(pdf_path.name))
        df = rows_to_output_dataframe(parsed.rows)
        stem = pdf_path.stem
        xlsx_path = OUTPUT_DIR / f"{stem}_commission_mapped.xlsx"
        log.info(
            "file=%s valid_table_found=%s rows=%d out=%s",
            pdf_path.name,
            bool(parsed.matched_headers),
            len(parsed.rows),
            xlsx_path.name,
        )

        summaries.append(f"{pdf_path.name}: matched={bool(parsed.matched_headers)} rows={len(parsed.rows)} -> {xlsx_path}")

        try:
            with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Commission")
        except Exception as e:
            log.exception("Failed writing Excel for %s: %s", pdf_path.name, e)
            summaries[-1] += f" WRITE_ERROR={e!s}"

    print("\nSummary")
    print("-------")
    for line in summaries:
        print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
