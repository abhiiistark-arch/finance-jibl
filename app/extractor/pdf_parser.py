"""
Extract tables from PDFs: every page/block whose columns match the commission schema
(≥5 fuzzy-matched columns) contributes rows; pdfplumber is tried before PyMuPDF.
"""

from __future__ import annotations

import logging
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator, Optional, Sequence

SchemaMatch = tuple[list[str], list[list[Any]], dict[str, str]]

import pdfplumber
import pymupdf

from fuzzywuzzy import fuzz

from app.utils.cleaner import clean_cell_string, normalize_header

logger = logging.getLogger(__name__)


EXPECTED_SCHEMA_LABELS = [
    "Dept Code",
    "Department Name",
    "Policy/ Endt number",
    "Insured Name",
    "Policy Expiry Date",
    "ELG Premium Amount",
    "Ineligible Amount",
    "Commission Amount",
    "Insured Type",
]


@dataclass
class ParsedTableResult:
    """Rows as dicts keyed by canonical PDF field names (schema labels)."""
    matched_headers: dict[str, str]
    rows: list[dict[str, object]]


def normalize_header_cells(row: Sequence[Any]) -> list[str]:
    return [clean_cell_string(c) if c is not None else "" for c in row]


def _fuzz_score(a_norm: str, b_norm: str) -> float:
    """Return similarity in [0, 1]: max of ratio / token / partial fuzzywuzzy signals."""
    if not a_norm or not b_norm:
        return 0.0
    base = fuzz.ratio(a_norm, b_norm)
    token = fuzz.token_sort_ratio(a_norm, b_norm)
    partial = fuzz.partial_ratio(a_norm, b_norm)
    return max(base, token, partial) / 100.0


def find_column_assignments(headers: Sequence[str]) -> tuple[dict[str, str], dict[str, str]]:
    """
    Bidirectional map: pdf column header -> canonical expected label,
    canonical -> pdf header (inverse). Greedy assignment by best global scores.
    """
    header_triples: list[tuple[str, str, str]] = []
    for raw in headers:
        r = clean_cell_string(raw)
        if not r:
            continue
        header_triples.append((r, normalize_header(r), r))

    candidates: list[tuple[float, str, str]] = []
    for exp in EXPECTED_SCHEMA_LABELS:
        en = normalize_header(exp)
        best_s = 0.0
        best_pdf = ""
        for _raw, hn, original in header_triples:
            s = max(_fuzz_score(en, hn), _fuzz_score(en, normalize_header(_raw)))
            if s > best_s:
                best_s = s
                best_pdf = original
        if best_pdf:
            candidates.append((best_s, exp, best_pdf))

    candidates.sort(key=lambda x: -x[0])
    used_pdf: set[str] = set()
    used_exp: set[str] = set()
    pdf_to_expected: dict[str, str] = {}
    CONFIDENCE_THRESHOLD = 0.65

    for score, expected, pdf_col in candidates:
        if pdf_col in used_pdf or expected in used_exp:
            continue
        if score < CONFIDENCE_THRESHOLD:
            continue
        used_pdf.add(pdf_col)
        used_exp.add(expected)
        pdf_to_expected[pdf_col] = expected

    expected_to_pdf = {v: k for k, v in pdf_to_expected.items()}
    return pdf_to_expected, expected_to_pdf


def validate_table(headers: Sequence[str]) -> tuple[bool, dict[str, str], dict[str, str]]:
    """Valid when at least 5 of 9 schema columns fuzzy-match headers."""
    pdf_to_expected, expected_to_pdf = find_column_assignments(headers)
    if len(expected_to_pdf) >= 5:
        return True, pdf_to_expected, expected_to_pdf
    return False, {}, {}


def _dedupe_header(headers: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for h in headers:
        raw = clean_cell_string(h)
        base = raw or "(empty)"
        n = seen.get(base, 0)
        seen[base] = n + 1
        out.append(h if n == 0 else f"{h} [{n}]")
    return out


_HEADER_SCAN_ROWS = 15

# Lone label row inserted before a second table block on some layouts; exclude it and rows below it.
_SECTION_BREAK_DEPARTMENT_NAME = normalize_header("Department Name")


def _row_starts_secondary_table_section(row: Sequence[Any]) -> bool:
    """True when the row is only a 'Department Name' heading (same label as schema column name)."""
    non_empty = [clean_cell_string(c) for c in row if clean_cell_string(c)]
    if not non_empty:
        return False
    if len(non_empty) != 1:
        return False
    return normalize_header(non_empty[0]) == _SECTION_BREAK_DEPARTMENT_NAME


def find_schema_match_in_matrix(matrix: Sequence[Sequence[Any]]) -> Optional[SchemaMatch]:
    """
    Scan top rows for the commission header line (handles header not row 0, merged/noise rows).
    """
    if len(matrix) < 1:
        return None
    scan = min(_HEADER_SCAN_ROWS, len(matrix))
    for i in range(scan):
        row_cells = normalize_header_cells(matrix[i])
        if not any(row_cells):
            continue
        headers_list = _dedupe_header(list(row_cells))
        ok, pdf_to_expected, _ = validate_table(headers_list)
        if not ok:
            continue
        body = list(matrix[i + 1 :])
        return headers_list, body, pdf_to_expected
    return None


def _row_to_canonical(
    pdf_row: Sequence[Any],
    hdr_list: Sequence[str],
    pdf_to_expected: dict[str, str],
) -> dict[str, object]:
    out: dict[str, object] = {}
    n = min(len(hdr_list), len(pdf_row))
    for i in range(n):
        pdf_h = hdr_list[i]
        canonical = pdf_to_expected.get(pdf_h)
        if canonical is None:
            continue
        out[canonical] = pdf_row[i]
    return out


# pdfplumber presets: ruled commission PDFs often need line-based lattice; fallback to text/text.
TABLE_PRESETS: list[dict[str, Any] | None] = [
    None,
    {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3,
        "join_tolerance": 4,
        "edge_min_length": 30,
        "text_tolerance": 2,
        "text_x_tolerance": 2,
        "text_y_tolerance": 2,
    },
    {"vertical_strategy": "lines", "horizontal_strategy": "text", "snap_tolerance": 3},
    {"vertical_strategy": "text", "horizontal_strategy": "lines", "snap_tolerance": 3},
    {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "edge_min_length": 20,
    },
]


def _iter_tables_pdfplumber(path: Path) -> Iterator[SchemaMatch]:
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            # Avoid the same logical table yielded once per preset (duplicate rows).
            page_done = False
            for preset in TABLE_PRESETS:
                if page_done:
                    break
                if preset:
                    rows = page.extract_tables(table_settings=preset)
                else:
                    rows = page.extract_tables()
                for t in rows or []:
                    if not t:
                        continue
                    hit = find_schema_match_in_matrix(t)
                    if hit is None:
                        continue
                    yield hit
                    page_done = True
                    break


def _iter_tables_pymupdf(path: Path) -> Iterator[SchemaMatch]:
    doc = pymupdf.open(str(path))
    try:
        for page in doc:
            find_tables = getattr(page, "find_tables", None)
            if not callable(find_tables):
                continue
            tabs = find_tables()
            if tabs is None:
                continue
            table_list = getattr(tabs, "tables", []) or []
            for tab in table_list:
                try:
                    data = tab.extract()
                except Exception as e:
                    logger.debug("pymupdf extract failed on a table: %s", e)
                    continue
                if not data:
                    continue
                hit = find_schema_match_in_matrix(data)
                if hit is None:
                    continue
                yield hit
                break
    finally:
        doc.close()


def extract_valid_table_rows(pdf_source: Path | bytes, source_name: str = "") -> ParsedTableResult:
    """
    Return all body rows from every extractable table that matches the schema (≥5 / 9 columns),
    in page order. Rows keyed by canonical labels.
    """
    tmp_path: Path | None = None
    if isinstance(pdf_source, bytes):
        fd, pname = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        Path(pname).write_bytes(pdf_source)
        path = Path(pname)
        tmp_path = path
    else:
        path = Path(pdf_source)

    name = source_name or path.name

    pipelines: list[tuple[str, Callable[[Path], Iterable[SchemaMatch]]]] = [
        ("pdfplumber", lambda p: _iter_tables_pdfplumber(p)),
        ("pymupdf", lambda p: _iter_tables_pymupdf(p)),
    ]

    try:
        for engine, iterate in pipelines:
            try:
                rows_out: list[dict[str, object]] = []
                matched_headers: dict[str, str] = {}
                table_index = 0
                for headers_list, body, pdf_to_expected in iterate(path):
                    table_index += 1
                    if not matched_headers:
                        matched_headers = pdf_to_expected
                    logger.info(
                        "Valid schema table %d: file=%s engine=%s cols=%s body_rows=%d",
                        table_index,
                        name,
                        engine,
                        list(pdf_to_expected.values()),
                        len(body),
                    )
                    for r in body:
                        if _row_starts_secondary_table_section(r):
                            logger.debug(
                                "Skipping subsection heading row 'Department Name' (continuing table) file=%s",
                                name,
                            )
                            continue
                        if not any(clean_cell_string(c) if c is not None else False for c in r):
                            continue
                        try:
                            row_dict = _row_to_canonical(r, headers_list, pdf_to_expected)
                        except Exception as e:
                            logger.warning("Corrupt row skipped in %s: %s", name, e)
                            continue
                        rows_out.append(row_dict)

                if table_index > 0:
                    logger.info(
                        "Extracted %d data row(s) from %d schema table(s) via %s file=%s",
                        len(rows_out),
                        table_index,
                        engine,
                        name,
                    )
                    return ParsedTableResult(matched_headers=matched_headers, rows=rows_out)

                logger.warning("No qualifying table extracted with pipeline %s (%s)", engine, name)

            except Exception as e:
                logger.exception("Pipeline %s failed for %s: %s", engine, name, e)
                continue

        logger.warning(
            "No valid schema-matching table in %s; returning headers-only Excel.",
            name,
        )
        return ParsedTableResult(matched_headers={}, rows=[])

    finally:
        if tmp_path is not None:
            try:
                tmp_path.unlink(missing_ok=True)
            except OSError:
                logger.debug("Temp PDF could not be removed: %s", tmp_path)


def extract_upload_bytes(raw: bytes, filename: str = "upload.pdf") -> ParsedTableResult:
    """Convenience for Flask uploads."""
    return extract_valid_table_rows(raw, source_name=filename)
