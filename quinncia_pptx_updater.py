"""Utilities for updating the Quinncia Appendix table in a PowerPoint deck.

This module updates slide 2 only: Appendix: Entire MSB.
It pulls values from the first CSV/Excel section named "Quinncia Metrics (All Students)".
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Dict, Iterable, List, Tuple

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor

SECTION_TITLE = "Quinncia Metrics (All Students)"
TARGET_SLIDE_INDEX = 1  # slide 2 in PowerPoint, zero-indexed in python-pptx
BLACK = RGBColor(0, 0, 0)

# PowerPoint display names on slide 2 -> names used in the Quinncia export.
PROGRAM_LOOKUP = {
    "HR": "Human Resource Management",
    "BSIS": "Information Systems (BS)",
    "MISM": "Information Systems (MS)",
    "Overall MSB": "Overall Marriott School",
}

# PowerPoint header text -> Quinncia CSV/Excel column name.
HEADER_TO_METRIC = {
    "Students": "enrolled_students",
    "Sign Ups": "quinncia_sign_ups",
    "Sign Up%": "signup_pct",
    "Resume Students": "resume_students",
    "Resume %": "resume_student_pct",
    "Total Resume Uploads": "resume_uploads",
    "Avg Resume Score": "avg_resume_score",
    "Max Resume Score": "max_resume_score",
    "LinkedIn Students": "linkedin_students",
    "Linked In%": "linkedin_student_pct",
    "Total Linked In Uploads": "linkedin_uploads",
    "Avg Linked In Score": "avg_linkedin_score",
    "Max Linked In Score": "max_linkedin_score",
    "Interview Students": "interview_students",
    "Interview %": "interview_student_pct",
    "Total Interviews": "interview_uploads",
    "Avg Interview Score": "avg_interview_score",
    "Max Interview Score": "max_interview_score",
    "All 3 Students": "all_three_students",
    "All 3 Student %": "all_three_student_pct",
}

COLUMN_ALIASES = {
    "program": "major",
    "major": "major",
    "students": "enrolled_students",
    "enrolled_students": "enrolled_students",
    "sign_ups": "quinncia_sign_ups",
    "quinncia_sign_ups": "quinncia_sign_ups",
    "sign_up": "signup_pct",
    "signup_pct": "signup_pct",
    "resume_students": "resume_students",
    "resume_student_pct": "resume_student_pct",
    "resume": "resume_student_pct",
    "total_resume_uploads": "resume_uploads",
    "resume_uploads": "resume_uploads",
    "avg_resume_score": "avg_resume_score",
    "max_resume_score": "max_resume_score",
    "linkedin_students": "linkedin_students",
    "linked_in_students": "linkedin_students",
    "linked_in": "linkedin_student_pct",
    "linkedin_student_pct": "linkedin_student_pct",
    "linked_in_pct": "linkedin_student_pct",
    "total_linked_in_uploads": "linkedin_uploads",
    "total_linkedin_uploads": "linkedin_uploads",
    "linkedin_uploads": "linkedin_uploads",
    "avg_linked_in_score": "avg_linkedin_score",
    "avg_linkedin_score": "avg_linkedin_score",
    "max_linked_in_score": "max_linkedin_score",
    "max_linkedin_score": "max_linkedin_score",
    "interview_students": "interview_students",
    "interview": "interview_student_pct",
    "interview_pct": "interview_student_pct",
    "interview_student_pct": "interview_student_pct",
    "total_interviews": "interview_uploads",
    "interview_uploads": "interview_uploads",
    "avg_interview_score": "avg_interview_score",
    "max_interview_score": "max_interview_score",
    "all_3_students": "all_three_students",
    "all_three_students": "all_three_students",
    "all_3_student": "all_three_student_pct",
    "all_3_student_pct": "all_three_student_pct",
    "all_three_student_pct": "all_three_student_pct",
}

REQUIRED_COLUMNS = {"major"} | set(HEADER_TO_METRIC.values())


@dataclass
class UpdateSummary:
    updated_rows: int
    updated_cells: int
    missing_programs: List[str]
    section_used: str = SECTION_TITLE
    slide_updated: int = 2


def _norm(text: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(text).strip().lower()).strip()


def _canon_col(name: object) -> str:
    cleaned = str(name).strip().lower()
    cleaned = cleaned.replace("%", "pct")
    cleaned = re.sub(r"[^a-z0-9]+", "_", cleaned).strip("_")
    return COLUMN_ALIASES.get(cleaned, cleaned)


def _decode_bytes(raw: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _extract_csv_section(raw: bytes) -> pd.DataFrame:
    """Extract the Quinncia Metrics (All Students) section from a multi-section CSV."""
    text = _decode_bytes(raw)
    lines = text.splitlines()
    target = _norm(SECTION_TITLE)

    start = None
    for i, line in enumerate(lines):
        # Strip delimiter-only noise that sometimes appears in CSV exports.
        cleaned = line.strip().strip(",")
        if _norm(cleaned) == target:
            start = i + 1
            break

    if start is None:
        # Fall back to treating the upload as a simple one-table CSV.
        return pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False)

    section_lines: List[str] = []
    for line in lines[start:]:
        if not line.strip():
            if section_lines:
                break
            continue
        section_lines.append(line)

    if not section_lines:
        raise ValueError(f"Found '{SECTION_TITLE}', but no table was found underneath it.")

    return pd.read_csv(io.StringIO("\n".join(section_lines)), dtype=str, keep_default_na=False)


def _read_excel_metrics(raw: bytes) -> pd.DataFrame:
    """Read metrics from an Excel workbook, if the export is uploaded as .xlsx."""
    excel = pd.ExcelFile(io.BytesIO(raw))
    target_norm = _norm(SECTION_TITLE)

    # Preferred: sheet is named exactly/approximately like the section.
    for sheet_name in excel.sheet_names:
        sheet_norm = _norm(sheet_name)
        if sheet_norm == target_norm or ("quinncia metrics" in sheet_norm and "all students" in sheet_norm):
            return pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, dtype=str, keep_default_na=False)

    # Fallback: scan sheets for either the section title or a header row starting with major.
    for sheet_name in excel.sheet_names:
        scan = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, header=None, dtype=str, keep_default_na=False)
        for row_idx in range(len(scan)):
            row_values = [str(x).strip() for x in scan.iloc[row_idx].tolist()]
            normalized_row = [_norm(x) for x in row_values]
            if target_norm in normalized_row:
                header_idx = row_idx + 1
            elif "major" in normalized_row:
                header_idx = row_idx
            else:
                continue

            header = [str(x).strip() for x in scan.iloc[header_idx].tolist()]
            data_rows = []
            for data_idx in range(header_idx + 1, len(scan)):
                row = scan.iloc[data_idx].tolist()
                if not any(str(x).strip() for x in row):
                    break
                data_rows.append(row)
            if data_rows:
                return pd.DataFrame(data_rows, columns=header)

    raise ValueError(
        "Could not find a sheet or section named 'Quinncia Metrics (All Students)' in the uploaded spreadsheet."
    )


def load_metrics_table(raw: bytes, filename: str) -> pd.DataFrame:
    """Load and normalize the Quinncia Metrics (All Students) table."""
    lower_name = filename.lower()
    if lower_name.endswith(".xlsx") or lower_name.endswith(".xlsm") or lower_name.endswith(".xls"):
        df = _read_excel_metrics(raw)
    else:
        df = _extract_csv_section(raw)

    df = df.copy()
    df.columns = [_canon_col(c) for c in df.columns]

    # Drop fully empty rows and trim string cells.
    df = df.replace({pd.NA: ""}).fillna("")
    for col in df.columns:
        df[col] = df[col].map(lambda x: str(x).strip())
    df = df[df.apply(lambda row: any(str(x).strip() for x in row), axis=1)]

    missing = sorted(REQUIRED_COLUMNS - set(df.columns))
    if missing:
        raise ValueError(
            "The metrics table is missing these required columns: " + ", ".join(missing)
        )

    return df


def format_metric_value(value: object) -> str:
    """Format values for PowerPoint exactly like the report table: no unnecessary .0 or trailing zeros."""
    if value is None:
        return ""
    raw = str(value).strip()
    if raw == "" or raw.lower() in {"nan", "none", "null", "<na>"}:
        return ""

    try:
        number = Decimal(raw)
    except (InvalidOperation, ValueError):
        return raw

    if number == number.to_integral_value():
        return str(int(number))

    # Avoid scientific notation and strip trailing zeros.
    return format(number.normalize(), "f").rstrip("0").rstrip(".")


def _set_cell_text(cell, text: str, black: bool = True) -> None:
    """Replace a table cell's text while keeping the existing paragraph/cell styling as much as possible."""
    text = "" if text is None else str(text)
    tf = cell.text_frame

    # Keep only the first paragraph.
    for paragraph in list(tf.paragraphs)[1:]:
        paragraph._element.getparent().remove(paragraph._element)

    paragraph = tf.paragraphs[0]

    # Keep only the first run, so alignment/font size from the template is retained.
    if paragraph.runs:
        run = paragraph.runs[0]
        run.text = text
        for extra_run in list(paragraph.runs)[1:]:
            extra_run._r.getparent().remove(extra_run._r)
    else:
        run = paragraph.add_run()
        run.text = text

    if black:
        run.font.color.rgb = BLACK


def _find_appendix_table(prs: Presentation):
    try:
        slide = prs.slides[TARGET_SLIDE_INDEX]
    except IndexError as exc:
        raise ValueError("The uploaded PowerPoint does not have a slide 2 to update.") from exc

    for shape in slide.shapes:
        if not shape.has_table:
            continue
        table = shape.table
        headers = [table.cell(0, c).text.strip() for c in range(len(table.columns))]
        header_set = set(headers)
        if {"Program", "Students", "Sign Ups", "Sign Up%"}.issubset(header_set):
            return table

    raise ValueError("Could not find the Appendix: Entire MSB table on slide 2.")


def update_powerpoint(pptx_bytes: bytes, metrics_bytes: bytes, metrics_filename: str) -> Tuple[bytes, UpdateSummary]:
    """Update slide 2 of a PowerPoint using Quinncia Metrics (All Students) data."""
    df = load_metrics_table(metrics_bytes, metrics_filename)
    data_by_major: Dict[str, pd.Series] = {
        _norm(row["major"]): row for _, row in df.iterrows() if str(row.get("major", "")).strip()
    }

    prs = Presentation(io.BytesIO(pptx_bytes))
    table = _find_appendix_table(prs)

    headers = [table.cell(0, c).text.strip() for c in range(len(table.columns))]
    missing_headers = [h for h in headers[1:] if h not in HEADER_TO_METRIC]
    if missing_headers:
        raise ValueError("Unrecognized table headers on slide 2: " + ", ".join(missing_headers))

    updated_rows = 0
    updated_cells = 0
    missing_programs: List[str] = []

    for r in range(1, len(table.rows)):
        display_program = table.cell(r, 0).text.strip()
        source_program = PROGRAM_LOOKUP.get(display_program, display_program)
        row = data_by_major.get(_norm(source_program))

        # Keep program labels from the PowerPoint and make body text black.
        _set_cell_text(table.cell(r, 0), display_program, black=True)

        if row is None:
            missing_programs.append(display_program)
            continue

        updated_rows += 1
        for c, ppt_header in enumerate(headers[1:], start=1):
            metric_col = HEADER_TO_METRIC[ppt_header]
            value = format_metric_value(row.get(metric_col, ""))
            _set_cell_text(table.cell(r, c), value, black=True)
            updated_cells += 1

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    summary = UpdateSummary(
        updated_rows=updated_rows,
        updated_cells=updated_cells,
        missing_programs=missing_programs,
    )
    return output.getvalue(), summary
