"""Utilities for updating Quinncia Appendix tables in a PowerPoint deck.

This module updates:
- Slide 1: KPI overview, from the Overall Marriott School row in "Quinncia Metrics (All Students)"
- Slide 2: Appendix: Entire MSB, from "Quinncia Metrics (All Students)"
- Slide 3: Class of 2027 KPI overview, from the Overall Marriott School row in "Quinncia Metrics (Class of 2027 and Above)"
- Slide 4: Appendix: Class of 2027, from "Quinncia Metrics (Class of 2027 and Above)"

It keeps the PowerPoint's program labels and formatting, while replacing only the values.
"""

from __future__ import annotations

import io
import re
import zipfile
from datetime import date
from xml.sax.saxutils import escape
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, List, Tuple

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt

ALL_STUDENTS_SECTION = "Quinncia Metrics (All Students)"
CLASS_2027_SECTION = "Quinncia Metrics (Class of 2027 and Above)"
BLACK = RGBColor(0, 0, 0)

# PowerPoint display names -> names used in the Quinncia export.
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
class SlideUpdateSummary:
    slide_number: int
    section_used: str
    updated_rows: int
    updated_cells: int
    missing_programs: List[str]


@dataclass
class UpdateSummary:
    slide_summaries: List[SlideUpdateSummary]
    kpi_slides_updated: List[int]

    @property
    def updated_rows(self) -> int:
        return sum(s.updated_rows for s in self.slide_summaries)

    @property
    def updated_cells(self) -> int:
        return sum(s.updated_cells for s in self.slide_summaries)

    @property
    def missing_programs(self) -> List[str]:
        missing: List[str] = []
        for s in self.slide_summaries:
            missing.extend([f"Slide {s.slide_number}: {p}" for p in s.missing_programs])
        return missing

    @property
    def slide_updated(self) -> str:
        slides = sorted(set(self.kpi_slides_updated + [s.slide_number for s in self.slide_summaries]))
        return ", ".join(str(slide) for slide in slides)


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


def _clean_metrics_table(df: pd.DataFrame, section_title: str) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_canon_col(c) for c in df.columns]

    df = df.replace({pd.NA: ""}).fillna("")
    for col in df.columns:
        df[col] = df[col].map(lambda x: str(x).strip())
    df = df[df.apply(lambda row: any(str(x).strip() for x in row), axis=1)]

    missing = sorted(REQUIRED_COLUMNS - set(df.columns))
    if missing:
        raise ValueError(
            f"The section '{section_title}' is missing these required columns: " + ", ".join(missing)
        )
    return df


def _extract_csv_section(raw: bytes, section_title: str) -> pd.DataFrame:
    text = _decode_bytes(raw)
    lines = text.splitlines()
    target = _norm(section_title)

    start = None
    for i, line in enumerate(lines):
        cleaned = line.strip().strip(",")
        if _norm(cleaned) == target:
            start = i + 1
            break

    if start is None:
        # Fallback only for the all-students section if a simple one-table CSV is uploaded.
        if _norm(section_title) == _norm(ALL_STUDENTS_SECTION):
            return pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False)
        raise ValueError(f"Could not find the section '{section_title}' in the uploaded CSV.")

    section_lines: List[str] = []
    for line in lines[start:]:
        if not line.strip():
            if section_lines:
                break
            continue
        section_lines.append(line)

    if not section_lines:
        raise ValueError(f"Found '{section_title}', but no table was found underneath it.")

    return pd.read_csv(io.StringIO("\n".join(section_lines)), dtype=str, keep_default_na=False)


def _read_excel_metrics(raw: bytes, section_title: str) -> pd.DataFrame:
    excel = pd.ExcelFile(io.BytesIO(raw))
    target_norm = _norm(section_title)

    for sheet_name in excel.sheet_names:
        sheet_norm = _norm(sheet_name)
        if sheet_norm == target_norm:
            return pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, dtype=str, keep_default_na=False)

    for sheet_name in excel.sheet_names:
        scan = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, header=None, dtype=str, keep_default_na=False)
        for row_idx in range(len(scan)):
            row_values = [str(x).strip() for x in scan.iloc[row_idx].tolist()]
            normalized_row = [_norm(x) for x in row_values]
            if target_norm in normalized_row:
                header_idx = row_idx + 1
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

    # Simple-table fallback for all-students section only.
    if _norm(section_title) == _norm(ALL_STUDENTS_SECTION):
        for sheet_name in excel.sheet_names:
            scan = pd.read_excel(io.BytesIO(raw), sheet_name=sheet_name, header=None, dtype=str, keep_default_na=False)
            for row_idx in range(len(scan)):
                normalized_row = [_norm(x) for x in scan.iloc[row_idx].tolist()]
                if "major" in normalized_row:
                    header = [str(x).strip() for x in scan.iloc[row_idx].tolist()]
                    data_rows = []
                    for data_idx in range(row_idx + 1, len(scan)):
                        row = scan.iloc[data_idx].tolist()
                        if not any(str(x).strip() for x in row):
                            break
                        data_rows.append(row)
                    if data_rows:
                        return pd.DataFrame(data_rows, columns=header)

    raise ValueError(f"Could not find the section '{section_title}' in the uploaded spreadsheet.")


def load_metrics_table(raw: bytes, filename: str, section_title: str = ALL_STUDENTS_SECTION) -> pd.DataFrame:
    lower_name = filename.lower()
    if lower_name.endswith((".xlsx", ".xlsm", ".xls")):
        df = _read_excel_metrics(raw, section_title)
    else:
        df = _extract_csv_section(raw, section_title)
    return _clean_metrics_table(df, section_title)


def format_metric_value(value: object) -> str:
    """Format values for PowerPoint: no unnecessary .0 or trailing zeros."""
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
    return format(number.normalize(), "f").rstrip("0").rstrip(".")


def _set_cell_text(cell, text: str, black: bool = True) -> None:
    text = "" if text is None else str(text)
    tf = cell.text_frame

    for paragraph in list(tf.paragraphs)[1:]:
        paragraph._element.getparent().remove(paragraph._element)

    paragraph = tf.paragraphs[0]
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
    if run.font.size is None:
        run.font.size = Pt(9)


def _find_appendix_table(prs: Presentation, slide_index: int, slide_number: int):
    try:
        slide = prs.slides[slide_index]
    except IndexError as exc:
        raise ValueError(f"The uploaded PowerPoint does not have a slide {slide_number} to update.") from exc

    for shape in slide.shapes:
        if not shape.has_table:
            continue
        table = shape.table
        headers = [table.cell(0, c).text.strip() for c in range(len(table.columns))]
        if {"Program", "Students", "Sign Ups", "Sign Up%"}.issubset(set(headers)):
            return table

    raise ValueError(f"Could not find the Appendix table on slide {slide_number}.")


def _update_one_table(prs: Presentation, df: pd.DataFrame, slide_index: int, slide_number: int, section_title: str) -> SlideUpdateSummary:
    data_by_major: Dict[str, pd.Series] = {
        _norm(row["major"]): row for _, row in df.iterrows() if str(row.get("major", "")).strip()
    }

    table = _find_appendix_table(prs, slide_index, slide_number)
    headers = [table.cell(0, c).text.strip() for c in range(len(table.columns))]
    missing_headers = [h for h in headers[1:] if h not in HEADER_TO_METRIC]
    if missing_headers:
        raise ValueError(f"Unrecognized table headers on slide {slide_number}: " + ", ".join(missing_headers))

    updated_rows = 0
    updated_cells = 0
    missing_programs: List[str] = []

    for r in range(1, len(table.rows)):
        display_program = table.cell(r, 0).text.strip()
        source_program = PROGRAM_LOOKUP.get(display_program, display_program)
        row = data_by_major.get(_norm(source_program))

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

    return SlideUpdateSummary(
        slide_number=slide_number,
        section_used=section_title,
        updated_rows=updated_rows,
        updated_cells=updated_cells,
        missing_programs=missing_programs,
    )



def _percent_text(value: object) -> str:
    """Format a decimal percentage as one display decimal, using normal half-up rounding."""
    raw = str(value).strip()
    if not raw:
        return "0.0%"
    try:
        number = Decimal(raw) * Decimal("100")
    except (InvalidOperation, ValueError):
        return raw
    number = number.quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    return f"{number}%"


def _overall_metrics_row(df: pd.DataFrame, section_title: str) -> pd.Series:
    for _, row in df.iterrows():
        if _norm(row.get("major", "")) == _norm("Overall Marriott School"):
            return row
    if len(df) > 0:
        # In the Quinncia export, the final row is the total row if the label changes.
        return df.iloc[-1]
    raise ValueError(f"The section '{section_title}' does not contain any data rows.")


def _ordinal_suffix(day: int) -> str:
    if 11 <= day % 100 <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def _date_text_parts(current_date: date) -> Tuple[str, str, str]:
    month = current_date.strftime("%B")
    return f"{month} {current_date.day}", _ordinal_suffix(current_date.day), f", {current_date.year}"


def _kpi_text_replacements(row: pd.Series, current_date: date) -> Dict[int, str]:
    month_day, suffix, year_text = _date_text_parts(current_date)
    return {
        1: _percent_text(row.get("signup_pct", "")),
        2: f"{format_metric_value(row.get('quinncia_sign_ups', ''))} / {format_metric_value(row.get('enrolled_students', ''))} ",
        4: _percent_text(row.get("resume_student_pct", "")),
        5: f"{format_metric_value(row.get('resume_students', ''))} ",
        7: f"{format_metric_value(row.get('resume_uploads', ''))} ",
        11: _percent_text(row.get("linkedin_student_pct", "")),
        12: f"{format_metric_value(row.get('linkedin_students', ''))} ",
        14: f"{format_metric_value(row.get('linkedin_uploads', ''))} ",
        18: _percent_text(row.get("interview_student_pct", "")),
        19: f"{format_metric_value(row.get('interview_students', ''))} ",
        21: f"{format_metric_value(row.get('interview_uploads', ''))} ",
        25: _percent_text(row.get("all_three_student_pct", "")),
        26: f"{format_metric_value(row.get('all_three_students', ''))} ",
        41: month_day,
        42: suffix,
        43: year_text,
    }


def _replace_text_nodes_by_index(xml_text: str, replacements: Dict[int, str]) -> str:
    pattern = re.compile(r"(<a:t[^>]*>)(.*?)(</a:t>)", flags=re.DOTALL)
    index = -1

    def replace(match: re.Match) -> str:
        nonlocal index
        index += 1
        if index not in replacements:
            return match.group(0)
        return match.group(1) + escape(str(replacements[index])) + match.group(3)

    return pattern.sub(replace, xml_text)


def _make_red_text_black(xml_text: str) -> str:
    return re.sub(r'(<a:srgbClr\b[^>]*\bval=")FF0000("[^>]*/?>)', r'\g<1>000000\2', xml_text, flags=re.IGNORECASE)


def _update_kpi_slides_xml(pptx_bytes: bytes, all_students_df: pd.DataFrame, class_2027_df: pd.DataFrame) -> bytes:
    replacements_by_slide = {
        1: _kpi_text_replacements(_overall_metrics_row(all_students_df, ALL_STUDENTS_SECTION), date.today()),
        3: _kpi_text_replacements(_overall_metrics_row(class_2027_df, CLASS_2027_SECTION), date.today()),
    }

    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zin, zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            match = re.match(r"ppt/slides/slide(1|3)\.xml$", item.filename)
            if match:
                slide_number = int(match.group(1))
                xml_text = data.decode("utf-8")
                xml_text = _replace_text_nodes_by_index(xml_text, replacements_by_slide[slide_number])
                xml_text = _make_red_text_black(xml_text)
                data = xml_text.encode("utf-8")
            zout.writestr(item, data)
    output.seek(0)
    return output.getvalue()

def update_powerpoint(pptx_bytes: bytes, metrics_bytes: bytes, metrics_filename: str) -> Tuple[bytes, UpdateSummary]:
    """Update KPI slides 1/3 and appendix table slides 2/4 using Quinncia metrics data."""
    all_students_df = load_metrics_table(metrics_bytes, metrics_filename, ALL_STUDENTS_SECTION)
    class_2027_df = load_metrics_table(metrics_bytes, metrics_filename, CLASS_2027_SECTION)

    prs = Presentation(io.BytesIO(pptx_bytes))
    summaries = [
        _update_one_table(prs, all_students_df, slide_index=1, slide_number=2, section_title=ALL_STUDENTS_SECTION),
        _update_one_table(prs, class_2027_df, slide_index=3, slide_number=4, section_title=CLASS_2027_SECTION),
    ]

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    updated_bytes = _update_kpi_slides_xml(output.getvalue(), all_students_df, class_2027_df)
    return updated_bytes, UpdateSummary(slide_summaries=summaries, kpi_slides_updated=[1, 3])
