```python
import re
from io import BytesIO
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

from fpdf import FPDF


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Attendance Data Transformer",
    layout="wide"
)


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_SESSION_STATE = {
    "processed": False,
    "summary_df": None,
    "detailed_dfs": {},
    "sorted_class_names": [],
    "working_days": None,
    "file_uploader_key": 0,
    "student_working_days": {},
    "student_late_days": {},
    "student_very_late_days": {},
    "student_absent_days": {},
    "file_signature": None,
}


for key, value in DEFAULT_SESSION_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# APPLICATION RESET
# ============================================================

def reset_application():
    """Reset all processed data and user overrides."""

    st.session_state.processed = False
    st.session_state.summary_df = None
    st.session_state.detailed_dfs = {}
    st.session_state.sorted_class_names = []
    st.session_state.working_days = None

    st.session_state.student_working_days = {}
    st.session_state.student_late_days = {}
    st.session_state.student_very_late_days = {}
    st.session_state.student_absent_days = {}

    st.session_state.file_signature = None

    st.session_state.file_uploader_key += 1


# ============================================================
# PAGE HEADER
# ============================================================

st.title("Attendance Data Transformer")

st.markdown(
    """
This app transforms attendance summary data into detailed student
attendance records.

Upload your attendance summary Excel file and the app will generate:

- Class summary statistics
- Detailed student attendance records
- Excel report with formatting and highlighting
- PDF attendance report
"""
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean_column_names(df):
    """Remove accidental spaces from column names."""

    df = df.copy()

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    return df


def find_best_column(df, target_name, fallback=None):
    """
    Find a column using:
    1. Exact match
    2. Case-insensitive match
    3. Fuzzy matching
    """

    available_columns = list(df.columns)

    if target_name in available_columns:
        return target_name

    target_lower = str(target_name).strip().lower()

    for col in available_columns:
        if str(col).strip().lower() == target_lower:
            return col

    if not available_columns:
        return fallback

    match = process.extractOne(
        target_name,
        available_columns,
        scorer=fuzz.token_sort_ratio
    )

    if match and match[1] >= 70:
        return match[0]

    return fallback


def natural_sort_class_names(class_names):
    """
    Sort classes naturally.

    Examples:
        GRADE 01
        GRADE 02 - A
        GRADE 02 - B
        GRADE 03
    """

    def sort_key(name):
        name = str(name)

        numbers = re.findall(r"\d+", name)

        grade_number = int(numbers[0]) if numbers else 999

        section_match = re.search(
            r"-\s*([A-Za-z])$",
            name
        )

        section = (
            section_match.group(1).upper()
            if section_match
            else ""
        )

        return (
            grade_number,
            section,
            name.upper()
        )

    return sorted(
        list(class_names),
        key=sort_key
    )


def safe_number(series):
    """Convert a pandas series to numeric safely."""

    return pd.to_numeric(
        series,
        errors="coerce"
    ).fillna(0)


def safe_pdf_text(value):
    """
    FPDF 1.7.2's built-in Arial font uses Latin-1.
    Replace unsupported characters instead of crashing.
    """

    text = str(value)

    return (
        text
        .encode("latin-1", errors="replace")
        .decode("latin-1")
    )


# ============================================================
# DATE RANGE FROM FILENAME
# ============================================================

def extract_date_range(filename):
    """
    Extract dates from filenames such as:

    Attendance Summary_2026-05-04_2026-05-08_20260512113609.xlsx
    """

    try:
        matches = re.findall(
            r"(\d{4}-\d{2}-\d{2})",
            str(filename)
        )

        if len(matches) >= 2:

            start_date = datetime.strptime(
                matches[0],
                "%Y-%m-%d"
            )

            end_date = datetime.strptime(
                matches[1],
                "%Y-%m-%d"
            )

            return (
                start_date.strftime("%d.%m.%y"),
                end_date.strftime("%d.%m.%y")
            )

    except Exception:
        pass

    return "N/A", "N/A"


# ============================================================
# WORKING DAYS DETECTION
# ============================================================

def detect_working_days(df):
    """
    Automatically determine working days from:

        Present + Absent

    without modifying the original dataframe.
    """

    try:
        present_col = find_best_column(
            df,
            "Present"
        )

        absent_col = find_best_column(
            df,
            "Absent"
        )

        if not present_col or not absent_col:
            return None

        present = safe_number(
            df[present_col]
        )

        absent = safe_number(
            df[absent_col]
        )

        total_days = present + absent

        if len(total_days) == 0:
            return None

        maximum = total_days.max()

        if pd.isna(maximum):
            return None

        return int(maximum)

    except Exception:
        return None


# ============================================================
# EXCEL STYLING
# ============================================================

def apply_excel_styling(
    worksheet,
    title,
    is_summary=False,
    late_threshold=0,
    very_late_threshold=0,
    absent_threshold=0,
):
    """Apply professional formatting to an Excel worksheet."""

    header_font = Font(
        name="Aptos Display",
        size=12,
        bold=True
    )

    data_font = Font(
        name="Aptos Display",
        size=11
    )

    header_fill = PatternFill(
        start_color="DDEBF7",
        end_color="DDEBF7",
        fill_type="solid"
    )

    yellow_fill = PatternFill(
        start_color="FFFF00",
        end_color="FFFF00",
        fill_type="solid"
    )

    red_fill = PatternFill(
        start_color="F94949",
        end_color="F94949",
        fill_type="solid"
    )

    alignment_center = Alignment(
        horizontal="center",
        vertical="center"
    )

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = alignment_center
        cell.border = thin_border

    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    if worksheet.max_row >= 2:

        for row in worksheet.iter_rows(
            min_row=2,
            max_row=worksheet.max_row
        ):

            late_value = 0
            very_late_value = 0
            absent_value = 0

            if not is_summary:

                try:
                    absent_value = float(
                        row[4].value or 0
                    )
                except Exception:
                    absent_value = 0

                try:
                    late_value = float(
                        row[5].value or 0
                    )
                except Exception:
                    late_value = 0

                try:
                    very_late_value = float(
                        row[6].value or 0
                    )
                except Exception:
                    very_late_value = 0

            is_absent = (
                absent_value >= absent_threshold
                if not is_summary
                else False
            )

            for index, cell in enumerate(row):

                cell.font = data_font
                cell.border = thin_border
                cell.alignment = alignment_center

                if not is_summary:

                    # Absent:
                    # highlight Student Name + Absent column
                    if is_absent and index in [1, 4]:
                        cell.fill = red_fill

                    # Late:
                    elif (
                        late_value >= late_threshold
                        and index == 5
                    ):
                        cell.fill = yellow_fill

                    # Very Late:
                    elif (
                        very_late_value >= very_late_threshold
                        and index == 6
                    ):
                        cell.fill = yellow_fill

    # --------------------------------------------------------
    # COLUMN WIDTHS
    # --------------------------------------------------------

    if is_summary:

        column_widths = {
            "A": 20,
            "B": 18,
            "C": 22,
            "D": 15,
            "E": 15,
            "F": 15,
            "G": 18,
            "H": 28,
        }

    else:

        column_widths = {
            "A": 15,
            "B": 40,
            "C": 15,
            "D": 10,
            "E": 10,
            "F": 10,
            "G": 12,
            "H": 15,
            "I": 14,
        }

    for col, width in column_widths.items():
        worksheet.column_dimensions[col].width = width

    # --------------------------------------------------------
    # PERCENTAGE FORMAT
    # --------------------------------------------------------

    if worksheet.max_row >= 2:

        for row_number in range(
            2,
            worksheet.max_row + 1
        ):

            worksheet[
                f"H{row_number}"
            ].number_format = "0.00"

    # --------------------------------------------------------
    # SUMMARY-SPECIFIC STYLING
    # --------------------------------------------------------

    if is_summary:

        worksheet.freeze_panes = "A3"

        top_attendance = -1
        top_row_index = None

        for row_index, row in enumerate(
            worksheet.iter_rows(
                min_row=2,
                max_row=worksheet.max_row
            ),
            start=2
        ):

            try:

                attendance = float(
                    row[7].value
                )

                if attendance > top_attendance:

                    top_attendance = attendance
                    top_row_index = row_index

            except Exception:
                continue

        for row_index, row in enumerate(
            worksheet.iter_rows(
                min_row=2,
                max_row=worksheet.max_row
            ),
            start=2
        ):

            if row_index == top_row_index:

                fill = PatternFill(
                    start_color="60D276",
                    end_color="60D276",
                    fill_type="solid"
                )

            else:

                fill = PatternFill(
                    start_color=(
                        "F7F9FC"
                        if row_index % 2 == 0
                        else "FFFFFF"
                    ),
                    end_color=(
                        "F7F9FC"
                        if row_index % 2 == 0
                        else "FFFFFF"
                    ),
                    fill_type="solid"
                )

            for cell in row:

                cell.fill = fill

                cell.font = Font(
                    name="Aptos Display",
                    size=11
                )

                cell.alignment = alignment_center
                cell.border = thin_border

    # --------------------------------------------------------
    # TITLE ROW
    # --------------------------------------------------------

    worksheet.insert_rows(1)

    if is_summary:
        worksheet.merge_cells("A1:H1")
    else:
        worksheet.merge_cells("A1:I1")

    title_cell = worksheet["A1"]

    title_cell.value = title

    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    if is_summary:

        title_cell.font = Font(
            name="Aptos Display",
            size=30,
            bold=True
        )

        worksheet.row_dimensions[1].height = 45

    else:

        title_cell.font = Font(
            name="Aptos Display",
            size=24,
            bold=True
        )

    return worksheet


# ============================================================
# EXCEL GENERATION
# ============================================================

def to_excel_bytes(
    summary_df,
    detailed_dfs,
    sorted_class_names,
    uploaded_filename,
    late_threshold,
    very_late_threshold,
    absent_threshold,
):
    """Generate the formatted Excel report."""

    workbook = Workbook()

    # Remove default worksheet
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    # --------------------------------------------------------
    # SUMMARY SHEET
    # --------------------------------------------------------

    summary_sheet = workbook.create_sheet(
        "Class Summary"
    )

    for row in dataframe_to_rows(
        summary_df,
        index=False,
        header=True
    ):
        summary_sheet.append(row)

    start_date, end_date = extract_date_range(
        uploaded_filename
    )

    summary_title = (
        f"ATTENDANCE SUMMARY - "
        f"{start_date} - {end_date}"
    )

    apply_excel_styling(
        summary_sheet,
        summary_title,
        is_summary=True
    )

    # --------------------------------------------------------
    # CLASS SHEETS
    # --------------------------------------------------------

    used_sheet_names = {
        "Class Summary"
    }

    for class_name in sorted_class_names:

        if class_name not in detailed_dfs:
            continue

        # Excel sheet names have a 31-character limit
        base_name = str(class_name)[:31]

        sheet_name = base_name
        counter = 2

        while sheet_name in used_sheet_names:

            suffix = f" ({counter})"

            sheet_name = (
                str(class_name)[:31 - len(suffix)]
                + suffix
            )

            counter += 1

        used_sheet_names.add(sheet_name)

        class_sheet = workbook.create_sheet(
            sheet_name
        )

        class_df = detailed_dfs[class_name]

        for row in dataframe_to_rows(
            class_df,
            index=False,
            header=True
        ):
            class_sheet.append(row)

        apply_excel_styling(
            class_sheet,
            str(class_name),
            is_summary=False,
            late_threshold=late_threshold,
            very_late_threshold=very_late_threshold,
            absent_threshold=absent_threshold,
        )

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# PDF GENERATION
# ============================================================

def generate_pdf_report(
    summary_df,
    detailed_dfs,
    sorted_class_names,
    uploaded_filename,
    late_highlight_threshold,
    very_late_highlight_threshold,
    absent_highlight_threshold,
):
    """Generate a PDF attendance report."""

    start_date, end_date = extract_date_range(
        uploaded_filename
    )

    pdf = FPDF()

    pdf.set_auto_page_break(
        auto=True,
        margin=10
    )

    # ========================================================
    # SUMMARY PAGE
    # ========================================================

    pdf.add_page()

    pdf.set_font(
        "Arial",
        "B",
        16
    )

    pdf.cell(
        0,
        10,
        safe_pdf_text(
            f"ATTENDANCE SUMMARY - "
            f"{start_date} - {end_date}"
        ),
        ln=True,
        align="C"
    )

    pdf.ln(5)

    headers = [
        "Class",
        "Students",
        "Days",
        "Present",
        "Absent",
        "Late",
        "V.Late",
        "Attendance %"
    ]

    col_widths = [
        35,
        20,
        18,
        22,
        22,
        18,
        20,
        30
    ]

    pdf.set_font(
        "Arial",
        "B",
        9
    )

    for index, header in enumerate(headers):

        pdf.cell(
            col_widths[index],
            10,
            safe_pdf_text(header),
            border=1,
            align="C"
        )

    pdf.ln()

    pdf.set_font(
        "Arial",
        "",
        8
    )

    for _, row in summary_df.iterrows():

        values = [
            row.get("Class", ""),
            row.get("Total_Students", ""),
            row.get("Total_Working_Days", ""),
            row.get("Avg_Present", ""),
            row.get("Avg_Absent", ""),
            row.get("Avg_Late", ""),
            row.get("Avg_Very_Late", ""),
            row.get("Avg_Attendance_Percentage", ""),
        ]

        for index, value in enumerate(values):

            if index == 7:

                try:
                    value = f"{float(value):.2f}"
                except Exception:
                    value = str(value)

            value = safe_pdf_text(value)

            pdf.cell(
                col_widths[index],
                9,
                value,
                border=1,
                align="C"
            )

        pdf.ln()

    # ========================================================
    # DETAILED CLASS PAGES
    # ========================================================

    for class_name in sorted_class_names:

        if class_name not in detailed_dfs:
            continue

        df_detail = detailed_dfs[class_name]

        pdf.add_page()

        pdf.set_font(
            "Arial",
            "B",
            15
        )

        pdf.cell(
            0,
            10,
            safe_pdf_text(
                f"{class_name} "
                f"({start_date} - {end_date})"
            ),
            ln=True,
            align="C"
        )

        pdf.ln(4)

        headers = [
            "Adm No",
            "Student Name",
            "W.D",
            "P",
            "A",
            "L",
            "V.L",
            "Att%"
        ]

        widths = [
            18,
            70,
            12,
            12,
            12,
            12,
            12,
            18
        ]

        pdf.set_font(
            "Arial",
            "B",
            8
        )

        for index, header in enumerate(headers):

            pdf.cell(
                widths[index],
                8,
                safe_pdf_text(header),
                border=1,
                align="C"
            )

        pdf.ln()

        pdf.set_font(
            "Arial",
            "",
            7
        )

        for _, row in df_detail.iterrows():

            try:
                absent_value = float(
                    row["Absent"]
                )
            except Exception:
                absent_value = 0

            try:
                late_value = float(
                    row["Late"]
                )
            except Exception:
                late_value = 0

            try:
                very_late_value = float(
                    row["Very_Late"]
                )
            except Exception:
                very_late_value = 0

            is_absent = (
                absent_value
                >= absent_highlight_threshold
            )

            is_late = (
                late_value
                >= late_highlight_threshold
                or
                very_late_value
                >= very_late_highlight_threshold
            )

            if is_absent:

                pdf.set_fill_color(
                    249,
                    73,
                    73
                )

            elif is_late:

                pdf.set_fill_color(
                    255,
                    255,
                    0
                )

            else:

                pdf.set_fill_color(
                    255,
                    255,
                    255
                )

            try:
                attendance = float(
                    row["Attendance %"]
                )

                attendance_text = (
                    f"{attendance:.2f}"
                )

            except Exception:
                attendance_text = "0.00"

            values = [
                row.get("Admission No", ""),
                str(
                    row.get(
                        "Student Name",
                        ""
                    )
                )[:40],
                row.get("Working_Days", ""),
                row.get("Present", ""),
                row.get("Absent", ""),
                row.get("Late", ""),
                row.get("Very_Late", ""),
                attendance_text,
            ]

            for index, value in enumerate(values):

                pdf.cell(
                    widths[index],
                    7,
                    safe_pdf_text(value),
                    border=1,
                    align="C",
                    fill=True
                )

            pdf.ln()

    # FPDF 1.7.2 returns a string when dest='S'
    pdf_output = pdf.output(
        dest="S"
    )

    return pdf_output.encode(
        "latin-1",
        errors="replace"
    )


# ============================================================
# READ EXCEL FILE
# ============================================================

@st.cache_data(ttl=600)
def read_excel(file_bytes):
    """Read an XLSX file safely."""

    try:

        return pd.read_excel(
            BytesIO(file_bytes),
            engine="openpyxl"
        )

    except Exception as error:

        raise ValueError(
            f"Could not read the Excel file: {error}"
        )


# ============================================================
# PROCESS ATTENDANCE DATA
# ============================================================

def process_real_data(
    source_df,
    class_list,
    course_column,
    class_mapping,
    working_days,
):
    """Transform source attendance data."""

    df = source_df.copy()

    # --------------------------------------------------------
    # CLEAN COLUMN NAMES
    # --------------------------------------------------------

    df = clean_column_names(df)

    # --------------------------------------------------------
    # REQUIRED COLUMNS
    # --------------------------------------------------------

    required_columns = [
        "Admission No",
        "Student Name",
        "Present",
        "Absent",
    ]

    available_columns = list(
        df.columns
    )

    column_mapping = {}

    for required_column in required_columns:

        matched = find_best_column(
            df,
            required_column
        )

        if matched is None:

            raise ValueError(
                f"Required column "
                f"'{required_column}' "
                f"could not be found."
            )

        column_mapping[
            matched
        ] = required_column

    # --------------------------------------------------------
    # OPTIONAL COLUMNS
    # --------------------------------------------------------

    late_column = find_best_column(
        df,
        "Late"
    )

    very_late_column = find_best_column(
        df,
        "Very_Late"
    )

    if very_late_column is None:

        very_late_column = find_best_column(
            df,
            "Very Late"
        )

    if late_column:
        column_mapping[
            late_column
        ] = "Late"

    if very_late_column:
        column_mapping[
            very_late_column
        ] = "Very_Late"

    # --------------------------------------------------------
    # RENAME
    # --------------------------------------------------------

    df = df.rename(
        columns=column_mapping
    )

    # --------------------------------------------------------
    # OPTIONAL COLUMNS DEFAULTS
    # --------------------------------------------------------

    if "Late" not in df.columns:
        df["Late"] = 0

    if "Very_Late" not in df.columns:
        df["Very_Late"] = 0

    # --------------------------------------------------------
    # NUMERIC CONVERSION
    # --------------------------------------------------------

    for column in [
        "Present",
        "Absent",
        "Late",
        "Very_Late",
    ]:

        df[column] = safe_number(
            df[column]
        )

    # --------------------------------------------------------
    # STUDENT OVERRIDES
    # --------------------------------------------------------

    admission_column = "Admission No"

    if st.session_state.student_late_days:

        df["Late"] = (
            df[admission_column]
            .map(
                st.session_state.student_late_days
            )
            .fillna(df["Late"])
        )

    if st.session_state.student_very_late_days:

        df["Very_Late"] = (
            df[admission_column]
            .map(
                st.session_state.student_very_late_days
            )
            .fillna(df["Very_Late"])
        )

    if st.session_state.student_absent_days:

        df["Absent"] = (
            df[admission_column]
            .map(
                st.session_state.student_absent_days
            )
            .fillna(df["Absent"])
        )

    # --------------------------------------------------------
    # WORKING DAYS
    # --------------------------------------------------------

    if st.session_state.student_working_days:

        df["Working_Days"] = (
            df[admission_column]
            .map(
                st.session_state.student_working_days
            )
            .fillna(working_days)
        )

    else:

        df["Working_Days"] = working_days

    df["Working_Days"] = safe_number(
        df["Working_Days"]
    )

    # Prevent zero/negative working days
    df["Working_Days"] = df[
        "Working_Days"
    ].clip(lower=1)

    # --------------------------------------------------------
    # ABSENT
    # --------------------------------------------------------

    df["Absent"] = safe_number(
        df["Absent"]
    )

    df["Absent"] = df[
        "Absent"
    ].clip(lower=0)

    df["Absent"] = np.minimum(
        df["Absent"],
        df["Working_Days"]
    )

    # --------------------------------------------------------
    # PRESENT
    # --------------------------------------------------------

    df["Present"] = (
        df["Working_Days"]
        - df["Absent"]
    )

    df["Present"] = df[
        "Present"
    ].clip(lower=0)

    # --------------------------------------------------------
    # ATTENDANCE %
    # --------------------------------------------------------

    df["Attendance %"] = np.where(
        df["Working_Days"] > 0,
        (
            df["Present"]
            / df["Working_Days"]
        ) * 100,
        0
    )

    # --------------------------------------------------------
    # CLASS MAPPING
    # --------------------------------------------------------

    if course_column is None:

        # If no course column was detected,
        # try to use an existing Class column.

        existing_class_column = (
            find_best_column(
                df,
                "Class"
            )
        )

        if existing_class_column:

            df["Class"] = (
                df[existing_class_column]
                .astype(str)
                .str.strip()
            )

        else:

            raise ValueError(
                "Could not determine the class "
                "for the students."
            )

    else:

        if course_column not in df.columns:

            raise ValueError(
                f"Course/class column "
                f"'{course_column}' "
                f"could not be found."
            )

        df["Class"] = (
            df[course_column]
            .map(class_mapping)
        )

        # Handle unmapped values
        df["Class"] = df[
            "Class"
        ].fillna("UNASSIGNED")

        df["Class"] = (
            df["Class"]
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # GRADE 02 SECTION SPLITTING
    # --------------------------------------------------------

    has_grade_02 = (
        df["Class"]
        .eq("GRADE 02")
        .any()
    )

    has_batch_id = (
        "batch_id" in df.columns
    )

    if has_grade_02 and has_batch_id:

        def split_grade_2(row):

            if row["Class"] != "GRADE 02":
                return row["Class"]

            batch_value = str(
                row["batch_id"]
            ).strip()

            if not batch_value:
                return "GRADE 02 - UNKNOWN"

            # Try common formats:
            #
            # XYZ-A
            # XYZ-B
            # something-A-extra
            #
            parts = batch_value.split("-")

            section = None

            for part in reversed(parts):

                cleaned = (
                    str(part)
                    .strip()
                    .upper()
                )

                if cleaned in ["A", "B"]:

                    section = cleaned
                    break

            if section:

                return (
                    f"GRADE 02 - {section}"
                )

            # Also support values ending in A/B
            match = re.search(
                r"\b([AB])\b$",
                batch_value.upper()
            )

            if match:

                return (
                    f"GRADE 02 - "
                    f"{match.group(1)}"
                )

            return "GRADE 02 - UNKNOWN"

        df["Class"] = df.apply(
            split_grade_2,
            axis=1
        )

    elif has_grade_02 and not has_batch_id:

        st.warning(
            "The file does not contain "
            "'batch_id'. Grade 02 students "
            "will remain under GRADE 02."
        )

    # --------------------------------------------------------
    # BUILD FINAL CLASS LIST
    # --------------------------------------------------------

    actual_classes = (
        df["Class"]
        .dropna()
        .astype(str)
        .str.strip()
        .tolist()
    )

    final_class_list = []

    # Start with user-selected classes
    for class_name in class_list:

        if class_name not in final_class_list:
            final_class_list.append(
                class_name
            )

    # Add classes actually found in data
    for class_name in actual_classes:

        if class_name not in final_class_list:
            final_class_list.append(
                class_name
            )

    # --------------------------------------------------------
    # GROUP BY CLASS
    # --------------------------------------------------------

    detailed_dfs = {}

    for class_name in final_class_list:

        class_data = df[
            df["Class"] == class_name
        ].copy()

        if class_data.empty:
            continue

        output_columns = [
            "Admission No",
            "Student Name",
            "Working_Days",
            "Present",
            "Absent",
            "Late",
            "Very_Late",
            "Attendance %",
            "Class",
        ]

        output_columns = [
            column
            for column in output_columns
            if column in class_data.columns
        ]

        class_data = class_data[
            output_columns
        ]

        detailed_dfs[
            class_name
        ] = class_data.reset_index(
            drop=True
        )

    return detailed_dfs


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "Upload your attendance summary Excel file",
    type=["xlsx"],
    key=(
        f"file_uploader_"
        f"{st.session_state.file_uploader_key}"
    ),
    help=(
        "Upload an .xlsx attendance summary file."
    ),
)


if uploaded_file is None:

    st.info(
        "Upload your attendance summary file "
        "to continue."
    )

    st.stop()


# ============================================================
# DETECT NEW FILE
# ============================================================

file_bytes = uploaded_file.getvalue()

file_signature = (
    uploaded_file.name,
    len(file_bytes),
    hash(file_bytes)
)

if (
    st.session_state.file_signature is not None
    and st.session_state.file_signature
    != file_signature
):

    # New file uploaded.
    # Clear old processed results and overrides.

    st.session_state.processed = False
    st.session_state.summary_df = None
    st.session_state.detailed_dfs = {}
    st.session_state.sorted_class_names = []

    st.session_state.student_working_days = {}
    st.session_state.student_late_days = {}
    st.session_state.student_very_late_days = {}
    st.session_state.student_absent_days = {}

    st.session_state.working_days = None


st.session_state.file_signature = file_signature


# ============================================================
# READ FILE
# ============================================================

try:

    df = read_excel(
        file_bytes
    )

    if df is None or df.empty:

        st.error(
            "The uploaded Excel file is empty."
        )

        st.stop()

    df = clean_column_names(df)

except Exception as error:

    st.error(
        f"Could not read the uploaded Excel file: "
        f"{error}"
    )

    st.stop()


# ============================================================
# PREVIEW
# ============================================================

st.subheader(
    "Preview of Your Data"
)

st.dataframe(
    df.head(20),
    use_container_width=True
)


# ============================================================
# TRANSFORMATION SETTINGS
# ============================================================

st.subheader(
    "Transformation Settings"
)


# ============================================================
# COURSE / CLASS DETECTION
# ============================================================

course_column_candidates = [
    "course_name",
    "Course Name",
    "Class",
    "Grade",
    "Section",
]

course_column = None

for candidate in course_column_candidates:

    match = find_best_column(
        df,
        candidate
    )

    if match:

        course_column = match
        break


class_mapping = {}

if course_column:

    st.success(
        f"Detected course/class column: "
        f"'{course_column}'"
    )

    unique_courses = (
        df[course_column]
        .drop_duplicates()
        .tolist()
    )

    # Remove duplicate NaN entries safely
    cleaned_courses = []

    for course in unique_courses:

        if pd.isna(course):

            if "UNASSIGNED" not in cleaned_courses:
                cleaned_courses.append(
                    "UNASSIGNED"
                )

        else:

            course_text = str(
                course
            ).strip()

            if course_text not in cleaned_courses:
                cleaned_courses.append(
                    course_text
                )

    st.write(
        f"Found {len(cleaned_courses)} "
        f"unique course/class values."
    )

    st.subheader(
        "Course to Class Mapping"
    )

    st.write(
        "Map each course name to the "
        "standardized class name."
    )

    default_classes = {
        "7th Year": "GRADE 07",
        "6th Year": "GRADE 06",
        "5th Year": "GRADE 05",
        "4th Year": "GRADE 04",
        "3rd Year": "GRADE 03",
        "2nd Year": "GRADE 02",
        "1st Year": "GRADE 01",
    }

    for course in cleaned_courses:

        default_class = None

        course_text = str(
            course
        )

        for key, value in default_classes.items():

            if key.lower() in course_text.lower():

                default_class = value
                break

        if default_class is None:

            # Try extracting a number
            number_match = re.search(
                r"\d+",
                course_text
            )

            if number_match:

                default_class = (
                    f"GRADE "
                    f"{int(number_match.group()):02d}"
                )

        if default_class is None:

            default_class = (
                "UNASSIGNED"
                if course == "UNASSIGNED"
                else course_text
            )

        mapped_class = st.text_input(
            f"Map '{course}' to class:",
            value=default_class,
            key=f"map_{str(course)}",
        )

        # Store mapping against actual source value
        if course == "UNASSIGNED":

            # There may be NaN values in the source
            for original_value in (
                df[course_column]
                .unique()
            ):

                if pd.isna(original_value):

                    class_mapping[
                        original_value
                    ] = mapped_class.strip()

        else:

            class_mapping[
                course
            ] = mapped_class.strip()

    class_list = list(
        dict.fromkeys(
            class_mapping.values()
        )
    )

else:

    st.warning(
        "Could not detect a course/class "
        "column automatically."
    )

    class_names = st.text_area(
        "Enter class names (one per line)",
        value=(
            "GRADE 01\n"
            "GRADE 02\n"
            "GRADE 03\n"
            "GRADE 04\n"
            "GRADE 05\n"
            "GRADE 06\n"
            "GRADE 07"
        ),
        help=(
            "Enter the class names that should "
            "appear in the output. "
            "One class per line."
        ),
    )

    class_list = [
        name.strip()
        for name in class_names.splitlines()
        if name.strip()
    ]

    # If an existing Class column is available,
    # process_real_data() will use it.
    if not find_best_column(df, "Class"):

        st.error(
            "No course/class column was found. "
            "Please provide a course/class column "
            "in the Excel file."
        )

        st.stop()


# ============================================================
# WORKING DAYS
# ============================================================

st.subheader(
    "Working Days Settings"
)

auto_working_days = detect_working_days(
    df
)

if auto_working_days:

    st.success(
        f"Auto-detected working days: "
        f"{auto_working_days}"
    )

else:

    st.warning(
        "Could not automatically detect "
        "working days. Please enter them manually."
    )


use_manual_working_days = st.checkbox(
    "Override working days manually"
)


if (
    use_manual_working_days
    or not auto_working_days
):

    previous_working_days = (
        st.session_state.working_days
        if st.session_state.working_days
        else (
            auto_working_days
            if auto_working_days
            else 1
        )
    )

    working_days = st.number_input(
        "Total working days*",
        min_value=1,
        max_value=365,
        value=int(previous_working_days),
        step=1,
        help=(
            "Enter the total number of "
            "working days manually."
        ),
    )

else:

    working_days = auto_working_days


# ============================================================
# HIGHLIGHT SETTINGS
# ============================================================

st.subheader(
    "Late Comer Highlight Settings"
)

late_highlight_threshold = st.number_input(
    "Highlight students if Late days are "
    "greater than or equal to:",
    min_value=0,
    max_value=365,
    value=4,
    step=1,
    help=(
        "Late cells will be highlighted "
        "when Late days reach this number."
    ),
)

very_late_highlight_threshold = st.number_input(
    "Highlight students if Very Late days are "
    "greater than or equal to:",
    min_value=0,
    max_value=365,
    value=1,
    step=1,
    help=(
        "Very Late cells will be highlighted "
        "when Very Late days reach this number."
    ),
)

absent_highlight_threshold = st.number_input(
    "Highlight students if Absent days are "
    "greater than or equal to:",
    min_value=0,
    max_value=365,
    value=3,
    step=1,
    help=(
        "Student name and Absent cells will "
        "be highlighted red when Absent days "
        "reach this number."
    ),
)


# ============================================================
# INDIVIDUAL WORKING DAYS OVERRIDE
# ============================================================

override_working_days = st.checkbox(
    "Override working days for specific students",
    help=(
        "Set different working days for "
        "individual students."
    ),
)


if override_working_days:

    st.subheader(
        "Set Individual Working Days"
    )

    admission_col = find_best_column(
        df,
        "Admission No"
    )

    student_col = find_best_column(
        df,
        "Student Name"
    )

    if not admission_col or not student_col:

        st.error(
            "Could not find Admission No "
            "or Student Name columns."
        )

    else:

        temp_df = df[
            [admission_col, student_col]
        ].copy()

        temp_df.columns = [
            "Admission No",
            "Student Name",
        ]

        temp_df["Working_Days"] = (
            working_days
        )

        edited_df = st.data_editor(
            temp_df,
            use_container_width=True,
            key="working_days_editor",
            column_config={
                "Working_Days":
                    st.column_config.NumberColumn(
                        "Working Days",
                        min_value=1,
                        max_value=365,
                        step=1,
                    )
            },
        )

        st.session_state.student_working_days = dict(
            zip(
                edited_df["Admission No"],
                edited_df["Working_Days"]
            )
        )


# ============================================================
# INDIVIDUAL LATE / VERY LATE OVERRIDE
# ============================================================

override_late_days = st.checkbox(
    "Override Late / Very Late days "
    "for specific students",
    help=(
        "Manually edit Late and Very Late "
        "days for individual students."
    ),
)


if override_late_days:

    st.subheader(
        "Set Individual Late / Very Late Days"
    )

    admission_col = find_best_column(
        df,
        "Admission No"
    )

    student_col = find_best_column(
        df,
        "Student Name"
    )

    late_col = find_best_column(
        df,
        "Late"
    )

    very_late_col = find_best_column(
        df,
        "Very_Late"
    )

    if very_late_col is None:

        very_late_col = find_best_column(
            df,
            "Very Late"
        )

    if not admission_col or not student_col:

        st.error(
            "Could not find Admission No "
            "or Student Name columns."
        )

    else:

        temp_late_df = df[
            [admission_col, student_col]
        ].copy()

        temp_late_df.columns = [
            "Admission No",
            "Student Name",
        ]

        if late_col:

            temp_late_df["Late"] = safe_number(
                df[late_col]
            )

        else:

            temp_late_df["Late"] = 0

        if very_late_col:

            temp_late_df[
                "Very_Late"
            ] = safe_number(
                df[very_late_col]
            )

        else:

            temp_late_df[
                "Very_Late"
            ] = 0

        edited_late_df = st.data_editor(
            temp_late_df,
            use_container_width=True,
            key="late_days_editor",
            column_config={
                "Late":
                    st.column_config.NumberColumn(
                        "Late",
                        min_value=0,
                        max_value=365,
                        step=1,
                    ),
                "Very_Late":
                    st.column_config.NumberColumn(
                        "Very Late",
                        min_value=0,
                        max_value=365,
                        step=1,
                    ),
            },
        )

        st.session_state.student_late_days = dict(
            zip(
                edited_late_df["Admission No"],
                edited_late_df["Late"]
            )
        )

        st.session_state.student_very_late_days = dict(
            zip(
                edited_late_df["Admission No"],
                edited_late_df["Very_Late"]
            )
        )


# ============================================================
# INDIVIDUAL ABSENT OVERRIDE
# ============================================================

override_absent_days = st.checkbox(
    "Override Absent days "
    "for specific students",
    help=(
        "Manually change Absent days. "
        "Present and Attendance % will "
        "recalculate automatically."
    ),
)


if override_absent_days:

    st.subheader(
        "Set Individual Absent Days"
    )

    st.caption(
        "Set Absent to 0 to mark a student "
        "as fully present for the selected "
        "working days."
    )

    admission_col = find_best_column(
        df,
        "Admission No"
    )

    student_col = find_best_column(
        df,
        "Student Name"
    )

    absent_col = find_best_column(
        df,
        "Absent"
    )

    if not admission_col or not student_col:

        st.error(
            "Could not find Admission No "
            "or Student Name columns."
        )

    else:

        temp_absent_df = df[
            [admission_col, student_col]
        ].copy()

        temp_absent_df.columns = [
            "Admission No",
            "Student Name",
        ]

        if absent_col:

            temp_absent_df["Absent"] = (
                safe_number(
                    df[absent_col]
                )
                .round()
                .astype(int)
            )

        else:

            temp_absent_df["Absent"] = 0

        edited_absent_df = st.data_editor(
            temp_absent_df,
            use_container_width=True,
            key="absent_days_editor",
            column_config={
                "Absent":
                    st.column_config.NumberColumn(
                        "Absent",
                        min_value=0,
                        max_value=365,
                        step=1,
                        help=(
                            "Change this student's "
                            "absent days. Present and "
                            "Attendance % will "
                            "recalculate automatically."
                        ),
                    )
            },
        )

        st.session_state.student_absent_days = dict(
            zip(
                edited_absent_df[
                    "Admission No"
                ],
                edited_absent_df[
                    "Absent"
                ]
            )
        )


# ============================================================
# PROCESS BUTTON
# ============================================================

st.markdown("---")

process_button = st.button(
    "Process Attendance Data",
    type="primary",
    use_container_width=True,
)


# ============================================================
# PROCESS DATA
# ============================================================

if process_button:

    if working_days is None:

        st.error(
            "Please enter the total number "
            "of working days."
        )

        st.stop()

    if working_days <= 0:

        st.error(
            "Working days must be at least 1."
        )

        st.stop()

    try:

        with st.spinner(
            "Processing attendance data..."
        ):

            st.session_state.working_days = (
                working_days
            )

            detailed_dfs = process_real_data(
                df,
                class_list,
                course_column,
                class_mapping,
                working_days,
            )

        if not detailed_dfs:

            st.error(
                "No data was processed. "
                "Please check the class mapping "
                "and input file."
            )

            st.stop()

        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------

        summary_data = []

        sorted_class_names = (
            natural_sort_class_names(
                detailed_dfs.keys()
            )
        )

        for class_name in sorted_class_names:

            df_detail = (
                detailed_dfs[class_name]
            )

            summary_data.append(
                {
                    "Class": class_name,
                    "Total_Students": len(
                        df_detail
                    ),
                    "Total_Working_Days": round(
                        df_detail[
                            "Working_Days"
                        ].mean(),
                        2,
                    ),
                    "Avg_Present": round(
                        df_detail[
                            "Present"
                        ].mean(),
                        2,
                    ),
                    "Avg_Absent": round(
                        df_detail[
                            "Absent"
                        ].mean(),
                        2,
                    ),
                    "Avg_Late": round(
                        df_detail[
                            "Late"
                        ].mean(),
                        2,
                    ),
                    "Avg_Very_Late": round(
                        df_detail[
                            "Very_Late"
                        ].mean(),
                        2,
                    ),
                    "Avg_Attendance_Percentage":
                        round(
                            df_detail[
                                "Attendance %"
                            ].mean(),
                            2,
                        ),
                }
            )

        summary_df = pd.DataFrame(
            summary_data
        )

        # ----------------------------------------------------
        # SAVE RESULTS
        # ----------------------------------------------------

        st.session_state.processed = True

        st.session_state.summary_df = (
            summary_df
        )

        st.session_state.detailed_dfs = (
            detailed_dfs
        )

        st.session_state.sorted_class_names = (
            sorted_class_names
        )

        st.success(
            "Attendance data processed successfully!"
        )

    except Exception as error:

        st.error(
            "An error occurred while processing "
            f"the attendance data:\n\n{error}"
        )

        with st.expander(
            "Technical error details"
        ):

            st.exception(error)


# ============================================================
# SHOW RESULTS
# ============================================================

if st.session_state.processed:

    st.markdown("---")

    st.subheader(
        "Preview of Processed Data"
    )

    tab1, tab2 = st.tabs(
        [
            "Summary",
            "Detailed View",
        ]
    )

    # --------------------------------------------------------
    # SUMMARY TAB
    # --------------------------------------------------------

    with tab1:

        st.write(
            "Class Summary"
        )

        st.dataframe(
            st.session_state.summary_df,
            use_container_width=True,
        )

    # --------------------------------------------------------
    # DETAIL TAB
    # --------------------------------------------------------

    with tab2:

        available_classes = (
            st.session_state.sorted_class_names
        )

        if available_classes:

            selected_class = st.selectbox(
                "Select class to view details",
                options=available_classes,
            )

            st.dataframe(
                st.session_state.detailed_dfs[
                    selected_class
                ],
                use_container_width=True,
            )

    # ========================================================
    # GENERATE EXCEL
    # ========================================================

    try:

        excel_bytes = to_excel_bytes(
            st.session_state.summary_df,
            st.session_state.detailed_dfs,
            st.session_state.sorted_class_names,
            uploaded_file.name,
            late_highlight_threshold,
            very_late_highlight_threshold,
            absent_highlight_threshold,
        )

    except Exception as error:

        excel_bytes = None

        st.error(
            f"Could not generate Excel report: "
            f"{error}"
        )

    # ========================================================
    # GENERATE PDF
    # ========================================================

    try:

        pdf_bytes = generate_pdf_report(
            st.session_state.summary_df,
            st.session_state.detailed_dfs,
            st.session_state.sorted_class_names,
            uploaded_file.name,
            late_highlight_threshold,
            very_late_highlight_threshold,
            absent_highlight_threshold,
        )

    except Exception as error:

        pdf_bytes = None

        st.error(
            f"Could not generate PDF report: "
            f"{error}"
        )

    # ========================================================
    # DOWNLOAD BUTTONS
    # ========================================================

    col1, col2 = st.columns(2)

    with col1:

        if excel_bytes is not None:

            st.download_button(
                label=(
                    "Download Detailed "
                    "Attendance Report (Excel)"
                ),
                data=excel_bytes,
                file_name=(
                    "detailed_attendance_report.xlsx"
                ),
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    with col2:

        if pdf_bytes is not None:

            st.download_button(
                label=(
                    "Download Attendance "
                    "Report (PDF)"
                ),
                data=pdf_bytes,
                file_name=(
                    "attendance_report.pdf"
                ),
                mime="application/pdf",
                use_container_width=True,
            )

    st.info(
        """
        **PDF Export Notice**

        The PDF is generated using FPDF and uses standard
        fonts. If your student names contain Sinhala,
        Tamil, Arabic, or other non-Latin characters,
        those characters may be replaced in the PDF.

        The Excel report is recommended when you need
        the original names and full formatting.
        """
    )

    # ========================================================
    # NEW FILE
    # ========================================================

    st.markdown("---")

    if st.button(
        "Add a New File",
        key="reset_button",
        use_container_width=True,
    ):

        reset_application()

        st.rerun()


# ============================================================
# INSTRUCTIONS
# ============================================================

st.markdown("---")

st.subheader(
    "Instructions"
)

st.markdown(
    """
### What this app does

1. Upload your attendance summary `.xlsx` file.
2. The app detects the course/class column.
3. Map course names to standardized class names.
4. Working days are automatically detected from
   **Present + Absent**.
5. You can manually override:
   - Working Days
   - Late Days
   - Very Late Days
   - Absent Days
6. Grade 02 can automatically be divided into
   **GRADE 02 - A** and **GRADE 02 - B** when
   `batch_id` is available.
7. The app calculates:
   - Present
   - Absent
   - Attendance %
   - Average Late
   - Average Very Late
8. The Excel report contains:
   - Class Summary
   - Individual class sheets
   - Highlighting for late students
   - Highlighting for absent students
   - Top-performing class
9. A PDF report can also be downloaded.
"""
)
```
