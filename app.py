# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from rapidfuzz import process, fuzz
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF
from datetime import datetime

st.set_page_config(page_title="Attendance Data Transformer", layout="wide")

st.title("Attendance Data Transformer")
st.markdown(
    """
This app transforms your attendance summary data into detailed student attendance records.
Upload your attendance summary Excel file and the app will generate detailed student records.
"""
)

# ---------- Session State Initialisation ----------
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'summary_df' not in st.session_state:
    st.session_state.summary_df = None
if 'detailed_dfs' not in st.session_state:
    st.session_state.detailed_dfs = {}
if 'sorted_class_names' not in st.session_state:
    st.session_state.sorted_class_names = []
if 'working_days' not in st.session_state:
    st.session_state.working_days = None
if 'file_uploader_key' not in st.session_state:
    st.session_state.file_uploader_key = 0
if 'student_working_days' not in st.session_state:
    st.session_state.student_working_days = {}
if 'student_late_days' not in st.session_state:
    st.session_state.student_late_days = {}
if 'student_very_late_days' not in st.session_state:
    st.session_state.student_very_late_days = {}
if 'student_absent_days' not in st.session_state:
    st.session_state.student_absent_days = {}


# ---------- Helper Functions ----------
def reset_application():
    """Reset all session state to allow a fresh upload."""
    st.session_state.processed = False
    st.session_state.summary_df = None
    st.session_state.detailed_dfs = {}
    st.session_state.sorted_class_names = []
    st.session_state.working_days = None
    st.session_state.file_uploader_key += 1
    st.session_state.student_working_days = {}
    st.session_state.student_late_days = {}
    st.session_state.student_very_late_days = {}
    st.session_state.student_absent_days = {}


def apply_excel_styling(worksheet, title, is_summary=False, late_threshold=0,
                        very_late_threshold=0, absent_threshold=0):
    """Apply consistent styling to an Excel worksheet."""
    header_font = Font(name='Aptos Display', size=12, bold=True)
    data_font = Font(name='Aptos Display', size=12)
    header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    red_fill = PatternFill(start_color="F94949", end_color="F94949", fill_type="solid")
    alignment_center = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # Header row
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = alignment_center
        cell.border = thin_border

    # Data rows
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        late_value = row[5].value if row[5].value is not None else 0
        very_late_value = row[6].value if row[6].value is not None else 0
        absent_value = row[4].value if row[4].value is not None else 0

        is_absent = absent_value >= absent_threshold

        for idx, cell in enumerate(row):
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = alignment_center

            if not is_summary:
                # Highlight absent student name + absent column in red
                if is_absent and idx in [1, 4]:
                    cell.fill = red_fill
                # Highlight late column in yellow
                elif late_value >= late_threshold and idx == 5:
                    cell.fill = yellow_fill
                # Highlight very late column in yellow
                elif very_late_value >= very_late_threshold and idx == 6:
                    cell.fill = yellow_fill

    # Column widths and freeze panes
    if is_summary:
        for col, width in {'A': 20, 'B': 18, 'C': 22, 'D': 15, 'E': 15,
                           'F': 15, 'G': 18, 'H': 28}.items():
            worksheet.column_dimensions[col].width = width
        worksheet.freeze_panes = "A3"
    else:
        for col, width in {'A': 15, 'B': 40, 'C': 15, 'D': 10, 'E': 10,
                           'F': 10, 'G': 12, 'H': 15, 'I': 14}.items():
            worksheet.column_dimensions[col].width = width

    # Format percentage column
    for row in range(2, worksheet.max_row + 1):
        worksheet[f'H{row}'].number_format = '0.00'

    # Insert and style title row
    worksheet.insert_rows(1)
    if is_summary:
        worksheet.merge_cells('A1:H1')
    else:
        worksheet.merge_cells('A1:I1')
    title_cell = worksheet['A1']
    title_cell.value = title
    title_cell.font = Font(name='Aptos Display', size=36 if is_summary else 24, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    if is_summary:
        worksheet.row_dimensions[1].height = 45

    # Additional summary styling: alternating row colours & top class highlight
    if is_summary:
        # Find row with highest attendance %
        top_attendance = -1
        top_row_index = None
        for idx, row in enumerate(worksheet.iter_rows(min_row=3, max_row=worksheet.max_row), start=3):
            try:
                attendance = float(row[7].value)
                if attendance > top_attendance:
                    top_attendance = attendance
                    top_row_index = idx
            except (ValueError, TypeError):
                pass

        for i, row in enumerate(worksheet.iter_rows(min_row=3, max_row=worksheet.max_row), start=3):
            if i == top_row_index:
                fill = PatternFill(start_color="60D276", end_color="60D276", fill_type="solid")
            else:
                fill = PatternFill(start_color="F7F9FC" if i % 2 == 0 else "FFFFFF",
                                   end_color="F7F9FC" if i % 2 == 0 else "FFFFFF", fill_type="solid")
            for cell in row:
                cell.fill = fill
                cell.font = Font(name='Aptos Display', size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border

    return worksheet


def detect_working_days(df):
    """Try to detect working days from Present + Absent columns."""
    try:
        if 'Present' in df.columns and 'Absent' in df.columns:
            df['__total_days__'] = df['Present'] + df['Absent']
            return int(df['__total_days__'].max())
    except Exception:
        pass
    return None


def find_best_column(df, target_name, fallback=None):
    """Fuzzy match a column name in the DataFrame."""
    available = df.columns.tolist()
    if target_name in available:
        return target_name
    for col in available:
        if str(col).strip().lower() == target_name.strip().lower():
            return col
    match = process.extractOne(target_name, available, scorer=fuzz.token_sort_ratio)
    if match and match[1] > 60:
        return match[0]
    return fallback


def extract_date_range(filename):
    """Extract start and end dates from filename (format: YYYY-MM-DD)."""
    try:
        matches = re.findall(r'(\d{4}-\d{2}-\d{2})', filename)
        if len(matches) >= 2:
            start = datetime.strptime(matches[0], "%Y-%m-%d")
            end = datetime.strptime(matches[1], "%Y-%m-%d")
            return start.strftime("%d.%m.%y"), end.strftime("%d.%m.%y")
    except Exception:
        pass
    return "N/A", "N/A"


def sort_class_names(class_names):
    """Sort class names naturally (GRADE 01, GRADE 02 - A, ...)."""
    def key(name):
        nums = re.findall(r'\d+', name)
        grade_num = int(nums[0]) if nums else 999
        section_match = re.search(r'-\s*([A-Z])$', name)
        section = section_match.group(1) if section_match else ''
        return grade_num, section
    return sorted(class_names, key=key)


# ---------- Core Processing ----------
def process_real_data(df, class_list, course_column, class_mapping, working_days):
    """Transform raw data into detailed per-class DataFrames."""
    detailed_dfs = {}

    # Ensure batch_id exists for Grade 02 splitting
    if 'batch_id' not in df.columns:
        st.error("The column 'batch_id' is required to split Grade 02 into sections.")
        st.stop()

    # Fuzzy match required columns
    required = ['Admission No', 'Student Name', 'Present', 'Absent']
    column_mapping = {}
    available = df.columns.tolist()
    for col in required:
        match = process.extractOne(col, available, scorer=fuzz.token_sort_ratio)
        column_mapping[col] = match[0] if match and match[1] > 60 else col

    # Optional columns
    for opt in ['Late', 'Very_Late', 'Very Late']:
        match = process.extractOne(opt, available, scorer=fuzz.token_sort_ratio)
        if match and match[1] > 60:
            column_mapping[opt] = match[0]

    df = df.rename(columns=column_mapping)

    # Fill missing optional columns with 0
    if 'Late' not in df.columns:
        df['Late'] = 0
    if 'Very_Late' not in df.columns:
        df['Very_Late'] = df['Very Late'] if 'Very Late' in df.columns else 0

    # Convert numeric columns
    for col in ['Present', 'Absent', 'Late', 'Very_Late']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Apply manual overrides from session state
    if st.session_state.student_late_days:
        df['Late'] = df['Admission No'].map(st.session_state.student_late_days).fillna(df['Late'])
    if st.session_state.student_very_late_days:
        df['Very_Late'] = df['Admission No'].map(st.session_state.student_very_late_days).fillna(df['Very_Late'])
    if st.session_state.student_absent_days:
        df['Absent'] = df['Admission No'].map(st.session_state.student_absent_days).fillna(df['Absent'])

    # Assign working days per student
    if st.session_state.student_working_days:
        df['Working_Days'] = df['Admission No'].map(st.session_state.student_working_days).fillna(working_days)
    else:
        df['Working_Days'] = working_days
    df['Working_Days'] = pd.to_numeric(df['Working_Days'], errors='coerce').fillna(working_days)

    # Ensure Absent stays within bounds
    df['Absent'] = df['Absent'].clip(lower=0)
    df['Absent'] = np.minimum(df['Absent'], df['Working_Days'])

    # Recalculate Present and Attendance %
    df['Present'] = (df['Working_Days'] - df['Absent']).clip(lower=0)
    df['Attendance %'] = np.where(df['Working_Days'] > 0,
                                  (df['Present'] / df['Working_Days']) * 100, 0)

    # Map course names to class names
    df['Class'] = df[course_column].map(class_mapping)

    # Split Grade 02 into sections using batch_id
    def split_grade2(row):
        if row['Class'] == 'GRADE 02':
            try:
                section = str(row['batch_id']).split('-')[1]
                return f"GRADE 02 - {section}"
            except (IndexError, AttributeError):
                return "GRADE 02 - UNKNOWN"
        return row['Class']

    df['Class'] = df.apply(split_grade2, axis=1)

    # Update class_list to include split sections while preserving order
    new_class_list = []
    for cls in class_list:
        if cls == 'GRADE 02':
            for sec in ['GRADE 02 - A', 'GRADE 02 - B']:
                if sec not in new_class_list:
                    new_class_list.append(sec)
        else:
            if cls not in new_class_list:
                new_class_list.append(cls)
    class_list = new_class_list

    # Build detailed DataFrames per class
    output_cols = ['Admission No', 'Student Name', 'Working_Days', 'Present',
                   'Absent', 'Late', 'Very_Late', 'Attendance %', 'Class']
    for class_name in class_list:
        class_data = df[df['Class'] == class_name].copy()
        if class_data.empty:
            continue
        class_data = class_data[[c for c in output_cols if c in class_data.columns]]
        detailed_dfs[class_name] = class_data

    return detailed_dfs


# ---------- File Export Functions ----------
def to_excel_bytes(summary_df, detailed_dfs, sorted_class_names,
                   late_threshold, very_late_threshold, absent_threshold,
                   uploaded_filename):
    """Create a styled Excel workbook in memory."""
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    # Summary sheet
    ws_summary = wb.create_sheet("Class Summary")
    for r in dataframe_to_rows(summary_df, index=False, header=True):
        ws_summary.append(r)
    start_date, end_date = extract_date_range(uploaded_filename)
    summary_title = f"ATTENDANCE SUMMARY - {start_date} - {end_date}"
    apply_excel_styling(ws_summary, summary_title, is_summary=True)

    # Detailed class sheets
    for class_name in sorted_class_names:
        if class_name in detailed_dfs:
            sheet_name = class_name[:31]  # Excel sheet name limit
            ws_class = wb.create_sheet(sheet_name)
            for r in dataframe_to_rows(detailed_dfs[class_name], index=False, header=True):
                ws_class.append(r)
            apply_excel_styling(ws_class, class_name, is_summary=False,
                                late_threshold=late_threshold,
                                very_late_threshold=very_late_threshold,
                                absent_threshold=absent_threshold)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def generate_pdf_report(summary_df, detailed_dfs, sorted_class_names,
                        uploaded_filename, late_threshold, very_late_threshold,
                        absent_threshold):
    """Create a PDF report and return bytes."""
    start_date, end_date = extract_date_range(uploaded_filename)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=10)

    # ---- Summary Page ----
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f"ATTENDANCE SUMMARY - {start_date} - {end_date}", ln=True, align='C')
    pdf.ln(5)

    headers = ["Class", "Students", "Days", "Present", "Absent", "Late", "V.Late", "Attendance %"]
    col_widths = [35, 20, 18, 22, 22, 18, 20, 30]
    pdf.set_font("Arial", 'B', 10)
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 10, h, border=1, align='C')
    pdf.ln()
    pdf.set_font("Arial", '', 9)
    for _, row in summary_df.iterrows():
        values = [str(row["Class"]), str(row["Total_Students"]),
                  str(row["Total_Working_Days"]), str(row["Avg_Present"]),
                  str(row["Avg_Absent"]), str(row["Avg_Late"]),
                  str(row["Avg_Very_Late"]), f'{row["Avg_Attendance_Percentage"]:.2f}']
        for i, v in enumerate(values):
            pdf.cell(col_widths[i], 9, v, border=1, align='C')
        pdf.ln()

    # ---- Detailed Class Pages ----
    for class_name in sorted_class_names:
        if class_name not in detailed_dfs:
            continue
        df_detail = detailed_dfs[class_name]
        pdf.add_page()
        pdf.set_font("Arial", 'B', 15)
        pdf.cell(0, 10, f"{class_name} ({start_date} - {end_date})", ln=True, align='C')
        pdf.ln(4)

        headers = ["Adm No", "Student Name", "W.D", "P", "A", "L", "V.L", "Att%"]
        widths = [18, 70, 12, 12, 12, 12, 12, 18]
        pdf.set_font("Arial", 'B', 8)
        for i, h in enumerate(headers):
            pdf.cell(widths[i], 8, h, border=1, align='C')
        pdf.ln()
        pdf.set_font("Arial", '', 7)

        for _, row in df_detail.iterrows():
            values = [str(row["Admission No"]), str(row["Student Name"])[:40],
                      str(row["Working_Days"]), str(row["Present"]),
                      str(row["Absent"]), str(row["Late"]),
                      str(row["Very_Late"]), f'{row["Attendance %"]:.2f}']
            is_absent = row["Absent"] >= absent_threshold
            is_late = (row["Late"] >= late_threshold or
                       row["Very_Late"] >= very_late_threshold)

            if is_absent:
                pdf.set_fill_color(249, 73, 73)  # red
            elif is_late:
                pdf.set_fill_color(255, 255, 0)  # yellow
            else:
                pdf.set_fill_color(255, 255, 255)

            for i, v in enumerate(values):
                pdf.cell(widths[i], 7, v, border=1, align='C', fill=True)
            pdf.ln()

    return pdf.output(dest='S').encode('latin1')


# ---------- Cached File Reader ----------
@st.cache_data(ttl=600)
def read_excel(file):
    """Read Excel file into a DataFrame."""
    try:
        return pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None


# ======================== MAIN APP ========================
uploaded_file = st.file_uploader(
    "Upload your attendance summary Excel file",
    type=["xls", "xlsx"],
    key=f"file_uploader_{st.session_state.file_uploader_key}"
)

if not uploaded_file:
    st.info("Upload your attendance summary file to continue.")
    st.stop()

df = read_excel(uploaded_file)
if df is None:
    st.stop()

st.subheader("Preview of your data")
st.dataframe(df.head())

# ---- User Configuration ----
st.subheader("Transformation Settings")

# Course/class column detection
course_column_candidates = ['course_name', 'Course Name', 'Class', 'Grade', 'Section']
course_column = next((c for c in course_column_candidates if c in df.columns), None)

if course_column:
    st.success(f"Detected course/class column: '{course_column}'")
    unique_courses = df[course_column].unique().tolist()
    st.write(f"Found {len(unique_courses)} unique course/class values:")
    st.write(unique_courses)

    st.subheader("Course to Class Mapping")
    st.write("Please map each course name to a standardized class name:")
    class_mapping = {}
    default_classes = {
        "7th Year": "GRADE 07", "6th Year": "GRADE 06", "5th Year": "GRADE 05",
        "4th Year": "GRADE 04", "3rd Year": "GRADE 03", "2nd Year": "GRADE 02",
        "1st Year": "GRADE 01"
    }
    for course in unique_courses:
        if pd.isna(course):
            default = "UNASSIGNED"
        else:
            default = next((v for k, v in default_classes.items() if k in str(course)), f"GRADE {course}")
        mapped = st.text_input(f"Map '{course}' to class:", value=default, key=f"map_{course}")
        class_mapping[course] = mapped.strip()
    class_list = list(set(class_mapping.values()))
else:
    st.warning("Could not detect a course/class column in your data.")
    class_names = st.text_area(
        "Enter class names (one per line)",
        value="GRADE 01\nGRADE 02\nGRADE 03\nGRADE 04\nGRADE 05\nGRADE 06\nGRADE 07",
        help="One class per line."
    )
    class_list = [name.strip() for name in class_names.split('\n') if name.strip()]
    class_mapping = {}

# Working days
st.subheader("Working Days Settings")
auto_working_days = detect_working_days(df)
if auto_working_days:
    st.success(f"Auto-detected working days from file: {auto_working_days}")
else:
    st.warning("Could not auto-detect working days. Please enter manually.")

use_manual = st.checkbox("Override working days manually")
if use_manual or not auto_working_days:
    working_days = st.number_input(
        "Total working days*", min_value=1, max_value=365,
        value=st.session_state.working_days if st.session_state.working_days else 1)
else:
    working_days = auto_working_days

# Highlight thresholds
st.subheader("Late Comer Highlight Settings")
late_threshold = st.number_input("Highlight if Late days >= ", min_value=0, max_value=365, value=4)
very_late_threshold = st.number_input("Highlight if Very Late days >= ", min_value=0, max_value=365, value=1)
absent_threshold = st.number_input("Highlight if Absent days >= ", min_value=0, max_value=365, value=3)

# ---- Override Options ----
override_working = st.checkbox("Override working days for specific students")
if override_working:
    st.subheader("Set Individual Working Days")
    temp_df = df[['Admission No', 'Student Name']].copy()
    temp_df['Working_Days'] = working_days
    edited_df = st.data_editor(temp_df, use_container_width=True,
                               column_config={"Working_Days": st.column_config.NumberColumn("Working Days",
                                                                                            min_value=1, max_value=365)})
    st.session_state.student_working_days = dict(zip(edited_df['Admission No'], edited_df['Working_Days']))

override_late = st.checkbox("Override Late / Very Late days for specific students")
if override_late:
    st.subheader("Set Individual Late / Very Late Days")
    temp_late = df[['Admission No', 'Student Name']].copy()
    temp_late['Late'] = df['Late'] if 'Late' in df.columns else 0
    temp_late['Very_Late'] = df['Very_Late'] if 'Very_Late' in df.columns else (df['Very Late'] if 'Very Late' in df.columns else 0)
    edited_late = st.data_editor(temp_late, use_container_width=True,
                                 column_config={
                                     "Late": st.column_config.NumberColumn("Late", min_value=0, max_value=365),
                                     "Very_Late": st.column_config.NumberColumn("Very Late", min_value=0, max_value=365)
                                 })
    st.session_state.student_late_days = dict(zip(edited_late['Admission No'], edited_late['Late']))
    st.session_state.student_very_late_days = dict(zip(edited_late['Admission No'], edited_late['Very_Late']))

override_absent = st.checkbox("Override Absent days for specific students")
if override_absent:
    st.subheader("Set Individual Absent Days")
    admission_col = find_best_column(df, 'Admission No', 'Admission No')
    student_col = find_best_column(df, 'Student Name', 'Student Name')
    absent_col = find_best_column(df, 'Absent', None)
    temp_absent = df[[admission_col, student_col]].copy()
    temp_absent.columns = ['Admission No', 'Student Name']
    temp_absent['Absent'] = pd.to_numeric(df[absent_col], errors='coerce').fillna(0).astype(int) if absent_col else 0
    edited_absent = st.data_editor(temp_absent, use_container_width=True,
                                   column_config={"Absent": st.column_config.NumberColumn("Absent", min_value=0, max_value=365)})
    st.session_state.student_absent_days = dict(zip(edited_absent['Admission No'], edited_absent['Absent']))

# ---- Process Data ----
process_button = st.button("Process Attendance Data")
if process_button:
    if working_days is None or working_days <= 0:
        st.error("Please enter a valid number of working days (minimum 1).")
        st.stop()
    st.session_state.working_days = working_days

    detailed_dfs = process_real_data(df, class_list, course_column, class_mapping, working_days)
    if not detailed_dfs:
        st.error("No data was processed. Please check your input and try again.")
        st.stop()

    sorted_class_names = sort_class_names(detailed_dfs.keys())

    # Build summary DataFrame
    summary_data = []
    for cname in sorted_class_names:
        d = detailed_dfs[cname]
        summary_data.append({
            "Class": cname,
            "Total_Students": len(d),
            "Total_Working_Days": round(d["Working_Days"].mean(), 2),
            "Avg_Present": round(d["Present"].mean(), 2),
            "Avg_Absent": round(d["Absent"].mean(), 2),
            "Avg_Late": round(d["Late"].mean(), 2),
            "Avg_Very_Late": round(d["Very_Late"].mean(), 2),
            "Avg_Attendance_Percentage": round(d["Attendance %"].mean(), 2)
        })
    summary_df = pd.DataFrame(summary_data)

    st.session_state.processed = True
    st.session_state.summary_df = summary_df
    st.session_state.detailed_dfs = detailed_dfs
    st.session_state.sorted_class_names = sorted_class_names

# ---- Display Results ----
if st.session_state.processed:
    st.subheader("Preview of Processed Data")
    tab1, tab2 = st.tabs(["Summary", "Detailed View"])
    with tab1:
        st.dataframe(st.session_state.summary_df)
    with tab2:
        selected = st.selectbox("Select class to view details", options=st.session_state.sorted_class_names)
        st.dataframe(st.session_state.detailed_dfs[selected])

    # Create download files
    excel_bytes = to_excel_bytes(
        st.session_state.summary_df,
        st.session_state.detailed_dfs,
        st.session_state.sorted_class_names,
        late_threshold,
        very_late_threshold,
        absent_threshold,
        uploaded_file.name
    )

    pdf_bytes = generate_pdf_report(
        st.session_state.summary_df,
        st.session_state.detailed_dfs,
        st.session_state.sorted_class_names,
        uploaded_file.name,
        late_threshold,
        very_late_threshold,
        absent_threshold
    )

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="Download Detailed Attendance Report (Excel)",
            data=excel_bytes,
            file_name="detailed_attendance_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        st.info(
            "**PDF Export Notice**\n\n"
            "The PDF export feature is currently under development. "
            "Some formatting issues may still be present.\n\n"
            "For the best experience, I recommend that Ustadh may use the Excel download option."
        )
        st.download_button(
            label="Download Attendance Report (PDF)",
            data=pdf_bytes,
            file_name="attendance_report.pdf",
            mime="application/pdf",
            help="PDF export is still under development. Formatting may not be perfect."
        )

    st.success("Attendance data processed successfully! Download the files above.")
    if st.button("Add a new file", key="reset_button"):
        reset_application()
        st.rerun()

elif not process_button:
    st.info("Click the button above to process your attendance data based on your settings.")

st.markdown("---")
st.subheader("Instructions")
st.markdown("""
The app will create:
- A summary sheet with class statistics
- Separate sheets for each class with detailed student attendance records (ordered by class name)
""")
