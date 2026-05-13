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
import tempfile
from fpdf import FPDF
import base64
from datetime import datetime

st.set_page_config(page_title="Attendance Data Transformer", layout="wide")

st.title("Attendance Data Transformer")
st.markdown(
    """
This app transforms your attendance summary data into detailed student attendance records.
Upload your attendance summary Excel file and the app will generate detailed student records.
"""
)

# Initialize session state variables
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

# Function to reset the application
def reset_application():
    st.session_state.processed = False
    st.session_state.summary_df = None
    st.session_state.detailed_dfs = {}
    st.session_state.sorted_class_names = []
    st.session_state.working_days = None
    st.session_state.file_uploader_key += 1  # Change the key to reset the file uploader


# Function to apply Excel styling
def apply_excel_styling(
    worksheet,
    title,
    is_summary=False,
    student_names=None,
    late_threshold=0,
    very_late_threshold=0,
    absent_threshold=0
):

    header_font = Font(name='Aptos Display', size=12, bold=True)
    data_font = Font(name='Aptos Display', size=12)

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

    # 🔴 RED for Absent
    red_fill = PatternFill(
        start_color="F94949",
        end_color="F94949",
        fill_type="solid"
    )

    alignment_center = Alignment(
        horizontal='center',
        vertical='center'
    )

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # HEADER STYLE
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = alignment_center
        cell.border = thin_border

    # DATA STYLE
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):

        late_value = 0
        very_late_value = 0
        absent_value = 0

        try:
            late_value = row[5].value if row[5].value is not None else 0
            very_late_value = row[6].value if row[6].value is not None else 0
            absent_value = row[4].value if row[4].value is not None else 0
        except:
            pass

        # Conditions
        is_absent = absent_value >= absent_threshold

        should_highlight = (
            late_value >= late_threshold or
            very_late_value >= very_late_threshold
        )

        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = alignment_center

            # 🔴 ABSENT (highest priority)
            if is_absent and not is_summary:
                cell.fill = red_fill

            # 🟡 LATE / VERY LATE
            elif should_highlight and not is_summary:
                cell.fill = yellow_fill

    # COLUMN WIDTHS
    if is_summary:
        column_widths = {
            'A': 20,  # Class
            'B': 18,  # Total Students
            'C': 22,  # Working Days
            'D': 15,  # Avg Present
            'E': 15,  # Avg Absent
            'F': 15,  # Avg Late
            'G': 18,  # Avg Very Late
            'H': 28   # Attendance %
        }
    
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
    
        # ❄️ Freeze header + title
        worksheet.freeze_panes = "A3"
    
        # 🏆 FIND TOP PERFORMING CLASS (HIGHEST ATTENDANCE %)
        top_attendance = -1
        top_row_index = None
    
        for idx, row in enumerate(
            worksheet.iter_rows(min_row=3, max_row=worksheet.max_row),
            start=3
        ):
            try:
                attendance = float(row[7].value)
                if attendance > top_attendance:
                    top_attendance = attendance
                    top_row_index = idx
            except:
                pass
    
        # 🎨 DASHBOARD ROW STYLING
        for i, row in enumerate(
            worksheet.iter_rows(min_row=3, max_row=worksheet.max_row),
            start=3
        ):
            
            # Base alternating row color
            base_fill = PatternFill(
                start_color="F7F9FC" if i % 2 == 0 else "FFFFFF",
                end_color="F7F9FC" if i % 2 == 0 else "FFFFFF",
                fill_type="solid"
            )
    
            # 🟢 TOP CLASS HIGHLIGHT
            if i == top_row_index:
                fill = PatternFill(
                    start_color="C6EFCE",
                    end_color="C6EFCE",
                    fill_type="solid"
                )
            else:
                fill = base_fill
    
            for cell in row:
                cell.fill = fill
                cell.font = Font(name='Aptos Display', size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
                
    else:
        column_widths = {
            'A': 15, 'B': 40, 'C': 15, 'D': 10,
            'E': 10, 'F': 10, 'G': 12, 'H': 15, 'I': 14
        }

    for col, width in column_widths.items():
        worksheet.column_dimensions[col].width = width

    # FORMAT PERCENTAGE COLUMN
    for row in range(2, worksheet.max_row + 1):
        worksheet[f'H{row}'].number_format = '0.00'

    # TITLE ROW
    worksheet.insert_rows(1)
    
    if is_summary:
        worksheet.merge_cells('A1:H1')
    else:
        worksheet.merge_cells('A1:I1')
    
    title_cell = worksheet['A1']
    title_cell.value = title
    
    # 🔥 SUMMARY TITLE STYLE
    if is_summary:
        title_cell.font = Font(
            name='Aptos Display',
            size=36,
            bold=True
        )
    
        worksheet.row_dimensions[1].height = 45
    
    else:
        title_cell.font = Font(
            name='Aptos Display',
            size=14,
            bold=True
        )
    
    title_cell.alignment = Alignment(
        horizontal='center',
        vertical='center'
    )
    
    return worksheet

# Function to detect working days automatically
def detect_working_days(df):
    try:
        if 'Present' in df.columns and 'Absent' in df.columns:
            df['__total_days__'] = df['Present'] + df['Absent']
            return int(df['__total_days__'].max())
    except:
        pass
    return None

# Function to extract date range from uploaded filename
def extract_date_range(filename):
    """
    Example filename:
    Attendance Summary_2026-05-04_2026-05-08_20260512113609.xlsx
    """

    try:
        matches = re.findall(r'(\d{4}-\d{2}-\d{2})', filename)

        if len(matches) >= 2:
            start_date = datetime.strptime(matches[0], "%Y-%m-%d")
            end_date = datetime.strptime(matches[1], "%Y-%m-%d")

            return (
                start_date.strftime("%d.%m.%y"),
                end_date.strftime("%d.%m.%y")
            )

    except:
        pass

    return ("N/A", "N/A")

# Function to generate PDF report
def generate_pdf_report(summary_df, detailed_dfs, sorted_class_names, uploaded_filename):

    start_date, end_date = extract_date_range(uploaded_filename)

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=10)

    # =========================
    # SUMMARY PAGE
    # =========================
    pdf.add_page()

    # Title
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(
        0,
        10,
        f"ATTENDANCE SUMMARY - {start_date} - {end_date}",
        ln=True,
        align='C'
    )

    pdf.ln(5)

    # Summary table
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

    col_widths = [35, 20, 18, 22, 22, 18, 20, 30]

    # Header
    pdf.set_font("Arial", 'B', 10)

    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, border=1, align='C')

    pdf.ln()

    # Data
    pdf.set_font("Arial", '', 9)

    for _, row in summary_df.iterrows():

        values = [
            str(row["Class"]),
            str(row["Total_Students"]),
            str(row["Total_Working_Days"]),
            str(row["Avg_Present"]),
            str(row["Avg_Absent"]),
            str(row["Avg_Late"]),
            str(row["Avg_Very_Late"]),
            f'{row["Avg_Attendance_Percentage"]:.2f}'
        ]

        for i, value in enumerate(values):
            pdf.cell(col_widths[i], 9, value, border=1, align='C')

        pdf.ln()

    # =========================
    # DETAILED CLASS PAGES
    # =========================
    for class_name in sorted_class_names:

        if class_name not in detailed_dfs:
            continue

        df_detail = detailed_dfs[class_name]

        pdf.add_page()

        # Class title
        pdf.set_font("Arial", 'B', 15)
        pdf.cell(
            0,
            10,
            f"{class_name} ({start_date} - {end_date})",
            ln=True,
            align='C'
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

        widths = [18, 70, 12, 12, 12, 12, 12, 18]

        # Header
        pdf.set_font("Arial", 'B', 8)

        for i, h in enumerate(headers):
            pdf.cell(widths[i], 8, h, border=1, align='C')

        pdf.ln()

        # Rows
        pdf.set_font("Arial", '', 7)

        for _, row in df_detail.iterrows():

            values = [
                str(row["Admission No"]),
                str(row["Student Name"])[:40],
                str(row["Working_Days"]),
                str(row["Present"]),
                str(row["Absent"]),
                str(row["Late"]),
                str(row["Very_Late"]),
                f'{row["Attendance %"]:.2f}'
            ]

            # Highlight logic
            is_absent = row["Absent"] >= absent_highlight_threshold

            is_late = (
                row["Late"] >= late_highlight_threshold or
                row["Very_Late"] >= very_late_highlight_threshold
            )

            if is_absent:
                pdf.set_fill_color(249, 73, 73)  # red
            elif is_late:
                pdf.set_fill_color(255, 255, 0)  # yellow
            else:
                pdf.set_fill_color(255, 255, 255)

            for i, value in enumerate(values):
                pdf.cell(
                    widths[i],
                    7,
                    value,
                    border=1,
                    align='C',
                    fill=True
                )

            pdf.ln()

    return pdf.output(dest='S').encode('latin1')
    
    # Save to bytes buffer
    pdf_bytes = pdf.output(dest='S').encode('latin1')
    return pdf_bytes

# --- Upload files
uploaded_file = st.file_uploader(
    "Upload your attendance summary Excel file", 
    type=["xls", "xlsx"],
    key=f"file_uploader_{st.session_state.file_uploader_key}"
)

if not uploaded_file:
    st.info("Upload your attendance summary file to continue.")
    st.stop()

# --- Read file
@st.cache_data(ttl=600)
def read_excel(file, sheet_name_hint=""):
    try:
        if sheet_name_hint:
            return pd.read_excel(file, sheet_name=sheet_name_hint, engine="openpyxl")
        else:
            return pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        try:
            return pd.read_excel(file, engine="openpyxl")
        except Exception as e2:
            st.error(f"Error reading file: {e2}")
            return None

try:
    df = read_excel(uploaded_file)
    
    if df is None:
        st.error("Could not read the uploaded Excel file.")
        st.stop()
        
except Exception as e:
    st.error(f"Could not read the uploaded Excel file: {e}")
    st.stop()

st.subheader("Preview of your data")
st.dataframe(df.head())

# --- Get user input for the transformation
st.subheader("Transformation Settings")

# Try to detect course_name column
course_column_candidates = ['course_name', 'Course Name', 'Class', 'Grade', 'Section']
course_column = None

for candidate in course_column_candidates:
    if candidate in df.columns:
        course_column = candidate
        break

if course_column:
    st.success(f"Detected course/class column: '{course_column}'")
    
    # Show unique values in the course column
    unique_courses = df[course_column].unique().tolist()
    st.write(f"Found {len(unique_courses)} unique course/class values:")
    st.write(unique_courses)
    
    # Create mapping from course names to standardized class names
    st.subheader("Course to Class Mapping")
    st.write("Please map each course name to a standardized class name:")
    
    class_mapping = {}
    default_classes = {
        "7th Year": "GRADE 07",
        "6th Year": "GRADE 06", 
        "5th Year": "GRADE 05",
        "4th Year": "GRADE 04",
        "3rd Year": "GRADE 03",
        "2nd Year": "GRADE 02",
        "1st Year": "GRADE 01"
    }
    
    for course in unique_courses:
        # Handle NaN values
        if pd.isna(course):
            default_class = "UNASSIGNED"
        else:
            # Try to find a default mapping
            default_class = None
            for key, value in default_classes.items():
                if key in str(course):
                    default_class = value
                    break
        
        # Let user confirm or change the mapping
        mapped_class = st.text_input(
            f"Map '{course}' to class:", 
            value=default_class if default_class else f"GRADE {course}",
            key=f"map_{course}"
        )
        class_mapping[course] = mapped_class.strip()
    
    # Get the list of classes from the mapping
    class_list = list(set(class_mapping.values()))
    
else:
    st.warning("Could not detect a course/class column in your data.")
    class_names = st.text_area(
        "Enter class names (one per line)", 
        value="GRADE 01\nGRADE 02\nGRADE 03\nGRADE 04\nGRADE 05\nGRADE 06\nGRADE 07",
        help="Enter the class names that should appear in the output. One class per line."
    )
    class_list = [name.strip() for name in class_names.split('\n') if name.strip()]
    class_mapping = {}

# --- Working Days Settings
st.subheader("Working Days Settings")

# Auto detect working days from Present + Absent
auto_working_days = detect_working_days(df)

if auto_working_days:
    st.success(f"Auto-detected working days from file: {auto_working_days}")
else:
    st.warning("Could not auto-detect working days. Please enter manually.")

# Manual override option
use_manual = st.checkbox("Override working days manually")

if use_manual or not auto_working_days:

    working_days = st.number_input(
        "Total working days*",
        min_value=1,
        max_value=365,
        value=st.session_state.working_days if st.session_state.working_days else 1,
        help="Enter total working days manually"
    )

else:
    working_days = auto_working_days

# --- Highlight settings
st.subheader("Late Comer Highlight Settings")

late_highlight_threshold = st.number_input(
    "Highlight students if Late days are greater than or equal to:",
    min_value=0,
    max_value=365,
    value=4,
    step=1,
    help="Student names will be highlighted in yellow if their Late days reach this number"
)

very_late_highlight_threshold = st.number_input(
    "Highlight students if Very Late days are greater than or equal to:",
    min_value=0,
    max_value=365,
    value=1,
    step=1,
    help="Student names will be highlighted in yellow if their Very Late days reach this number"
)

absent_highlight_threshold = st.number_input(
    "Highlight students if Absent days are greater than or equal to:",
    min_value=0,
    max_value=365,
    value=3,
    step=1,
    help="Student rows will be highlighted in red if Absent days reach this number"
)

# Optional: Override working days for specific students
override_working_days = st.checkbox(
    "Override working days for specific students",
    help="Enable this to set different working days for selected students"
)

if override_working_days:
    st.subheader("Set Individual Working Days")

    temp_df = df[['Admission No', 'Student Name']].copy()
    temp_df['Working_Days'] = working_days

    edited_df = st.data_editor(
        temp_df,
        use_container_width=True,
        column_config={
            "Working_Days": st.column_config.NumberColumn(
                "Working Days",
                min_value=1,
                max_value=365,
                step=1
            )
        }
    )

    # Store per-student working days
    st.session_state.student_working_days = dict(
        zip(edited_df['Admission No'], edited_df['Working_Days'])
    )

# Function to sort class names in natural order (GRADE 01, GRADE 02, etc.)
def sort_class_names(class_names):

    def sort_key(name):
        # Extract grade number
        numbers = re.findall(r'\d+', name)
        grade_num = int(numbers[0]) if numbers else 999

        # Extract section (A, B, etc.)
        match = re.search(r'-\s*([A-Z])$', name)
        section = match.group(1) if match else ''

        return (grade_num, section)

    return sorted(class_names, key=sort_key)

#Function to detect working days
def detect_working_days(df):
    try:
        if 'Present' in df.columns and 'Absent' in df.columns:
            df['__total_days__'] = df['Present'] + df['Absent']
            return int(df['__total_days__'].max())
    except:
        pass
    return None
    
# Function to create Excel file
def to_excel_bytes(
    summary_df,
    detailed_dfs,
    sorted_class_names,
    late_threshold,
    very_late_threshold):
    # Create a workbook
    wb = Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Add summary sheet
    ws_summary = wb.create_sheet("Class Summary")
    
    # Write summary data
    for r in dataframe_to_rows(summary_df, index=False, header=True):
        ws_summary.append(r)
    
    # Apply styling to summary sheet
    # Extract date range from uploaded filename
    start_date, end_date = extract_date_range(uploaded_file.name)
    
    summary_title = f"ATTENDANCE SUMMARY - {start_date} - {end_date}"
    
    ws_summary = apply_excel_styling(
        ws_summary,
        summary_title,
        is_summary=True
    )
    
    # Add class sheets
    for class_name in sorted_class_names:
        if class_name in detailed_dfs:
            # Shorten sheet name if too long for Excel
            sheet_name = class_name[:31] if len(class_name) > 31 else class_name
            ws_class = wb.create_sheet(sheet_name)  # Fixed: Changed create() to create_sheet()
            
            # Write class data
            for r in dataframe_to_rows(detailed_dfs[class_name], index=False, header=True):
                ws_class.append(r)
            
            # Get student names for this class to calculate optimal column width
            student_names = detailed_dfs[class_name]['Student Name'].tolist() if 'Student Name' in detailed_dfs[class_name].columns else []
            
            # Apply styling to class sheet
            ws_class = apply_excel_styling(
                ws_class,
                class_name,
                is_summary=False,
                student_names=student_names,
                late_threshold=late_highlight_threshold,
                very_late_threshold=very_late_highlight_threshold,
                absent_threshold=absent_highlight_threshold
            )
    
    # Save to bytes
    towrite = BytesIO()
    wb.save(towrite)
    towrite.seek(0)
    return towrite

# --- Process the real data
def process_real_data(df, class_list, course_column, class_mapping, working_days):
    detailed_dfs = {}
    
    # Ensure batch_id exists
    if 'batch_id' not in df.columns:
        st.error("The column 'batch_id' is required to split Grade 02 into sections.")
        st.stop()
    
    # Ensure we have the required columns
    required_columns = ['Admission No', 'Student Name', 'Present', 'Absent']
    available_columns = df.columns.tolist()
    
    # Fuzzy match required columns
    column_mapping = {}
    for req_col in required_columns:
        match = process.extractOne(req_col, available_columns, scorer=fuzz.token_sort_ratio)
        if match and match[1] > 60:
            column_mapping[req_col] = match[0]
        else:
            column_mapping[req_col] = req_col
            st.warning(f"Could not find a matching column for '{req_col}'.")
    
    # Optional columns
    optional_columns = ['Late', 'Very_Late', 'Very Late']
    for opt_col in optional_columns:
        match = process.extractOne(opt_col, available_columns, scorer=fuzz.token_sort_ratio)
        if match and match[1] > 60:
            column_mapping[opt_col] = match[0]
    
    df = df.rename(columns=column_mapping)
    
    # Fill missing optional columns
    if 'Late' not in df.columns:
        df['Late'] = 0
    if 'Very_Late' not in df.columns:
        if 'Very Late' in df.columns:
            df['Very_Late'] = df['Very Late']
        else:
            df['Very_Late'] = 0
    
    # Working days logic
    if 'student_working_days' in st.session_state and st.session_state.student_working_days:
        df['Working_Days'] = df['Admission No'].map(
            st.session_state.student_working_days
        ).fillna(working_days)
    else:
        df['Working_Days'] = working_days
    
    # Attendance %
    df['Attendance %'] = (df['Present'] / df['Working_Days']) * 100

    # --- CLASS MAPPING ---
    df['Class'] = df[course_column].map(class_mapping)

    # 🔥 NEW: Split Grade 02 using batch_id
    def split_grade_2(row):
        if row['Class'] == 'GRADE 02':
            try:
                section = str(row['batch_id']).split('-')[1]  # Extract A or B
                return f"GRADE 02 - {section}"
            except:
                return "GRADE 02 - UNKNOWN"
        return row['Class']

    df['Class'] = df.apply(split_grade_2, axis=1)

    # 🔥 Update class list dynamically
    updated_class_list = []
    for cls in class_list:
        if cls == 'GRADE 02':
            updated_class_list.extend(['GRADE 02 - A', 'GRADE 02 - B'])
        else:
            updated_class_list.append(cls)

    # Preserve order without using set
    class_list = []
    for cls in updated_class_list:
        if cls not in class_list:
            class_list.append(cls)

    # --- GROUPING ---
    for class_name in class_list:
        class_data = df[df['Class'] == class_name].copy()
        
        if class_data.empty:
            continue
            
        output_columns = ['Admission No', 'Student Name', 'Working_Days', 'Present', 
                         'Absent', 'Late', 'Very_Late', 'Attendance %', 'Class']
        
        output_columns = [col for col in output_columns if col in class_data.columns]
        class_data = class_data[output_columns]
        
        detailed_dfs[class_name] = class_data

    return detailed_dfs

# --- Generate the detailed data
process_button = st.button("Process Attendance Data")

if process_button:
    # Check if working days is provided and valid
    if working_days is None:
        st.error("Please enter the total number of working days.")
        st.stop()
    
    if working_days <= 0:
        st.error("Please enter a valid number of working days (minimum 1)")
        st.stop()
    
    # Store working days in session state
    st.session_state.working_days = working_days
    
    detailed_dfs = process_real_data(df, class_list, course_column, class_mapping, working_days)
    
    if not detailed_dfs:
        st.error("No data was processed. Please check your input and try again.")
        st.stop()
    
    # Create a summary sheet
    summary_data = []
    
    # Sort class names in natural order
    sorted_class_names = sort_class_names(detailed_dfs.keys())
    
    for class_name in sorted_class_names:
        df_detail = detailed_dfs[class_name]
        summary_data.append({
            "Class": class_name,
            "Total_Students": len(df_detail),
            "Total_Working_Days": working_days,
            "Avg_Present": round(df_detail["Present"].mean(), 2),
            "Avg_Absent": round(df_detail["Absent"].mean(), 2),
            "Avg_Late": round(df_detail["Late"].mean(), 2),
            "Avg_Very_Late": round(df_detail["Very_Late"].mean(), 2),
            "Avg_Attendance_Percentage": round(df_detail["Attendance %"].mean(), 2)
        })
    
    summary_df = pd.DataFrame(summary_data)
    
    # Store results in session state
    st.session_state.processed = True
    st.session_state.summary_df = summary_df
    st.session_state.detailed_dfs = detailed_dfs
    st.session_state.sorted_class_names = sorted_class_names

# Show results if data has been processed
if st.session_state.processed:
    # --- Display preview
    st.subheader("Preview of Processed Data")
    
    tab1, tab2 = st.tabs(["Summary", "Detailed View"])
    
    with tab1:
        st.write("Class Summary")
        st.dataframe(st.session_state.summary_df)
    
    with tab2:
        selected_class = st.selectbox("Select class to view details", options=st.session_state.sorted_class_names)
        st.dataframe(st.session_state.detailed_dfs[selected_class])
    
    # --- Download buttons
    excel_bytes = to_excel_bytes(
    st.session_state.summary_df,
    st.session_state.detailed_dfs,
    st.session_state.sorted_class_names,
    late_highlight_threshold,
    very_late_highlight_threshold)
    
    # Create two columns for the download buttons
    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            label="Download Detailed Attendance Report (Excel)",
            data=excel_bytes,
            file_name="detailed_attendance_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col2:
        # Generate PDF report
        pdf_bytes = generate_pdf_report(
            st.session_state.summary_df,
            st.session_state.detailed_dfs,
            st.session_state.sorted_class_names,
            uploaded_file.name
        )        
        # Show info message about PDF being in development
        st.info(
            """
            **PDF Export Notice**
            
            The PDF export feature is currently under development. 
            Some formatting issues may still be present.
            
            For the best experience, I recommend that Ustadh may use the Excel download option.
            """
        )
        
        # Enabled PDF download button
        st.download_button(
            label="Download Attendance Report (PDF)",
            data=pdf_bytes,
            file_name="attendance_report.pdf",
            mime="application/pdf",
            help="PDF export is still under development. Formatting may not be perfect."
        )
    
    st.success("Attendance data processed successfully! Download the files above.")
    
    # --- Add a button to reset the application and upload a new file
    if st.button("Add a new file", key="reset_button"):
        reset_application()
        st.rerun()

elif not process_button:
    st.info("Click the button above to process your attendance data based on your settings.")

# --- Instructions
st.markdown("---")
st.subheader("Instructions")
st.markdown("""

The app will create:
- A summary sheet with class statistics
- Separate sheets for each class with detailed student attendance records (ordered by class name)
""")




