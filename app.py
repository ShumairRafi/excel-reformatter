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

st.set_page_config(page_title="Attendance Data Transformer", layout="wide")

st.title("Attendance Data Transformer")
st.markdown(
    """
This app transforms your attendance summary data into detailed student attendance records.
Upload your attendance summary Excel file and the app will generate detailed student records.
"""
)

# Show Developer Message
st.info(
"""
**Message from the Developer**
            
✨ Assalamu Alaikum Warahmatullahi Wabarakatuh, Ustadh! ✨

It is with great pleasure that I present this Attendance Data Transformer application, hoping to simplify Ustadh's administrative tasks and bring efficiency in managing student attendance records.

Should Ustadh encounter any challenges or have suggestions for refinement, please know that Ustadh's guidance is deeply valued and will be gratefully received to further improve this tool.

I hope this application will help bring ease to Ustadh's important work. Jazakallah Khairan!
""")

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
if 'highlights' not in st.session_state:
    st.session_state.highlights = {}

# Function to reset the application
def reset_application():
    st.session_state.processed = False
    st.session_state.summary_df = None
    st.session_state.detailed_dfs = {}
    st.session_state.sorted_class_names = []
    st.session_state.working_days = None
    st.session_state.highlights = {}
    st.session_state.file_uploader_key += 1  # Change the key to reset the file uploader

# Function to apply Excel styling
def apply_excel_styling(worksheet, title, is_summary=False, student_names=None):
    # Define styles
    header_font = Font(name='Aptos Display', size=12, bold=True)
    data_font = Font(name='Aptos Display', size=12)
    header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    alignment_center = Alignment(horizontal='center', vertical='center')
    alignment_left = Alignment(horizontal='left', vertical='center')
    
    # Thin border
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Apply styles to header row
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = alignment_center
        cell.border = thin_border
    
    # Apply styles to data rows
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            
            # For summary sheet, alignment
            if is_summary:
                    cell.alignment = alignment_center
            
            else:
                # For class sheets, use the original alignment
                if cell.column == 1:  # Admission No
                    cell.alignment = alignment_center
                elif cell.column == 2:  # Student Name
                    cell.alignment = alignment_left
                else:
                    cell.alignment = alignment_center
    
    # Adjust column widths
    if is_summary:
        column_widths = {
            'A': 18,  # Class
            'B': 20,  # Total_Students
            'C': 23,  # Total_Working_Days
            'D': 15,  # Avg_Present
            'E': 15,  # Avg_Absent
            'F': 15,  # Avg_Late
            'G': 15,  # Avg_Very_Late
            'H': 30   # Avg_Attendance_Percentage
        }
    else:
            
        column_widths = {
            'A': 15,  # Admission No
            'B': 40,  # Student Name
            'C': 15,  # Working Days
            'D': 10,  # Present
            'E': 10,  # Absent
            'F': 10,  # Late
            'G': 12,  # Very Late
            'H': 15,  # Attendance %
            'I': 12   # Class
        }
    
    for col, width in column_widths.items():
        worksheet.column_dimensions[col].width = width
    
    # Format percentage column
    for row in range(2, worksheet.max_row + 1):
        if is_summary:
            cell = worksheet[f'H{row}']
        else:
            cell = worksheet[f'H{row}']
        cell.number_format = '0.00'
    
    # Add title
    worksheet.insert_rows(1)
    if is_summary:
        worksheet.merge_cells('A1:H1')
        title_cell = worksheet['A1']
    else:
        worksheet.merge_cells('A1:I1')
        title_cell = worksheet['A1']
    title_cell.value = title
    title_cell.font = Font(name='Aptos Display', size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    return worksheet

# Function to generate highlights and insights
def generate_highlights(detailed_dfs, summary_df):
    highlights = {}
    
    # Combine all student data
    all_students = pd.concat(detailed_dfs.values(), ignore_index=True)
    
    # 1. Students with zero absence
    zero_absence_students = all_students[all_students['Absent'] == 0]
    highlights['zero_absence_count'] = len(zero_absence_students)
    highlights['zero_absence_students'] = zero_absence_students[['Student Name', 'Class']].head(10)  # Top 10
    
    # 2. Students with highest number of absences
    high_absence_students = all_students.sort_values('Absent', ascending=False)
    highlights['high_absence_count'] = min(10, len(high_absence_students))
    highlights['high_absence_students'] = high_absence_students[['Student Name', 'Class', 'Absent']].head(10)
    
    # 3. Class with best attendance (highest average attendance percentage)
    best_class = summary_df.loc[summary_df['Avg_Attendance_Percentage'].idxmax()]
    highlights['best_class'] = {
        'class_name': best_class['Class'],
        'attendance_percentage': best_class['Avg_Attendance_Percentage']
    }
    
    # 4. Class with lowest attendance
    worst_class = summary_df.loc[summary_df['Avg_Attendance_Percentage'].idxmin()]
    highlights['worst_class'] = {
        'class_name': worst_class['Class'],
        'attendance_percentage': worst_class['Avg_Attendance_Percentage']
    }
    
    # 5. Class with highest attendance rate but with absentees
    # Filter classes that have some absentees but good attendance
    classes_with_absentees = summary_df[summary_df['Avg_Absent'] > 0]
    if not classes_with_absentees.empty:
        best_class_with_absentees = classes_with_absentees.loc[classes_with_absentees['Avg_Attendance_Percentage'].idxmax()]
        highlights['best_class_with_absentees'] = {
            'class_name': best_class_with_absentees['Class'],
            'attendance_percentage': best_class_with_absentees['Avg_Attendance_Percentage'],
            'avg_absent': best_class_with_absentees['Avg_Absent']
        }
        
        # Get actual absent students from this class
        class_data = detailed_dfs[best_class_with_absentees['Class']]
        absent_students = class_data[class_data['Absent'] > 0][['Student Name', 'Absent']].sort_values('Absent', ascending=False)
        highlights['absent_students_in_best_class'] = absent_students
    else:
        highlights['best_class_with_absentees'] = None
        highlights['absent_students_in_best_class'] = pd.DataFrame()
    
    return highlights

# Function to generate PDF report
def generate_pdf_report(summary_df, detailed_dfs, sorted_class_names, highlights):
    """Generate a PDF version of the attendance report"""
    
    # Create PDF object
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Add a page for the summary
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "ATTENDANCE SUMMARY", 0, 1, 'C')
    pdf.ln(10)
    
    # Add highlights section
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "KEY HIGHLIGHTS", 0, 1, 'L')
    pdf.ln(5)
    
    pdf.set_font("Arial", '', 12)
    # Zero absence students
    pdf.cell(0, 10, f"Students with Perfect Attendance (Zero Absence): {highlights['zero_absence_count']}", 0, 1, 'L')
    
    # Highest absence students
    pdf.cell(0, 10, f"Students with Highest Absences (Top {highlights['high_absence_count']}):", 0, 1, 'L')
    
    # Best and worst classes
    pdf.cell(0, 10, f"Class with Best Attendance: {highlights['best_class']['class_name']} ({highlights['best_class']['attendance_percentage']:.2f}%)", 0, 1, 'L')
    pdf.cell(0, 10, f"Class with Lowest Attendance: {highlights['worst_class']['class_name']} ({highlights['worst_class']['attendance_percentage']:.2f}%)", 0, 1, 'L')
    
    if highlights['best_class_with_absentees']:
        pdf.cell(0, 10, f"Class with Highest Attendance Rate (with absentees): {highlights['best_class_with_absentees']['class_name']} ({highlights['best_class_with_absentees']['attendance_percentage']:.2f}%)", 0, 1, 'L')
    
    pdf.ln(10)
    
    # Add summary table
    pdf.set_font("Arial", 'B', 12)
    columns = list(summary_df.columns)
    
    # Set column widths
    col_widths = [40, 30, 35, 25, 25, 25, 25, 40]
    
    # Add header
    for i, column in enumerate(columns):
        pdf.cell(col_widths[i], 10, column, 1, 0, 'C')
    pdf.ln()
    
    # Add data rows
    pdf.set_font("Arial", '', 10)
    for _, row in summary_df.iterrows():
        for i, column in enumerate(columns):
            value = str(row[column])
            pdf.cell(col_widths[i], 10, value, 1, 0, 'C')
        pdf.ln()
    
    # Add detailed class reports
    for class_name in sorted_class_names:
        if class_name in detailed_dfs:
            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(0, 10, f"CLASS: {class_name}", 0, 1, 'C')
            pdf.ln(10)
            
            df_detail = detailed_dfs[class_name]
            columns = list(df_detail.columns)
            
            # Set column widths for detail table
            detail_col_widths = [30, 60, 25, 20, 20, 20, 25, 25, 20]
            
            # Add header
            pdf.set_font("Arial", 'B', 10)
            for i, column in enumerate(columns):
                pdf.cell(detail_col_widths[i], 10, column, 1, 0, 'C')
            pdf.ln()
            
            # Add data rows
            pdf.set_font("Arial", '', 8)
            for _, row in df_detail.iterrows():
                for i, column in enumerate(columns):
                    value = str(row[column])
                    pdf.cell(detail_col_widths[i], 10, value, 1, 0, 'C')
                pdf.ln()
    
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

# Get working days with default value as None
working_days = st.number_input(
    "Total working days*", 
    min_value=1, 
    max_value=365, 
    value=st.session_state.working_days,
    help="Enter the total number of working days for the period. This field is required.",
    placeholder="Enter a number between 1 and 365"
)

# Function to sort class names in natural order (GRADE 01, GRADE 02, etc.)
def sort_class_names(class_names):
    def extract_number(name):
        # Extract numbers from class name
        numbers = re.findall(r'\d+', name)
        return int(numbers[0]) if numbers else float('inf')
    
    return sorted(class_names, key=extract_number)

# Function to create Excel file
def to_excel_bytes(summary_df, detailed_dfs, sorted_class_names, highlights):
    # Create a workbook
    wb = Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Add highlights sheet
    ws_highlights = wb.create_sheet("Highlights")
    
    # Write highlights data
    ws_highlights.append(["ATTENDANCE HIGHLIGHTS"])
    ws_highlights.append([])
    
    # Zero absence students
    ws_highlights.append(["Students with Perfect Attendance (Zero Absence)", f"Total: {highlights['zero_absence_count']}"])
    if highlights['zero_absence_count'] > 0:
        ws_highlights.append(["Student Name", "Class"])
        for _, student in highlights['zero_absence_students'].iterrows():
            ws_highlights.append([student['Student Name'], student['Class']])
    else:
        ws_highlights.append(["No students with perfect attendance"])
    ws_highlights.append([])
    
    # Highest absence students
    ws_highlights.append([f"Students with Highest Absences (Top {highlights['high_absence_count']})"])
    ws_highlights.append(["Student Name", "Class", "Absent Days"])
    for _, student in highlights['high_absence_students'].iterrows():
        ws_highlights.append([student['Student Name'], student['Class'], student['Absent']])
    ws_highlights.append([])
    
    # Class performance
    ws_highlights.append(["Class Performance Summary"])
    ws_highlights.append(["Best Attendance Class", f"{highlights['best_class']['class_name']} ({highlights['best_class']['attendance_percentage']:.2f}%)"])
    ws_highlights.append(["Lowest Attendance Class", f"{highlights['worst_class']['class_name']} ({highlights['worst_class']['attendance_percentage']:.2f}%)"])
    
    if highlights['best_class_with_absentees']:
        ws_highlights.append(["Best Class with Absentees", f"{highlights['best_class_with_absentees']['class_name']} ({highlights['best_class_with_absentees']['attendance_percentage']:.2f}%)"])
        if not highlights['absent_students_in_best_class'].empty:
            ws_highlights.append([])
            ws_highlights.append(["Absent Students in Best Class"])
            ws_highlights.append(["Student Name", "Absent Days"])
            for _, student in highlights['absent_students_in_best_class'].iterrows():
                ws_highlights.append([student['Student Name'], student['Absent']])
    
    # Apply styling to highlights sheet
    ws_highlights = apply_excel_styling(ws_highlights, "ATTENDANCE HIGHLIGHTS", is_summary=True)
    
    # Add summary sheet
    ws_summary = wb.create_sheet("Class Summary")
    
    # Write summary data
    for r in dataframe_to_rows(summary_df, index=False, header=True):
        ws_summary.append(r)
    
    # Apply styling to summary sheet
    ws_summary = apply_excel_styling(ws_summary, "ATTENDANCE SUMMARY", is_summary=True)
    
    # Add class sheets
    for class_name in sorted_class_names:
        if class_name in detailed_dfs:
            # Shorten sheet name if too long for Excel
            sheet_name = class_name[:31] if len(class_name) > 31 else class_name
            ws_class = wb.create_sheet(sheet_name)
            
            # Write class data
            for r in dataframe_to_rows(detailed_dfs[class_name], index=False, header=True):
                ws_class.append(r)
            
            # Get student names for this class to calculate optimal column width
            student_names = detailed_dfs[class_name]['Student Name'].tolist() if 'Student Name' in detailed_dfs[class_name].columns else []
            
            # Apply styling to class sheet
            ws_class = apply_excel_styling(ws_class, class_name, is_summary=False, student_names=student_names)
    
    # Save to bytes
    towrite = BytesIO()
    wb.save(towrite)
    towrite.seek(0)
    return towrite

# --- Process the real data
def process_real_data(df, class_list, course_column, class_mapping, working_days):
    detailed_dfs = {}
    
    # Ensure we have the required columns
    required_columns = ['Admission No', 'Student Name', 'Present', 'Absent']
    available_columns = df.columns.tolist()
    
    # Try to map existing columns to required columns using fuzzy matching
    column_mapping = {}
    for req_col in required_columns:
        match = process.extractOne(req_col, available_columns, scorer=fuzz.token_sort_ratio)
        if match and match[1] > 60:  # If similarity score > 60%
            column_mapping[req_col] = match[0]
        else:
            column_mapping[req_col] = req_col
            st.warning(f"Could not find a matching column for '{req_col}'. Please ensure your data has this column.")
    
    # Also try to map optional columns like Late and Very Late
    optional_columns = ['Late', 'Very_Late', 'Very Late']
    for opt_col in optional_columns:
        match = process.extractOne(opt_col, available_columns, scorer=fuzz.token_sort_ratio)
        if match and match[1] > 60:
            column_mapping[opt_col] = match[0]
            st.info(f"Mapped '{match[0]}' to '{opt_col}'")
    
    # Rename columns for consistency
    df = df.rename(columns=column_mapping)
    
    # Add missing columns with default values
    if 'Late' not in df.columns:
        df['Late'] = 0
    if 'Very_Late' not in df.columns:
        # Also check for 'Very Late' with space
        if 'Very Late' in df.columns:
            df['Very_Late'] = df['Very Late']
        else:
            df['Very_Late'] = 0
    
    # Calculate attendance percentage
    df['Working_Days'] = working_days
    df['Attendance %'] = (df['Present'] / working_days) * 100
    
    # If we have a course column, use it to map to classes
    if course_column and class_mapping:
        # Apply class mapping
        df['Class'] = df[course_column].map(class_mapping)
        
        # Filter for classes in our list
        df = df[df['Class'].isin(class_list)]
        
        # Group by class
        for class_name in class_list:
            class_data = df[df['Class'] == class_name].copy()
            
            if class_data.empty:
                st.warning(f"No students found for class: {class_name}")
                continue
                
            # Select and order columns for output
            output_columns = ['Admission No', 'Student Name', 'Working_Days', 'Present', 
                             'Absent', 'Late', 'Very_Late', 'Attendance %', 'Class']
            
            # Keep only columns that exist in the dataframe
            output_columns = [col for col in output_columns if col in class_data.columns]
            class_data = class_data[output_columns]
            
            detailed_dfs[class_name] = class_data
    else:
        # Fallback: If no course column, use the class list as provided
        st.warning("No course column detected. Using manual class assignment.")
        
        # Distribute students evenly among classes
        students_per_class = len(df) // len(class_list)
        remainder = len(df) % len(class_list)
        
        start_idx = 0
        for i, class_name in enumerate(class_list):
            end_idx = start_idx + students_per_class + (1 if i < remainder else 0)
            class_data = df.iloc[start_idx:end_idx].copy()
            class_data['Class'] = class_name
            
            # Select and order columns for output
            output_columns = ['Admission No', 'Student Name', 'Working_Days', 'Present', 
                             'Absent', 'Late', 'Very_Late', 'Attendance %', 'Class']
            
            # Keep only columns that exist in the dataframe
            output_columns = [col for col in output_columns if col in class_data.columns]
            class_data = class_data[output_columns]
            
            detailed_dfs[class_name] = class_data
            start_idx = end_idx
    
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
    
    # Generate highlights
    highlights = generate_highlights(detailed_dfs, summary_df)
    
    # Store results in session state
    st.session_state.processed = True
    st.session_state.summary_df = summary_df
    st.session_state.detailed_dfs = detailed_dfs
    st.session_state.sorted_class_names = sorted_class_names
    st.session_state.highlights = highlights

# Show results if data has been processed
if st.session_state.processed:
    # --- Display preview
    st.subheader("Preview of Processed Data")
    
    tab1, tab2, tab3 = st.tabs(["Summary", "Detailed View", "Highlights"])
    
    with tab1:
        st.write("Class Summary")
        st.dataframe(st.session_state.summary_df)
    
    with tab2:
        selected_class = st.selectbox("Select class to view details", options=st.session_state.sorted_class_names)
        st.dataframe(st.session_state.detailed_dfs[selected_class])
    
    with tab3:
        st.subheader("Attendance Highlights")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.metric(
                "Students with Perfect Attendance", 
                st.session_state.highlights['zero_absence_count']
            )
            
            st.metric(
                "Class with Best Attendance", 
                f"{st.session_state.highlights['best_class']['class_name']} ({st.session_state.highlights['best_class']['attendance_percentage']:.2f}%)"
            )
        
        with col2:
            st.metric(
                "Students with Highest Absences", 
                st.session_state.highlights['high_absence_count']
            )
            
            st.metric(
                "Class with Lowest Attendance", 
                f"{st.session_state.highlights['worst_class']['class_name']} ({st.session_state.highlights['worst_class']['attendance_percentage']:.2f}%)"
            )
        
        # Display detailed highlights
        st.subheader("Detailed Highlights")
        
        # Students with zero absence
        st.write("**Students with Perfect Attendance (Zero Absence):**")
        if st.session_state.highlights['zero_absence_count'] > 0:
            st.dataframe(st.session_state.highlights['zero_absence_students'])
        else:
            st.write("No students with perfect attendance")
        
        # Students with highest absences
        st.write(f"**Students with Highest Absences (Top {st.session_state.highlights['high_absence_count']}):**")
        st.dataframe(st.session_state.highlights['high_absence_students'])
        
        # Class with highest attendance rate but with absentees
        if st.session_state.highlights['best_class_with_absentees']:
            st.write("**Class with Highest Attendance Rate (with absentees):**")
            st.write(f"{st.session_state.highlights['best_class_with_absentees']['class_name']} ({st.session_state.highlights['best_class_with_absentees']['attendance_percentage']:.2f}%)")
            
            if not st.session_state.highlights['absent_students_in_best_class'].empty:
                st.write("**Absent Students in this Class:**")
                st.dataframe(st.session_state.highlights['absent_students_in_best_class'])
    
    # --- Download buttons
    excel_bytes = to_excel_bytes(st.session_state.summary_df, st.session_state.detailed_dfs, st.session_state.sorted_class_names, st.session_state.highlights)
    
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
        pdf_bytes = generate_pdf_report(st.session_state.summary_df, st.session_state.detailed_dfs, st.session_state.sorted_class_names, st.session_state.highlights)
        
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
1. Upload your attendance summary Excel file
2. The app will detect the 'course_name' column and show you the unique values
3. Map each course name to a standardized class name (e.g., "7th Year" → "GRADE 07")
4. **Set the total working days** (this field is required and must be greater than 0)
5. Click "Process Attendance Data"
6. Review the preview and download the generated file
7. Use the "Add a new file" button to reset the application and upload a different file

The app will create:
- A summary sheet with class statistics
- Separate sheets for each class with detailed student attendance records (ordered by class name)
- A highlights sheet with key insights including:
  - Students with perfect attendance (zero absence)
  - Students with highest number of absences
  - Class with best attendance rate
  - Class with lowest attendance rate
  - Class with highest attendance rate that still has absentees

**Note:** Your data should include at least these columns (or similar):
- Admission No
- Student Name  
- Present
- Absent
- course_name (or similar column indicating student's course/class)

If your columns have different names, the app will try to match them automatically.
""")
