# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Attendance Data Transformer", layout="wide")

st.title("Attendance Data Transformer")
st.markdown(
    """
This app transforms your attendance summary data into detailed student attendance records.
Upload your attendance summary file and the app will generate detailed student records.
"""
)

# --- Upload files
uploaded_file = st.file_uploader("Upload your attendance summary Excel file", type=["xls", "xlsx"])

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

# Get class names from the data if possible, otherwise allow manual input
if 'Class' in df.columns:
    existing_classes = df['Class'].unique().tolist()
    class_names = st.text_area(
        "Enter class names (one per line)", 
        value="\n".join(existing_classes),
        help="Edit the class names if needed. One class per line."
    )
else:
    class_names = st.text_area(
        "Enter class names (one per line)", 
        value="GRADE 01\nGRADE 02\nGRADE 03\nGRADE 04\nGRADE 05\nGRADE 07",
        help="Enter the class names that should appear in the output. One class per line."
    )

class_list = [name.strip() for name in class_names.split('\n') if name.strip()]

# Get admission number ranges for each class
st.subheader("Admission Number Ranges by Class")
admission_ranges = {}

for class_name in class_list:
    col1, col2 = st.columns(2)
    with col1:
        min_admission = st.number_input(
            f"Starting admission number for {class_name}", 
            min_value=1, 
            max_value=10000, 
            value=276 if class_name == "GRADE 01" else 279,
            key=f"min_{class_name}"
        )
    with col2:
        max_admission = st.number_input(
            f"Ending admission number for {class_name}", 
            min_value=min_admission, 
            max_value=10000, 
            value=297 if class_name == "GRADE 01" else 294,
            key=f"max_{class_name}"
        )
    admission_ranges[class_name] = (min_admission, max_admission)

# Get working days
working_days = st.number_input(
    "Total working days", 
    min_value=1, 
    max_value=365, 
    value=82
)

# --- Process the real data
def process_real_data(df, class_list, admission_ranges, working_days):
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
    
    # Rename columns for consistency
    df = df.rename(columns=column_mapping)
    
    # Add missing columns with default values
    if 'Late' not in df.columns:
        df['Late'] = 0
    if 'Very_Late' not in df.columns:
        df['Very_Late'] = 0
    
    # Calculate attendance percentage
    df['Working_Days'] = working_days
    df['Attendance %'] = (df['Present'] / working_days) * 100
    
    # Filter students by admission number for each class
    for class_name in class_list:
        min_admission, max_admission = admission_ranges[class_name]
        
        # Filter students in this admission range
        class_data = df[
            (df['Admission No'] >= min_admission) & 
            (df['Admission No'] <= max_admission)
        ].copy()
        
        if class_data.empty:
            st.warning(f"No students found in admission range {min_admission}-{max_admission} for class: {class_name}")
            continue
            
        # Add class information
        class_data['Class'] = class_name
        
        # Select and order columns for output
        output_columns = ['Admission No', 'Student Name', 'Working_Days', 'Present', 
                         'Absent', 'Late', 'Very_Late', 'Attendance %', 'Class']
        
        # Keep only columns that exist in the dataframe
        output_columns = [col for col in output_columns if col in class_data.columns]
        class_data = class_data[output_columns]
        
        detailed_dfs[class_name] = class_data
    
    return detailed_dfs

# --- Generate the detailed data
if st.button("Process Attendance Data"):
    detailed_dfs = process_real_data(df, class_list, admission_ranges, working_days)
    
    if not detailed_dfs:
        st.error("No data was processed. Please check your input and try again.")
        st.stop()
    
    # Create a summary sheet
    summary_data = []
    for class_name, df_detail in detailed_dfs.items():
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
    
    # --- Display preview
    st.subheader("Preview of Processed Data")
    
    tab1, tab2 = st.tabs(["Summary", "Detailed View"])
    
    with tab1:
        st.write("Class Summary")
        st.dataframe(summary_df)
    
    with tab2:
        selected_class = st.selectbox("Select class to view details", options=list(detailed_dfs.keys()))
        st.dataframe(detailed_dfs[selected_class])
    
    # --- Download button
    def to_excel_bytes(summary_df, detailed_dfs):
        towrite = BytesIO()
        with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Class Summary")
            
            for class_name, df_detail in detailed_dfs.items():
                # Shorten sheet name if too long for Excel
                sheet_name = class_name[:31] if len(class_name) > 31 else class_name
                df_detail.to_excel(writer, index=False, sheet_name=sheet_name)
        
        towrite.seek(0)
        return towrite
    
    excel_bytes = to_excel_bytes(summary_df, detailed_dfs)
    
    st.download_button(
        label="Download Detailed Attendance Report",
        data=excel_bytes,
        file_name="detailed_attendance_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.success("Attendance data processed successfully! Download the file above.")

else:
    st.info("Click the button above to process your attendance data based on your settings.")

# --- Instructions
st.markdown("---")
st.subheader("Instructions")
st.markdown("""
1. Upload your attendance summary Excel file
2. The app will try to detect class names from your data, or you can enter them manually
3. For each class, specify the range of admission numbers
4. Set the total working days
5. Click "Process Attendance Data"
6. Review the preview and download the generated file

The app will create:
- A summary sheet with class statistics
- Separate sheets for each class with detailed student attendance records

**Note:** Your data should include at least these columns (or similar):
- Admission No
- Student Name  
- Present
- Absent

If your columns have different names, the app will try to match them automatically.
""")
