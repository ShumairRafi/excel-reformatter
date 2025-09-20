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

# Get working days
working_days = st.number_input(
    "Total working days", 
    min_value=1, 
    max_value=365, 
    value=82
)

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
if st.button("Process Attendance Data"):
    detailed_dfs = process_real_data(df, class_list, course_column, class_mapping, working_days)
    
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
2. The app will detect the 'course_name' column and show you the unique values
3. Map each course name to a standardized class name (e.g., "7th Year" → "GRADE 07")
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
- course_name (or similar column indicating student's course/class)

If your columns have different names, the app will try to match them automatically.
""")
