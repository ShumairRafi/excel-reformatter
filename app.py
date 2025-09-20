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
Upload your summary file and the app will generate detailed student records.
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

# Get class names
class_names = st.text_area(
    "Enter class names (one per line)", 
    value="GRADE 01\nGRADE 02\nGRADE 03\nGRADE 04\nGRADE 05\nGRADE 07",
    help="Enter the class names that should appear in the output. One class per line."
)

class_list = [name.strip() for name in class_names.split('\n') if name.strip()]

# Get number of students per class
students_per_class = {}
st.write("Enter number of students for each class:")
for class_name in class_list:
    students_per_class[class_name] = st.number_input(
        f"Number of students in {class_name}", 
        min_value=1, 
        max_value=100, 
        value=22 if class_name == "GRADE 01" else 16 if class_name == "GRADE 02" else 8,
        key=f"students_{class_name}"
    )

# Get working days
working_days = st.number_input(
    "Total working days", 
    min_value=1, 
    max_value=365, 
    value=82
)

# --- Generate sample student data
def generate_student_data(class_name, num_students, working_days):
    students = []
    
    # Sample names for demonstration - in a real app, you might want to import a list of names
    sample_names = [
        "Huzaifa usama", "Mohamed Ramzeen Thahir", "Muhammad Rafi Muhammad Shumair",
        "Muhammad Shifas Muhammadh", "Abdullah Dilshard", "Arham Aasif", "Mohamed Bilaal",
        "Mohammed Aakif Mohammed Ameer", "Ahmed Fairooz", "Muhammed Khalid",
        "Mohamad Faraj Mohamad Firnas", "Atheequr Rahman", "Aabidh Zackey", "Yasir Faleel",
        "Mahamath Yusuf Cassim", "Mohammadh Rifai Abdur Rahman", "Muhammad Fiaz",
        "Muhammed Arham Imran Haladeen", "Abdullah Firdous", "Abdur Rahman Fazme",
        "Saadh Firdous", "Yusuf Iqbal"
    ]
    
    # Generate admission numbers starting from a base
    base_admission = 276 if class_name == "GRADE 01" else 279
    
    for i in range(num_students):
        admission_no = base_admission + i
        student_name = sample_names[i % len(sample_names)] if i < len(sample_names) else f"Student {i+1}"
        
        # Generate realistic attendance data
        present = np.random.randint(working_days * 0.4, working_days * 0.9)
        absent = working_days - present
        late = np.random.randint(0, present * 0.6)
        very_late = np.random.randint(0, late * 0.3)
        
        # Calculate attendance percentage
        attendance_pct = (present / working_days) * 100
        
        students.append({
            "Admission No": admission_no,
            "Student Name": student_name,
            "Working_Days": working_days,
            "Present": present,
            "Absent": absent,
            "Late": late,
            "Very_Late": very_late,
            "Attendance %": round(attendance_pct, 2),
            "Class": class_name
        })
    
    return pd.DataFrame(students)

# --- Generate the detailed data
if st.button("Generate Detailed Student Records"):
    detailed_dfs = {}
    
    for class_name in class_list:
        num_students = students_per_class[class_name]
        detailed_dfs[class_name] = generate_student_data(class_name, num_students, working_days)
    
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
    st.subheader("Preview of Generated Data")
    
    tab1, tab2 = st.tabs(["Summary", "Detailed View"])
    
    with tab1:
        st.write("Class Summary")
        st.dataframe(summary_df)
    
    with tab2:
        selected_class = st.selectbox("Select class to view details", options=class_list)
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
    
    st.success("Detailed attendance report generated successfully! Download the file above.")

else:
    st.info("Click the button above to generate detailed student records based on your settings.")

# --- Instructions
st.markdown("---")
st.subheader("Instructions")
st.markdown("""
1. Upload your attendance summary Excel file
2. Enter the class names (one per line)
3. Specify the number of students in each class
4. Set the total working days
5. Click "Generate Detailed Student Records"
6. Review the preview and download the generated file

The app will create:
- A summary sheet with class statistics
- Separate sheets for each class with detailed student attendance records
- Realistic sample data for demonstration purposes

Note: In a real application, you would replace the sample data with your actual student database.
""")
