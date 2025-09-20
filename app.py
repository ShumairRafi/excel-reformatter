# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Excel Reformatter", layout="wide")

st.title("Excel Reformatter — Multi-Sheet Attendance Formatting")
st.markdown(
    """
Upload your **old Excel file** (with multiple sheets for each grade) and your **new Excel file** (single sheet with all data). 
The app will separate the data by grade and reformat it to match your old file's structure.
"""
)

# --- Upload files
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("Upload OLD Excel file (multi-sheet format)", type=["xls", "xlsx"], key="old")
with col2:
    new_file = st.file_uploader("Upload NEW Excel file (single sheet data)", type=["xls", "xlsx"], key="new")

if not (old_file and new_file):
    st.info("Upload both the old file (format) and the new file (data) to continue.")
    st.stop()

# --- Read files
@st.cache_data(ttl=600)
def read_excel_sheets(file):
    try:
        xl = pd.ExcelFile(file)
        return {sheet_name: pd.read_excel(file, sheet_name=sheet_name) for sheet_name in xl.sheet_names}
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

@st.cache_data(ttl=600)
def read_excel_single(file, sheet_name_hint=""):
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
    # Read all sheets from old file
    old_sheets = read_excel_sheets(old_file)
    
    # Read new file (single sheet)
    df_new = read_excel_single(new_file)
    
    if old_sheets is None or df_new is None:
        st.error("Could not read one of the uploaded Excel files.")
        st.stop()
        
except Exception as e:
    st.error(f"Could not read one of the uploaded Excel files: {e}")
    st.stop()

# Show available sheets in old file
st.subheader("Sheets in old file")
st.write(list(old_sheets.keys()))

# Let user select which sheet to use as reference format
ref_sheet = st.selectbox("Select reference sheet to use as format template", options=list(old_sheets.keys()))
df_old = old_sheets[ref_sheet]

st.subheader("Preview — columns")

c1, c2 = st.columns(2)
with c1:
    st.write(f"Old file - {ref_sheet} sheet (first 5 rows)")
    st.dataframe(df_old.head())
    st.write(f"Shape: {df_old.shape}")
with c2:
    st.write("New file (first 5 rows)")
    st.dataframe(df_new.head())
    st.write(f"Shape: {df_new.shape}")

# Prepare column lists
old_cols = list(df_old.columns.astype(str))
new_cols = list(df_new.columns.astype(str))

# --- Automatic mapping using fuzzy matching
st.subheader("Column mapping")
st.write("Below are suggestions. Adjust any mapping using the dropdowns. If a column has no source, choose `-- (blank)` to fill blanks.")

# Build suggestions
suggestions = {}
for oc in old_cols:
    # use rapidfuzz to find best match among new_cols
    match = process.extractOne(oc, new_cols, scorer=fuzz.token_sort_ratio)
    if match:
        matched_col, score, _ = match
    else:
        matched_col, score = None, 0
        
    if score >= 70:  # Fixed threshold for simplicity
        suggestions[oc] = (matched_col, score)
    else:
        suggestions[oc] = (None, score)

# UI: let user change
mapping = {}
cols_for_select = ["-- (blank)"] + new_cols

for oc in old_cols:
    suggested_col, score = suggestions[oc]
    label = f"Map target column `{oc}`"
    if suggested_col:
        label += f" (suggested: `{suggested_col}` — score {int(score)})"
    else:
        label += f" (no good match — best score {int(score)})"
    
    default_index = 0
    if suggested_col and suggested_col in cols_for_select:
        default_index = cols_for_select.index(suggested_col)
        
    choice = st.selectbox(
        label, 
        options=cols_for_select, 
        index=default_index, 
        key=f"map_{hash(oc)}"
    )
    mapping[oc] = None if choice == "-- (blank)" else choice

# Check if the user mapped multiple target cols to the same new column
reverse_map = {}
for tcol, src in mapping.items():
    if src is None:
        continue
    reverse_map.setdefault(src, []).append(tcol)

conflicts = {src: tcols for src, tcols in reverse_map.items() if len(tcols) > 1}
if conflicts:
    st.warning("Some source columns are mapped to multiple target columns. That's allowed, but please confirm it's intentional.")
    for src, tcols in conflicts.items():
        st.write(f"Source column `{src}` → target columns: {', '.join('`'+t+'`' for t in tcols)}")

# --- Build the reformatted DataFrame
def build_reformatted(df_new, mapping, old_cols):
    out_df = pd.DataFrame()
    
    for oc in old_cols:
        src = mapping.get(oc)
        
        if src is None:
            # fill with blank / NaN
            out_df[oc] = pd.NA
        else:
            # if src exists in df_new, copy; else blank
            if src in df_new.columns:
                out_df[oc] = df_new[src]
            else:
                out_df[oc] = pd.NA
    
    return out_df

reformatted_df = build_reformatted(df_new, mapping, old_cols)

st.subheader("Reformatted preview (first 10 rows)")
st.dataframe(reformatted_df.head(10))
st.write(f"Reformatted shape: {reformatted_df.shape}")

# --- Multi-sheet processing for attendance data
st.markdown("### Multi-sheet Processing")
st.write("Your old file has multiple sheets for different grades. Would you like to split the new data by grade?")

# Try to detect grade information in the new data
grade_columns = [col for col in df_new.columns if any(term in str(col).lower() for term in ['grade', 'class', 'batch', 'course'])]

if grade_columns:
    grade_column = st.selectbox("Select column containing grade/class information", options=grade_columns)
    
    # Get unique grades/classes
    unique_grades = df_new[grade_column].dropna().unique()
    st.write(f"Found {len(unique_grades)} unique grades/classes: {unique_grades}")
    
    process_multisheet = st.checkbox("Process as multi-sheet Excel file", value=True)
else:
    st.info("No obvious grade/class column found in the new data.")
    process_multisheet = False

# --- Download button
def to_excel_bytes(df):
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reformatted")
    towrite.seek(0)
    return towrite

def to_multisheet_excel_bytes(df, grade_column):
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        # First sheet with all data
        df.to_excel(writer, index=False, sheet_name="All Data")
        
        # Sheets for each grade
        for grade in df[grade_column].dropna().unique():
            grade_df = df[df[grade_column] == grade]
            sheet_name = f"Grade {grade}" if len(str(grade)) < 20 else f"Grade_{hash(grade)}"
            grade_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])  # Excel sheet name limit
    
    towrite.seek(0)
    return towrite

if process_multisheet and grade_columns:
    excel_bytes = to_multisheet_excel_bytes(reformatted_df, grade_column)
    file_name = "reformatted_multisheet.xlsx"
else:
    excel_bytes = to_excel_bytes(reformatted_df)
    file_name = "reformatted.xlsx"

st.download_button(
    label="Download reformatted Excel",
    data=excel_bytes,
    file_name=file_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("Done — download the file above. The new data has been reformatted to match your old file's structure.")
