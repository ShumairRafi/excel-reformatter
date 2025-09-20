# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Excel Reformatter", layout="wide")

st.title("Excel Reformatter — Map New Data to Old File Format")
st.markdown(
    """
Upload your **old Excel file** (the file whose format you want to match) and your **new Excel file** (data to be reformatted). 
The app will suggest column mappings and let you adjust them before exporting a reformatted Excel file.
"""
)

# --- Sidebar: options
with st.sidebar:
    st.header("Options")
    fuzz_threshold = st.slider(
        "Auto-mapping similarity threshold", 0, 100, 70,
        help="Minimum similarity (0-100) for a suggested automatic match. Lower = more aggressive matching."
    )
    sheet_old = st.text_input("Old file sheet name (leave blank = first sheet)", value="")
    sheet_new = st.text_input("New file sheet name (leave blank = first sheet)", value="")
    preserve_old_index = st.checkbox("Preserve old file's index structure", value=False)
    case_sensitive = st.checkbox("Case-sensitive matching", value=False)

# --- Upload files
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("Upload OLD Excel file (target format)", type=["xls", "xlsx"], key="old")
with col2:
    new_file = st.file_uploader("Upload NEW Excel file (data to convert)", type=["xls", "xlsx"], key="new")

if not (old_file and new_file):
    st.info("Upload both the old file (format) and the new file (data) to continue.")
    st.stop()

# --- Read files
@st.cache_data(ttl=600)
def read_excel(file, sheet_name_hint):
    try:
        if sheet_name_hint:
            return pd.read_excel(file, sheet_name=sheet_name_hint, engine="openpyxl")
        else:
            return pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        # If sheet_name hint fails, try default
        try:
            return pd.read_excel(file, engine="openpyxl")
        except Exception as e2:
            st.error(f"Error reading file: {e2}")
            return None

try:
    df_old = read_excel(old_file, sheet_old)
    df_new = read_excel(new_file, sheet_new)
    
    if df_old is None or df_new is None:
        st.error("Could not read one of the uploaded Excel files. Please check the sheet names.")
        st.stop()
        
except Exception as e:
    st.error(f"Could not read one of the uploaded Excel files: {e}")
    st.stop()

# Store original dtypes from old file to preserve them
old_dtypes = df_old.dtypes.to_dict()

st.subheader("Preview — columns")

c1, c2 = st.columns(2)
with c1:
    st.write("Old file (target format) — first 5 rows")
    st.dataframe(df_old.head())
    st.write(f"Old file shape: {df_old.shape}")
with c2:
    st.write("New file (incoming data) — first 5 rows")
    st.dataframe(df_new.head())
    st.write(f"New file shape: {df_new.shape}")

# Prepare column lists
old_cols = list(df_old.columns.astype(str))
new_cols = list(df_new.columns.astype(str))

if not case_sensitive:
    old_cols_lower = [col.lower() for col in old_cols]
    new_cols_lower = [col.lower() for col in new_cols]

# --- Automatic mapping using fuzzy matching
st.subheader("Column mapping")
st.write("Below are suggestions. Adjust any mapping using the dropdowns. If a column has no source, choose `-- (blank)` to fill blanks.")

# Build suggestions
suggestions = {}
for i, oc in enumerate(old_cols):
    if case_sensitive:
        search_space = new_cols
        target_col = oc
    else:
        search_space = new_cols_lower
        target_col = oc.lower()
    
    # use rapidfuzz to find best match among new_cols
    match = process.extractOne(target_col, search_space, scorer=fuzz.token_sort_ratio)
    if match:
        matched_col, score, _ = match
        if not case_sensitive:
            # Get the original case version of the matched column
            matched_col = new_cols[new_cols_lower.index(matched_col)]
    else:
        matched_col, score = None, 0
        
    if score >= fuzz_threshold:
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
        key=f"map_{hash(oc)}"  # More unique key
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
def build_reformatted(df_new, mapping, old_cols, old_dtypes):
    out_df = pd.DataFrame()
    
    for oc in old_cols:
        src = mapping.get(oc)
        
        if src is None:
            # fill with blank / NaN but preserve original dtype if possible
            out_df[oc] = pd.Series(dtype=old_dtypes.get(oc, object))
        else:
            # if src exists in df_new, copy; else blank
            if src in df_new.columns:
                out_df[oc] = df_new[src]
                
                # Try to convert to original dtype if possible
                try:
                    if old_dtypes[oc] != out_df[oc].dtype:
                        out_df[oc] = out_df[oc].astype(old_dtypes[oc])
                except (ValueError, TypeError):
                    # If conversion fails, keep as is
                    pass
            else:
                out_df[oc] = pd.Series(dtype=old_dtypes.get(oc, object))
    
    return out_df

reformatted_df = build_reformatted(df_new, mapping, old_cols, old_dtypes)

st.subheader("Reformatted preview (first 10 rows)")
st.dataframe(reformatted_df.head(10))
st.write(f"Reformatted shape: {reformatted_df.shape}")

# --- Post-processing options
st.markdown("### Post-processing (optional)")
st.write("You can do basic cleaning here. For more advanced transforms, download and modify manually.")

col1, col2 = st.columns(2)
with col1:
    do_dropna = st.checkbox("Drop rows where all values are blank", value=False)
    strip_whitespace = st.checkbox("Strip whitespace from text columns", value=True)
with col2:
    reset_index = st.checkbox("Reset index to match old file", value=True)
    fill_na = st.selectbox("Fill NaN values with", ["", "0", "Empty string", "Previous value"])

if do_dropna:
    reformatted_df = reformatted_df.dropna(how='all')

if strip_whitespace:
    # Strip whitespace from string columns
    string_cols = reformatted_df.select_dtypes(include=['object']).columns
    for col in string_cols:
        reformatted_df[col] = reformatted_df[col].str.strip()

if reset_index and preserve_old_index:
    reformatted_df.index = df_old.index[:len(reformatted_df)]

if fill_na:
    if fill_na == "0":
        reformatted_df = reformatted_df.fillna(0)
    elif fill_na == "Empty string":
        reformatted_df = reformatted_df.fillna("")
    elif fill_na == "Previous value":
        reformatted_df = reformatted_df.fillna(method='ffill')

# --- Download button
def to_excel_bytes(df):
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reformatted")
    towrite.seek(0)
    return towrite

excel_bytes = to_excel_bytes(reformatted_df)

st.download_button(
    label="Download reformatted Excel",
    data=excel_bytes,
    file_name="reformatted.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("Done — download the file above. The new data has been reformatted to match your old file's structure.")
