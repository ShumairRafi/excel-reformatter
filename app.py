# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Excel Reformatter", layout="wide")

st.title("Excel Reformatter — map new data to old file format")
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
    allow_partial = st.checkbox("Allow leaving unmapped columns (fills with blank)", value=True)

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
            raise

try:
    df_old = read_excel(old_file, sheet_old)
    df_new = read_excel(new_file, sheet_new)
except Exception as e:
    st.error(f"Could not read one of the uploaded Excel files: {e}")
    st.stop()

st.subheader("Preview — columns")

c1, c2 = st.columns(2)
with c1:
    st.write("Old file (target format) — first 5 rows")
    st.dataframe(df_old.head())
with c2:
    st.write("New file (incoming data) — first 5 rows")
    st.dataframe(df_new.head())

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
    choice = st.selectbox(label, options=cols_for_select, index=(cols_for_select.index(suggested_col) if suggested_col in cols_for_select else 0), key=f"map_{oc}")
    mapping[oc] = None if choice == "-- (blank)" else choice

# Optionally show reverse mapping conflicts
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
def build_reformatted(df_new, mapping, old_cols, allow_partial=True):
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
    # Keep the same dtypes where possible? We'll attempt to align simple numeric/date types
    return out_df

reformatted_df = build_reformatted(df_new, mapping, old_cols, allow_partial=allow_partial)

st.subheader("Reformatted preview (first 10 rows)")
st.dataframe(reformatted_df.head(10))

# --- Post-processing helpers
st.markdown("### Post-processing (basic)")
st.write("You can do basic cleaning here (optional). For more advanced transforms, download and modify manually.")

do_dropna = st.checkbox("Drop rows where all values are blank", value=False)
if do_dropna:
    reformatted_df = reformatted_df.dropna(how="all")

# --- Download button
def to_excel_bytes(df):
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        writer.save()
    towrite.seek(0)
    return towrite

excel_bytes = to_excel_bytes(reformatted_df)

st.download_button(
    label="Download reformatted Excel",
    data=excel_bytes,
    file_name="reformatted.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("Done — download the file above. If you need more advanced transformations (column splits, date parsing, lookups), tell me and I can extend the app.")
