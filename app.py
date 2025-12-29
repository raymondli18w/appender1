import streamlit as st
import pandas as pd
from io import BytesIO

def merge_excel_files(uploaded_files, match_by_name=True):
    dataframes = []
    all_columns_sets = []

    for file in uploaded_files:
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = df.columns.astype(str).str.strip()  # Normalize column names
        dataframes.append(df)
        all_columns_sets.append(set(df.columns))

    if not dataframes:
        return pd.DataFrame()

    # Use intersection: only columns in ALL files
    common_columns = set.intersection(*all_columns_sets) if all_columns_sets else set()

    if not common_columns:
        st.warning("No common columns found across all files.")
        return pd.DataFrame()

    aligned_dfs = []
    for df in dataframes:
        aligned_df = df.reindex(columns=sorted(common_columns))
        aligned_dfs.append(aligned_df)

    merged_df = pd.concat(aligned_dfs, ignore_index=True)
    return merged_df

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Merged Data')
    return output.getvalue()

# === Streamlit UI ===
st.title("📊 Excel File Merger – Column Name Matching")
st.write("""
Upload multiple `.xlsx` files. The app will:
- Match columns by **name** (order doesn’t matter).
- Keep only columns that appear in **all files**.
- Combine all data into one table with a single header.
""")

uploaded_files = st.file_uploader(
    "📁 Upload Excel files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📄 {len(uploaded_files)} file(s) uploaded.")

    if len(uploaded_files) == 1:
        st.info("Please upload at least two files to merge.")
    else:
        with st.spinner("Merging files..."):
            merged_df = merge_excel_files(uploaded_files)

        if not merged_df.empty:
            st.success(f"✅ Merged! Final data has {merged_df.shape[0]} rows and {merged_df.shape[1]} columns.")
            st.dataframe(merged_df)

            excel_data = to_excel_bytes(merged_df)
            st.download_button(
                label="📥 Download Merged Excel File",
                data=excel_data,
                file_name="merged_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ Failed to merge: no common columns found.")
else:
    st.info("👆 Upload Excel files to begin.")