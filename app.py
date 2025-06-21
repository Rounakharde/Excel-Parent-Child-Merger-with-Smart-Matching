import streamlit as st
import pandas as pd
import io
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import time
from PIL import Image
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("📊 Excel Parent-Child Merger with Smart Matching & Multi-Format Export")

# Function to find the best matching column name using fuzzy matching
def find_best_match(col, columns):
    best_match, score = process.extractOne(col, columns, scorer=fuzz.token_sort_ratio)
    return best_match if score > 70 else None

# Automatically detect correct header row
def detect_header_row(df, max_rows=5):
    for i in range(max_rows):
        if df.iloc[i].isnull().sum() < len(df.columns) * 0.5:
            df.columns = df.iloc[i]
            return df[i+1:].reset_index(drop=True)
    df.columns = [f"Column_{i}" for i in range(df.shape[1])]
    return df.reset_index(drop=True)

# Optimized search filter function using vectorized operations
def fast_search(df, term):
    mask = df.astype(str).apply(lambda x: x.str.contains(term, case=False, na=False))
    return df[mask.any(axis=1)]

# Export functions
def download_file(df, file_type, file_name):
    if file_type == 'CSV':
        return df.to_csv(index=False).encode('utf-8')
    elif file_type == 'XLSX':
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()
    elif file_type == 'PDF':
        fig, ax = plt.subplots(figsize=(12, len(df)//2 + 1))
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, loc='center')
        plt.tight_layout()
        output = io.BytesIO()
        plt.savefig(output, format='pdf')
        plt.close(fig)
        return output.getvalue()
    elif file_type == 'JPG':
        fig, ax = plt.subplots(figsize=(12, len(df)//2 + 1))
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, loc='center')
        plt.tight_layout()
        output = io.BytesIO()
        plt.savefig(output, format='jpeg')
        plt.close(fig)
        return output.getvalue()
    else:
        return None

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=[".xlsx"])

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names

        parent_df, child_df = None, None
        if len(sheet_names) >= 2:
            sheet_dfs = []
            for name in sheet_names:
                raw_df = xl.parse(name, header=None)
                cleaned_df = detect_header_row(raw_df)
                cleaned_df = cleaned_df.dropna(how='all').dropna(axis=1, how='all')
                sheet_dfs.append((name, cleaned_df))

            sorted_sheets = sorted(sheet_dfs, key=lambda x: len(x[1]), reverse=True)
            parent_name, parent_df = sorted_sheets[0]
            child_name, child_df = sorted_sheets[1]
        else:
            st.error("At least two sheets required for Parent and Child.")

        if parent_df is not None and child_df is not None:
            st.subheader(f"📁 Parent Sheet: {parent_name}")
            parent_df = parent_df.dropna(how='all').dropna(axis=1, how='all')
            search_term_parent = st.text_input("Search in Parent Table", key="parent_search")
            if search_term_parent:
                parent_filtered = fast_search(parent_df, search_term_parent)
            else:
                parent_filtered = parent_df
            st.dataframe(parent_filtered, use_container_width=True, height=300)

            st.subheader(f"📁 Child Sheet: {child_name}")
            child_df = child_df.dropna(how='all').dropna(axis=1, how='all')
            search_term_child = st.text_input("Search in Child Table", key="child_search")
            if search_term_child:
                child_filtered = fast_search(child_df, search_term_child)
            else:
                child_filtered = child_df
            st.dataframe(child_filtered, use_container_width=True, height=300)

            # Smart column name matching
            parent_cols = parent_df.columns.astype(str).tolist()
            child_cols = child_df.columns.astype(str).tolist()
            match_found = False
            for p_col in parent_cols:
                best_child_col = find_best_match(p_col, child_cols)
                if best_child_col:
                    try:
                        merged_df = pd.merge(parent_df, child_df, left_on=p_col, right_on=best_child_col, how='inner')
                        match_found = True
                        match_info = f"Merging on Parent Column: **{p_col}** ↔ Child Column: **{best_child_col}**"
                        break
                    except Exception as e:
                        st.warning(f"Merge attempt failed on {p_col} and {best_child_col}: {e}")

            if match_found:
                st.subheader("🔗 Merged Result")
                search_term_merge = st.text_input("Search in Merged Table", key="merged_search")
                if search_term_merge:
                    merged_filtered = fast_search(merged_df, search_term_merge)
                else:
                    merged_filtered = merged_df
                st.markdown(match_info)
                st.dataframe(merged_filtered, use_container_width=True, height=400)

                st.subheader("📥 Download Merged File")
                file_type = st.selectbox("Choose Format", ["CSV", "XLSX", "PDF", "JPG"])
                if st.button("Download"):
                    content = download_file(merged_filtered, file_type, "Merged_Result")
                    if content:
                        mime = {
                            'CSV': 'text/csv',
                            'XLSX': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            'PDF': 'application/pdf',
                            'JPG': 'image/jpeg'
                        }[file_type]
                        st.download_button(f"⬇️ Download as {file_type}", content, file_name=f"Merged_Result.{file_type.lower()}", mime=mime)
                    else:
                        st.warning("Unsupported format selected.")

            else:
                st.warning("⚠️ No matching column names found for merging. Please check headers.")

    except Exception as e:
        st.error(f"Error reading file: {e}")
