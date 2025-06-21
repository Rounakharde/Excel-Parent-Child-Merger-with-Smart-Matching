import streamlit as st
import pandas as pd
import io
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

st.set_page_config(layout="wide")
st.title("📊 Excel Parent-Child Merger with Smart Matching")

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

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=[".xlsx"])

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names

        # Automatically detect Parent and Child sheets based on size
        parent_df, child_df = None, None
        if len(sheet_names) >= 2:
            sheet_dfs = []
            for name in sheet_names:
                raw_df = xl.parse(name, header=None)
                cleaned_df = detect_header_row(raw_df)
                cleaned_df = cleaned_df.dropna(how='all')
                cleaned_df = cleaned_df.dropna(axis=1, how='all')
                sheet_dfs.append((name, cleaned_df))

            sorted_sheets = sorted(sheet_dfs, key=lambda x: len(x[1]), reverse=True)
            parent_name, parent_df = sorted_sheets[0]
            child_name, child_df = sorted_sheets[1]
        else:
            st.error("At least two sheets required for Parent and Child.")

        if parent_df is not None and child_df is not None:
            st.subheader(f"📁 Parent Sheet: {parent_name}")
            parent_df = parent_df.dropna(how='all').dropna(axis=1, how='all')
            search_term_parent = st.text_input("Search in Parent Table")
            if search_term_parent:
                parent_filtered = parent_df[parent_df.apply(lambda row: row.astype(str).str.contains(search_term_parent, case=False).any(), axis=1)]
            else:
                parent_filtered = parent_df
            st.dataframe(parent_filtered, use_container_width=True)

            st.subheader(f"📁 Child Sheet: {child_name}")
            child_df = child_df.dropna(how='all').dropna(axis=1, how='all')
            search_term_child = st.text_input("Search in Child Table")
            if search_term_child:
                child_filtered = child_df[child_df.apply(lambda row: row.astype(str).str.contains(search_term_child, case=False).any(), axis=1)]
            else:
                child_filtered = child_df
            st.dataframe(child_filtered, use_container_width=True)

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
                search_term_merge = st.text_input("Search in Merged Table")
                if search_term_merge:
                    merged_filtered = merged_df[merged_df.apply(lambda row: row.astype(str).str.contains(search_term_merge, case=False).any(), axis=1)]
                else:
                    merged_filtered = merged_df
                st.markdown(match_info)
                st.dataframe(merged_filtered, use_container_width=True)

                # Download merged file
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    parent_df.to_excel(writer, sheet_name='Parent', index=False)
                    child_df.to_excel(writer, sheet_name='Child', index=False)
                    merged_df.to_excel(writer, sheet_name='Merged', index=False)
                st.download_button("📥 Download Merged Excel", output.getvalue(), file_name="Merged_Result.xlsx")

            else:
                st.warning("⚠️ No matching column names found for merging. Please check headers.")

    except Exception as e:
        st.error(f"Error reading file: {e}")
