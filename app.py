import streamlit as st
import pandas as pd
import numpy as np
import io
import base64
from io import BytesIO
import matplotlib.pyplot as plt
from fpdf import FPDF
from difflib import SequenceMatcher

st.set_page_config(layout="wide")
st.title("📁 Excel Parent-Child Merger & Filter Tool")

# Function to find best matching column using similarity
def find_best_match(col1, col2):
    best_match = None
    highest_ratio = 0
    for c1 in col1:
        for c2 in col2:
            ratio = SequenceMatcher(None, c1.lower(), c2.lower()).ratio()
            if ratio > highest_ratio:
                highest_ratio = ratio
                best_match = (c1, c2)
    return best_match if highest_ratio > 0.5 else (None, None)

# File uploader
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    # Auto-select sheet names
    if len(sheet_names) >= 2:
        parent_sheet = sheet_names[0]
        child_sheet = sheet_names[1]
    else:
        st.error("Please upload an Excel file with at least two sheets (Parent and Child).")
        st.stop()

    # Read sheets with dynamic header detection
    def read_with_header_guess(sheet):
        df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
        header_row_index = df_raw.first_valid_index()
        for i in range(min(5, len(df_raw))):
            if df_raw.iloc[i].nunique() == len(df_raw.columns):
                header_row_index = i
                break
        df = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row_index)
        return df

    parent_df = read_with_header_guess(parent_sheet)
    child_df = read_with_header_guess(child_sheet)

    st.subheader("📊 Parent Table")
    st.dataframe(parent_df, use_container_width=True)

    st.subheader("📊 Child Table")
    st.dataframe(child_df, use_container_width=True)

    # Auto match columns
    parent_col, child_col = find_best_match(parent_df.columns, child_df.columns)

    if not parent_col or not child_col:
        st.error("⚠️ No matching columns found for merging. Please check headers.")
        st.stop()

    merged_df = pd.merge(parent_df, child_df, left_on=parent_col, right_on=child_col, how="inner")

    # ---------- COLUMN SELECTOR ----------
    st.subheader("🧩 Column Selector")
    selected_columns = st.multiselect("Select columns to display", merged_df.columns.tolist(), default=merged_df.columns.tolist())
    filtered_merged_df = merged_df[selected_columns]

    # ---------- ROW SEARCH ----------
    st.subheader("🔍 Search Rows")
    search_text = st.text_input("Enter text to search in all rows")
    if search_text:
        filtered_merged_df = filtered_merged_df[filtered_merged_df.apply(lambda row: row.astype(str).str.contains(search_text, case=False).any(), axis=1)]

    # ---------- FILTER UI IN SHORT FORM ----------
    with st.expander("🔧 Column Filters"):
        col1, col2, col3, col4 = st.columns(4)
        filter_column = col1.selectbox("Column", filtered_merged_df.columns)
        condition = col2.selectbox("Cond", ["=", "!=", ">", "<", ">=", "<="])
        filter_value = col3.text_input("Value")
        if col4.button("Filter"):
            try:
                val_type = type(filtered_merged_df[filter_column].dropna().iloc[0])
                val = val_type(filter_value)
                if condition == "=":
                    filtered_merged_df = filtered_merged_df[filtered_merged_df[filter_column] == val]
                elif condition == "!=":
                    filtered_merged_df = filtered_merged_df[filtered_merged_df[filter_column] != val]
                elif condition == ">":
                    filtered_merged_df = filtered_merged_df[filtered_merged_df[filter_column] > val]
                elif condition == "<":
                    filtered_merged_df = filtered_merged_df[filtered_merged_df[filter_column] < val]
                elif condition == ">=":
                    filtered_merged_df = filtered_merged_df[filtered_merged_df[filter_column] >= val]
                elif condition == "<=":
                    filtered_merged_df = filtered_merged_df[filtered_merged_df[filter_column] <= val]
            except Exception as e:
                st.warning(f"Filter error: {e}")

    st.subheader("🔗 Merged Table")
    st.dataframe(filtered_merged_df, use_container_width=True)

    st.subheader("📥 Download Merged File")

    def convert_df_to_csv(df):
        return df.to_csv(index=False).encode("utf-8")

    def convert_df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Merged")
        return output.getvalue()

    def convert_df_to_pdf(df):
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.add_page()
        pdf.set_font("Arial", size=10)
        pdf.set_auto_page_break(auto=True, margin=10)
        col_width = pdf.w / (len(df.columns) + 1)
        row_height = 8
        for col in df.columns:
            pdf.cell(col_width, row_height * 1.25, str(col), border=1)
        pdf.ln(row_height * 1.25)
        for i in range(len(df)):
            for col in df.columns:
                value = str(df.iloc[i][col])
                pdf.cell(col_width, row_height, value[:30], border=1)
            pdf.ln(row_height)
        return pdf.output(dest='S').encode('latin-1')

    colA, colB, colC = st.columns(3)

    with colA:
        csv = convert_df_to_csv(filtered_merged_df)
        st.download_button("⬇️ Download CSV", csv, file_name="merged.csv", mime="text/csv")

    with colB:
        excel = convert_df_to_excel(filtered_merged_df)
        st.download_button("⬇️ Download XLSX", excel, file_name="merged.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with colC:
        pdf_bytes = convert_df_to_pdf(filtered_merged_df)
        st.download_button("⬇️ Download PDF", pdf_bytes, file_name="merged.pdf", mime="application/pdf")
