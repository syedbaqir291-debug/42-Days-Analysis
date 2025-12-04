import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from tempfile import NamedTemporaryFile

st.title("Excel Statistical Rows Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Read file
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names

    st.write("Sheets found:", sheets)

    start_col = st.text_input("Enter Start Column (e.g. B)")
    end_col = st.text_input("Enter End Column (e.g. E)")

    if st.button("Process File"):
        if not start_col or not end_col:
            st.error("Please enter both start and end columns.")
        else:
            # Load workbook for editing formulas
            wb = load_workbook(uploaded_file)

            for sheet_name in sheets:
                ws = wb[sheet_name]

                # Convert column letters to numbers
                start_idx = ws[start_col + "1"].column
                end_idx = ws[end_col + "1"].column

                # Find last row with data in selected column
                last_row = ws.max_row

                # Column to write results = next column after end column
                result_col = get_column_letter(end_idx + 1)

                # Build selected range
                selected_range = f"{start_col}1:{end_col}{last_row}"

                # Write formulas below existing data
                stats = [
                    ("Mean", "AVERAGE"),
                    ("Median", "MEDIAN"),
                    ("Minimum", "MIN"),
                    ("Maximum", "MAX"),
                    ("Standard Deviation", "STDEV.S")
                ]

                write_row = last_row + 2  # Leave one blank row

                for label, fx in stats:
                    ws[f"{result_col}{write_row}"] = f"={fx}({selected_range})"
                    ws[f"{start_col}{write_row}"] = label
                    write_row += 1

            # Save updated file
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                wb.save(tmp.name)
                tmp.seek(0)
                st.success("File processed successfully!")
                st.download_button(
                    "Download Updated Excel",
                    data=tmp.read(),
                    file_name="updated_file.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
