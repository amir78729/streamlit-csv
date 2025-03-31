import pandas as pd
import streamlit as st
import io
import os
from utils.date import is_holiday
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook

# Constants
SEP = ';'
ENCODING = 'ISO-8859-1'
TARGET_SHEETS_HIGH = ["Sjuk", "Sjuk >år", "Sjuk Fortsättningsnivå", "Sjuk Fortsättningsnivå>År"]
TARGET_SHEETS_LOW = ["Sjuk", "Sjuk >år"]
SHEET_NAMES = TARGET_SHEETS_HIGH + ["Sjukers", "Arbetsskadelivränta"]

# Styling
header_font = Font(bold=True)
bold_font = Font(bold=True)
center_alignment = Alignment(horizontal='center')
border_style = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)
purple_font = Font(color="800080", bold=True)
orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

# Functions
def apply_cell_style(cell, col_name, value):
    cell.border = border_style
    if col_name == "Omf":
        try:
            cell.value = round(float(str(value).replace(',', '.')), 3)
        except:
            pass
    elif col_name == "Verklig lön":
        try:
            number_value = int(str(value).split(',')[0])
            cell.alignment = center_alignment
            cell.font = purple_font if number_value >= 49000 else bold_font
            cell.value = number_value
            cell.number_format = '#,##0'
        except:
            pass
    elif col_name == "Sem Grp":
        if str(value) in ['8', '9']:
            cell.fill = orange_fill

def copy_rows_to_sheet(sheet, headers, rows, start_row=2):
    for r_offset, row in enumerate(rows, start=start_row):
        for c_idx, value in enumerate(row, 1):
            col_name = headers[c_idx - 1]
            cell = sheet.cell(row=r_offset, column=c_idx, value=value)
            apply_cell_style(cell, col_name, value)

def extract_matching_rows(df, headers, condition_fn):
    matched = []
    for row in dataframe_to_rows(df, index=False, header=True):
        if row == headers:
            continue
        row_dict = dict(zip(headers, row))
        if condition_fn(row_dict):
            matched.append(row)
    return matched

# Streamlit UI
st.title("S2_10: Sjuk mer än 365 dagar")
keep_all_data = st.checkbox("Keep original data as a worksheet", value=False)
uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])
st.info("Don't worry! all the calculations are running in your machine and the data is not being shared anywhere.")

if uploaded_file is not None:
    base_filename = os.path.splitext(uploaded_file.name)[0]
    excel_result_filename = f"{base_filename}-result.xlsx"

    with st.status("Reading CSV File"):
        df = pd.read_csv(uploaded_file, sep=SEP, encoding=ENCODING)
        st.write(df)

    wb = Workbook()

    if keep_all_data:
        ws_all = wb.active
        ws_all.title = "All data"
    else:
        wb.remove(wb.active)
        ws_all = None  # So we can skip writing to it later
    
    sheets = {name: wb.create_sheet(title=name) for name in SHEET_NAMES}
    
    # set the active sheet    
    wb.active = wb[sheets["Sjuk"].title]


    with st.status("Update cell values"):
        headers = df.columns.tolist()

        # Write all data + styles
        if keep_all_data:
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    col_name = headers[c_idx - 1]
                    cell = ws_all.cell(row=r_idx, column=c_idx, value=value)
                    if r_idx == 1:
                        cell.font = header_font
                    else:
                        apply_cell_style(cell, col_name, value)

        # Apply headers to all sheets
        for sheet in sheets.values():
            for c_idx, header in enumerate(headers, 1):
                cell = sheet.cell(row=1, column=c_idx, value=header)
                cell.font = header_font
                cell.border = border_style

        # Match logic
        high_income_rows = extract_matching_rows(
            df, headers,
            lambda row: int(str(row.get("Verklig lön", "0")).split(',')[0]) >= 49000
            and row.get("Frn Ors", "").strip() in ["FörlSj>2", "Sjuk", "Sj>2år", "Sjukförl"]
        )
        low_income_rows = extract_matching_rows(
            df, headers,
            lambda row: int(str(row.get("Verklig lön", "0")).split(',')[0]) < 49000
            and row.get("Frn Ors", "").strip() in [
                "Arbsk>2", "FörlSj>2", "Sj>180", "Sj>180Se", "Sj>2år", "SjFK>år", "Sjuk"
            ]
        )

        # Copy matched rows
        for name in TARGET_SHEETS_HIGH:
            copy_rows_to_sheet(sheets[name], headers, high_income_rows, start_row=2)
        for name in TARGET_SHEETS_LOW:
            existing_rows = sheets[name].max_row
            copy_rows_to_sheet(sheets[name], headers, low_income_rows, start_row=existing_rows + 1)

    # Auto-filters
    if ws_all:
        ws_all.auto_filter.ref = ws_all.dimensions
    
    for sheet in sheets.values():
        sheet.auto_filter.ref = sheet.dimensions

    with st.status("Creating the excel file"):
        # Save and export
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

    st.download_button(
        label="Download Result Excel",
        data=output,
        file_name=excel_result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
