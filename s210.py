import pandas as pd
import streamlit as st
import io
import os
from utils.date import is_holiday
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook

SEP = ';'
ENCODING = 'ISO-8859-1'

st.title("S2_10: Sjuk mer än 365 dagar")

uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])

st.info("Don't worry! all the calculations are running in your machine and the data is not being shared anywhere.")

if uploaded_file is not None:
    base_filename = os.path.splitext(uploaded_file.name)[0]
    excel_result_filename = f"{base_filename}-result.xlsx"

    with st.status("Reading CSV File"):
        df = pd.read_csv(uploaded_file, sep=SEP, encoding=ENCODING)
        st.write(df)
        
    wb = Workbook()
    ws = wb.active
    ws.title = "all data"
    
    sheet_names = [
        "Sjuk",
        "Sjuk >år",
        "Sjuk Fortsättningsnivå",
        "Sjuk Fortsättningsnivå>År",
        "Sjukers",
        "Arbetsskadelivränta"
    ]
    
    sheets = {name: wb.create_sheet(title=name) for name in sheet_names}
        
    with st.status("Update cell styles"):
        header_font = Font(bold=True)
        bold_font = Font(bold=True)
        center_alignment = Alignment(horizontal='center')
        border_style = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )
        purple_font = Font(color="800080", bold=True)
        orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.border = border_style

                if r_idx == 1:
                    cell.font = header_font
                    headers = row 
                else:
                    col_letter = ws.cell(row=1, column=c_idx).value

                    if col_letter == "Omf":
                        cell.value = round(float(value.replace(',', '.')), 3)

                    elif col_letter == "Verklig lön":
                        cell.alignment = center_alignment
                        number_value = int(value.split(',')[0])
                        cell.font = purple_font if number_value >= 49000 else bold_font
                        cell.value = number_value
                        cell.number_format = '#,##0'

                    elif col_letter == "Sem Grp":
                        if str(value) in ['8', '9']:
                            cell.fill = orange_fill

        st.write(df)

    ws.auto_filter.ref = ws.dimensions

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="Download Result Excel",
        data=output,
        file_name=excel_result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )