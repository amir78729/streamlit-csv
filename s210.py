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
    ws_all = wb.active
    ws_all.title = "all data"

    # Create additional sheets
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

        # Track matching rows
        matched_rows = []

        # Write data to "all data" and capture matching rows
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws_all.cell(row=r_idx, column=c_idx, value=value)
                cell.border = border_style

                if r_idx == 1:
                    cell.font = header_font
                    headers = row  # Save headers for reuse
                else:
                    col_letter = ws_all.cell(row=1, column=c_idx).value

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

            # Check matching condition
            if r_idx > 1:
                row_dict = dict(zip(headers, row))
                try:
                    lon_val = int(row_dict.get("Verklig lön", "0").split(',')[0])
                    frn_ors = str(row_dict.get("Frn Ors", "")).strip()
                    if lon_val >= 49000 and frn_ors in ["FörlSj>2", "Sjuk", "Sj>2år", "Sjukförl"]:
                        matched_rows.append(row)
                except Exception as e:
                    pass

        extra_matched_rows = []
        for row in dataframe_to_rows(df, index=False, header=True):
            row_dict = dict(zip(headers, row))
            try:
                lon_val = int(row_dict.get("Verklig lön", "0").split(',')[0])
                frn_ors = str(row_dict.get("Frn Ors", "")).strip()
                if lon_val < 49000 and frn_ors in [
                    "Arbsk>2", "FörlSj>2", "Sj>180", "Sj>180Se", "Sj>2år", "SjFK>år", "Sjuk"
                ]:
                    extra_matched_rows.append(row)
            except:
                pass

        # Copy headers to each sheet
        for sheet in sheets.values():
            for c_idx, header in enumerate(headers, 1):
                cell = sheet.cell(row=1, column=c_idx, value=header)
                cell.font = header_font
                cell.border = border_style

        # Copy matched rows into 4 specific sheets
        target_sheets = [
            "Sjuk",
            "Sjuk >år",
            "Sjuk Fortsättningsnivå",
            "Sjuk Fortsättningsnivå>År"
        ]
        for sheet_name in target_sheets:
            sheet = sheets[sheet_name]
            for r_offset, row in enumerate(matched_rows, start=2):
                for c_idx, value in enumerate(row, 1):
                    col_name = headers[c_idx - 1]
                    cell = sheet.cell(row=r_offset, column=c_idx, value=value)
                    cell.border = border_style

                    # Apply same logic for each relevant column
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
                            
        for sheet_name in ["Sjuk", "Sjuk >år"]:
            sheet = sheets[sheet_name]
            existing_row_count = sheet.max_row
            for r_offset, row in enumerate(extra_matched_rows, start=existing_row_count + 1):
                for c_idx, value in enumerate(row, 1):
                    col_name = headers[c_idx - 1]
                    cell = sheet.cell(row=r_offset, column=c_idx, value=value)
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

    ws_all.auto_filter.ref = ws_all.dimensions
    for sheet in sheets.values():
        sheet.auto_filter.ref = sheet.dimensions

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="Download Result Excel",
        data=output,
        file_name=excel_result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
