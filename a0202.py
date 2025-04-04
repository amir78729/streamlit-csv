import pandas as pd
import streamlit as st
import io
import os
from openpyxl.styles import Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from datetime import datetime

SEP = ';'
ENCODING = 'ISO-8859-1'

# Styling
header_font = Font(bold=True)
border_style = Border(left=Side(style='thin'), right=Side(style='thin'),
                      top=Side(style='thin'), bottom=Side(style='thin'))

# Streamlit UI
st.title("A02_02: Sjuk o frånvaro med omf")
keep_all_data = st.checkbox("Keep original data as a worksheet", value=False)
uploaded_file = st.file_uploader("Upload your file", type=["xlsx"])
month = st.radio("Select Month", options=[str(i).zfill(2) for i in range(1, 13)])  # Months as 01, 02, ... 12
is_before_payment = st.checkbox("Is this process before payment?", value=True)

st.info("Don't worry! All calculations are local and data is not shared.")

if uploaded_file is not None:
    base_filename = os.path.splitext(uploaded_file.name)[0]
    excel_result_filename = f"{base_filename}-result.xlsx"

    with st.status("Reading File"):
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        st.write(df)
        df['Sem Gfom'] = df['Sem Gfom'].astype(str)

    # Filter data
    df_filtered = pd.DataFrame()

    with st.status("Step 1: Filtering data"):
        df_filtered = df[
            (df['Lbertom'] == 0)  # Check Lbertom value
            & ((df['Semester'] == 'Semsjuk') | (df['Semester'] == 'Semsjuk1'))  # Check if Semester is 'Semsjuk' or 'Semsjuk1'
            & (df['Sjuk Korr'] == 1)  # Check Sjuk Korr value
            & ((df['Sem Omf'] + df['Annan Omf']) == 1)  # Sum of 'Sem Omf' and 'Annan Omf' equals 1
            | ((df['Sem Omf'] + df['Annan Omf']) != 1) & (df['Sem Korr'] == 2)  # OR condition for Sem Korr == 2
        ]

        # Filter by month
        df_filtered = df_filtered[df_filtered['Sem Gfom'].str[:6] == f"{datetime.now().year}{month}"]

        st.write('Rows that match filtering criteria:')
        st.write(df_filtered)

        # Merge to remove filtered data from df
        try:
            df = pd.merge(df, df_filtered, how="outer", indicator=True).query("_merge != 'both'").drop('_merge', axis=1).reset_index(drop=True)
        except: 
            pass

        st.write('Filtered data removed, resulting data:')
        st.write(df)

    if is_before_payment:
        with st.status("Filter for before payment"):
            df_filtered = df[((df['Sem Omf'] + df['Annan Omf']) == 1)]

            # Filter by month
            df_filtered = df_filtered[df_filtered['Sem Gfom'].str[:6] == f"{datetime.now().year}{month}"]

            st.write('Rows that match filtering criteria:')
            st.write(df_filtered)

            # Merge to remove filtered data from df
            try:
                df = pd.merge(df, df_filtered, how="outer", indicator=True).query("_merge != 'both'").drop('_merge', axis=1).reset_index(drop=True)
            except: 
                pass

            st.write('Filtered data removed, resulting data:')
            st.write(df)

    with st.status("Filter based on Semsjuk and Semsjuk1 Semesters"):
        df_filtered = df[
            ((df['Semester'] == 'Semsjuk') | (df['Semester'] == 'Semsjuk1'))  # Check if Semester is 'Semsjuk' or 'Semsjuk1'
            & ((df['Sem Omf'] + df['Annan Omf']) == 1)  # Sum of 'Sem Omf' and 'Annan Omf' equals 1
        ]

        # Filter by month
        df_filtered = df_filtered[df_filtered['Sem Gfom'].str[:6] == f"{datetime.now().year}{month}"]

        st.write('Rows that match filtering criteria:')
        st.write(df_filtered)

        # Merge to remove filtered data from df
        try:
            df = pd.merge(df, df_filtered, how="outer", indicator=True).query("_merge != 'both'").drop('_merge', axis=1).reset_index(drop=True)
        except: 
            pass

        st.write('Filtered data removed, resulting data:')
        st.write(df)

    with st.status("Filter based on Sembet and Sembet1 Semesters"):
        df_filtered = df[
            ((df['Semester'] == 'Sembet') | (df['Semester'] == 'Sembet1'))  # Check if Semester is 'Semsjuk' or 'Semsjuk1'
            & ((df['Sjuk Korr']) == 1)
            & ((df['Sem Korr']) == 2)
        ]

        # Filter by month
        df_filtered = df_filtered[df_filtered['Sem Gfom'].str[:6] == f"{datetime.now().year}{month}"]

        st.write('Rows that match filtering criteria:')
        st.write(df_filtered)

        # Merge to remove filtered data from df
        try:
            df = pd.merge(df, df_filtered, how="outer", indicator=True).query("_merge != 'both'").drop('_merge', axis=1).reset_index(drop=True)
        except: 
            pass

        st.write('Filtered data removed, resulting data:')
        st.write(df)
    with st.status("Filter Sjuk Korr = 2"):
        df['Pnr'] = df['Pnr'].astype(str).str.strip()

        # After filtering by 'Sjuk Korr' == 1
        df_filtered = df[df['Sjuk Korr'] == 1]

        # Reset index to avoid issues with groupby
        df_filtered = df_filtered.reset_index(drop=True)

        # Count occurrences of each 'Pnr'
        pnr_counts = df_filtered.groupby('Pnr').size()

        # Filter out rows where 'Pnr' appears only once (unique Pnr)
        df_filtered = df_filtered[df_filtered['Pnr'].isin(pnr_counts[pnr_counts > 1].index)]

        st.write('Rows with unique Pnr removed:')
        st.write(df_filtered)

        # Merge to remove filtered data from df
        try:
            df = pd.merge(df, df_filtered, how="outer", indicator=True).query("_merge != 'both'").drop('_merge', axis=1).reset_index(drop=True)
        except: 
            pass

        st.write('Filtered data removed, resulting data:')
        st.write(df)
        
        

    with st.status("Filter based on Förv Bolag value"):
        df_filtered = df[
            ((df['Förv Bolag'] == 'Arbvux') | (df['Förv Bolag'] == 'AMA'))
        ]

        # Filter by month
        df_filtered = df_filtered[df_filtered['Sem Gfom'].str[:6] == f"{datetime.now().year}{month}"]

        st.write('Rows that match filtering criteria:')
        st.write(df_filtered)

        # Merge to remove filtered data from df
        try:
            df = pd.merge(df, df_filtered, how="outer", indicator=True).query("_merge != 'both'").drop('_merge', axis=1).reset_index(drop=True)
        except: 
            pass

        st.write('Filtered data removed, resulting data:')
        st.write(df)
    
    with st.status("Creating Excel File"):
        output = io.BytesIO()

        # Create an Excel file using openpyxl
        wb = Workbook()
        ws_all = wb.active
        ws_all.title = "Filtered Data"

        # Write the DataFrame to the sheet
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws_all.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    cell.font = header_font
                    cell.border = border_style

        # Apply auto-filters
        ws_all.auto_filter.ref = ws_all.dimensions

        # Save the workbook to the output stream
        wb.save(output)
        processed_data = output.getvalue()

    # Provide download button
    st.download_button(
        label="Download Excel",
        data=processed_data,
        file_name=excel_result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
