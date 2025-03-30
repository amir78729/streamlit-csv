import pandas as pd
import streamlit as st
import io
import os
from utils.date import is_holiday

SEP = ';'
ENCODING = 'utf-8'

st.title("A05 07 List Noll Bearb")

uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])

st.info("Don't worry! all the calculations are running in your machine and the data is not being shared anywhere.")

col = {
    'Förv/Bolag': 'Förv/Bolag',
    'Semgrp': 'Semgrp',
    'Arbhel': 'Arbhel',
    'Gfom': 'Gfom',
    'Orsak': 'Orsak',
    'Orstxt': 'Orstxt',
    'Frånvaro': 'Frånvaro',
    'Omf': 'Omf',
    'Specantal': 'Specantal',
    'Antal': 'Antal'
}

# This block runs only if a file has been uploaded
if uploaded_file is not None:
    with st.status("Reading CSV File"):
        df = pd.read_csv(uploaded_file, sep=SEP, encoding=ENCODING)
        st.write(df)
        
    with st.status("Cleaning Data"):
        st.write("Remove quetes...")
        df = df.replace('"', '')
        
        try:
            st.write("Removing `{}` column...".format(col['Orsak']))
            df = df.drop(col['Orsak'], axis=1)
        except:
            pass

        try:
            st.write("Renaming `{}` column to `{}`...".format(col['Orsak'], col['Frånvaro']))
            df = df.rename(columns={col['Orstxt']: col['Frånvaro']})
        except:
            pass

        st.write("Removing redundant values in `{}` column...".format(col['Frånvaro']))
        df = df[~df[col['Frånvaro']].str.lower().isin([
            'sjukfk semgr fer/ upp1-45',
            'sjuk från dag 46',
            'sjuk obetald ferielön',
            'sjuk från dag 46',
        ])]
        
        for c in [col['Arbhel'], col["Omf"], col['Specantal'], col['Antal']]:
            st.write('Replace `,` to `.` in floating numbers for `{}`...'.format(c))
            df[c] = df[c].replace(',', '.', regex=True)
            st.write(df)
        
        for c in [col['Arbhel'], col["Omf"], col['Specantal'], col['Antal']]:
            st.write('Round `{}` values...'.format(c))
            df[c] = df[c].astype(float).round(2)
        st.write(df)

    with st.status("Filtering File"):
        st.write('Filtering `{}`...'.format(col['Förv/Bolag']))
        df_filtered = df[df[col['Förv/Bolag']].str.strip() != 'KulturN']
        st.write(df_filtered)

        st.write('Filtering `{}`...'.format(col['Semgrp']))
        df_filtered = df_filtered[~df_filtered[col['Semgrp']].isin(['1', '8', '9', '22'])]        
        st.write(df_filtered)

        st.write('Filtering `{}`...'.format(col['Arbhel']))
        df_filtered = df_filtered[df_filtered[col['Arbhel']].astype(str).str.startswith(('35', '40'))]
        st.write(df_filtered)

        st.write('Filtering `{}`...'.format(col['Gfom']))
        df_filtered = df_filtered[df_filtered[col['Gfom']].apply(is_holiday)]
        st.write(df_filtered)

        st.write('Following rows should be removed:')
        st.write(df_filtered)

        try:
            df_result = pd.merge(df, df_filtered, how="outer", indicator=True).query("_merge != 'both'").drop('_merge', axis=1).reset_index(drop=True)
        except: 
            df_result = df
            
    with st.status("Creating Excel File"):
        st.write("## Filtered Data", df_result)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, index=False, sheet_name='FilteredData')
        processed_data = output.getvalue()

    base_filename = os.path.splitext(uploaded_file.name)[0]
    csv_result_filename = f"{base_filename}-result.csv"
    excel_result_filename = f"{base_filename}-result.xlsx"

    # Replace the download buttons with:
    st.download_button(
        "Download Result CSV",
        df_result.to_csv(sep=SEP, index=False, encoding=ENCODING),
        file_name=csv_result_filename,
    )

    st.download_button(
        label="Download Result Excel",
        data=processed_data,
        file_name=excel_result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
