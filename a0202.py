import pandas as pd
import streamlit as st
from datetime import datetime
import io
import os



SEP = ';'
ENCODING = 'latin1'

st.title("A02_02: Sjuk o frånvaro med omf")
st.write("Implementing...")

# uploaded_file = st.file_uploader("Upload a CSV file", type=["csv"])
# # This block runs only if a file has been uploaded
# if uploaded_file is not None:
#     with st.status("Reading Excel File"):
#       df = pd.read_csv(uploaded_file, sep=SEP, encoding=ENCODING)
#       st.write(df)
      
#     # df['Verklig lön'] = pd.to_numeric(df['Verklig lön'], errors='coerce').fillna(0).astype(int)
    
#     base_filename = os.path.splitext(uploaded_file.name)[0]
#     excel_result_filename = f"{base_filename}-result.xlsx"
    
#     with pd.ExcelWriter(excel_result_filename, engine='xlsxwriter') as writer:
#         df.to_excel(writer, index=False, sheet_name='Sheet1')
#         workbook = writer.book
#         worksheet = writer.sheets['Sheet1']

#         # Define formats
#         number_format = workbook.add_format({'align': 'center', 'bold': True})
#         purple_format = workbook.add_format({'align': 'center', 'bold': True, 'font_color': 'purple'})

#         # Apply formats row by row
#         for row_num, value in enumerate(df['Verklig lön'], start=1):  # Start from the first data row
#             if value >= 47750:
#                 worksheet.write(row_num, df.columns.get_loc('Verklig lön'), value, purple_format)
#             else:
#                 worksheet.write(row_num, df.columns.get_loc('Verklig lön'), value, number_format)
#      # File downloader
#     with open(excel_result_filename, "rb") as file:
#         st.download_button(
#             label="Download Formatted Excel File",
#             data=file,
#             file_name=excel_result_filename,
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )