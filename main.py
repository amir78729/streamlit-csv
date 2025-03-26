# Import necessary libraries
import pandas as pd  # Used for handling CSV data
import streamlit as st  # Used to create the web interface
from datetime import datetime  # Helps work with dates

# Define some constants
SEP = ';'  # The character that separates values in the CSV file (semicolon)
ENCODING = 'utf-8'  # Encoding used to read/write files

# List of public holidays – add dates here in "YYYYMMDD" format, like "20231225" for Dec 25, 2023
PUBLIC_HOLIDAYS = [
    # 2025
    "20250101",
    "20250106",
    "20250418",
    "20250420",
    "20250421",
    "20250501",
    "20250529",
    "20250606",
    "20250608",
    "20250621",
    "20251101",
    "20251225",
    "20251226",

    # 2026
    "20260101",
    "20260106",
    "20260403",
    "20260405",
    "20260406",
    "20260501",
    "20260514",
    "20260524",
    "20260606",
    "20260620",
    "20261031",
    "20261225",
    "20261226",
]

# This function checks if a date is a weekend or a public holiday
def is_holiday(date_str):
    # Check if the date is in the list of public holidays
    is_public_holiday = str(date_str) in PUBLIC_HOLIDAYS
    if is_public_holiday:
        st.write('- `{}` is a holiday 🎄, removing it...'.format(date_str))
        return True

    # Convert the date from string to a datetime object
    date = datetime.strptime(str(date_str), "%Y%m%d")

    # Check if it's Saturday (5) or Sunday (6)
    is_weekend = date.weekday() in [5, 6]

    if is_weekend:
        st.write('- `{}` is a weekend 😴, removing it...'.format(date_str))
        return True

    # If it's not a holiday or weekend, keep the date
    st.write('- `{}` is a workday 💼, keeping it...'.format(date_str))
    return False

# Set the title of the web app
st.title("CSV Filter Tool")

# Let the user upload a CSV file
uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])

# Show an info message to let users know their data is safe
st.info("Don't worry! all the calculations are running in your machine and the data is not being shared anywhere.")

# These are the names of the columns we will use for filtering
col = {
    'Förv/Bolag': 'Förv/Bolag',
    'Semgrp': 'Semgrp',
    'Arbhel': 'Arbhel',
    'Gfom': 'Gfom',
}

# This block runs only if a file has been uploaded
if uploaded_file is not None:
    # Show a status box while processing
    with st.status("Processing File"):
        st.write('### Reading file...')
        # Read the CSV file into a DataFrame
        df = pd.read_csv(uploaded_file, sep=SEP, encoding=ENCODING)
        st.write('- done!')

        st.write("### Cleaning Data...")
        # Remove double quotes from the data
        df = df.replace('"', '')
        st.write('- done!')

        # Filter out rows where the 'Förv/Bolag' column is equal to 'KulturN'
        st.write('### Filtering `{}`...'.format(col['Förv/Bolag']))
        df = df[df[col['Förv/Bolag']].str.strip() != 'KulturN']
        st.write('- done!')

        # Filter out rows where 'Semgrp' column is in the list ['1', '8', '9', '22']
        st.write('### Filtering `{}`...'.format(col['Semgrp']))
        df = df[~df[col['Semgrp']].isin(['1', '8', '9', '22'])]
        st.write('- done!')

        # Filter out rows where 'Arbhel' column is in the list ['35', '40']
        st.write('### Filtering `{}`...'.format(col['Arbhel']))
        df = df[~df[col['Arbhel']].isin(['35', '40'])]
        st.write('- done!')

        # Use the is_holiday function to remove rows with holiday or weekend dates
        st.write('### Filtering `{}`...'.format(col['Gfom']))
        df = df[~df[col['Gfom']].apply(is_holiday)]
        st.write('- done!')
    
    # Celebrate when the process is done!
    st.balloons()

    # Show the filtered data to the user
    st.write("## Filtered Data", df)

    # Let the user download the filtered CSV file
    st.download_button(
        "Download Result CSV",  # Button text
        df.to_csv(sep=SEP, index=False, encoding=ENCODING),  # Convert DataFrame back to CSV format
        file_name="filtered-result.csv",  # Name of the downloaded file
        mime='text/csv'  # File type
    )
