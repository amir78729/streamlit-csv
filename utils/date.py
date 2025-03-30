import streamlit as st
from datetime import datetime

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
        st.write('- `{}` is a holiday 🎄'.format(date_str))
        return True

    # Convert the date from string to a datetime object
    date = datetime.strptime(str(date_str), "%Y%m%d")

    # Check if it's Saturday (5) or Sunday (6)
    is_weekend = date.weekday() in [5, 6]

    if is_weekend:
        st.write('- `{}` is a weekend 😴'.format(date_str))
        return True

    # If it's not a holiday or weekend, keep the date
    st.write('- `{}` is a workday 💼'.format(date_str))
    return False