import streamlit as st
# icons source: https://fonts.google.com/icons

pg = st.navigation([
  st.Page("a0507.py", title="A05_07: List Noll Bearb"),
  st.Page("a0202.py", title="A02_02: Sjuk o frånvaro med omf"),
  st.Page("s210.py", title="S2_10:Sjuk mer än 365 dagar"),
])
st.set_page_config(page_title="CSV Tools", page_icon=":material/filter_alt:")
pg.run()