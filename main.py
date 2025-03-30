import streamlit as st
# icons source: https://fonts.google.com/icons

pg = st.navigation([
  st.Page("a0507.py", title="A05 07 List Noll Bearb"),
  st.Page("a0202.py", title="A02 02 Sjuk o frånvaro med omf")
])
st.set_page_config(page_title="CSV Tools", page_icon=":material/filter_alt:")
pg.run()