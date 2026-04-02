import streamlit as st

from invoice_app import render_database_page, render_page_header


st.set_page_config(page_title="Receipts Database", page_icon=":page_facing_up:", layout="wide")
token, _, _ = render_page_header("Receipts Database")
render_database_page(token)
