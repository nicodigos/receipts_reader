import streamlit as st

from invoice_app import render_page_header, render_process_page


st.set_page_config(page_title="Invoice Process", page_icon=":page_facing_up:", layout="wide")
token, _, _ = render_page_header("Upload New Invoices")
render_process_page(token)
