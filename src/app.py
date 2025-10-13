import streamlit as st
from datetime import datetime

# === PAGE CONFIG ===
st.set_page_config(
    page_title="Auto Curing Worklist Helper",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === HEADER ===
st.title("🔁 Auto Curing Worklist Helper")
st.caption(f"Last updated: {datetime.now().strftime('%B %d, %Y')}")

st.markdown("---")

# === INTRODUCTION ===
st.markdown(
    """
    ### 👋 Welcome to the **Auto Curing Worklist Helper**
    This automation tool was built to **simplify and speed up** your daily **TAD updates**, 
    file validation, and **Masterlist synchronization**.

    It helps ensure that your **Active List** stays updated, clean, and aligned — 
    removing repetitive manual Excel work.

    ---
    """
)
# === SIDEBAR ===
