import streamlit as st
import pandas as pd

# Page configuration
st.set_page_config(
    page_title="Audit Analyzer",
    page_icon="📊",
    layout="centered"
)

# Minimal Shadcn-like CSS
st.markdown("""
    <style>
    /* Main container styling */
    .block-container {
        padding-top: 3rem;
        padding-bottom: 3rem;
    }
    
    /* Typography */
    h1 {
        font-weight: 700;
        letter-spacing: -0.025em;
        color: #0f172a;
    }
    p {
        color: #64748b;
    }

    /* Subheader spacing */
    .stSubheader {
        margin-top: 1.5rem;
        font-weight: 600;
        color: #1e293b;
    }

    /* Minimalist upload styling */
    .stFileUploader {
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 10px;
        background-color: #ffffff;
    }

    /* Footer/Info styling */
    .stAlert {
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }
    </style>
    """, unsafe_allow_html=True)

# Header
st.title("Audit Analyzer")
st.write("Upload your reports to start the analysis.")

st.divider()

# Upload sections
st.subheader("Audit report from Shinsa")
shinsa_file = st.file_uploader(
    "Choose a file", 
    type=["xlsx", "xls", "csv"],
    key="shinsa_uploader",
    label_visibility="collapsed"
)

st.subheader("MTD Successful visit Report")
mtd_file = st.file_uploader(
    "Choose a file", 
    type=["xlsx", "xls", "csv"],
    key="mtd_uploader",
    label_visibility="collapsed"
)

# Placeholder for data processing
if shinsa_file and mtd_file:
    st.success("Files uploaded successfully.")
elif shinsa_file or mtd_file:
    st.info("Please upload both files to proceed.")
