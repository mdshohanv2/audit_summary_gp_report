import streamlit as st
import pandas as pd

# Page configuration
st.set_page_config(
    page_title="Audit Analyzer",
    page_icon="📊",
    layout="wide"
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

# Helper function to process Shinsa report
def get_shinsa_summary(file):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            # For Excel files, we might need openpyxl
            df = pd.read_excel(file)
        
        # Look for the target column (case-insensitive)
        target_col = 'visit_pos_status'
        actual_col = next((c for c in df.columns if str(c).strip().lower() == target_col), None)
        
        if actual_col:
            # Normalize data: lowercase, stripped, handle nulls
            status_series = df[actual_col].astype(str).str.strip().str.lower()
            
            # Define what counts as "empty"
            empty_markers = ['nan', '', 'none', 'null']
            status_series = status_series.apply(lambda x: 'unknown/empty' if x in empty_markers else x)
            
            # Count occurrences
            counts = status_series.value_counts().reset_index()
            counts.columns = ['Status', 'Count']
            
            # Calculate total
            total = counts['Count'].sum()
            
            return counts, total, None
        else:
            return None, 0, f"Column '{target_col}' not found in the uploaded file."
    except Exception as e:
        return None, 0, f"Error processing file: {str(e)}"

# Helper function to process MTD report
def get_mtd_summary(file):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        
        # 1. Look for Successful Visits column
        target_col = 'mtd successful visits'
        actual_col = next((c for c in df.columns if str(c).strip().lower().replace('\n', ' ') == target_col), None)
        
        total_sum = 0
        sum_error = None
        if actual_col:
            numeric_values = pd.to_numeric(df[actual_col], errors='coerce')
            total_sum = numeric_values.sum()
        else:
            sum_error = "Column '[MTD Successful Visits]' not found."

        # 2. Look for 'To Date' column to extract period
        date_col_name = 'to date'
        actual_date_col = next((c for c in df.columns if str(c).strip().lower() == date_col_name), None)
        extracted_period = None
        
        # 2. Look for 'To Date' column to extract period
        # Using .replace('\n', ' ') to handle multi-line headers shown in the user image
        actual_date_col = next((c for c in df.columns if str(c).strip().lower().replace('\n', ' ') == 'to date'), None)
        
        extracted_period = None
        if actual_date_col:
            # Try to convert the whole column to datetime
            dates = pd.to_datetime(df[actual_date_col], errors='coerce').dropna()
            if not dates.empty:
                # Format as 13-02-2026 (DD-MM-YYYY)
                extracted_period = dates.max().strftime('%d-%m-%Y')
            else:
                # Fallback for text-based dates
                last_val = str(df[actual_date_col].dropna().iloc[-1]).strip()
                import re
                # Match YYYY-MM-DD
                match_iso = re.search(r'(\d{4})-(\d{2})-(\d{2})', last_val)
                # Match DD-MM-YYYY
                match_dmy = re.search(r'(\d{2})-(\d{2})-(\d{4})', last_val)
                
                if match_iso:
                    yyyy, mm, dd = match_iso.groups()
                    extracted_period = f"{dd}-{mm}-{yyyy}"
                elif match_dmy:
                    extracted_period = last_val
                else:
                    extracted_period = last_val
        
        return total_sum, extracted_period, sum_error
    except Exception as e:
        return 0, None, f"Error processing file: {str(e)}"

# Helper function to create final summary table
def create_final_summary(shinsa_counts, total_shinsa, mtd_total):
    # Mapping table label to internal normalized status keys
    mapping = [
        ("Visit Status Open", "open"),
        ("Visit Status temporary closed", "temporarily_closed"),
        ("Visit Status permanently closed", "permanently_closed"),
        ("Visit status Moved", "moved"),
        ("Visit status Not found", "pos_not_found")
    ]
    
    data = []
    for label, status_key in mapping:
        # Get count from shinsa_counts if available
        count = 0
        if shinsa_counts is not None and not shinsa_counts.empty:
            match = shinsa_counts[shinsa_counts['Status'] == status_key]
            if not match.empty:
                count = match.iloc[0]['Count']
        
        # Calculations (following image logic: Succ visits for rows = 0)
        succ_visits = 0
        visit_ach = 0.0
        audit_ach = (count / total_shinsa) if total_shinsa > 0 else 0.0
        coverage = (count / mtd_total) if mtd_total > 0 else 0.0
        
        data.append({
            "Visit Status": label,
            "Successful Visits": succ_visits,
            "Visit Ach%": f"{visit_ach:.2%}",
            "Audited Visits": int(count),
            "Audit Ach%": f"{audit_ach:.2%}",
            "Coverage %": f"{coverage:.2%}"
        })
    
    # Calculation for Grand Total row
    grand_total_coverage = (total_shinsa / mtd_total) if mtd_total > 0 else 0.0 # Wait, image shows Audited / Succ
    # Actually, in image: 14,220 / 2,74,950 = 5.17%
    grand_total_coverage = (total_shinsa / mtd_total) if mtd_total > 0 else 0.0
    
    grand_total = {
        "Visit Status": "Grand Total",
        "Successful Visits": int(mtd_total),
        "Visit Ach%": "0%", # As per image
        "Audited Visits": int(total_shinsa),
        "Audit Ach%": "100.00%",
        "Coverage %": f"{grand_total_coverage:.2%}"
    }
    
    return pd.DataFrame(data), grand_total

# Header
st.title("Audit Analyzer")
st.write("Upload your reports to start the analysis.")

st.divider()

# Upload sections
col1, col2 = st.columns(2)

with col1:
    st.subheader("Audit report from Shinsa")
    shinsa_file = st.file_uploader(
        "Choose Shinsa file", 
        type=["xlsx", "xls", "csv"],
        key="shinsa_uploader",
        label_visibility="collapsed"
    )

with col2:
    st.subheader("MTD Successful visit Report")
    mtd_file = st.file_uploader(
        "Choose MTD file", 
        type=["xlsx", "xls", "csv"],
        key="mtd_uploader",
        label_visibility="collapsed"
    )

st.divider()

# Variables to store analysis results
shinsa_status_df = None
shinsa_total = 0
mtd_total_sum = 0
auto_period = None

# Processing Shinsa
if shinsa_file:
    with st.spinner("Analyzing Shinsa..."):
        shinsa_status_df, shinsa_total, error = get_shinsa_summary(shinsa_file)
        if error:
            st.error(error)

# Processing MTD
# PROCESSING logic for MTD report
if mtd_file:
    with st.spinner("Analyzing MTD..."):
        mtd_total_sum, auto_period, error = get_mtd_summary(mtd_file)
        if error:
            st.error(error)

# Final Summary Table (Only shows when both are uploaded)
if shinsa_file and mtd_file and not error:
    # Use the extracted date from the 'To Date' column, or a placeholder if missing
    display_date = auto_period if auto_period else "Date Not Found"
    
    st.markdown(f"""
        <div style='background-color: #333333; padding: 10px; border-radius: 5px; margin-bottom: 20px; text-align: center;'>
            <h2 style='color: white; margin: 0;'>Audit Summary Report [{display_date}]</h2>
        </div>
    """, unsafe_allow_html=True)
    
    final_df, grand_total_row = create_final_summary(shinsa_status_df, shinsa_total, mtd_total_sum)
    
    # Add grand total row to dataframe for display
    display_df = pd.concat([final_df, pd.DataFrame([grand_total_row])], ignore_index=True)
    
    # Custom styling for the summary table
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )
    
    # Highlighting the final coverage status
    st.info(f"Final Coverage for {display_date}: **{grand_total_row['Coverage %']}**")

# Overall status help
if not shinsa_file or not mtd_file:
    st.info("Please upload both files to generate the final Audit Summary Report.")
