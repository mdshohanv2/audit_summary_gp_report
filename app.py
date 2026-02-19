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
    }
    
    /* Subheader spacing */
    .stSubheader {
        margin-top: 1.5rem;
        font-weight: 600;
    }

    /* Minimalist upload styling with theme-aware borders */
    .stFileUploader {
        border: 1px solid rgba(128, 128, 128, 0.2);
        border-radius: 8px;
        padding: 10px;
    }

    /* Footer/Info styling */
    .stAlert {
        border-radius: 8px;
        border: 1px solid rgba(128, 128, 128, 0.2);
        padding: 0.75rem 1rem;
        min-height: 52px;
        display: flex;
        align-items: center;
    }
    
    /* Match Download button style to Alert box */
    .stDownloadButton, .stDownloadButton > button {
        width: 100%;
        min-height: 52px;
        border-radius: 8px !important;
        border: 1px solid rgba(128, 128, 128, 0.2) !important;
        background-color: transparent !important;
        transition: all 0.2s ease;
    }
    .stDownloadButton > button:hover {
        background-color: rgba(128, 128, 128, 0.1) !important;
        border-color: rgba(128, 128, 128, 0.4) !important;
    }
    
    /* Divider styling */
    hr {
        margin-top: 2rem;
        margin-bottom: 2rem;
        opacity: 0.1;
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
        
        # 1. Total Successful Visits
        target_col = 'mtd successful visits'
        actual_col = next((c for c in df.columns if str(c).strip().lower().replace('\n', ' ') == target_col), None)
        total_sum = pd.to_numeric(df[actual_col], errors='coerce').sum() if actual_col else 0

        # 2. Individual Status Sums (New Requirement)
        status_mappings = {
            "open": "visit status open",
            "temporarily_closed": "visit status temporary closed",
            "permanently_closed": "visit status permanently closed",
            "moved": "visit status moved",
            "pos_not_found": "visit status not found"
        }
        status_sums = {}
        for key, header in status_mappings.items():
            col = next((c for c in df.columns if str(c).strip().lower().replace('\n', ' ') == header), None)
            status_sums[key] = pd.to_numeric(df[col], errors='coerce').sum() if col else 0

        # 3. Date extraction
        actual_date_col = next((c for c in df.columns if str(c).strip().lower().replace('\n', ' ') == 'to date'), None)
        extracted_period = None
        if actual_date_col:
            dates = pd.to_datetime(df[actual_date_col], errors='coerce').dropna()
            if not dates.empty:
                extracted_period = dates.max().strftime('%d-%B-%Y')
            else:
                last_val = str(df[actual_date_col].dropna().iloc[-1]).strip()
                import re
                match = re.search(r'(\d{4})-(\d{2})-(\d{2})', last_val)
                if match:
                    extracted_period = pd.to_datetime(last_val).strftime('%d-%B-%Y')
                else:
                    extracted_period = last_val
        
        return total_sum, status_sums, extracted_period, None
    except Exception as e:
        return 0, {}, None, f"Error processing file: {str(e)}"

# Helper function to create final summary table
def create_final_summary(shinsa_counts, total_shinsa, mtd_total, mtd_status_sums):
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
        # Shinsa counts
        count = 0
        if shinsa_counts is not None and not shinsa_counts.empty:
            match = shinsa_counts[shinsa_counts['Status'] == status_key]
            if not match.empty:
                count = match.iloc[0]['Count']
        
        # MTD counts for this category
        mtd_val = mtd_status_sums.get(status_key, 0)
        
        # Calculations based on standard audit logic
        visit_ach = (mtd_val / mtd_total) if mtd_total > 0 else 0.0
        audit_ach = (count / total_shinsa) if total_shinsa > 0 else 0.0
        row_coverage = (count / mtd_val) if mtd_val > 0 else 0.0
        
        data.append({
            "Visit Status": label,
            "Successful Visits": int(mtd_val),
            "Visit Ach%": f"{visit_ach:.2%}",
            "Audited Visits": int(count),
            "Audit Ach%": f"{audit_ach:.2%}",
            "Coverage %": f"{row_coverage:.2%}"
        })
    
    # Calculation for Grand Total row
    grand_total_coverage = (total_shinsa / mtd_total) if mtd_total > 0 else 0.0
    
    grand_total = {
        "Visit Status": "Grand Total",
        "Successful Visits": int(mtd_total),
        "Visit Ach%": "100.00%", 
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
mtd_status_sums = {}
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
        mtd_total_sum, mtd_status_sums, auto_period, error = get_mtd_summary(mtd_file)
        if error:
            st.error(error)

# Helper function to export to Excel
import io
def export_to_excel(df, title):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Summary', startrow=1)
        workbook = writer.book
        worksheet = writer.sheets['Summary']
        
        # Add title
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        header_fill = PatternFill(start_color='333333', end_color='333333', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True, size=14)
        
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
        cell = worksheet.cell(row=1, column=1)
        cell.value = title
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
        
        # Style header row of table
        table_header_fill = PatternFill(start_color='E0E0E0', end_color='E0E0E0', fill_type='solid')
        for col_num, value in enumerate(df.columns, 1):
            cell = worksheet.cell(row=2, column=col_num)
            cell.fill = table_header_fill
            cell.font = Font(bold=True)
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # Style data rows
        for row_num in range(3, len(df) + 3):
            for col_num in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if row_num == len(df) + 2: # Grand Total row
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')

        # Auto-adjust column widths
        from openpyxl.utils import get_column_letter
        for i, col in enumerate(worksheet.columns, 1):
            max_length = 0
            column_letter = get_column_letter(i)
            for cell in col:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 5)
            worksheet.column_dimensions[column_letter].width = adjusted_width

    return output.getvalue()

# Final Summary Table (Only shows when both are uploaded)
if shinsa_file and mtd_file and not error:
    # Use the extracted date from the 'To Date' column, or a placeholder if missing
    display_date = auto_period if auto_period else "Date Not Found"
    report_title = f"Audit Summary Report [{display_date}]"
    
    st.markdown(f"""
        <div style='background-color: #333333; padding: 10px; border-radius: 5px; margin-bottom: 20px; text-align: center;'>
            <h2 style='color: white; margin: 0;'>{report_title}</h2>
        </div>
    """, unsafe_allow_html=True)
    
    final_df, grand_total_row = create_final_summary(shinsa_status_df, shinsa_total, mtd_total_sum, mtd_status_sums)
    
    # Add grand total row to dataframe for display
    display_df = pd.concat([final_df, pd.DataFrame([grand_total_row])], ignore_index=True)
    
    # Custom styling for the summary table
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )
    
    # Display as clean, professional text instead of a box
    st.markdown(f"### Final Coverage for {display_date}: **{grand_total_row['Coverage %']}**")

    # Download button directly below, aligned to the left
    col_dl, col_empty = st.columns([1, 3])
    with col_dl:
        excel_data = export_to_excel(display_df, report_title)
        st.download_button(
            label="📊 Download Excel Report",
            data=excel_data,
            file_name=f"Audit_Summary_{display_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Overall status help
if not shinsa_file or not mtd_file:
    st.info("Please upload both files to generate the final Audit Summary Report.")
