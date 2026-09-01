import streamlit as st
import openpyxl
from io import BytesIO
import re

# Set Streamlit page configuration
st.set_page_config(page_title="MDM Register Automator", layout="centered")

st.title("🏫 MDM Excel Auto-Filler")
st.write("Upload the blank **MDM MONTHLY REPORT.xlsx** template and paste your extracted text to automatically populate the register.")

# 1. File Upload
uploaded_file = st.file_uploader("Upload MDM MONTHLY REPORT.xlsx", type=["xlsx"])

# 2. Text Input Area
raw_text = st.text_area("Paste the Extracted Data Here", height=250, 
                        help="You can paste the raw mangled text or a clean Markdown table.")

if st.button("Generate Filled Excel") and uploaded_file and raw_text:
    try:
        # Load the workbook and select the active sheet
        wb = openpyxl.load_workbook(uploaded_file)
        sheet = wb.active
        
        data_rows = []
        total_row = None
        
        # Logic 1: Parse the continuous, unformatted string (Fallback)
        if "Days" in raw_text and "|" not in raw_text and "\t" not in raw_text:
            # Regex extracts: Date (dd.mm.yy) | Class V (2 digits) | Class VI-VIII (3 digits) | Items (Bengali) | Cost (4 digits)
            pattern = r'(\d{2}\.\d{2}\.\d{2})(\d{1,2})(\d{2,3})([^\d]+)(\d{3,4})'
            matches = re.findall(pattern, raw_text)
            for m in matches:
                data_rows.append([m[0], m[1], m[2], m[3].strip(), m[4]])
            
            # Extract the total row from the very end of the string
            total_match = re.search(r'(\d{2})\s*Days(\d{3})(\d{4})(\d{5})', raw_text)
            if total_match:
                total_row = [f"{total_match.group(1)} Days", total_match.group(2), total_match.group(3), "", total_match.group(4)]
                
        # Logic 2: Parse standard Markdown tables or Tab-Separated Data
        else:
            lines = raw_text.strip().split('\n')
            for line in lines:
                if '|' in line:
                    parts = [p.strip() for p in line.split('|') if p.strip()]
                    # Skip headers and dividers
                    if len(parts) >= 5 and '---' not in parts[0] and 'Date' not in parts[0] and 'তারিখ' not in parts[0]:
                        if "Days" in parts[0]:
                            total_row = parts
                        else:
                            data_rows.append(parts)
                else:
                    parts = line.split('\t')
                    if len(parts) >= 5 and 'Date' not in parts[0]:
                        if "Days" in parts[0]:
                            total_row = parts
                        else:
                            data_rows.append(parts)
        
        # 3. Write Daily Data to Excel (Starting at Row 13 for Sl. No. 1)
        current_row = 13
        for row in data_rows:
            sheet.cell(row=current_row, column=2).value = row[0]  # Date
            sheet.cell(row=current_row, column=3).value = int(row[1]) if row[1].isdigit() else row[1]  # Class V
            sheet.cell(row=current_row, column=4).value = int(row[2]) if row[2].isdigit() else row[2]  # Class VI-VIII
            sheet.cell(row=current_row, column=5).value = row[3]  # Items Served
            
            # Clean and insert the cost
            cost = re.sub(r'[^\d]', '', row[4])
            if cost:
                sheet.cell(row=current_row, column=6).value = int(cost)
                
            current_row += 1
            
        # 4. Write Total Data (Row 39 based on the Excel template structure)
        if total_row:
            total_row_idx = 39 
            v_total = re.sub(r'\D', '', total_row[1]) if len(total_row) > 1 else ""
            vi_viii_total = re.sub(r'\D', '', total_row[2]) if len(total_row) > 2 else ""
            
            # Locate cost in Markdown tables (usually index 4) vs standard arrays
            cost_idx = 4 if len(total_row) > 4 and total_row[4] else (3 if len(total_row) > 3 else -1)
            total_cost = re.sub(r'\D', '', total_row[cost_idx]) if cost_idx != -1 else ""
            
            if v_total: sheet.cell(row=total_row_idx, column=3).value = int(v_total)
            if vi_viii_total: sheet.cell(row=total_row_idx, column=4).value = int(vi_viii_total)
            if total_cost: sheet.cell(row=total_row_idx, column=6).value = int(total_cost)
            
        # 5. Save and Download
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.success("✅ Excel file populated successfully!")
        st.download_button(
            label="📥 Download Filled MDM Report",
            data=output,
            file_name="Filled_MDM_MONTHLY_REPORT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"An error occurred during processing: {e}")
