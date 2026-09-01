import streamlit as st
import pandas as pd
import io
import openpyxl
from io import StringIO

st.title("MDM Monthly Report Generator")
st.write("Upload your blank MDM Monthly Report Excel template and paste the daily CSV data to generate the final report.")

# 1. Upload Template
template_file = st.file_uploader("Upload Blank Excel Template (.xlsx)", type=["xlsx"])

# 2. Input Data
data_input = st.text_area(
    "Paste your CSV Data (Include headers)", 
    height=300, 
    placeholder="Date,No of student availing MDM Class - V,No of student availing MDM Class - VI to VIII,Items served..."
)

if st.button("Generate Final Report"):
    if template_file is not None and data_input:
        try:
            # Parse the CSV data
            # Use StringIO to read the raw text as a dataframe
            df_data = pd.read_csv(StringIO(data_input.strip()))
            
            # Load the Excel workbook using openpyxl to preserve formatting
            wb = openpyxl.load_workbook(template_file)
            sheet = wb["MDM Report"] # Make sure this matches the sheet name in your file
            
            # The daily data starts at Excel row 13
            start_row = 13
            
            for index, row in df_data.iterrows():
                current_row = start_row + index
                
                # Column A (1): Serial Number (exclude for the final summary row)
                if "Days" not in str(row.iloc[0]):
                    sheet.cell(row=current_row, column=1).value = index + 1
                
                # Column B (2) through H (8) mapping to the CSV columns
                sheet.cell(row=current_row, column=2).value = row.iloc[0] # Date
                sheet.cell(row=current_row, column=3).value = row.iloc[1] # Class V
                sheet.cell(row=current_row, column=4).value = row.iloc[2] # Class VI to VIII
                sheet.cell(row=current_row, column=5).value = row.iloc[3] # Items served
                sheet.cell(row=current_row, column=6).value = row.iloc[4] # Cooking cost
                sheet.cell(row=current_row, column=7).value = row.iloc[5] # Food Grains Class V
                sheet.cell(row=current_row, column=8).value = row.iloc[6] # Food Grains Class VI-VIII

            # Save the modified workbook to a BytesIO object
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success("Report generated successfully!")
            
            # Download Button
            st.download_button(
                label="Download Populated MDM Report",
                data=output,
                file_name="Completed_MDM_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("Please upload the template and paste the data before generating.")
