import streamlit as st
import pandas as pd
import io
import openpyxl
from io import StringIO

st.title("MDM Monthly Report Generator")
st.write("Paste your CSV data to load it, or type directly into the editable grid below.")

# Define the dropdown options with English prefixes for easy typing
FOOD_OPTIONS = [
    "a - আলু",
    "b - বেগুন বরবটি",
    "d - ডিম",
    "k - কুমড়ো",
    "kp - কচু পুঁই",
    "p - পটল",
    "s - সয়াবিন",
    "none - (শুধু ভাত ও ডাল)"
]

# Helper function to convert raw CSV strings to dropdown format
def to_dropdown_format(item):
    item = str(item).strip()
    if "কুমড়ো" in item: return "k - কুমড়ো"
    if "কচু পুঁই" in item: return "kp - কচু পুঁই"
    if "পটল" in item: return "p - পটল"
    if "বেগুন বরবটি" in item: return "b - বেগুন বরবটি"
    if "ডিম" in item: return "d - ডিম"
    if "আলু" in item: return "a - আলু"
    if "সয়াবিন" in item: return "s - সয়াবিন"
    if item == "ভাত, ডাল": return "none - (শুধু ভাত ও ডাল)"
    return item # Fallback for unmapped items

# Helper function to convert dropdown format back to official MDM format
def to_excel_format(item):
    item = str(item).strip()
    if item == "a - আলু": return "ভাত, ডাল, আলু"
    if item == "b - বেগুন বরবটি": return "ভাত, ডাল, বেগুন বরবটি"
    if item == "d - ডিম": return "ভাত, ডাল, ডিম"
    if item == "k - কুমড়ো": return "ভাত, ডাল, কুমড়ো"
    if item == "kp - কচু পুঁই": return "ভাত, ডাল, কচু পুঁই"
    if item == "p - পটল": return "ভাত, ডাল, পটল"
    if item == "s - সয়াবিন": return "ভাত, ডাল, সয়াবিন"
    if item == "none - (শুধু ভাত ও ডাল)": return "ভাত, ডাল"
    return item # Fallback

# Initialize empty dataframe in session state if it doesn't exist
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(columns=[
        "Date", 
        "No of student availing MDM Class - V", 
        "No of student availing MDM Class - VI to VIII", 
        "Items served", 
        "Cooking cost for Class - V to VIII", 
        "Food Grains for Class - V", 
        "Food grains for Class - VI to VIII"
    ])

# 1. Optional CSV Data Loader
with st.expander("Step 1: Paste CSV Data (Optional)", expanded=True):
    data_input = st.text_area(
        "Paste your CSV Data here and click Load", 
        height=150, 
        placeholder="Date,No of student availing MDM Class - V,No of student availing MDM Class - VI to VIII,Items served..."
    )
    
    if st.button("Load Pasted Data"):
        if data_input:
            try:
                # Parse CSV
                new_df = pd.read_csv(StringIO(data_input.strip()))
                # Convert the "Items served" column to the dropdown UI format
                if 'Items served' in new_df.columns:
                    new_df['Items served'] = new_df['Items served'].apply(to_dropdown_format)
                
                st.session_state.df = new_df
                st.rerun() # Refresh the app to show loaded data in the editor
            except Exception as e:
                st.error(f"Error parsing data: {e}")

# 2. Interactive Excel-Like Grid
st.write("### Step 2: Edit Data")
st.caption("You can edit cells directly. Click the 'Items served' column and type 'k' for কুমড়ো, 'd' for ডিম, etc.")

edited_df = st.data_editor(
    st.session_state.df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Items served": st.column_config.SelectboxColumn(
            "Items served (Side Dish)",
            help="Select the side dish. 'ভাত, ডাল' is added automatically in the final Excel.",
            options=FOOD_OPTIONS,
            required=True
        )
    }
)

# 3. Generate Final Report
if st.button("Generate Final Excel Report", type="primary"):
    try:
        # Load the hardcoded Excel workbook from the repository
        template_path = "MDM MONTHLY REPORT_2.xlsx"
        wb = openpyxl.load_workbook(template_path)
        sheet = wb["MDM Report"] # Make sure this matches the sheet name
        
        # The daily data starts at Excel row 13
        start_row = 13
        
        # Fill data into the Excel sheet
        for index, row in edited_df.iterrows():
            current_row = start_row + index
            
            # Handle empty rows if added via the grid
            if pd.isna(row["Date"]) and pd.isna(row["No of student availing MDM Class - V"]):
                continue
                
            # Column A (1): Serial Number (exclude for the final summary row)
            if "Days" not in str(row["Date"]):
                sheet.cell(row=current_row, column=1).value = index + 1
            
            # Populate Columns B (2) through H (8)
            sheet.cell(row=current_row, column=2).value = row["Date"]
            sheet.cell(row=current_row, column=3).value = row["No of student availing MDM Class - V"]
            sheet.cell(row=current_row, column=4).value = row["No of student availing MDM Class - VI to VIII"]
            
            # Convert UI dropdown value back to full Bengali string with Vat and Dal
            final_food_string = to_excel_format(row["Items served"])
            sheet.cell(row=current_row, column=5).value = final_food_string
            
            sheet.cell(row=current_row, column=6).value = row["Cooking cost for Class - V to VIII"]
            sheet.cell(row=current_row, column=7).value = row["Food Grains for Class - V"]
            sheet.cell(row=current_row, column=8).value = row["Food grains for Class - VI to VIII"]

        # Save to memory for download
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.success("Report generated successfully!")
        
        # Provide Download Button
        st.download_button(
            label="Download Populated MDM Report",
            data=output,
            file_name="Completed_MDM_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except FileNotFoundError:
        st.error("Template file 'MDM MONTHLY REPORT_2.xlsx' not found. Ensure it is uploaded to the same folder in GitHub.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
