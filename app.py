import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from io import StringIO
import avro

st.set_page_config(page_title="MDM Monthly Report", layout="wide")
st.title("MDM Monthly Report Generator")
st.write("Type English phonetic text in the grid (e.g. 'dim', 'kumro') to auto-convert to Bengali, configure rates, and generate your A4 Excel report.")

# Helper function to transliterate English text to Bengali phonetically using Avro
def phonetic_convert(item):
    if pd.isna(item): return ""
    item_str = str(item).strip()
    if not item_str: return ""
    
    # If it already contains Bengali characters, return as is
    if any('\u0980' <= c <= '\u09ff' for c in item_str):
        return item_str
        
    try:
        converted = avro.parse(item_str)
        return converted
    except Exception:
        return item_str

# Initialize empty dataframe with exact column names matching headers
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(columns=[
        "Date",
        "No of student availing MDM Class - V",
        "No of student availing MDM Class - VI to VIII",
        "Items served",
        "Cooking cost for Class - V to VIII",
        "Food Grains for Class - V (Kg.)",
        "Food grains for Class - VI to VIII (Kg.)"
    ])

# 1. Optional CSV Data Loader
with st.expander("Step 1: Paste CSV Data (Optional)", expanded=False):
    data_input = st.text_area(
        "Paste your CSV Data here and click Load",
        height=150,
        placeholder="Date,No of student availing MDM Class - V,No of student availing MDM Class - VI to VIII,Items served..."
    )

    if st.button("Load Pasted Data"):
        if data_input:
            try:
                new_df = pd.read_csv(StringIO(data_input.strip()))
                # If loading legacy CSV columns without (Kg.), safely map them
                if "Food Grains for Class - V" in new_df.columns and "Food Grains for Class - V (Kg.)" not in new_df.columns:
                    new_df.rename(columns={
                        "Food Grains for Class - V": "Food Grains for Class - V (Kg.)",
                        "Food grains for Class - VI to VIII": "Food grains for Class - VI to VIII (Kg.)"
                    }, inplace=True)
                
                if 'Items served' in new_df.columns:
                    new_df['Items served'] = new_df['Items served'].apply(phonetic_convert)
                st.session_state.df = new_df
                st.rerun()
            except Exception as e:
                st.error(f"Error parsing data: {e}")

# 2. Interactive Excel-Like Grid & Rate Configuration
st.write("### Step 2: Configure Rates, Edit Daily Data & Auto-Calculate")

# Rate Configuration Boxes
st.markdown("##### ⚙️ Calculation Rates Configuration")
r_col1, r_col2, r_col3, r_col4 = st.columns(4)
with r_col1:
    cost_rate_v = st.number_input("Class V Cost (Rs/student)", value=6.78, step=0.01, format="%.2f")
with r_col2:
    cost_rate_vi = st.number_input("Class VI-VIII Cost (Rs/student)", value=10.17, step=0.01, format="%.2f")
with r_col3:
    grain_rate_v = st.number_input("Class V Grains (Kg/student)", value=0.100, step=0.005, format="%.3f")
with r_col4:
    grain_rate_vi = st.number_input("Class VI-VIII Grains (Kg/student)", value=0.150, step=0.005, format="%.3f")

if st.button("🧮 Auto-Calculate Costs & Grains", help="Calculates Cooking Cost and Food Grains using the rates above"):
    st.session_state.df["Cooking cost for Class - V to VIII"] = st.session_state.df["Cooking cost for Class - V to VIII"].astype(object)
    st.session_state.df["Food Grains for Class - V (Kg.)"] = st.session_state.df["Food Grains for Class - V (Kg.)"].astype(object)
    st.session_state.df["Food grains for Class - VI to VIII (Kg.)"] = st.session_state.df["Food grains for Class - VI to VIII (Kg.)"].astype(object)

    for idx, row in st.session_state.df.iterrows():
        if pd.notna(row["No of student availing MDM Class - V"]):
            try:
                v_stu = float(row["No of student availing MDM Class - V"])
                vi_stu = float(row["No of student availing MDM Class - VI to VIII"])
                
                if "Days" not in str(row["Date"]) and "Total" not in str(row["Date"]):
                    st.session_state.df.at[idx, "Cooking cost for Class - V to VIII"] = round((v_stu * cost_rate_v) + (vi_stu * cost_rate_vi), 2)
                    st.session_state.df.at[idx, "Food Grains for Class - V (Kg.)"] = round(v_stu * grain_rate_v, 3)
                    st.session_state.df.at[idx, "Food grains for Class - VI to VIII (Kg.)"] = round(vi_stu * grain_rate_vi, 3)
            except ValueError:
                pass
    st.rerun()

st.caption("Type any English word in 'Items served' (e.g. 'bhat, dal, dim') and it will automatically phonetic-convert to Bengali.")

edited_df = st.data_editor(
    st.session_state.df,
    num_rows="dynamic",
    use_container_width=True
)

# 3. Financial Data Entry Boxes
st.write("### Step 3: Financial Details (Opening Balances & Received Funds)")
st.caption("Enter the amounts below. The app will auto-calculate Total Available Funds and Closing Balances in the final Excel file.")
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        ob_cooking = st.number_input("1. Opening Balance Cooking Cost (Rs.)", value=0.0)
        fund_received = st.number_input("3. Fund for Cooking Cost Received (Rs.)", value=0.0)
        st.markdown("---")
        ob_grains = st.number_input("2. Opening Balance Food Grains (Kg.)", value=0.0)
        grains_received = st.number_input("5. Food Grains Received (Kg.)", value=0.0)
    with col2:
        hon_received = st.number_input("4. Honorarium Received (Rs.)", value=0.0)
        hon_expenditure = st.number_input("10. Total Expenditure Honorarium (Rs.)", value=0.0)


# ---------------------------------------------------------------------------
# Excel Template Builder (Strict A4 Scale & PDF Accuracy)
# ---------------------------------------------------------------------------
def create_excel_template(c_v, c_vi):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MDM Report"

    # ---- Strict A4 Print Scaling ----
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_options.horizontalCentered = True
    
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.5, header=0.3, footer=0.3)

    f_title = Font(name="Arial", size=9, bold=True)        
    f_enroll = Font(name="Arial", size=8, bold=True)        
    f_hdr9 = Font(name="Arial", size=9, bold=True)          
    f_hdr8 = Font(name="Arial", size=8, bold=True)          
    f_data_bold = Font(name="Calibri", size=11, bold=True)  

    a_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    a_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    a_center = Alignment(horizontal="center", vertical="center")
    a_right = Alignment(horizontal="right", vertical="center")

    thin = Side(style="thin")
    full_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows(min_row=1, max_row=48, min_col=1, max_col=8):
        for cell in row:
            cell.border = full_border

    for col in range(1, 9):
        ws.cell(row=49, column=col).border = Border(top=thin, left=thin if col==1 else None, right=thin if col==8 else None)
        ws.cell(row=50, column=col).border = Border(bottom=thin, left=thin if col==1 else None, right=thin if col==8 else None)

    ws.column_dimensions['A'].width = 5.0
    ws.column_dimensions['B'].width = 11.0
    ws.column_dimensions['C'].width = 11.5
    ws.column_dimensions['D'].width = 11.5
    ws.column_dimensions['E'].width = 25.0
    ws.column_dimensions['F'].width = 11.5
    ws.column_dimensions['G'].width = 11.5
    ws.column_dimensions['H'].width = 12.0

    row_heights = {
        1: 18.0, 2: 18.0, 3: 16.0, 4: 16.0,
        5: 14.5, 6: 14.5, 7: 14.5, 8: 14.5, 9: 14.5, 10: 14.5, 11: 14.5,
        12: 52.0, 
        **{r: 14.5 for r in range(13, 39)},
        39: 16.0, 40: 20.0, 41: 26.0, 42: 16.0,
        43: 14.5, 44: 14.5, 45: 14.5, 46: 14.5, 47: 14.5, 48: 14.5,
        49: 65.0, 50: 30.0, 
    }
    for r, h in row_heights.items():
        ws.row_dimensions[r].height = h

    merges = [
        "A1:H1", "A2:E2", "F2:G2", "A3:E3", "F3:G3",
        "B4:E4", "B5:E5", "B6:E6", "B7:E7", "B8:E8", "B9:E9", "B10:E10", "B11:E11",
        "A39:B39", "A40:H40", "A41:H41",
        "B42:E42", "B43:E43", "B44:E44", "B45:E45", "B46:E46", "B47:E47", "B48:E48",
        "A49:H49", "A50:D50", "E50:H50",
    ]
    for m in merges:
        ws.merge_cells(m)

    c = ws.cell(row=1, column=1, value="MONTHLY REPORT FOR COOKING COST & FOOD GRAINS UNDER MID - DAY - MEAL FOR THE MONTH OF -")
    c.font = f_title; c.alignment = a_left_wrap

    c = ws.cell(row=2, column=1, value="Name of the School : SHYAMBAZAR D N HIGH SCHOOL")
    c.font = f_title; c.alignment = a_left_wrap
    
    c = ws.cell(row=2, column=6, value="Total Enrolment : Class V = 60")
    c.font = f_enroll; c.alignment = a_center_wrap

    c = ws.cell(row=3, column=1, value="Gram Sansad No. - Bhallagram , Circle : Monglkote-II, Block : Mongalkote")
    c.font = f_title; c.alignment = a_left_wrap
    
    c = ws.cell(row=3, column=6, value="Class VI to VIII = 322")
    c.font = f_hdr9; c.alignment = a_center_wrap

    header_row4 = {1: "Sl.", 2: "Particular", 6: "Class -V", 7: "VI to VIII", 8: "Total"}
    for col, val in header_row4.items():
        c = ws.cell(row=4, column=col, value=val)
        c.font = f_hdr9; c.alignment = a_center_wrap

    particulars_top = [
        "Opening Balance Cooking Cost at the beginning of the month",
        "Opening Balance Food grains at the beginning of the month",
        "Fund for cooking cost received this month (for",
        "Honorarium received this month (for",
        "Food Grains received this month (for",
        "Total available Fund (1+3) (other grant)",
        "Total Food Grains available (2+5) AMOUNT",
    ]
    for idx, item in enumerate(particulars_top):
        r = 5 + idx
        c = ws.cell(row=r, column=1, value=idx + 1)
        c.font = f_data_bold; c.alignment = a_center
        c = ws.cell(row=r, column=2, value=item)
        c.font = f_hdr9; c.alignment = a_left_wrap

    headers = [
        "Sl. No.", "Date", "No of student\navailing MDM\nClass -V",
        "No of student\navailing MDM\nClass VI to VIII", "Items served",
        "Cooking cost\nfor Class - V to VIII", "Food Grains for\nClass - V (Kg.)",
        "Food grains for\nClass VI to VIII (Kg.)"
    ]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=12, column=col_num, value=header)
        cell.font = f_hdr8; cell.alignment = a_center_wrap

    for i in range(26):
        r = 13 + i
        c = ws.cell(row=r, column=1, value=i + 1)
        c.font = f_hdr9; c.alignment = a_center_wrap
        for col in range(2, 9):
            ws.cell(row=r, column=col).font = f_data_bold
            ws.cell(row=r, column=col).alignment = a_center_wrap

    c = ws.cell(row=39, column=1, value="Total")
    c.font = f_hdr9; c.alignment = a_center_wrap
    for col in range(3, 9):
        ws.cell(row=39, column=col).font = f_data_bold
        ws.cell(row=39, column=col).alignment = a_center_wrap

    formula_note = f"*(Total student of Class - V X {c_v}) + Total student of Class - VI to VIII X {c_vi})"
    c = ws.cell(row=40, column=1, value=formula_note)
    c.font = f_hdr8; c.alignment = a_left_wrap

    c = ws.cell(row=41, column=1, value="Name of SHG (Cook-cum-Helper) MC - APPROVED_Name of Bank with Branch CANARA BANK, Mathrun Branch._ A/c no_3641101002012_worked during month.......")
    c.font = f_hdr8; c.alignment = a_left_wrap

    header_row42 = {1: "Sl.", 2: "Particular", 6: "Class - V", 7: "VI to VIII", 8: "Total"}
    for col, val in header_row42.items():
        c = ws.cell(row=42, column=col, value=val)
        c.font = f_hdr9; c.alignment = a_center_wrap

    particulars_bottom = [
        "Total expenditure (Cooking Cost)",
        "Total Utilised (Food grains) in Kg.",
        "Total expenditure (Honorarium)",
        "Closing Balance Fund (7-8)",
        "Closing Balance Food grains (7-9)",
        "Closing Balance Honorarium (4-10)",
    ]
    for idx, item in enumerate(particulars_bottom):
        r = 43 + idx
        c = ws.cell(row=r, column=1, value=idx + 8)
        c.font = f_hdr9; c.alignment = a_center_wrap
        c = ws.cell(row=r, column=2, value=item)
        c.font = f_hdr9; c.alignment = a_left_wrap

    c = ws.cell(row=50, column=1, value="Signature of MDM Nodal Teacher")
    c.font = f_hdr9; c.alignment = a_center
    c = ws.cell(row=50, column=5, value="Signature of H M with seal")
    c.font = f_hdr9; c.alignment = a_right

    return wb, ws


# 4. Generate Final Report
st.markdown("---")
if st.button("Download Final Excel Report", type="primary"):
    try:
        wb, sheet = create_excel_template(cost_rate_v, cost_rate_vi)
        start_row = 13

        sheet.cell(row=5, column=8).value = ob_cooking
        sheet.cell(row=6, column=8).value = ob_grains
        sheet.cell(row=7, column=8).value = fund_received
        sheet.cell(row=8, column=8).value = hon_received
        sheet.cell(row=9, column=8).value = grains_received
        
        total_fund = ob_cooking + fund_received
        total_grains = ob_grains + grains_received
        sheet.cell(row=10, column=8).value = total_fund
        sheet.cell(row=11, column=8).value = total_grains

        total_cooking_exp = 0.0
        total_grains_exp = 0.0

        for index, row in edited_df.iterrows():
            if pd.isna(row["Date"]): 
                continue

            is_summary = ("Days" in str(row["Date"]) or "Total" in str(row["Date"]))

            if is_summary:
                sheet.cell(row=39, column=1).value = str(row["Date"])
                sheet.cell(row=39, column=3).value = row.get("No of student availing MDM Class - V", "")
                sheet.cell(row=39, column=4).value = row.get("No of student availing MDM Class - VI to VIII", "")
                sheet.cell(row=39, column=6).value = row.get("Cooking cost for Class - V to VIII", "")
                sheet.cell(row=39, column=7).value = row.get("Food Grains for Class - V (Kg.)", "")
                sheet.cell(row=39, column=8).value = row.get("Food grains for Class - VI to VIII (Kg.)", "")
                
                try:
                    total_cooking_exp = float(row.get("Cooking cost for Class - V to VIII", 0))
                    g_v = float(row.get("Food Grains for Class - V (Kg.)", 0))
                    g_vi = float(row.get("Food grains for Class - VI to VIII (Kg.)", 0))
                    total_grains_exp = g_v + g_vi
                except ValueError:
                    pass

            else:
                current_row = start_row + index
                if current_row >= 39: 
                    break 

                sheet.cell(row=current_row, column=1).value = index + 1
                sheet.cell(row=current_row, column=2).value = row["Date"]
                sheet.cell(row=current_row, column=3).value = row["No of student availing MDM Class - V"]
                sheet.cell(row=current_row, column=4).value = row["No of student availing MDM Class - VI to VIII"]

                raw_item = row["Items served"]
                final_food_string = phonetic_convert(raw_item)
                if final_food_string:
                    sheet.cell(row=current_row, column=5).value = final_food_string

                sheet.cell(row=current_row, column=6).value = row["Cooking cost for Class - V to VIII"]
                sheet.cell(row=current_row, column=7).value = row["Food Grains for Class - V (Kg.)"]
                sheet.cell(row=current_row, column=8).value = row["Food grains for Class - VI to VIII (Kg.)"]

        sheet.cell(row=43, column=8).value = total_cooking_exp
        sheet.cell(row=44, column=8).value = total_grains_exp
        sheet.cell(row=45, column=8).value = hon_expenditure
        
        sheet.cell(row=46, column=8).value = round(total_fund - total_cooking_exp, 2)
        sheet.cell(row=47, column=8).value = round(total_grains - total_grains_exp, 3)
        sheet.cell(row=48, column=8).value = round(hon_received - hon_expenditure, 2)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("A4 Report generated successfully!")

        st.download_button(
            label="⬇️ Download Populated MDM Report",
            data=output,
            file_name="Completed_MDM_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
