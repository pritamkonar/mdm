import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from io import StringIO

st.title("MDM Monthly Report Generator")
st.write("Paste your CSV data to load it, or type directly into the editable grid below. The app will generate the full Excel file from scratch.")

# Define the dropdown options with English prefixes for easy typing
def to_dropdown_format(item):
    if pd.isna(item): return None
    item = str(item).strip()
    if "কুমড়ো" in item: return "k - কুমড়ো"
    if "কচু পুঁই" in item: return "kp - কচু পুঁই"
    if "পটল" in item: return "p - পটল"
    if "বেগুন বরবটি" in item: return "b - বেগুন বরবটি"
    if "ডিম" in item: return "d - ডিম"
    if "আলু" in item: return "a - আলু"
    if "সয়াবিন" in item: return "s - সয়াবিন"
    if item == "ভাত, ডাল": return "none - (শুধু ভাত ও ডাল)"
    
    # SAFEGUARD: Return None instead of the raw text so the dropdown doesn't crash
    return None

# Helper function to convert raw CSV strings to dropdown format
def to_dropdown_format(item):
    if pd.isna(item): return None
    item = str(item).strip()
    if "কুমড়ো" in item: return "k - কুমড়ো"
    if "কচু পুঁই" in item: return "kp - কচু পুঁই"
    if "পটল" in item: return "p - পটল"
    if "বেগুন বরবটি" in item: return "b - বেগুন বরবটি"
    if "ডিম" in item: return "d - ডিম"
    if "আলু" in item: return "a - আলু"
    if "সয়াবিন" in item: return "s - সয়াবিন"
    if item == "ভাত, ডাল": return "none - (শুধু ভাত ও ডাল)"
    return item

# Helper function to convert dropdown format back to official MDM format
def to_excel_format(item):
    if pd.isna(item): return ""
    item = str(item).strip()
    if item == "a - আলু": return "ভাত, ডাল, আলু"
    if item == "b - বেগুন বরবটি": return "ভাত, ডাল, বেগুন বরবটি"
    if item == "d - ডিম": return "ভাত, ডাল, ডিম"
    if item == "k - কুমড়ো": return "ভাত, ডাল, কুমড়ো"
    if item == "kp - কচু পুঁই": return "ভাত, ডাল, কচু পুঁই"
    if item == "p - পটল": return "ভাত, ডাল, পটল"
    if item == "s - সয়াবিন": return "ভাত, ডাল, সয়াবিন"
    if item == "none - (শুধু ভাত ও ডাল)": return "ভাত, ডাল"
    return item if item != "nan" else ""

# Initialize empty dataframe in session state
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
                new_df = pd.read_csv(StringIO(data_input.strip()))
                if 'Items served' in new_df.columns:
                    new_df['Items served'] = new_df['Items served'].apply(to_dropdown_format)
                st.session_state.df = new_df
                st.rerun()
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
            options=FOOD_OPTIONS,
            required=False # Changed to False to allow empty values for the summary row
        )
    }
)

# ---------------------------------------------------------------------------
# Function to build the Excel structure as an EXACT clone of the original
# MDM_MONTHLY_REPORT.xlsx template
# ---------------------------------------------------------------------------
def create_excel_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MDM Report"

    # ---- Fonts ----
    f_title = Font(name="Arial", size=9, bold=True)        
    f_enroll = Font(name="Arial", size=7, bold=True)        
    f_hdr9 = Font(name="Arial", size=9, bold=True)          
    f_hdr8 = Font(name="Arial", size=8, bold=True)          
    f_data_bold = Font(name="Calibri", size=11, bold=True)  
    f_data_plain = Font(name="Calibri", size=11, bold=False)  

    # ---- Alignments ----
    a_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    a_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    a_center = Alignment(horizontal="center")
    a_right = Alignment(horizontal="right")

    thin = Side(style="thin")
    full_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    top_only_border = Border(top=thin)
    top_left_border = Border(left=thin, top=thin)
    top_right_border = Border(right=thin, top=thin)
    bottom_only_border = Border(bottom=thin)
    bottom_left_border = Border(left=thin, bottom=thin)
    bottom_right_border = Border(right=thin, bottom=thin)

    # ---- Base grid: thin border on every cell A1:H48 ----
    for row in ws.iter_rows(min_row=1, max_row=48, min_col=1, max_col=8):
        for cell in row:
            cell.border = full_border

    # Row 49 (blank divider): top border only, open sides except outer edges
    for col in range(1, 9):
        cell = ws.cell(row=49, column=col)
        if col == 1:
            cell.border = top_left_border
        elif col == 8:
            cell.border = top_right_border
        else:
            cell.border = top_only_border
        cell.font = Font(name="Calibri", size=11, bold=(col == 1))

    # Row 50 (signature line): bottom border only
    for col in range(1, 9):
        cell = ws.cell(row=50, column=col)
        if col == 1:
            cell.border = bottom_left_border
        elif col in (5, 8):
            cell.border = bottom_right_border
        else:
            cell.border = bottom_only_border
        cell.font = f_data_plain

    # ---- Column widths ----
    ws.column_dimensions['A'].width = 5.42578125
    ws.column_dimensions['B'].width = 10.42578125
    ws.column_dimensions['E'].width = 28.0
    ws.column_dimensions['F'].width = 11.42578125
    ws.column_dimensions['H'].width = 12.42578125

    # ---- Row heights ----
    row_heights = {
        1: 18.0, 2: 21.75, 3: 15.95, 4: 15.95,
        5: 14.1, 6: 14.1, 7: 14.1, 8: 14.1, 9: 14.1, 10: 14.1, 11: 14.1,
        12: 57.0,
        **{r: 13.5 for r in range(13, 39)},
        39: 15.0, 40: 18.75, 41: 25.5, 42: 15.0,
        43: 14.1, 44: 14.1, 45: 14.1, 46: 14.1, 47: 14.1,
        48: 13.5, 49: 52.5, 50: 36.75,
    }
    for r, h in row_heights.items():
        ws.row_dimensions[r].height = h

    # ---- Merges ----
    merges = [
        "A1:H1",
        "A2:E2", "F2:G2",
        "A3:E3", "F3:G3",
        "B4:E4",
        "B5:E5", "B6:E6", "B7:E7", "B8:E8", "B9:E9", "B10:E10", "B11:E11",
        "A39:B39",
        "A40:H40",
        "A41:H41",
        "B42:E42",
        "B43:E43", "B44:E44", "B45:E45", "B46:E46", "B47:E47", "B48:E48",
        "A49:H49",
        "A50:D50", "E50:H50",
    ]
    for m in merges:
        ws.merge_cells(m)

    # ---- Row 1: title ----
    c = ws.cell(row=1, column=1, value="MONTHLY REPORT FOR COOKING COST & FOOD GRAINS UNDER MID - DAY - MEAL FOR THE MONTH OF - ")
    c.font = f_title
    c.alignment = a_left_wrap

    # ---- Row 2: school name / enrolment ----
    c = ws.cell(row=2, column=1, value="Name of the School : SHYAMBAZAR D N HIGH SCHOOL")
    c.font = f_title
    c.alignment = a_left_wrap
    c = ws.cell(row=2, column=6, value="Total Enrolment : Class V = 60")
    c.font = f_enroll
    c.alignment = a_center_wrap
    ws.cell(row=2, column=8).font = f_title

    # ---- Row 3: gram sansad / class VI-VIII enrolment ----
    c = ws.cell(row=3, column=1, value="Gram Sansad No. - Bhallagram , Circle : Monglkote-II, Block : Mongalkote")
    c.font = f_title
    c.alignment = a_left_wrap
    c = ws.cell(row=3, column=6, value="Class VI to VIII = 322")
    c.font = f_hdr9
    c.alignment = a_center_wrap
    ws.cell(row=3, column=8).font = f_title

    # ---- Row 4: opening-balance table header ----
    header_row4 = {1: "Sl.", 2: "Particular", 6: "Class -V", 7: "VI to VIII", 8: "Total"}
    for col, val in header_row4.items():
        c = ws.cell(row=4, column=col, value=val)
        c.font = f_hdr9
        c.alignment = a_center_wrap

    # ---- Rows 5-11: opening-balance particulars ----
    particulars_top = [
        "Opening Balance Cooking Cost at the beginning of the month",
        "Opening Balance Food grains at the beginning of the month",
        "Fund for cooking cost received this month (for",
        "Honorarium received this month (for",
        "Food Grains received this month (for",
        "Total available Fund (1 + 3) (other grant)",
        "Total Food Grains available (2 + 5) AMOUNT",
    ]
    for idx, item in enumerate(particulars_top):
        r = 5 + idx
        c = ws.cell(row=r, column=1, value=idx + 1)
        c.font = f_data_bold
        c.alignment = a_center
        c = ws.cell(row=r, column=2, value=item)
        c.font = f_hdr9
        c.alignment = a_left_wrap
        for col in (6, 7, 8):
            ws.cell(row=r, column=col).font = f_data_bold

    # ---- Row 12: daily-data column headers ----
    headers = [
        "Sl. No.", "Date", "No of student availing MDM Class -V",
        "No of student availing MDM Class - VI to VIII", "Items served",
        "Cooking cost for Class - V to VIII", "Food Grains for Class - V",
        "Food grains for Class - VI to VIII"
    ]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=12, column=col_num, value=header)
        cell.font = f_hdr8
        cell.alignment = a_center_wrap

    # ---- Rows 13-38: daily data (blank, serial numbers only) ----
    for i in range(26):
        r = 13 + i
        c = ws.cell(row=r, column=1, value=i + 1)
        c.font = f_hdr9
        c.alignment = a_center_wrap
        for col in range(2, 9):
            ws.cell(row=r, column=col).font = f_data_bold

    # ---- Row 39: Total ----
    c = ws.cell(row=39, column=1, value="Total")
    c.font = f_hdr9
    c.alignment = a_center_wrap
    for col in range(3, 9):
        ws.cell(row=39, column=col).font = f_data_bold

    # ---- Row 40: cooking-cost formula note ----
    c = ws.cell(row=40, column=1,
                value="* (Total student of Class - V X 6.78) + Total student of Class - VI to VIII X 10.17)")
    c.font = f_hdr8
    c.alignment = a_left_wrap

    # ---- Row 41: SHG / bank note ----
    c = ws.cell(row=41, column=1,
                value="Name of SHG (Cook-cum-Helper) MC - APPROVED _Name of Bank with Branch CANARA BANK, "
                      "Mathrun Branch._ A/c no_3641101002012_ worked during month.......")
    c.font = f_hdr8
    c.alignment = a_left_wrap

    # ---- Row 42: closing-balance table header ----
    header_row42 = {1: "Sl.", 2: "Particular", 6: "Class - V", 7: "VI to VIII", 8: "Total"}
    for col, val in header_row42.items():
        c = ws.cell(row=42, column=col, value=val)
        c.font = f_hdr9
        c.alignment = a_center_wrap

    # ---- Rows 43-48: closing-balance particulars ----
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
        c.font = f_hdr9
        c.alignment = a_center_wrap
        c = ws.cell(row=r, column=2, value=item)
        c.font = f_hdr9
        c.alignment = a_left_wrap
        for col in (6, 7, 8):
            ws.cell(row=r, column=col).font = f_data_bold

    # ---- Row 50: signature line ----
    c = ws.cell(row=50, column=1, value="Signature of MDM Nodal Teacher")
    c.font = f_hdr9
    c.alignment = a_center
    c = ws.cell(row=50, column=5, value="Signature of H M with seal")
    c.font = f_hdr9
    c.alignment = a_right

    return wb, ws


# 3. Generate Final Report
if st.button("Generate Final Excel Report", type="primary"):
    try:
        wb, sheet = create_excel_template()
        start_row = 13

        for index, row in edited_df.iterrows():
            # Check if this is the summary/total row (e.g. "22 Days" or "Total")
            is_summary = pd.notna(row["Date"]) and ("Days" in str(row["Date"]) or "Total" in str(row["Date"]))

            if is_summary:
                # Force the summary row exactly to Row 39 
                # Write specifically to Column 1 to avoid crashing on the A39:B39 merged cell
                sheet.cell(row=39, column=1).value = row["Date"]
                
                sheet.cell(row=39, column=3).value = row["No of student availing MDM Class - V"]
                sheet.cell(row=39, column=4).value = row["No of student availing MDM Class - VI to VIII"]
                sheet.cell(row=39, column=6).value = row["Cooking cost for Class - V to VIII"]
                sheet.cell(row=39, column=7).value = row["Food Grains for Class - V"]
                sheet.cell(row=39, column=8).value = row["Food grains for Class - VI to VIII"]
            else:
                current_row = start_row + index
                
                # Prevent overflow if the user adds too many days
                if current_row >= 39: 
                    break

                if pd.isna(row["Date"]) and pd.isna(row["No of student availing MDM Class - V"]):
                    continue

                sheet.cell(row=current_row, column=1).value = index + 1
                sheet.cell(row=current_row, column=2).value = row["Date"]
                sheet.cell(row=current_row, column=3).value = row["No of student availing MDM Class - V"]
                sheet.cell(row=current_row, column=4).value = row["No of student availing MDM Class - VI to VIII"]

                final_food_string = to_excel_format(row["Items served"])
                if final_food_string:
                    sheet.cell(row=current_row, column=5).value = final_food_string

                sheet.cell(row=current_row, column=6).value = row["Cooking cost for Class - V to VIII"]
                sheet.cell(row=current_row, column=7).value = row["Food Grains for Class - V"]
                sheet.cell(row=current_row, column=8).value = row["Food grains for Class - VI to VIII"]

        # Save to memory for download
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("Report generated successfully!")

        st.download_button(
            label="Download Populated MDM Report",
            data=output,
            file_name="Completed_MDM_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
