import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
import io
import os

# --- CONFIGURATION ---
LOGO_PATH = "logo.png" 

st.set_page_config(page_title="Official COA Generator", layout="centered")
st.title("🛡️ Official Test Certificate Automator")
st.write("Upload QC and Packing data (**Excel or CSV**) to generate the Schluter COA.")

# --- SMART FILE LOADER ---
# This function handles both CSV/Excel and hunts for the correct header row
def load_data_smart(file):
    filename = file.name
    try:
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            df_temp = pd.read_excel(file, header=None)
            # Search first 20 rows for key column names
            for i, row in df_temp.iterrows():
                row_values = [str(val).strip() for val in row.values]
                if 'Channelid' in row_values or 'PrimarySrNo' in row_values:
                    return pd.read_excel(file, header=i)
            return pd.read_excel(file, header=1) 
        else:
            # For CSV files
            file.seek(0)
            df_temp = pd.read_csv(file, header=None, sep=None, engine='python')
            for i, row in df_temp.iterrows():
                row_values = [str(val).strip() for val in row.values]
                if 'Channelid' in row_values or 'PrimarySrNo' in row_values:
                    file.seek(0)
                    return pd.read_csv(file, header=i, sep=None, engine='python')
            file.seek(0)
            return pd.read_csv(file, header=1, sep=None, engine='python')
    except Exception as e:
        st.error(f"❌ Error reading {filename}: {e}")
        return None

# --- 1. FILE UPLOADERS ---
col1, col2 = st.columns(2)
with col1:
    qc_file = st.file_uploader("1. Upload QC Inspection Report", type=["xlsx", "csv"])
with col2:
    pack_file = st.file_uploader("2. Upload Packing Scanning Data", type=["xlsx", "csv"])

if qc_file and pack_file:
    try:
        # Load Data Smartly
        df_qc = load_data_smart(qc_file)
        df_pack = load_data_smart(pack_file)

        if df_qc is not None and df_pack is not None:
            df_qc.columns = df_qc.columns.astype(str).str.strip()
            df_pack.columns = df_pack.columns.astype(str).str.strip()

            df_qc['Channelid'] = df_qc['Channelid'].astype(str).str.strip()
            df_pack['PrimarySrNo'] = df_pack['PrimarySrNo'].astype(str).str.strip()
            
            # Reconciliation
            merged = df_qc.merge(df_pack, left_on='Channelid', right_on='PrimarySrNo', how='inner')

            if merged.empty:
                st.warning("⚠️ No matches found. Ensure IDs in both files are correct.")
            else:
                # --- 2. BUILD THE REPLICA ---
                wb = Workbook()
                ws = wb.active
                ws.title = "Test Certificate"
                
                bold_f = Font(bold=True, size=10)
                header_f = Font(bold=True, size=12)
                center = Alignment(horizontal='center', vertical='center', wrap_text=True)
                thin_b = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                top=Side(style='thin'), bottom=Side(style='thin'))

                # --- LOGO ---
                if os.path.exists(LOGO_PATH):
                    img = OpenpyxlImage(LOGO_PATH)
                    img.width, img.height = 140, 70
                    ws.add_image(img, 'A1')

                # Header Matter
                ws.merge_cells('D2:F2'); ws['D2'] = "SCHEDULE 3"; ws['D2'].font = header_f; ws['D2'].alignment = center
                ws.merge_cells('D3:F3'); ws['D3'] = "CERTIFICATE OF ANALYSIS"; ws['D3'].font = header_f; ws['D3'].alignment = center
                ws['G7'] = "TPL/SCH/TC-06"
                
                # Metadata
                info = merged.iloc[0]
                ws['A8'] = f"Schluter's purchase order no. {info.get('OrderNo', 'N/A')}"
                ws['A10'] = f"Date of shipment : {info.get('TransDate', 'N/A')}"
                ws['A11'] = "with Updated Termination:"
                ws['A12'] = "Better Transparent Silicone Sealant and Heat Shrink tubing with better adhesive (transparent glue);"

                # --- 3. TABLE HEADERS ---
                h_map = [
                    ("A14:A15", "Sl.No."),
                    ("B14:B15", "Heating Cable Sl.No."),
                    ("C14:C15", "Model No."),
                    ("D14:G14", "CABLE DC resistance Ohms"),
                ]
                
                for cell_range, text in h_map:
                    ws.merge_cells(cell_range)
                    cell = ws[cell_range.split(':')[0]]
                    cell.value = text; cell.font = bold_f; cell.alignment = center
                    start, end = cell_range.split(':')
                    for row_cells in ws[start:end]:
                        for c in row_cells: c.border = thin_b

                # Sub-Headers Row 15
                sub_headers = {'D15': "Min", 'E15': "Max", 'F15': "Actual", 'G15': "PASS / FAIL"}
                for cell_ref, text in sub_headers.items():
                    ws[cell_ref] = text
                    ws[cell_ref].border = thin_b
                    ws[cell_ref].font = bold_f
                    ws[cell_ref].alignment = center

                # --- 4. DATA FILLING ---
                for i, row in merged.iterrows():
                    r = 16 + i
                    v_min, v_max, v_act = row['Expectedminout'], row['Expectedmaxout'], row['Actualminout']
                    status = "PASS" if v_min <= v_act <= v_max else "FAIL"

                    data = [i+1, row['Channelid'], row['CustomerCode'], v_min, v_max, v_act, status]
                    
                    for c_idx, val in enumerate(data, 1):
                        cell = ws.cell(row=r, column=c_idx, value=val)
                        cell.border = thin_b
                        cell.alignment = center

                # --- 5. EXPORT ---
                output = io.BytesIO()
                wb.save(output)
                st.success(f"✅ Created COA with {len(merged)} matched records.")
                st.download_button("📥 Download Final COA", output.getvalue(), f"COA_{info.get('OrderNo', 'Batch')}.xlsx")

    except Exception as e:
        st.error(f"❌ System Error: {e}")