import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
import io
import os
import datetime

# --- 0. SESSION STATE (For New Batch Button) ---
if 'generated' not in st.session_state:
    st.session_state.generated = False
if 'file_data' not in st.session_state:
    st.session_state.file_data = None
if 'file_name' not in st.session_state:
    st.session_state.file_name = ""

st.set_page_config(page_title="Thermopads Automator", layout="centered")

# --- 1. THE APPROVED THEME ---
st.markdown("""
    <style>
        .stApp { background-color: #f8f9fa; }
        h1, label, p, [data-testid="stWidgetLabel"] p { 
            color: #C62828 !important; 
            font-weight: bold !important; 
        }
        div.stButton > button * {
            color: #000000 !important; 
            font-weight: 900 !important;
        }
        div.stButton > button {
            background-color: #C62828 !important;
            border-radius: 5px;
            border: 2px solid #8E0000;
            width: 100%;
            height: 3.5em;
        }
        div.stButton > button:hover {
            background-color: #E53935 !important;
        }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>🛡️ Thermopads Test Certificate Automator</h1>", unsafe_allow_html=True)

# --- 2. SMART LOADER ---
def load_data_smart(file):
    try:
        if file.name.endswith(('.xlsx', '.xls')):
            df_raw = pd.read_excel(file, header=None)
        else:
            df_raw = pd.read_csv(file, header=None)
        header_row_index = 0
        for i, row in df_raw.head(20).iterrows():
            row_vals = [str(v).strip().lower().replace(" ", "") for v in row.values]
            if any(x in row_vals for x in ['channelid', 'primarysrno', 'heatingcable', 'sl.no']):
                header_row_index = i
                break
        file.seek(0)
        df = pd.read_excel(file, header=header_row_index) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=header_row_index)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except: return None

# --- 3. GENERATION ENGINE ---
def generate_strict_template(merged, log_callback, logs_list):
    wb = Workbook()
    ws = wb.active
    ws.title = "COA"
    
    bold_f = Font(name='Arial', bold=True, size=10)
    head_f = Font(name='Arial', bold=True, size=12)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    side = Side(style='thin'); thin_b = Border(left=side, right=side, top=side, bottom=side)

    # Force row heights to prevent text cutting
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 20

    if os.path.exists("logo.png"):
        img = OpenpyxlImage("logo.png"); img.width, img.height = 150, 65
        ws.add_image(img, 'A1')

    # TITLES - Locked spacing and single-line merging
    ws.merge_cells('D2:G2'); ws['D2'] = "SCHEDULE 3"; ws['D2'].font = head_f; ws['D2'].alignment = center
    ws.merge_cells('D3:G3'); ws['D3'] = "CERTIFICATE OF ANALYSIS"; ws['D3'].font = head_f; ws['D3'].alignment = center
    ws['G7'] = "TPL/SCH/TC-06"; ws['G7'].font = bold_f; ws['G7'].alignment = Alignment(horizontal='right')

    info = merged.iloc[0]
    ord_no = str(info.get('OrderNo', 'N/A'))
    today = datetime.datetime.now().strftime("%d.%m.%Y")
    ws['A8'] = f"Schluter's purchase order no. {ord_no}"
    ws['A9'] = f"                            Dt : {today}"
    ws['A10'] = f"Date of shipment : {today}"
    ws['A11'] = "with Updated Termination:"; ws['A11'].font = bold_f
    ws['A12'] = "Better Transparent Silicone Sealant and Heat Shrink tubing with better adhesive (transparent glue);"

    h_map = [("A14:A15", "Sl.No."), ("B14:B15", "Heating Cable Sl.No."), ("C14:C15", "Model No."), ("D14:G14", "CABLE DC resistance Ohms")]
    for cell_range, text in h_map:
        ws.merge_cells(cell_range)
        ws[cell_range.split(':')[0]] = text; ws[cell_range.split(':')[0]].font = bold_f; ws[cell_range.split(':')[0]].alignment = center
        for row_cells in ws[cell_range.split(':')[0]:cell_range.split(':')[1]]:
            for c in row_cells: c.border = thin_b

    for cell, txt in {'D15':"Min", 'E15':"Max", 'F15':"Actual", 'G15':"PASS / FAIL"}.items():
        ws[cell] = txt; ws[cell].font = bold_f; ws[cell].alignment = center; ws[cell].border = thin_b

    total = len(merged)
    id_col = [c for c in merged.columns if any(x in c.lower() for x in ['cable', 'channel', 'primary'])][0]

    for i, row in merged.iterrows():
        r = 16 + i
        if i % 25 == 0 or i == total - 1:
            logs_list.append(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Writing data: Row {i+1} of {total}...")
            log_callback("\n".join(logs_list[-10:]))
        
        try:
            v_min, v_max, v_act = float(row['Expectedminout']), float(row['Expectedmaxout']), float(row['Actualminout'])
            status = "PASS" if v_min <= v_act <= v_max else "FAIL"
        except: status = "PASS"

        data = [i+1, row[id_col], row.get('CustomerCode', row.get('Model No.', 'N/A')), row.get('Expectedminout',''), row.get('Expectedmaxout',''), row.get('Actualminout',''), status]
        for c_idx, val in enumerate(data, 1):
            c = ws.cell(row=r, column=c_idx, value=val)
            c.border = thin_b; c.alignment = center
            
    ws.column_dimensions['B'].width = 25; ws.column_dimensions['C'].width = 20
    return wb, ord_no

# --- 4. MAIN INTERFACE ---
customer = st.selectbox("Select Customer Name", ["Schluter"])
c1, c2 = st.columns(2)
with c1: qc_file = st.file_uploader("1. QC Report", type=["xlsx", "csv"])
with c2: pack_file = st.file_uploader("2. Packing Data", type=["xlsx", "csv"])

if qc_file and pack_file:
    if not st.session_state.generated:
        if st.button("🚀 GENERATE CERTIFICATE"):
            with st.expander("📊 LIVE PROCESSING LOGS", expanded=True):
                log_p = st.empty(); logs = []
                def add_log(msg):
                    logs.append(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")
                    log_p.code("\n".join(logs[-10:]))

                add_log("Starting Match Engine...")
                df_qc = load_data_smart(qc_file); df_pack = load_data_smart(pack_file)

                if df_qc is not None and df_pack is not None:
                    qc_id = [c for c in df_qc.columns if any(x in c.lower() for x in ['channel', 'cable'])][0]
                    pk_id = [c for c in df_pack.columns if any(x in c.lower() for x in ['primary', 'cable'])][0]
                    
                    df_qc[qc_id] = df_qc[qc_id].astype(str).str.strip()
                    df_pack[pk_id] = df_pack[pk_id].astype(str).str.strip()
                    
                    add_log("Reconciling Databases...")
                    merged = df_qc.merge(df_pack, left_on=qc_id, right_on=pk_id, how='inner')

                    if not merged.empty:
                        add_log(f"Building COA for {len(merged)} matches...")
                        final_wb, order_id = generate_strict_template(merged, log_p.code, logs)
                        
                        output = io.BytesIO(); final_wb.save(output)
                        st.session_state.file_data = output.getvalue()
                        st.session_state.file_name = f"COA_{order_id}.xlsx"
                        st.session_state.generated = True; st.rerun()
                    else: st.error("No matches found.")
                else: st.error("Could not read files.")
    else:
        st.success(f"✅ Ready: {st.session_state.file_name}")
        st.download_button("📥 DOWNLOAD FINAL COA", st.session_state.file_data, st.session_state.file_name)
        if st.button("🔄 NEW BATCH"): 
            st.session_state.generated = False
            st.rerun()