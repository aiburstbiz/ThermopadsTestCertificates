import sys
import subprocess
import os
import io
import datetime
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font

def install_missing():
    try:
        import xlrd, openpyxl
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "xlrd", "openpyxl"])
install_missing()

# --- 0. SESSION STATE ---
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = []

st.set_page_config(page_title="Thermopads Automator", layout="wide")

# --- 1. THEME ---
st.markdown("""
    <style>
        .stApp { background-color: #f8f9fa; }
        h1, label, p, [data-testid="stWidgetLabel"] p { color: #C62828 !important; font-weight: bold !important; }

        div.stButton > button {
            background-color: #C62828 !important;
            border-radius: 5px; border: 2px solid #8E0000;
            width: 100%; height: 3.5em;
            cursor: pointer !important;
        }
        div.stButton > button * { color: #000000 !important; font-weight: 900 !important; }
        div.stButton > button:hover { background-color: #A52121 !important; }

        /* Make ALL buttons in the two small columns equally compact */
        div[data-testid="stHorizontalBlock"]:has(> div:nth-child(3)) div.stButton > button {
            height: 1.9em !important;
            min-height: unset !important;
            font-size: 0.76rem !important;
            padding: 0 10px !important;
            border-width: 1px !important;
        }

        div[data-baseweb="select"], div[data-baseweb="select"] * { cursor: pointer !important; }
        .stMultiSelect div[data-baseweb="tag"] { background-color: #C62828 !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>🛡️ Thermopads Test Certificate Automator</h1>", unsafe_allow_html=True)

# --- 2. CACHED SMART LOADER ---
@st.cache_data(show_spinner=False)
def load_data_smart(file_bytes, fname):
    try:
        fname = fname.lower()
        if fname.endswith('.xlsx'):
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str, engine='openpyxl')
        elif fname.endswith('.xls'):
            try:
                df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str, engine='xlrd')
            except:
                text = file_bytes.decode('utf-8', errors='ignore')
                df_raw = pd.read_csv(io.StringIO(text), header=None, sep=r'\t+', engine='python', dtype=str)
        else:
            text = file_bytes.decode('utf-8', errors='ignore')
            sep = r'\t+' if '\t' in text else ','
            df_raw = pd.read_csv(io.StringIO(text), header=None, sep=sep, engine='python', dtype=str)

        header_row_index = 0
        keywords = ['orderno', 'primarysrno', 'channelid', 'matdesc', 'heatingcable', 'sl.no', 'materialno']
        for i, row in df_raw.head(100).iterrows():
            row_vals = [str(v).strip().lower().replace(" ", "").replace("\t", "") for v in row.values if v is not None]
            if any(k in row_vals for k in keywords):
                header_row_index = i
                break
        df = df_raw.iloc[header_row_index:].copy()
        df.columns = [str(c).strip() for c in df.iloc[0]]
        df = df.iloc[1:].reset_index(drop=True)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed|^nan|^None|^$', na=False)]
        for col in df.columns:
            df[col] = df[col].astype(str).apply(lambda x: x.strip().replace('\t', '') if x != 'nan' else "")
        return df
    except:
        return None

def get_col(df, keywords):
    for col in df.columns:
        if any(k in col.lower().replace(" ", "").replace("_", "") for k in keywords):
            return col
    return None

def extract_watts(desc):
    match = re.search(r'(\d+)\s*W', str(desc), re.IGNORECASE)
    return match.group(1) if match else ""


# ============================================================
# ENGINE A: WARMUP  (anchor row 18, header at A9–G12)
# ============================================================
def generate_warmup_engine(df_subset, template_path, o_date, s_date, product_type, packing_filename, log_p, log_h):
    cust_col = None
    for col in df_subset.columns:
        if 'customername' in col.lower().replace(" ", "").replace("_", "") and col.endswith('_y'):
            cust_col = col; break
    if cust_col is None:
        cust_col = get_col(df_subset, ['customername', 'customer'])

    po_col = get_col(df_subset, ['orderno', 'ponumber', 'po'])
    id_col = get_col(df_subset, ['primary', 'channel', 'cable'])

    df_subset = df_subset.sort_values(by=id_col).reset_index(drop=True)
    wb = load_workbook(template_path)
    ws = wb.active

    thin         = Side(style='thin', color="000000")
    table_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font_black   = Font(color="000000", size=10)

    file_digits   = "".join(re.findall(r'\d+', packing_filename))[:4]
    sno_label     = f"{'SFHM' if product_type == 'MAT' else 'Wup'}-Wup-{file_digits}"
    customer_name = str(df_subset.iloc[0].get(cust_col, 'M/s. Warmup')).strip()

    sample = df_subset.iloc[0]
    # Warmup template header layout: rows 9–12, right side G11/G12
    ws['A9']  = f"Customer  :  {customer_name}"
    ws['A10'] = f"PO No : {sample.get(po_col, '')}"
    ws['A11'] = f"PO Dt : {o_date}"
    ws['G11'] = f"S.NO : {sno_label}"
    ws['A12'] = f"Product : {sample.get('MatDesc', sample.get('ProductName', ''))}"
    ws['G12'] = f"DATE : {s_date}"

    anchor_row   = 18
    footer_start = 1000
    for r in range(anchor_row, 300):
        row_combined = ' '.join(str(ws.cell(row=r, column=c).value or '').upper() for c in range(1, 9))
        if any(x in row_combined for x in ["PREPARED", "CHECKED", "APPROVED", "THERMOPADS"]):
            footer_start = r; break

    available  = footer_start - anchor_row
    batch_size = len(df_subset)
    if batch_size > available:
        ws.insert_rows(footer_start - 1, amount=batch_size - available)
    elif batch_size < available:
        ws.delete_rows(anchor_row + batch_size, amount=available - batch_size)

    for i, row in df_subset.iterrows():
        curr = anchor_row + i
        ts   = datetime.datetime.now().strftime('%H:%M:%S')
        log_h.append(f"[{ts}] MAPPING: {row.get(id_col)} -> Row {curr}")
        log_p.code("\n".join(log_h[-8:]))
        ws.cell(row=curr, column=1, value=i + 1).border = table_border
        ws.cell(row=curr, column=2, value=row.get(id_col, '')).border = table_border
        ws.cell(row=curr, column=3, value="".join(re.findall(r'\d+\.?\d*', str(row.get('Size', ''))))).border = table_border
        watts = extract_watts(row.get('MatDesc', ''))
        ws.cell(row=curr, column=4, value=watts if watts else "").border = table_border
        ws.cell(row=curr, column=5, value=row.get('Actualminout', '')).border = table_border
        ws.cell(row=curr, column=6, value="> 500").border = table_border
        ws.cell(row=curr, column=7, value="No Breakdown").border = table_border
        for c in range(1, 8):
            ws.cell(row=curr, column=c).alignment = Alignment(horizontal='center')
            ws.cell(row=curr, column=c).font = font_black
    return wb


# ============================================================
# ENGINE B: WARMLY  (anchor row 20, header at rows 11–14, G13/G14)
# ============================================================
def generate_warmly_engine(df_subset, template_path, o_date, s_date, product_type, packing_filename, log_p, log_h):
    cust_col = None
    for col in df_subset.columns:
        if 'customername' in col.lower().replace(" ", "").replace("_", "") and col.endswith('_y'):
            cust_col = col; break
    if cust_col is None:
        cust_col = get_col(df_subset, ['customername', 'customer'])

    po_col = get_col(df_subset, ['orderno', 'ponumber', 'po'])
    id_col = get_col(df_subset, ['primary', 'channel', 'cable'])

    df_subset = df_subset.sort_values(by=id_col).reset_index(drop=True)
    wb = load_workbook(template_path)
    ws = wb.active

    thin         = Side(style='thin', color="000000")
    table_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font_black   = Font(color="000000", size=10)

    file_digits   = "".join(re.findall(r'\d+', packing_filename))[:4]
    sno_label     = f"{'SFHM' if product_type == 'MAT' else 'Wup'}-Warmly-{file_digits}"
    customer_name = str(df_subset.iloc[0].get(cust_col, 'M/s. Warmly')).strip()

    sample = df_subset.iloc[0]
    # Warmly template header layout: rows 11–14, right side G13/G14
    # ONLY update the dynamic fields — leave all other template content untouched
    ws['A11'] = f"CUSTOMER : {customer_name}"
    ws['A12'] = f"PO No : {sample.get(po_col, '')}"
    # G12 = ORIGINAL — DO NOT TOUCH (already in template)
    ws['A13'] = f"PO Dt : {o_date}"
    ws['G13'] = f"S.NO : {sno_label}"
    ws['A14'] = f"Product : {sample.get('MatDesc', sample.get('ProductName', ''))}"
    ws['G14'] = f"DATE : {s_date}"

    anchor_row   = 20
    footer_start = 1000
    for r in range(anchor_row, 300):
        row_combined = ' '.join(str(ws.cell(row=r, column=c).value or '').upper() for c in range(1, 9))
        if any(x in row_combined for x in ["PREPARED", "CHECKED", "APPROVED", "THERMOPADS"]):
            footer_start = r; break

    available  = footer_start - anchor_row
    batch_size = len(df_subset)
    if batch_size > available:
        ws.insert_rows(footer_start - 1, amount=batch_size - available)
    elif batch_size < available:
        ws.delete_rows(anchor_row + batch_size, amount=available - batch_size)

    for i, row in df_subset.iterrows():
        curr = anchor_row + i
        ts   = datetime.datetime.now().strftime('%H:%M:%S')
        log_h.append(f"[{ts}] MAPPING: {row.get(id_col)} -> Row {curr}")
        log_p.code("\n".join(log_h[-8:]))
        ws.cell(row=curr, column=1, value=i + 1).border = table_border
        ws.cell(row=curr, column=2, value=row.get(id_col, '')).border = table_border
        ws.cell(row=curr, column=3, value="".join(re.findall(r'\d+\.?\d*', str(row.get('Size', ''))))).border = table_border
        watts = extract_watts(row.get('MatDesc', ''))
        ws.cell(row=curr, column=4, value=watts if watts else "").border = table_border
        ws.cell(row=curr, column=5, value=row.get('Actualminout', '')).border = table_border
        ws.cell(row=curr, column=6, value="> 500").border = table_border
        ws.cell(row=curr, column=7, value="No Breakdown").border = table_border
        for c in range(1, 8):
            ws.cell(row=curr, column=c).alignment = Alignment(horizontal='center')
            ws.cell(row=curr, column=c).font = font_black
    return wb


# ============================================================
# ENGINE C: SCHLUTER
# ============================================================
def generate_schluter_engine(merged, o_date, s_date, log_p, log_h):
    template_path = "schluter_template.xlsx"
    if not os.path.exists(template_path):
        st.error(f"CRITICAL: '{template_path}' missing from app folder!")
        return None, None

    wb     = load_workbook(template_path)
    ws     = wb.active
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    side   = Side(style='thin')
    thin_b = Border(left=side, right=side, top=side, bottom=side)

    ord_no    = str(merged.iloc[0].get('OrderNo', 'N/A'))
    ws['A8']  = f"Purchase order no. {ord_no}"
    ws['A9']  = f"Dt : {o_date}"
    ws['A10'] = f"Date of shipment : {s_date}"

    mod_col    = 'CustomerCode' if 'CustomerCode' in merged.columns else 'MaterialNo'
    cable_cols = [c for c in merged.columns if any(x in c.lower() for x in ['cable', 'channel', 'primary'])]
    cable_col  = cable_cols[0] if cable_cols else merged.columns[1]

    merged = merged.copy()
    merged['_sort_key'] = merged[mod_col].astype(str).str[-3:]
    merged = merged.sort_values(by=['_sort_key', cable_col]).drop(columns='_sort_key').reset_index(drop=True)

    current_excel_row = 17
    previous_model    = None

    for i, row in merged.iterrows():
        current_model = str(row[mod_col])
        if previous_model is not None and current_model != previous_model:
            ws.insert_rows(current_excel_row)
            for cell in ws[current_excel_row]:
                cell.border = Border()
            current_excel_row += 1

        ts = datetime.datetime.now().strftime('%H:%M:%S')
        log_h.append(f"[{ts}] MAPPING: {current_model} -> Row {current_excel_row}")
        log_p.code("\n".join(log_h[-10:]))

        row_data = [
            row[mod_col], row[cable_col],
            row.get('Expectedminout', ''), row.get('Expectedmaxout', ''), row.get('Actualminout', '')
        ]
        for c_idx, val in enumerate(row_data, 1):
            cell           = ws.cell(row=current_excel_row, column=c_idx, value=str(val))
            cell.border    = thin_b
            cell.alignment = center

        previous_model    = current_model
        current_excel_row += 1

    return wb, ord_no


# ============================================================
# --- 4. MAIN UI ---
# ============================================================
st.subheader("🏢 Batch Processing")
customer_choice = st.selectbox("Select Customer", ["Warmup", "Warmly", "Schluter"])

c1, c2 = st.columns(2)
with c1: ui_o_date = st.text_input("Order Date (PO Dt)",   value=datetime.datetime.now().strftime("%d.%m.%Y"))
with c2: ui_s_date = st.text_input("Shipment Date (DATE)", value=datetime.datetime.now().strftime("%d.%m.%Y"))

qc_file = st.file_uploader("1. QC Report",    type=["xlsx", "xls", "csv"])
pk_file = st.file_uploader("2. Packing Data", type=["xlsx", "xls", "csv"])

if qc_file and pk_file:
    df_qc = load_data_smart(qc_file.getvalue(), qc_file.name)
    df_pk = load_data_smart(pk_file.getvalue(), pk_file.name)

    if df_qc is not None and df_pk is not None:
        q_col  = get_col(df_qc, ['channel', 'cable', 'primary'])
        p_col  = get_col(df_pk,  ['primary', 'cable', 'channel'])
        merged = df_qc.merge(df_pk, left_on=q_col, right_on=p_col, how='inner')

        if not merged.empty:

            # ---- WARMUP / WARMLY ----
            if customer_choice in ["Warmup", "Warmly"]:
                matdesc_col = 'MatDesc' if 'MatDesc' in merged.columns else get_col(merged, ['matdesc'])
                p_options   = sorted(merged[matdesc_col].unique().tolist())

                # Equal-size compact buttons in equal-width columns
                sa_col, cb_col, _ = st.columns([1, 1, 8])
                with sa_col:
                    if st.button("✅ All", key="sel_all", help="Select all products"):
                        st.session_state['ms_selected'] = p_options
                with cb_col:
                    if st.button("❌ Clear", key="clr_all", help="Clear selection"):
                        st.session_state['ms_selected'] = []

                default_sel = st.session_state.get('ms_selected', [])
                selected = st.multiselect(
                    "📦 Select Product Batches to Generate",
                    options=p_options,
                    default=default_sel,
                    help="Pick products individually, or use ✅ All above."
                )
                st.session_state['ms_selected'] = selected

                if selected:
                    st.info(f"📋 {len(selected)} of {len(p_options)} product(s) selected — {len(selected)} certificate(s) will be generated.")
                else:
                    st.warning("⚠️ No products selected. Pick at least one to generate certificates.")

                if selected and st.button("🚀 GENERATE OFFICIAL CERTIFICATES"):
                    st.session_state.generated_files = []
                    with st.expander("📊 LIVE PROCESSING LOGS", expanded=True):
                        log_p = st.empty()
                        log_h = [f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Starting {customer_choice} — {len(selected)} product(s)..."]
                        for prod in selected:
                            sub    = merged[merged[matdesc_col] == prod]
                            is_mat = any(x in prod.upper() for x in ["MAT", "STICKY", "SFHMT"])
                            p_type = "MAT" if is_mat else "CABLE"

                            if customer_choice == "Warmly":
                                template = f"warmly_{'mat' if is_mat else 'cable'}_template.xlsx"
                                wb = generate_warmly_engine(
                                    sub, template, ui_o_date, ui_s_date,
                                    p_type, pk_file.name, log_p, log_h
                                )
                            else:  # Warmup
                                template = f"warmup_{'mat' if is_mat else 'cable'}_template.xlsx"
                                wb = generate_warmup_engine(
                                    sub, template, ui_o_date, ui_s_date,
                                    p_type, pk_file.name, log_p, log_h
                                )

                            if wb:
                                buf   = io.BytesIO()
                                wb.save(buf)
                                clean = re.sub(r'[\\/*?:"<>|]', "", prod)[:25]
                                st.session_state.generated_files.append({
                                    "name": f"Certificate_{customer_choice}_{clean}.xlsx",
                                    "data": buf.getvalue()
                                })
                    st.rerun()

            # ---- SCHLUTER ----
            else:
                if st.button("🚀 GENERATE SCHLUTER CERTIFICATE"):
                    st.session_state.generated_files = []
                    with st.expander("📊 LIVE LOGS", expanded=True):
                        log_p = st.empty()
                        log_h = [f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Starting Schluter..."]
                        wb, oid = generate_schluter_engine(merged, ui_o_date, ui_s_date, log_p, log_h)
                        if wb:
                            buf = io.BytesIO()
                            wb.save(buf)
                            st.session_state.generated_files.append({
                                "name": f"COA_{oid}.xlsx",
                                "data": buf.getvalue()
                            })
                    st.rerun()

        else:
            st.error("⚠️ No matching serial numbers found between QC and Packing files.")

# ---- DOWNLOAD SECTION ----
if st.session_state.generated_files:
    st.write("---")
    st.success(f"✅ {len(st.session_state.generated_files)} certificate(s) ready!")
    for i, f in enumerate(st.session_state.generated_files):
        st.download_button(
            label=f"📥 DOWNLOAD: {f['name']}",
            data=f['data'],
            file_name=f['name'],
            key=f"btn_{i}"
        )
    if st.button("🔄 START NEW BATCH"):
        st.session_state.generated_files = []
        st.session_state['ms_selected']  = []
        st.rerun()