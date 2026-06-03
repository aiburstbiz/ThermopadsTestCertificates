import base64
import math
import zipfile
import sys
import subprocess
import os
import io
import datetime
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill

def install_missing():
    try:
        import openpyxl
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
install_missing()

# ─── SESSION STATE ────────────────────────────────────────────────────────────
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = []

st.set_page_config(page_title="Thermopads Automator", layout="wide")

# ─── THEME ────────────────────────────────────────────────────────────────────
# ─── THEME ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* ── Base ── */
    .stApp {
        background-color: #f8f9fa;
    }

    h1, label, p, [data-testid="stWidgetLabel"] p {
        color: #C62828 !important;
        font-weight: bold !important;
    }

    /* ── Inputs ── */
    div[data-testid="stTextInput"] input {
        border: 2px solid #BDBDBD !important;
        border-radius: 6px !important;
        background-color: white !important;
    }
    div[data-testid="stTextInput"] input:focus {
        border: 2px solid #9E9E9E !important;
        box-shadow: none !important;
    }

    /* ── Selects ── */
    div[data-testid="stSelectbox"] > div,
    div[data-testid="stMultiSelect"] > div,
    div[data-baseweb="select"] {
        border: 2px solid #BDBDBD !important;
        border-radius: 6px !important;
        background-color: white !important;
        cursor: pointer !important;
    }
    div[data-baseweb="select"] * { cursor: pointer !important; }

    /* ── Multi-select tags ── */
    .stMultiSelect div[data-baseweb="tag"] {
        background-color: #C62828 !important;
        color: white !important;
        border-radius: 4px !important;
    }

    /* ── Buttons ── */
    div.stButton > button {
        background-color: #C62828 !important;
        border-radius: 5px;
        border: 2px solid #8E0000;
        width: 100%;
        height: 3.5em;
        cursor: pointer !important;
    }
    div.stButton > button * {
        color: #000000 !important;
        font-weight: 900 !important;
    }
    div.stButton > button:hover { background-color: #A52121 !important; }

    div[data-testid="stHorizontalBlock"]:has(> div:nth-child(3))
    div.stButton > button {
        height: 1.9em !important;
        min-height: unset !important;
        font-size: 0.76rem !important;
        padding: 0 10px !important;
        border-width: 1px !important;
    }

    /* ── Expander ── */
    .stExpander {
        border: 2px solid #D0D0D0 !important;
        border-radius: 8px !important;
    }

    /* ══════════════════════════════════════════════════════════════
       FILE UPLOADER
       ══════════════════════════════════════════════════════════════ */

    div[data-testid="stFileUploader"] {
        border: 2px solid #9E9E9E !important;
        border-radius: 8px !important;
        padding: 6px !important;
        background-color: #EDEDED !important;
    }

    section[data-testid="stFileUploadDropzone"] {
        border: 2px dashed #9E9E9E !important;
        border-radius: 8px !important;
        background-color: #E0E0E0 !important;
        min-height: 72px !important;
        cursor: pointer !important;
    }

    /* Hide Browse files button */
    section[data-testid="stFileUploadDropzone"] button,
    div[data-testid="stFileUploader"] button {
        display: none !important;
        visibility: hidden !important;
        width: 0 !important;
        height: 0 !important;
        overflow: hidden !important;
        position: absolute !important;
    }

    /* Uploaded file name prefix */
    div[data-testid="stFileUploaderFileName"]::before {
        content: "Uploaded: ";
        font-weight: bold;
        color: #2E7D32;
    }

    /* ── Footer copyright bar ── */
    .tp-footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background: #1a1a1a;
        color: #aaaaaa;
        text-align: center;
        padding: 8px 0 7px 0;
        font-size: 0.78rem;
        letter-spacing: 0.03em;
        z-index: 9999;
        border-top: 2px solid #C62828;
        font-family: 'Segoe UI', Arial, sans-serif;
    }
    .tp-footer span.brand {
        color: #ffffff;
        font-weight: 600;
    }
    .tp-footer span.powered {
        color: #C62828;
        font-weight: 600;
    }

    /* push content up so footer doesn't overlap last widget */
    .block-container { padding-bottom: 60px !important; }
</style>
""", unsafe_allow_html=True)

# ─── HEADER: logo top-left, title beside it ──────────────────────────────────
logo_path = os.path.join(os.getcwd(), "logo.png")
if os.path.exists(logo_path):
    with open(logo_path, "rb") as img_file:
        logo_b64 = base64.b64encode(img_file.read()).decode()
    st.markdown(f"""
        <div style='position:relative; display:flex; align-items:center; justify-content:center;
                    padding:8px 0; border-bottom:2px solid #e0e0e0; margin-bottom:18px; min-height:80px;'>
            <img src='data:image/png;base64,{logo_b64}'
                 style='height:110px; width:auto; position:absolute; left:0; top:-20px;'/>
            <h1 style='margin:0; color:#C62828; font-size:2rem; text-align:center;'>
                Thermopads Test Certificate Generation and Validation
            </h1>
        </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
        <div style='position:relative; display:flex; align-items:center; justify-content:center;
                    padding:8px 0; border-bottom:2px solid #e0e0e0; margin-bottom:18px; min-height:80px;'>
            <div style='font-size:2rem; position:absolute; left:0; top:0;'>🛡️</div>
            <h1 style='margin:0; color:#C62828; font-size:2rem; text-align:center;'>
                Thermopads Test Certificate Generation and Validation
            </h1>
        </div>
    """, unsafe_allow_html=True)

import json

@st.cache_data(show_spinner=False)
def load_template_registry():
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        script_dir = os.getcwd()
    path = os.path.join(script_dir, "templates.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

TEMPLATE_REGISTRY = load_template_registry()

# Derived automatically from JSON — no manual maintenance
CUSTOMER_TEMPLATES = {"Schluter": []}
FOLDER_MAP = {}
for key, cfg in TEMPLATE_REGISTRY.items():
    cust = cfg.get("customer", "Unknown")
    CUSTOMER_TEMPLATES.setdefault(cust, []).append(key)
    FOLDER_MAP[cust] = cfg.get("folder", cust.upper())
# ─────────────────────────────────────────────────────────────────────────────
# CHANGE 1 — Load Schluter resistance limits from CSV
# Place schluter_limits.csv in the same folder as app.py.
# Columns required: ModelCode, MinResistance, MaxResistance
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_schluter_limits():
    """
    Reads schluter_limits.csv from the app directory.
    Returns a dict  { "DHEHK12011": {"min": 101.55, "max": 118.51}, ... }
    Falls back to an empty dict if the file is missing, and shows a warning.
    """
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        script_dir = os.getcwd()

    candidates = [
        os.path.join(script_dir, "schluter_limits.csv"),
        os.path.join(os.getcwd(), "schluter_limits.csv"),
    ]
    csv_path = next((p for p in candidates if os.path.exists(p)), None)

    if csv_path is None:
        st.warning(
            "⚠️ `schluter_limits.csv` not found next to app.py. "
            "All Schluter resistance values will be shown in BLACK (no range check)."
        )
        return {}

    try:
        df = pd.read_csv(csv_path, dtype=str)
        required = {"ModelCode", "MinResistance", "MaxResistance"}
        if not required.issubset(set(df.columns)):
            st.error(
                f"❌ `schluter_limits.csv` must have columns: {required}. "
                f"Found: {list(df.columns)}"
            )
            return {}

        limits = {}
        for _, row in df.iterrows():
            key = str(row["ModelCode"]).strip().upper()
            try:
                limits[key] = {
                    "min": float(row["MinResistance"]),
                    "max": float(row["MaxResistance"]),
                }
            except ValueError:
                pass   # skip rows with non-numeric values
        return limits

    except Exception as e:
        st.error(f"❌ Failed to read `schluter_limits.csv`: {e}")
        return {}


# ─── Helpers ─────────────────────────────────────────────────────────────────
def fuzzy_find_template(folder_path, target_filename):
    exact_path = os.path.join(folder_path, target_filename)
    if os.path.exists(exact_path):
        return exact_path
    if not os.path.isdir(folder_path):
        return None
    def clean(s):
        return re.sub(r'[\s\-_]', '', str(s)).lower()
    target_base = clean(os.path.splitext(target_filename)[0])
    for fname in os.listdir(folder_path):
        fbase, fext = os.path.splitext(fname)
        if fext.lower() not in ('.xlsx', '.xls'):
            continue
        if clean(fbase) == target_base:
            return os.path.join(folder_path, fname)
    for fname in os.listdir(folder_path):
        fbase, fext = os.path.splitext(fname)
        if fext.lower() not in ('.xlsx', '.xls'):
            continue
        cb = clean(fbase)
        if target_base in cb or cb in target_base:
            return os.path.join(folder_path, fname)
    return None

def get_template_folder(customer):
    sub = FOLDER_MAP.get(customer, customer.upper())
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        script_dir = os.getcwd()
    candidates = [
        os.path.join(script_dir, "TC", sub),
        os.path.join(script_dir, sub),
        os.path.join(os.getcwd(), "TC", sub),
        os.path.join(os.getcwd(), sub),
        os.path.join(os.getcwd(), "TC_templates", "TC", sub),
        os.getcwd(),
    ]
    for c in candidates:
        if os.path.isdir(c):
            return c
    return os.path.join(os.getcwd(), "TC", sub)

@st.cache_data(show_spinner=False)
def load_data_smart(file_bytes, fname):
    try:
        fname_l = fname.lower()
        if fname_l.endswith('.xlsx'):
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str, engine='openpyxl')
        elif fname_l.endswith('.xls'):
            try:
                df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str, engine='xlrd')
            except Exception:
                text = file_bytes.decode('utf-8', errors='ignore')
                df_raw = pd.read_csv(io.StringIO(text), header=None, sep=r'\t+', engine='python', dtype=str)
        else:
            text = file_bytes.decode('utf-8', errors='ignore')
            sep = r'\t+' if '\t' in text else ','
            df_raw = pd.read_csv(io.StringIO(text), header=None, sep=sep, engine='python', dtype=str)

        keywords = ['orderno', 'primarysrno', 'channelid', 'matdesc', 'heatingcable', 'sl.no', 'materialno',
                    'sno', 'm/cno', 'mcno', 'srno', 'serialno', 'cableno']
        header_row_index = 0
        for i, row in df_raw.head(100).iterrows():
            row_vals = [str(v).strip().lower().replace(" ", "").replace("\t", "") for v in row.values if v is not None]
            if any(k in row_vals for k in keywords):
                header_row_index = i
                break

        df = df_raw.iloc[header_row_index:].copy()
        df.columns = [str(c).strip() for c in df.iloc[0]]
        df = df.iloc[1:].reset_index(drop=True)
        df = df.loc[:, ~df.columns.str.contains(r'^Unnamed|^nan|^None|^$', na=False)]
        for col in df.columns:
            df[col] = df[col].astype(str).apply(lambda x: x.strip().replace('\t', '') if x != 'nan' else "")
        return df
    except Exception:
        return None

def normalize_qc_columns(df):
    """
    Supports both QC formats:

    Format 1:
        Channelid
        Actualminout
        ProductName
        Size

    Format 2:
        M/C NO
        CCR
        PN
        M/C Size
    """

    rename_map = {}

    for col in df.columns:
        col_clean = str(col).strip().upper()

        if col_clean == "M/C NO":
            rename_map[col] = "Channelid"

        elif col_clean == "CCR":
            rename_map[col] = "Actualminout"

        elif col_clean == "PN":
            rename_map[col] = "ProductName"

        elif col_clean == "M/C SIZE":
            rename_map[col] = "Size"

    return df.rename(columns=rename_map)


def get_col(df, keywords):
    for k in keywords:
        for col in df.columns:
            norm = col.lower().replace(" ", "").replace("_", "").replace("/", "")
            if k in norm:
                return col
    return None

def resolve_matdesc_col(df):
    if 'MatDesc' in df.columns:
        return 'MatDesc'
    if 'ProductName' in df.columns:
        return 'ProductName'
    if 'PN' in df.columns:
        return 'PN'
    for col in df.columns:
        norm = col.lower().replace(" ", "").replace("_", "")
        if 'matdesc' in norm:
            return col
        if 'productname' in norm:
            return col
        if norm == 'pn':
            return col
    return None

def extract_watts(desc):
    match = re.search(r'(\d+(?:\.\d+)?)\s*W', str(desc), re.IGNORECASE)
    return match.group(1) if match else ""

def find_footer_start(ws, anchor):
    footer_text_row = None
    for r in range(anchor, ws.max_row + 1):
        combined = ' '.join(str(ws.cell(row=r, column=c).value or '').upper() for c in range(1, 10))
        if any(x in combined for x in ("PREPARED", "CHECKED", "APPROVED", "FOR THERMOPADS", "THERMOPADS PVT")):
            footer_text_row = r
            break
    if footer_text_row is None:
        return ws.max_row + 1
    block_start = footer_text_row
    for r in range(footer_text_row - 1, anchor - 1, -1):
        all_blank = all(ws.cell(row=r, column=c).value is None for c in range(1, 10))
        if all_blank:
            block_start = r
        else:
            break
    return block_start

def write_item_name_row(ws, row_num, item_name, num_cols, is_cladswiss=False):
    thin = Side(style='thin', color="000000")
    label_text = f"Item Name : {item_name}"
    if is_cladswiss:
        label_text += "  (Nominal, -5%,+10%)"
    fill       = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    bold_font  = Font(bold=True, size=10, color="1F3864")
    left_align = Alignment(horizontal='left', vertical='center')
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in range(1, num_cols + 1):
        cell            = ws.cell(row=row_num, column=c)
        cell.value      = label_text if c == 1 else None
        cell.fill       = fill
        cell.font       = bold_font
        cell.alignment  = left_align
        cell.border     = border
    try:
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=num_cols)
    except Exception:
        pass

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE WATERMARK FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════
def inject_page_watermarks_with_breaks(xlsx_bytes):
    import io
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PILImage, ImageDraw, ImageFont

    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'portrait'
    ws.page_margins.top       = 1.0
    ws.page_margins.bottom    = 1.0
    ws.page_margins.left      = 0.75
    ws.page_margins.right     = 0.75
    ws.page_margins.header    = 0.3
    ws.page_margins.footer    = 0.5

    ws.oddFooter.center.text  = '&K808080&"Arial,Italic"&10Page &P of &N'
    ws.evenFooter.center.text = '&K808080&"Arial,Italic"&10Page &P of &N'
    ws.sheet_view.view = 'normal'

    row_heights = {}
    for r in range(1, ws.max_row + 1):
        row_heights[r] = ws.row_dimensions[r].height or 15.0

    data_start = 1
    for r in range(1, ws.max_row + 1):
        cell_val = str(ws.cell(row=r, column=1).value or '').strip()
        if cell_val == '1' or 'Item Name' in cell_val:
            data_start = r
            break

    page_height_pts = 648.0
    header_height   = sum(row_heights.get(r, 15.0) for r in range(1, data_start))

    running_height  = header_height
    page_break_rows = []

    for r in range(data_start, ws.max_row + 1):
        h = row_heights.get(r, 15.0)
        running_height += h
        if running_height > page_height_pts:
            page_break_rows.append(r - 1)
            running_height = header_height + h

    total_pages = len(page_break_rows) + 1
    boundaries  = [1] + [pb + 1 for pb in page_break_rows] + [ws.max_row + 1]

    def make_watermark(text):
        width, height = 420, 280
        base = PILImage.new('RGBA', (width, height), (255, 255, 255, 0))
        draw = ImageDraw.Draw(base)
        font = None
        for fp in [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "C:/Windows/Fonts/arial.ttf", "arial.ttf",
        ]:
            try:
                font = ImageFont.truetype(fp, 38)
                break
            except Exception:
                continue
        if font is None:
            font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text(((width - tw) // 2, (height - th) // 2), text, fill=(80, 80, 80, 110), font=font)
        rotated = base.rotate(30, expand=False, resample=PILImage.BICUBIC)
        buf = io.BytesIO()
        rotated.save(buf, format='PNG')
        buf.seek(0)
        return buf

    def get_mid_row(start_row, end_row):
        total_h = sum(row_heights.get(r, 15.0) for r in range(start_row, end_row + 1))
        target  = total_h * 0.45
        acc     = 0
        for r in range(start_row, end_row + 1):
            acc += row_heights.get(r, 15.0)
            if acc >= target:
                return r
        return end_row

    for pg in range(total_pages):
        p_start = boundaries[pg]
        p_end   = boundaries[pg + 1] - 1
        mid_row = get_mid_row(p_start, p_end)
        label   = f"Page {pg + 1} of {total_pages}"
        wm_buf  = make_watermark(label)
        img        = XLImage(wm_buf)
        img.width  = 290
        img.height = 170
        img.anchor = f'C{mid_row}'
        ws.add_image(img)

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# UNIVERSAL ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def generate_universal_engine(
    df_merged, template_key, o_date, s_date,
    sno_value, selected_products,
    customer_name_override, po_number,
    matdesc_col, log_p, log_h,
    template_folder,
):
    cfg = TEMPLATE_REGISTRY[template_key]
    template_path = fuzzy_find_template(template_folder, cfg["path"])
    if not template_path or not os.path.exists(template_path):
        st.error(f"❌ CRITICAL ERROR: Template file not found.\nLooking for: `{cfg['path']}` inside `{template_folder}`")
        return None

    log_h.append(f"Loaded template: {os.path.basename(template_path)}")
    log_p.code("\n".join(log_h[-8:]))

    id_col   = get_col(df_merged, ['primarysrno', 'primary', 'channel', 'cable', 'm/cno', 'mcno', 'sno', 'srno'])
    res_col  = get_col(df_merged, ['actualminout', 'ccr'])
    size_col = get_col(df_merged, ['matsize', 'm/csize', 'mcsize', 'size'])
    if size_col and 'batch' in size_col.lower():
        size_col = None

    try:
        wb = load_workbook(template_path)
        ws = wb.active
    except Exception as e:
        st.error(f"❌ Failed to open template file:\n{os.path.basename(template_path)}")
        st.error(str(e))
        return None

    ws.sheet_view.view = 'normal'

    thin         = Side(style='thin', color="000000")
    table_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font_black   = Font(color="000000", size=10)
    center_align = Alignment(horizontal='center', vertical='center')

    anchor   = cfg["anchor"]
    num_cols = 7
    for c in range(1, 15):
        if ws.cell(row=anchor - 1, column=c).value is not None:
            num_cols = c
    num_cols = max(num_cols, 7)

    def _set_cell_preserving_prefix(ws, row, col, new_value, fallback_prefix, pattern):
        cell    = ws.cell(row=row, column=col)
        old_val = str(cell.value or '')
        m       = re.match(pattern, old_val, re.IGNORECASE)
        prefix  = m.group(0) if m else fallback_prefix
        cell.value = f"{prefix}{new_value}"

    _set_cell_preserving_prefix(ws, cfg["cust_row"], cfg["cust_col"], customer_name_override, "Customer : ", r'^(CUSTOMER\s*:?\s*|Customer\s*:?\s*)')
    _set_cell_preserving_prefix(ws, cfg["po_row"],   cfg["cust_col"], po_number,               "PO No : ",   r'^(PO\s*No\.?\s*:?\s*)')
    _set_cell_preserving_prefix(ws, cfg["podt_row"], cfg["cust_col"], o_date,                  "PO Dt : ",   r'^(PO\s*Dt\.?\s*:?\s*)')
    _set_cell_preserving_prefix(ws, cfg["sno_row"],  cfg["sno_col"],  sno_value,               "S.NO : ",    r'^(S\.?\s*No\.?\s*:?\s*)')
    _set_cell_preserving_prefix(ws, cfg["date_row"], cfg["date_col"], s_date,                  "DATE : ",    r'^(DATE\s*:?\s*)')


    footer_block_start = find_footer_start(ws, anchor)
    to_unmerge = [str(mc) for mc in list(ws.merged_cells.ranges) if mc.min_row >= footer_block_start]
    for ref in to_unmerge:
        ws.unmerge_cells(ref)

    footer_snapshot = []
    for r in range(footer_block_start, ws.max_row + 1):
        row_cells = []
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=c)
            try:
                rgb = cell.font.color.rgb if cell.font and cell.font.color and cell.font.color.type == 'rgb' else '000000'
            except Exception:
                rgb = '000000'
            row_cells.append({
                'value':     cell.value,
                'font':      Font(color=rgb, size=cell.font.size or 10, bold=bool(cell.font.bold)) if cell.font else None,
                'alignment': Alignment(wrap_text=True, vertical='center', horizontal=(cell.alignment.horizontal or 'left')) if cell.alignment else None,
            })
        footer_snapshot.append(row_cells)

    data_merges = [str(mc) for mc in list(ws.merged_cells.ranges) if anchor <= mc.min_row < footer_block_start]
    for ref in data_merges:
        ws.unmerge_cells(ref)

    rows_to_delete = ws.max_row - anchor + 1
    if rows_to_delete > 0:
        ws.delete_rows(anchor, rows_to_delete)

    current_row = anchor
    for prod in selected_products:
        df_prod = df_merged[df_merged[matdesc_col] == prod].copy()
        if id_col and id_col in df_prod.columns:
            df_prod = df_prod.sort_values(by=id_col).reset_index(drop=True)

        ts = datetime.datetime.now().strftime('%H:%M:%S')
        log_h.append(f"[{ts}] Product: {prod} — {len(df_prod)} rows")
        log_p.code("\n".join(log_h[-8:]))

        write_item_name_row(ws, current_row, prod, num_cols, is_cladswiss=cfg["is_cladswiss"])
        current_row += 1

        local_sno = 1
        for _, row in df_prod.iterrows():
            id_val      = row.get(id_col,      '') if id_col   else ''
            sz_val      = str(row.get(size_col, '')) if size_col else ''
            res_val     = row.get(res_col,      '') if res_col  else ''
            matdesc_val = row.get(matdesc_col,  '')
            watts       = extract_watts(matdesc_val)

            ws.cell(row=current_row, column=1, value=local_sno).border     = table_border
            ws.cell(row=current_row, column=2, value=id_val).border         = table_border
            ws.cell(row=current_row, column=3, value=sz_val).border         = table_border
            ws.cell(row=current_row, column=4, value=watts or "").border    = table_border
            ws.cell(row=current_row, column=5, value=res_val).border        = table_border
            ws.cell(row=current_row, column=6, value="> 500").border        = table_border
            ws.cell(row=current_row, column=7, value="No Breakdown").border = table_border

            for c in range(1, 8):
                ws.cell(row=current_row, column=c).alignment = center_align
                ws.cell(row=current_row, column=c).font       = font_black
            current_row += 1
            local_sno  += 1

        for c in range(1, num_cols + 1):
            ws.cell(row=current_row, column=c).value  = None
            ws.cell(row=current_row, column=c).border = Border()
        current_row += 1

    footer_write_start = current_row
    for r_off, row_cells in enumerate(footer_snapshot):
        wr = footer_write_start + r_off
        for c_idx, cd in enumerate(row_cells, start=1):
            cell = ws.cell(row=wr, column=c_idx)
            cell.value = cd['value']
            if cd['font']:       cell.font      = cd['font']
            if cd['alignment']:  cell.alignment = cd['alignment']

    return wb


# ─────────────────────────────────────────────────────────────────────────────
# ENGINE: SCHLUTER
# ─────────────────────────────────────────────────────────────────────────────
from openpyxl.formatting.formatting import ConditionalFormattingList
def generate_schluter_engine(merged, o_date, s_date, log_p, log_h):
    """
    CHANGES vs old version
    ──────────────────────
    1. SCHLUTER_LIMITS is now loaded from schluter_limits.csv via load_schluter_limits().
    2. Range check + colour assignment are done in ONE place:
         actual_in_range = (min_limit < truncated_val < max_limit)
       Equal-to-limit → NOT in range → RED.
       Strictly between limits → in range → BLACK.
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, Border, Side

    # ── CHANGE 1: load limits from CSV ───────────────────────────────────────
    SCHLUTER_LIMITS = load_schluter_limits()

    template_path = "schluter_template.xlsx"
    if not os.path.exists(template_path):
        st.error(f"CRITICAL: '{template_path}' missing from app folder!")
        return None, None

    def schluter_sort_key(model_str):
        model_str = str(model_str).strip()
        voltage = 0 if '120' in model_str else (1 if '240' in model_str else 2)
        m = re.search(r'(\d+)$', model_str)
        suffix = int(m.group(1)) if m else 0
        return (voltage, suffix, model_str)

    try:
        wb = load_workbook(template_path)
        ws = wb.active
    except Exception as e:
        st.error("❌ Failed to open Schluter template.")
        st.error(str(e))
        return None, None

    ws.sheet_view.view = 'normal'

    side   = Side(style='thin')
    thin_b = Border(left=side, right=side, top=side, bottom=side)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    mod_col    = 'CustomerCode' if 'CustomerCode' in merged.columns else 'MaterialNo'
    cable_cols = [c for c in merged.columns if any(
        x in c.lower().replace(' ', '').replace('/', '')
        for x in ['primary', 'cable', 'channel', 'mcno', 'sno', 'srno']
    )]
    cable_col = cable_cols[0] if cable_cols else merged.columns[1]

    actual_col = None
    for c in merged.columns:
        norm = c.lower().replace(' ', '').replace('_', '')
        if 'actualminout' in norm or 'ccr' in norm:
            actual_col = c
            break

    ord_no    = str(merged.iloc[0].get('OrderNo', 'N/A'))
    ws['A8']  = f"Schluter's purchase order no.{ord_no}"
    ws['A9']  = f"                        Dt : {o_date}"
    ws['A10'] = f"Date of shipment : {s_date}"

    merged = merged.copy()
    merged[mod_col] = merged[mod_col].astype(str).str.strip().str.upper()
    merged['_sk']   = merged[mod_col].apply(schluter_sort_key)
    merged          = merged.sort_values(by=['_sk', cable_col]).reset_index(drop=True)

    ANCHOR     = 16
    footer_row = None
    for r in range(ANCHOR, ws.max_row + 1):
        for c in range(1, 12):
            v = str(ws.cell(row=r, column=c).value or '')
            if 'However' in v or 'Actual Measured' in v or 'APPROVED' in v:
                footer_row = r
                break
        if footer_row:
            break

    footer_block_start = footer_row or (ws.max_row + 1)
    for r in range(footer_block_start - 1, ANCHOR, -1):
        if all(ws.cell(row=r, column=c).value is None for c in range(1, 12)):
            footer_block_start = r
        else:
            break

    for ref in [str(mc) for mc in list(ws.merged_cells.ranges) if mc.min_row >= footer_block_start]:
        ws.unmerge_cells(ref)

    footer_snapshot = []
    for r in range(footer_block_start, ws.max_row + 1):
        row_cells = []
        for c in range(1, 12):
            cell = ws.cell(row=r, column=c)
            row_cells.append({
                'value':   cell.value,
                'bold':    bool(cell.font.bold) if cell.font else False,
                'size':    cell.font.size or 10 if cell.font else 10,
                'h_align': cell.alignment.horizontal or 'left' if cell.alignment else 'left',
            })
        footer_snapshot.append(row_cells)

    ws.delete_rows(ANCHOR, ws.max_row - ANCHOR + 1)

    # ── CRITICAL FIX: Remove ALL conditional formatting from the sheet ──
    # The template has CF rules on col F that override our font colors in Excel.
    # We control colors manually via font color, so CF must be wiped completely.
    ws.conditional_formatting = ConditionalFormattingList()

    seen, models_ordered = set(), []
    for m in merged[mod_col]:
        if m not in seen:
            models_ordered.append(m)
            seen.add(m)

    for model in models_ordered:
        if model.strip().upper() not in SCHLUTER_LIMITS:
            st.warning(
                f"⚠️ Model `{model}` is not in schluter_limits.csv — "
                f"resistance values will be shown in BLACK (no range check applied)."
            )

    current_row = ANCHOR
    for model in models_ordered:
        df_model  = merged[merged[mod_col] == model].sort_values(by=cable_col)
        model_key = model.strip().upper()
        limits    = SCHLUTER_LIMITS.get(model_key, {"min": None, "max": None})

        for local_sno, (_, row) in enumerate(df_model.iterrows(), start=1):
            raw = str(row.get(actual_col, '') if actual_col else '').strip()

            # ── Parse value + decide colour ──────────────────────────────
            from decimal import Decimal, ROUND_DOWN, InvalidOperation

            min_limit = limits.get("min")
            max_limit = limits.get("max")

            try:
                # Use original value for range comparison

                actual_val = Decimal(str(raw))

                # Truncated value only for display
                display_val = actual_val.quantize(
                Decimal("0.01"),
                rounding=ROUND_DOWN
                )

                truncated_val = float(display_val)

                if min_limit is not None and max_limit is not None:
                    d_min = Decimal(str(min_limit))
                    d_max = Decimal(str(max_limit))

                    # Compare ORIGINAL value
                    actual_in_range = (actual_val > d_min) and (actual_val < d_max)
                else:
                    actual_in_range = True

                font_color = 'FF000000' if actual_in_range else 'FFFF0000'

            except (InvalidOperation, ValueError, Exception):
                # Non-numeric → show as-is in black
                truncated_val = raw
                actual_in_range = True
                font_color = 'FF000000'
            # ── end colour block ──────────────────────────────────────────

            row_vals = {
                1:  local_sno,
                2:  str(row[cable_col]),
                3:  model,
                4:  limits['min'] if limits['min'] is not None else '',
                5:  limits['max'] if limits['max'] is not None else '',
                6:  truncated_val,
                7:  'PASS',
                8:  'PASS',
                9:  '',
                10: '> 5 Gohms',
                11: 'PASS',
            }

            for c, val in row_vals.items():
                cell           = ws.cell(row=current_row, column=c)
                cell.value     = val
                cell.border    = thin_b
                cell.alignment = center

                if c == 6:
                    # Apply the colour decided above — single source of truth
                    cell.number_format = '0.00'
                    cell.font = Font(name='Arial', size=10, bold=False, color=font_color)
                    log_h.append(
                        f"ROW={current_row} MODEL={model} "
                        f"ACTUAL={truncated_val if isinstance(truncated_val, str) else f'{truncated_val:.2f}'} "
                        f"IN_RANGE={actual_in_range} COLOR={'BLACK' if actual_in_range else 'RED'}"
                    )
                else:
                    cell.font = Font(name='Arial', size=10, bold=False, color='FF000000')

            log_p.code("\n".join(log_h[-12:]))
            current_row += 1

        # blank separator row between models
        for c in range(1, 12):
            ws.cell(row=current_row, column=c).value  = None
            ws.cell(row=current_row, column=c).border = Border()
        current_row += 1

    footer_write_start = current_row
    for r_off, row_cells in enumerate(footer_snapshot):
        wr = footer_write_start + r_off
        for c_idx, cd in enumerate(row_cells, start=1):
            cell           = ws.cell(row=wr, column=c_idx)
            cell.value     = cd['value']
            cell.font      = Font(color='FF000000', size=cd['size'], bold=cd['bold'])
            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal=cd['h_align'])

    for off_start, off_end in [(2, 3), (4, 5), (6, 7)]:
        r1 = footer_write_start + off_start
        r2 = footer_write_start + off_end
        ws.merge_cells(f'B{r1}:J{r2}')
        ws.cell(row=r1, column=2).alignment = Alignment(wrap_text=True, vertical='center')

    return wb, ord_no


# ─────────────────────────────────────────────────────────────────────────────
# MAIN UI
# ─────────────────────────────────────────────────────────────────────────────
customer_choice = st.selectbox(" Select Customer", ["Warmup", "Warmly", "Schluter", "Cenika", "CladSwiss"])

c1, c2 = st.columns(2)
with c1:
    ui_o_date = st.text_input(" PO Dt (Order Date)",   value=datetime.datetime.now().strftime("%d.%m.%Y"))
with c2:
    ui_s_date = st.text_input(" DATE (Shipment Date)", value=datetime.datetime.now().strftime("%d.%m.%Y"))

selected_template_key = None
if customer_choice != "Schluter":
    avail_templates = CUSTOMER_TEMPLATES.get(customer_choice, [])
    if avail_templates:
        selected_template_key = st.selectbox(" Select Template", avail_templates)
        cfg_preview = TEMPLATE_REGISTRY.get(selected_template_key, {})
        st.caption(f"📄 Template file: `{cfg_preview.get('path', '')}`")
    else:
        st.warning(f"No templates registered for {customer_choice}.")

ui_sno = ""
if customer_choice != "Schluter" and selected_template_key:
    cfg_prev        = TEMPLATE_REGISTRY[selected_template_key]
    template_folder = get_template_folder(customer_choice)
    tpl_path        = fuzzy_find_template(template_folder, cfg_prev["path"])
    default_sno = ""
    if tpl_path and os.path.exists(tpl_path):
        try:
            _wb  = load_workbook(tpl_path, data_only=True)
            _ws  = _wb.active
            _raw = str(_ws.cell(row=cfg_prev["sno_row"], column=cfg_prev["sno_col"]).value or '')
            _m   = re.search(r'S\.?\s*No\.?\s*:?\s*(.*)', _raw, re.IGNORECASE)
            default_sno = _m.group(1).strip() if _m else _raw.strip()
        except Exception:
            st.warning(f"⚠️ Could not read template details from {os.path.basename(tpl_path)}")
    ui_sno = st.text_input(" S.NO (Certificate Serial Number)", value=default_sno)

# ─── FILE UPLOADERS ───────────────────────────────────────────────────────────
qc_file_1 = st.file_uploader(" QC Test Report", type=["xlsx", "xls", "csv"])

has_second_qc = st.radio("Was testing performed on multiple workstations?", ["Yes", "No"], index=1, horizontal=True)

qc_file_2 = None
if has_second_qc == "Yes":
    qc_file_2 = st.file_uploader(" QC Test Report 2", type=["xlsx", "xls", "csv"])

pk_file = st.file_uploader(" Packing Data", type=["xlsx", "xls", "csv"])

if qc_file_1 and pk_file:

    # -------------------------
    # QC FILE 1
    # -------------------------
    df_qc1 = load_data_smart(
        qc_file_1.getvalue(),
        qc_file_1.name
    )

    if df_qc1 is None:
        st.error("❌ Could not read QC File 1")
        st.stop()

    df_qc1 = normalize_qc_columns(df_qc1)

    # -------------------------
    # PACKING FILE
    # -------------------------
    df_pk = load_data_smart(
        pk_file.getvalue(),
        pk_file.name
    )

    if df_pk is None:
        st.error("❌ Could not read Packing File")
        st.stop()

    # -------------------------
    # STEP 1: STANDARDIZE SERIAL COL NAME
    # so both QC files have 'Channelid' before appending
    # -------------------------
    def standardize_serial_col(df):
        if 'Channelid' in df.columns:
            return df
        for col in df.columns:
            norm = col.lower().replace(" ", "").replace("_", "").replace("/", "")
            if any(k in norm for k in ['mcno', 'channel', 'cable', 'primary', 'srno']):
                return df.rename(columns={col: 'Channelid'})
        return df

    df_qc1 = standardize_serial_col(df_qc1)

    # -------------------------
    # STEP 2: LOAD + APPEND QC2 IF PROVIDED
    # -------------------------
    if has_second_qc == "Yes" and qc_file_2:
        df_qc2 = load_data_smart(qc_file_2.getvalue(), qc_file_2.name)
        if df_qc2 is not None:
            df_qc2 = normalize_qc_columns(df_qc2)
            df_qc2 = standardize_serial_col(df_qc2)
            df_qc  = pd.concat([df_qc1, df_qc2], ignore_index=True)
            st.info(f"Appended QC data: {len(df_qc1)} + {len(df_qc2)} = {len(df_qc)} rows")
        else:
            st.warning("⚠️ QC File 2 uploaded but could not be read.")
            df_qc = df_qc1.copy()
    else:
        df_qc = df_qc1.copy()

    # -------------------------
    # STEP 3: MERGE APPENDED QC WITH PACKING
    # -------------------------
    p_col = get_col(df_pk, ['primary', 'cable', 'channel'])

    if 'Channelid' not in df_qc.columns:
        st.error("❌ QC file missing Channel / M/C Number column.")
        st.stop()

    if not p_col:
        st.error("❌ Packing file missing Primary Serial Number column.")
        st.stop()

    df_qc = df_qc.copy()
    df_pk = df_pk.copy()
    df_qc['Channelid'] = df_qc['Channelid'].astype(str).str.replace(r"\s+", "", regex=True)
    df_pk[p_col]       = df_pk[p_col].astype(str).str.replace(r"\s+", "", regex=True)

    merged = df_qc.merge(df_pk, left_on='Channelid', right_on=p_col, how='inner')

    if merged.empty:
            st.error(
            "⚠️ No matching serial numbers found between QC and Packing files."
            )
            st.stop()
   

    if not merged.empty:
            if customer_choice != "Schluter" and selected_template_key:
                matdesc_col = resolve_matdesc_col(merged)
                cust_col_m  = None
                for col in merged.columns:
                    norm = col.lower().replace(" ", "").replace("_", "")
                    if 'customername' in norm and col.endswith('_y'):
                        cust_col_m = col
                        break
                if cust_col_m is None:
                    cust_col_m = get_col(merged, ['customername', 'customer'])

                po_col_m     = get_col(merged, ['orderno', 'ponumber', 'po'])
                sample0      = merged.iloc[0]
                default_cust = str(sample0.get(cust_col_m, '')).strip() if cust_col_m else ""
                default_po   = str(sample0.get(po_col_m,    '')).strip() if po_col_m   else ""

                cc1, cc2 = st.columns(2)
                with cc1:
                    ui_customer = st.text_input(" Enter Customer Name for the test certificate ", value=default_cust)
                with cc2:
                    ui_po = st.text_input(" Enter PO Number for the test certificate", value=default_po)

                if matdesc_col:
                    p_options = sorted(merged[matdesc_col].dropna().unique().tolist())
                else:
                    st.error("⚠️ Could not find MatDesc column in uploaded data.")
                    p_options = []

                sa_col, cb_col, _ = st.columns([1, 1, 8])
                with sa_col:
                    if st.button("✅ All", key="sel_all"):  st.session_state['ms_selected'] = p_options
                with cb_col:
                    if st.button("❌ Clear", key="clr_all"): st.session_state['ms_selected'] = []

                default_sel = st.session_state.get('ms_selected', [])
                default_sel = [x for x in default_sel if x in p_options]

                selected = st.multiselect(" Select Product Batches", options=p_options, default=default_sel)
                st.session_state['ms_selected'] = selected

                if selected and st.button(" GENERATE TEST CERTIFICATE"):
                    st.session_state.generated_files = []
                    wb = None

                    with st.expander(" LIVE PROCESSING LOGS", expanded=True):
                        log_p = st.empty()
                        log_h = [f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Starting {customer_choice}..."]
                        try:
                            tf = get_template_folder(customer_choice)
                            wb = generate_universal_engine(
                                df_merged=merged, template_key=selected_template_key,
                                o_date=ui_o_date, s_date=ui_s_date, sno_value=ui_sno,
                                selected_products=selected, customer_name_override=ui_customer,
                                po_number=ui_po, matdesc_col=matdesc_col,
                                log_p=log_p, log_h=log_h, template_folder=tf
                            )
                        except Exception as e:
                            import traceback
                            st.error(f"❌ ERROR: {e}\n\n{traceback.format_exc()}")

                    if wb:
                        buf = io.BytesIO()
                        wb.save(buf)
                        processed_bytes = inject_page_watermarks_with_breaks(buf.getvalue())
                        st.session_state.generated_files.append({
                            "name": f"{selected_template_key}.xlsx",
                            "data": processed_bytes,
                        })
                    st.rerun()

            else:
                if st.button("GENERATE SCHLUTER CERTIFICATE"):
                    st.session_state.generated_files = []
                    with st.expander(" LIVE LOGS", expanded=True):
                        log_p = st.empty()
                        log_h = [f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Starting Schluter..."]
                        wb, oid = generate_schluter_engine(merged, ui_o_date, ui_s_date, log_p, log_h)
                        if wb:
                            buf = io.BytesIO()
                            wb.save(buf)
                            processed_bytes = inject_page_watermarks_with_breaks(buf.getvalue())
                            st.session_state.generated_files.append({
                                "name": f"COA_{oid}.xlsx",
                                "data": processed_bytes,
                            })
                            st.rerun()
    else:
        st.error("⚠️ No matching serial numbers found between QC Report and Packing Data.")
        st.info("Please verify that both files contain the same cable/serial numbers.")

# ─── DOWNLOAD SECTION ─────────────────────────────────────────────────────────
if st.session_state.generated_files:
    st.write("---")
    st.success(f"✅ {len(st.session_state.generated_files)} certificate(s) ready!")
    for i, f in enumerate(st.session_state.generated_files):
        st.download_button(
            label=f" DOWNLOAD: {f['name']}",
            data=f['data'],
            file_name=f['name'],
            key=f"btn_{i}"
        )
    if st.button("START NEW BATCH"):
        st.session_state.generated_files = []
        st.session_state['ms_selected']  = []
        st.rerun()

# ─── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="tp-footer">
    &copy; 2026 <span class="brand">Thermopads Pvt. Ltd.</span>
    &nbsp;|&nbsp; All Rights Reserved
    &nbsp;|&nbsp; Powered by <span class="powered">AIBurst.Biz</span>
</div>
""", unsafe_allow_html=True)