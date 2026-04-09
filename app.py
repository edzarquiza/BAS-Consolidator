import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import io
import copy
from datetime import datetime

st.set_page_config(
    page_title="Dexterous · BAS Consolidator",
    page_icon="📋",
    layout="centered"
)

# ── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

  .stApp { background: #f8f9fb; }

  /* Header */
  .dex-header {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    border-radius: 12px;
    padding: 28px 32px 22px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 18px;
  }
  .dex-header-icon { font-size: 2.4rem; }
  .dex-header-text h1 {
    color: #ffffff;
    font-size: 1.55rem;
    font-weight: 700;
    margin: 0 0 4px;
  }
  .dex-header-text p {
    color: #a0aec0;
    font-size: 0.88rem;
    margin: 0;
  }

  /* Cards */
  .card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 22px 24px;
    margin-bottom: 20px;
  }
  .card-title {
    font-size: 0.78rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #718096;
    margin-bottom: 14px;
  }

  /* Section badges */
  .badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 0.72rem;
    font-weight: 600;
    margin-left: 8px;
  }
  .badge-required { background: #fed7d7; color: #c53030; }
  .badge-optional { background: #e9d8fd; color: #6b46c1; }
  .badge-cash     { background: #bee3f8; color: #2b6cb0; }

  /* Upload zones */
  [data-testid="stFileUploaderDropzone"] {
    border: 2px dashed #cbd5e0 !important;
    border-radius: 8px !important;
    background: #f7fafc !important;
  }
  [data-testid="stFileUploaderDropzone"]:hover {
    border-color: #667eea !important;
    background: #ebf4ff !important;
  }

  /* Generate button */
  .stButton > button {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
    color: white !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 14px 36px !important;
    font-size: 1rem !important;
    width: 100%;
    transition: opacity 0.2s;
  }
  .stButton > button:hover { opacity: 0.88; }

  /* Download button */
  [data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #38a169 0%, #2f855a 100%) !important;
    color: white !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 14px 36px !important;
    font-size: 1rem !important;
    width: 100%;
  }

  /* Status pills */
  .pill-ok  { background:#c6f6d5; color:#22543d; padding:3px 12px; border-radius:20px; font-size:0.8rem; font-weight:600; }
  .pill-err { background:#fed7d7; color:#742a2a; padding:3px 12px; border-radius:20px; font-size:0.8rem; font-weight:600; }

  /* Footer */
  .footer {
    text-align: center;
    color: #a0aec0;
    font-size: 0.78rem;
    margin-top: 40px;
    padding-top: 18px;
    border-top: 1px solid #e2e8f0;
  }

  /* Hide Streamlit branding */
  #MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="dex-header">
  <div class="dex-header-icon">📋</div>
  <div class="dex-header-text">
    <h1>BAS Consolidator</h1>
    <p>dexterous · Merge Xero exports into a single BAS workbook</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Month options ───────────────────────────────────────────────────────────────
MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
CURRENT_YEAR = datetime.now().year
YEARS = [str(y) for y in range(CURRENT_YEAR - 2, CURRENT_YEAR + 3)]

# ── Helper: detect sheet type from first-row content ───────────────────────────
def detect_report_type(wb: openpyxl.Workbook) -> str | None:
    """Return one of: activity_statement | balance_sheet | profit_loss | payroll | ar | ap | unknown"""
    for sn in wb.sheetnames:
        ws = wb[sn]
        first_vals = []
        for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
            for cell in row:
                if cell:
                    first_vals.append(str(cell).strip().lower())
        combined = " ".join(first_vals)
        if "activity statement" in combined:
            return "activity_statement"
        if "transactions by tax rate" in combined:
            return "activity_statement"
        if "balance sheet" in combined:
            return "balance_sheet"
        if "profit and loss" in combined or "profit & loss" in combined:
            return "profit_loss"
        if "payroll activity" in combined:
            return "payroll"
        if "aged receivables" in combined:
            return "ar"
        if "aged payables" in combined:
            return "ap"
    return "unknown"

def sheet_names_in(wb):
    return [s.lower() for s in wb.sheetnames]

def find_sheet(wb, candidates):
    """Find first sheet matching any candidate name (case-insensitive)."""
    for sn in wb.sheetnames:
        if sn.lower() in [c.lower() for c in candidates]:
            return wb[sn]
    return None

def copy_sheet_data(src_ws, dst_ws):
    """Copy all cell values and basic formatting from src to dst."""
    from openpyxl.cell.cell import MergedCell

    # Clear destination first (skip merged cells)
    for row in dst_ws.iter_rows():
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None

    # Unmerge destination before writing
    for merged in list(dst_ws.merged_cells.ranges):
        dst_ws.unmerge_cells(str(merged))

    for row in src_ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            dst_cell = dst_ws.cell(row=cell.row, column=cell.column)
            dst_cell.value = cell.value
            if cell.has_style:
                try:
                    dst_cell.font = copy.copy(cell.font)
                    dst_cell.fill = copy.copy(cell.fill)
                    dst_cell.border = copy.copy(cell.border)
                    dst_cell.alignment = copy.copy(cell.alignment)
                    dst_cell.number_format = cell.number_format
                except Exception:
                    pass

    # Column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        try:
            dst_ws.column_dimensions[col_letter].width = dim.width
        except Exception:
            pass

    # Row heights
    for row_idx, dim in src_ws.row_dimensions.items():
        try:
            dst_ws.row_dimensions[row_idx].height = dim.height
        except Exception:
            pass

    # Merged cells
    for merged in list(src_ws.merged_cells.ranges):
        try:
            dst_ws.merge_cells(str(merged))
        except Exception:
            pass

def build_output(
    client_name: str,
    month: str,
    year: str,
    accounting_method: str,
    payg: str,
    frequency: str,
    activity_wb,
    balance_wb,
    pl_wb,
    payroll_wb,
    ar_wb,
    ap_wb,
    template_accrual_bytes: bytes,
    template_cash_bytes: bytes,
) -> bytes:

    # Load the right template
    template_bytes = template_cash_bytes if accounting_method == "Cash Basis" else template_accrual_bytes
    out_wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)

    # ── Helper to get/create destination sheet ─────────────────────────────────
    def get_dst(preferred_name, fallbacks=None):
        names = [preferred_name] + (fallbacks or [])
        for n in names:
            for sn in out_wb.sheetnames:
                if sn.strip().lower() == n.strip().lower():
                    return out_wb[sn]
        # Create if not found
        return out_wb.create_sheet(preferred_name)

    # ── 1. Activity Statement → 3 sheets ──────────────────────────────────────
    if activity_wb:
        # Identify the 3 report sheets by their first-row header
        sheet_map = {}
        for sn in activity_wb.sheetnames:
            ws = activity_wb[sn]
            header = ""
            for row in ws.iter_rows(min_row=1, max_row=2, values_only=True):
                for cell in row:
                    if cell:
                        header += str(cell).lower() + " "
            if "activity statement" in header:
                sheet_map["gst_summary"] = ws
            elif "transactions by tax rate" in header:
                sheet_map["gst_detail"] = ws
            elif "transactions by bas field" in header:
                sheet_map["bas_field"] = ws

        if "gst_summary" in sheet_map:
            dst = get_dst("GST Summary")
            copy_sheet_data(sheet_map["gst_summary"], dst)

        if "gst_detail" in sheet_map:
            dst = get_dst("GST Detail")
            copy_sheet_data(sheet_map["gst_detail"], dst)

        if "bas_field" in sheet_map:
            dst = get_dst("BAS field", ["BAS Field"])
            copy_sheet_data(sheet_map["bas_field"], dst)

    # ── 2. Balance Sheet ───────────────────────────────────────────────────────
    if balance_wb:
        src = balance_wb.active
        dst = get_dst("BS", ["BS "])
        copy_sheet_data(src, dst)

    # ── 3. Profit & Loss ───────────────────────────────────────────────────────
    if pl_wb:
        src = pl_wb.active
        dst = get_dst("PL", ["P&L", "P&L "])
        copy_sheet_data(src, dst)

    # ── 4. Payroll ─────────────────────────────────────────────────────────────
    if payroll_wb:
        src = payroll_wb.active
        dst = get_dst("PAYROLL", ["Payroll Activity Smry"])
        copy_sheet_data(src, dst)

    # ── 5. AR (Cash only) ──────────────────────────────────────────────────────
    if ar_wb and accounting_method == "Cash Basis":
        src = ar_wb.active
        dst = get_dst("AR")
        copy_sheet_data(src, dst)

    # ── 6. AP (Cash only) ──────────────────────────────────────────────────────
    if ap_wb and accounting_method == "Cash Basis":
        src = ap_wb.active
        dst = get_dst("AP")
        copy_sheet_data(src, dst)

    # ── 7. Queries sheet ───────────────────────────────────────────────────────
    # Build period code e.g. DEC2025_BAS or DEC2025_BAS Qtr
    period_tag = f"{month}{year}_BAS"
    file_name_val = f"{month}{year}_BAS {client_name}"
    freq_suffix = " Qtr" if frequency == "Quarterly" else ""
    display_period = f"{month}{year}_BAS{freq_suffix}"

    queries_ws = None
    for sn in out_wb.sheetnames:
        if sn.lower() == "queries":
            queries_ws = out_wb[sn]
            break

    if queries_ws is None:
        queries_ws = out_wb.create_sheet("Queries", 0)

    # Set standard Queries fields
    queries_ws["A1"] = "Client Name"
    queries_ws["B1"] = client_name
    queries_ws["D1"] = "Period"
    queries_ws["E1"] = display_period

    queries_ws["A2"] = "Accounting Method"
    queries_ws["B2"] = accounting_method
    queries_ws["D2"] = "Completed by: "

    queries_ws["A3"] = "PAYG"
    queries_ws["B3"] = payg

    queries_ws["A4"] = "File Name"
    queries_ws["B4"] = file_name_val

    queries_ws["A5"] = "Note"
    queries_ws["A6"] = "Email sent for queries and confrimation"
    queries_ws["A7"] = "Subject reference"
    queries_ws["A8"] = file_name_val

    # ── Return bytes ───────────────────────────────────────────────────────────
    buf = io.BytesIO()
    out_wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────

# ── Step 1: Client & Period ───────────────────────────────────────────────────
st.markdown('<div class="card"><div class="card-title">① Client Details</div>', unsafe_allow_html=True)

col1, col2 = st.columns([2, 1])
with col1:
    client_name_input = st.text_input("Client Name", placeholder="e.g. UPSTAGE WORLD PTY LTD")
with col2:
    frequency = st.selectbox("Frequency", ["Quarterly", "Monthly"])

col3, col4, col5 = st.columns(3)
with col3:
    month_sel = st.selectbox("Month", MONTHS, index=MONTHS.index("DEC"))
with col4:
    year_sel = st.selectbox("Year", YEARS, index=YEARS.index(str(CURRENT_YEAR)))
with col5:
    accounting_method = st.selectbox("Accounting Method", ["Cash Basis", "Accrual Basis"])

payg_opts = ["Monthly", "Quarterly", "No Payroll"]
payg = st.selectbox("PAYG Instalment", payg_opts)

# Auto-fill from tracker
st.markdown("**Or auto-fill from Tracker file**", help="Upload your BAS Monthly Automation Tracker to auto-populate client settings")
tracker_file = st.file_uploader("BAS Monthly Automation Tracker (.xlsx)", type=["xlsx"], key="tracker")

if tracker_file and client_name_input:
    try:
        tracker_wb = load_workbook(io.BytesIO(tracker_file.read()), data_only=True)
        ws = tracker_wb["BAS"] if "BAS" in tracker_wb.sheetnames else tracker_wb.active
        headers = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]

        def col_idx(name):
            for i, h in enumerate(headers):
                if name.lower() in h.lower():
                    return i
            return None

        name_col = col_idx("XeroClientName") or col_idx("ClientName") or 0
        freq_col = col_idx("Frequency")
        payg_col = col_idx("PAYGI")
        acct_col = col_idx("GSTAccountingMethod") or col_idx("Accounting")

        for row in ws.iter_rows(min_row=2, values_only=True):
            cell_name = str(row[name_col]).strip().upper() if row[name_col] else ""
            if client_name_input.strip().upper() in cell_name or cell_name in client_name_input.strip().upper():
                matched_freq = str(row[freq_col]).strip() if freq_col is not None and row[freq_col] else None
                matched_payg = str(row[payg_col]).strip() if payg_col is not None and row[payg_col] else None
                matched_acct = str(row[acct_col]).strip() if acct_col is not None and row[acct_col] else None
                info_parts = []
                if matched_freq: info_parts.append(f"Frequency: **{matched_freq}**")
                if matched_payg: info_parts.append(f"PAYG: **{matched_payg}**")
                if matched_acct: info_parts.append(f"Accounting: **{matched_acct}**")
                if info_parts:
                    st.success("✅ Client found in tracker — " + " · ".join(info_parts))
                break
    except Exception as e:
        st.warning(f"Could not read tracker: {e}")

st.markdown('</div>', unsafe_allow_html=True)

# ── Preview filename ──────────────────────────────────────────────────────────
if client_name_input:
    freq_suffix = " Qtr" if frequency == "Quarterly" else ""
    preview_name = f"{month_sel}{str(year_sel)[-2:]}_BAS{freq_suffix} {client_name_input.strip().upper()}.xlsm"
    st.info(f"📁 Output filename: **{preview_name}**")

# ── Step 2: Templates ─────────────────────────────────────────────────────────
st.markdown('<div class="card"><div class="card-title">② BAS Templates</div>', unsafe_allow_html=True)
st.caption("Upload both template files once — they are used based on the selected Accounting Method above.")
col_t1, col_t2 = st.columns(2)
with col_t1:
    st.markdown("**Accrual Basis Template** (.xlsm)")
    tmpl_accrual = st.file_uploader("", type=["xlsm", "xlsx"], key="tmpl_accrual")
with col_t2:
    st.markdown("**Cash Basis Template** (.xlsm)")
    tmpl_cash = st.file_uploader("", type=["xlsm", "xlsx"], key="tmpl_cash")
st.markdown('</div>', unsafe_allow_html=True)

# ── Step 3: Source Reports ─────────────────────────────────────────────────────
st.markdown('<div class="card"><div class="card-title">③ Source Report Files</div>', unsafe_allow_html=True)

st.markdown("""
<table style="width:100%;font-size:0.83rem;border-collapse:collapse;margin-bottom:12px">
<tr style="background:#f7fafc;">
  <th style="padding:7px 10px;text-align:left;border-bottom:1px solid #e2e8f0">Upload File</th>
  <th style="padding:7px 10px;text-align:left;border-bottom:1px solid #e2e8f0">Goes into Sheet(s)</th>
  <th style="padding:7px 10px;text-align:left;border-bottom:1px solid #e2e8f0">Notes</th>
</tr>
<tr><td style="padding:6px 10px">Activity Statement</td><td>GST Summary · GST Detail · BAS Field</td><td>Contains 3 tabs</td></tr>
<tr style="background:#f7fafc"><td style="padding:6px 10px">Balance Sheet</td><td>BS</td><td>1 tab</td></tr>
<tr><td style="padding:6px 10px">Profit & Loss</td><td>PL</td><td>1 tab</td></tr>
<tr style="background:#f7fafc"><td style="padding:6px 10px">Payroll Activity Summary</td><td>PAYROLL</td><td>Optional</td></tr>
<tr><td style="padding:6px 10px">Accounts Receivable <span class="badge badge-cash">Cash only</span></td><td>AR</td><td>Optional</td></tr>
<tr style="background:#f7fafc"><td style="padding:6px 10px">Accounts Payable <span class="badge badge-cash">Cash only</span></td><td>AP</td><td>Optional</td></tr>
</table>
""", unsafe_allow_html=True)

col_a, col_b = st.columns(2)
with col_a:
    st.markdown("**Activity Statement** <span class='badge badge-required'>Required</span>", unsafe_allow_html=True)
    file_activity = st.file_uploader("", type=["xlsx", "xlsm", "xls"], key="activity")

    st.markdown("**Balance Sheet** <span class='badge badge-required'>Required</span>", unsafe_allow_html=True)
    file_balance = st.file_uploader("", type=["xlsx", "xlsm", "xls"], key="balance")

    st.markdown("**Profit & Loss** <span class='badge badge-required'>Required</span>", unsafe_allow_html=True)
    file_pl = st.file_uploader("", type=["xlsx", "xlsm", "xls"], key="pl")

with col_b:
    st.markdown("**Payroll Activity Summary** <span class='badge badge-optional'>Optional</span>", unsafe_allow_html=True)
    file_payroll = st.file_uploader("", type=["xlsx", "xlsm", "xls"], key="payroll")

    if accounting_method == "Cash Basis":
        st.markdown("**Accounts Receivable** <span class='badge badge-cash'>Cash Basis</span>", unsafe_allow_html=True)
        file_ar = st.file_uploader("", type=["xlsx", "xlsm", "xls"], key="ar")

        st.markdown("**Accounts Payable** <span class='badge badge-cash'>Cash Basis</span>", unsafe_allow_html=True)
        file_ap = st.file_uploader("", type=["xlsx", "xlsm", "xls"], key="ap")
    else:
        file_ar = None
        file_ap = None

st.markdown('</div>', unsafe_allow_html=True)

# ── Validation summary ────────────────────────────────────────────────────────
def pill(ok, label):
    cls = "pill-ok" if ok else "pill-err"
    icon = "✓" if ok else "✗"
    return f'<span class="{cls}">{icon} {label}</span>&nbsp;'

checks_html = ""
all_ok = True

ok_client = bool(client_name_input.strip())
checks_html += pill(ok_client, "Client Name")
if not ok_client: all_ok = False

ok_tmpl_a = tmpl_accrual is not None
ok_tmpl_c = tmpl_cash is not None
if accounting_method == "Accrual Basis":
    checks_html += pill(ok_tmpl_a, "Accrual Template")
    if not ok_tmpl_a: all_ok = False
else:
    checks_html += pill(ok_tmpl_c, "Cash Template")
    if not ok_tmpl_c: all_ok = False

checks_html += pill(file_activity is not None, "Activity Statement")
checks_html += pill(file_balance is not None, "Balance Sheet")
checks_html += pill(file_pl is not None, "P&L")

if not file_activity or not file_balance or not file_pl:
    all_ok = False

st.markdown(f'<div style="margin-bottom:18px">{checks_html}</div>', unsafe_allow_html=True)

# ── Generate ──────────────────────────────────────────────────────────────────
if st.button("⚡  Generate BAS Workbook"):
    if not all_ok:
        st.error("Please fill in all required fields and upload required files before generating.")
    else:
        with st.spinner("Consolidating files…"):
            try:
                def read_wb(f):
                    if f is None:
                        return None
                    return load_workbook(io.BytesIO(f.read()), keep_vba=True, data_only=False)

                act_wb = read_wb(file_activity)
                bal_wb = read_wb(file_balance)
                pl_wb_obj = read_wb(file_pl)
                pay_wb = read_wb(file_payroll)
                ar_wb_obj = read_wb(file_ar) if file_ar else None
                ap_wb_obj = read_wb(file_ap) if file_ap else None

                tmpl_accrual_bytes = tmpl_accrual.read() if tmpl_accrual else b""
                tmpl_cash_bytes = tmpl_cash.read() if tmpl_cash else b""

                output_bytes = build_output(
                    client_name=client_name_input.strip().upper(),
                    month=month_sel,
                    year=year_sel,
                    accounting_method=accounting_method,
                    payg=payg,
                    frequency=frequency,
                    activity_wb=act_wb,
                    balance_wb=bal_wb,
                    pl_wb=pl_wb_obj,
                    payroll_wb=pay_wb,
                    ar_wb=ar_wb_obj,
                    ap_wb=ap_wb_obj,
                    template_accrual_bytes=tmpl_accrual_bytes,
                    template_cash_bytes=tmpl_cash_bytes,
                )

                freq_suffix = " Qtr" if frequency == "Quarterly" else ""
                yr_short = str(year_sel)[-2:]
                download_name = f"{month_sel}{yr_short}_BAS{freq_suffix} {client_name_input.strip().upper()}.xlsm"

                st.session_state["output_bytes"] = output_bytes
                st.session_state["download_name"] = download_name
                st.success(f"✅ Workbook generated successfully: **{download_name}**")

            except Exception as e:
                st.error(f"❌ Error generating workbook: {e}")
                import traceback
                st.code(traceback.format_exc())

if "output_bytes" in st.session_state:
    st.download_button(
        label="⬇  Download BAS Workbook",
        data=st.session_state["output_bytes"],
        file_name=st.session_state["download_name"],
        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
    )

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  <span>dexterous · BAS Consolidator</span>&nbsp;·&nbsp;
  <span>Merge Xero exports into a BAS workbook</span>
</div>
""", unsafe_allow_html=True)
