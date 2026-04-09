import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import io
import copy
from datetime import datetime

st.set_page_config(
    page_title="Dexterous · BAS Consolidator",
    page_icon="📋",
    layout="centered"
)

st.markdown("""
<style>
  div.stButton > button {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
    color: white !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 8px !important;
    width: 100%;
    padding: 0.6rem 1.5rem !important;
    font-size: 1rem !important;
  }
  div.stDownloadButton > button {
    background: linear-gradient(135deg, #38a169 0%, #2f855a 100%) !important;
    color: white !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 8px !important;
    width: 100%;
    padding: 0.6rem 1.5rem !important;
    font-size: 1rem !important;
  }
  #MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <div style="background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);
                border-radius:12px;padding:28px 32px 22px;margin-bottom:8px;
                display:flex;align-items:center;gap:18px;">
      <span style="font-size:2.4rem;">&#128203;</span>
      <div>
        <div style="color:#fff;font-size:1.55rem;font-weight:700;margin-bottom:4px;">BAS Consolidator</div>
        <div style="color:#a0aec0;font-size:0.88rem;">dexterous &middot; Merge Xero exports into a single BAS workbook</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Month/Year options ──────────────────────────────────────────────────────────
MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
CURRENT_YEAR = datetime.now().year
YEARS = [str(y) for y in range(CURRENT_YEAR - 2, CURRENT_YEAR + 3)]

# ── Helpers ────────────────────────────────────────────────────────────────────
def copy_sheet_data(src_ws, dst_ws):
    from openpyxl.cell.cell import MergedCell
    for row in dst_ws.iter_rows():
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None
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
                    dst_cell.font      = copy.copy(cell.font)
                    dst_cell.fill      = copy.copy(cell.fill)
                    dst_cell.border    = copy.copy(cell.border)
                    dst_cell.alignment = copy.copy(cell.alignment)
                    dst_cell.number_format = cell.number_format
                except Exception:
                    pass
    for col_letter, dim in src_ws.column_dimensions.items():
        try: dst_ws.column_dimensions[col_letter].width = dim.width
        except Exception: pass
    for row_idx, dim in src_ws.row_dimensions.items():
        try: dst_ws.row_dimensions[row_idx].height = dim.height
        except Exception: pass
    for merged in list(src_ws.merged_cells.ranges):
        try: dst_ws.merge_cells(str(merged))
        except Exception: pass


def get_dst(out_wb, preferred_name, fallbacks=None):
    names = [preferred_name] + (fallbacks or [])
    for n in names:
        for sn in out_wb.sheetnames:
            if sn.strip().lower() == n.strip().lower():
                return out_wb[sn]
    return out_wb.create_sheet(preferred_name)


def build_output(client_name, month, year, accounting_method, payg, frequency,
                 activity_wb, balance_wb, pl_wb, payroll_wb, ar_wb, ap_wb,
                 template_accrual_bytes, template_cash_bytes):

    template_bytes = template_cash_bytes if accounting_method == "Cash Basis" else template_accrual_bytes
    out_wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)

    # Activity Statement → 3 sheets
    if activity_wb:
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
            copy_sheet_data(sheet_map["gst_summary"], get_dst(out_wb, "GST Summary"))
        if "gst_detail" in sheet_map:
            copy_sheet_data(sheet_map["gst_detail"], get_dst(out_wb, "GST Detail"))
        if "bas_field" in sheet_map:
            copy_sheet_data(sheet_map["bas_field"], get_dst(out_wb, "BAS field", ["BAS Field"]))

    # Balance Sheet
    if balance_wb:
        copy_sheet_data(balance_wb.active, get_dst(out_wb, "BS", ["BS "]))

    # Profit & Loss
    if pl_wb:
        copy_sheet_data(pl_wb.active, get_dst(out_wb, "PL", ["P&L", "P&L "]))

    # Payroll
    if payroll_wb:
        copy_sheet_data(payroll_wb.active, get_dst(out_wb, "PAYROLL", ["Payroll Activity Smry"]))

    # AR / AP (Cash only)
    if ar_wb and accounting_method == "Cash Basis":
        copy_sheet_data(ar_wb.active, get_dst(out_wb, "AR"))
    if ap_wb and accounting_method == "Cash Basis":
        copy_sheet_data(ap_wb.active, get_dst(out_wb, "AP"))

    # Queries sheet
    freq_suffix   = " Qtr" if frequency == "Quarterly" else ""
    display_period = f"{month}{year}_BAS{freq_suffix}"
    file_name_val  = f"{month}{year}_BAS {client_name}"

    queries_ws = None
    for sn in out_wb.sheetnames:
        if sn.lower() == "queries":
            queries_ws = out_wb[sn]
            break
    if queries_ws is None:
        queries_ws = out_wb.create_sheet("Queries", 0)

    queries_ws["A1"] = "Client Name";        queries_ws["B1"] = client_name
    queries_ws["D1"] = "Period";             queries_ws["E1"] = display_period
    queries_ws["A2"] = "Accounting Method";  queries_ws["B2"] = accounting_method
    queries_ws["D2"] = "Completed by: "
    queries_ws["A3"] = "PAYG";               queries_ws["B3"] = payg
    queries_ws["A4"] = "File Name";          queries_ws["B4"] = file_name_val
    queries_ws["A5"] = "Note"
    queries_ws["A6"] = "Email sent for queries and confrimation"
    queries_ws["A7"] = "Subject reference"
    queries_ws["A8"] = file_name_val

    buf = io.BytesIO()
    out_wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────

# ── Step 1: Client Details ────────────────────────────────────────────────────
st.divider()
st.subheader("① Client Details")

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

payg = st.selectbox("PAYG Instalment", ["Monthly", "Quarterly", "No Payroll"])

tracker_file = st.file_uploader(
    "BAS Monthly Automation Tracker (.xlsx) — optional, auto-fills client settings",
    type=["xlsx"], key="tracker"
)
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
                parts = []
                if freq_col is not None and row[freq_col]: parts.append(f"Frequency: **{row[freq_col]}**")
                if payg_col is not None and row[payg_col]: parts.append(f"PAYG: **{row[payg_col]}**")
                if acct_col is not None and row[acct_col]: parts.append(f"Accounting: **{row[acct_col]}**")
                if parts:
                    st.success("✅ Client found in tracker — " + " · ".join(parts))
                break
    except Exception as e:
        st.warning(f"Could not read tracker: {e}")

if client_name_input:
    freq_suffix  = " Qtr" if frequency == "Quarterly" else ""
    yr_short     = str(year_sel)[-2:]
    preview_name = f"{month_sel}{yr_short}_BAS{freq_suffix} {client_name_input.strip().upper()}.xlsm"
    st.info(f"📁 Output filename: **{preview_name}**")

# ── Step 2: Templates ─────────────────────────────────────────────────────────
st.divider()
st.subheader("② BAS Templates")
st.caption("Upload both templates — the correct one is chosen automatically based on the Accounting Method selected above.")

col_t1, col_t2 = st.columns(2)
with col_t1:
    tmpl_accrual = st.file_uploader("Accrual Basis Template (.xlsm)", type=["xlsm", "xlsx"], key="tmpl_accrual")
with col_t2:
    tmpl_cash = st.file_uploader("Cash Basis Template (.xlsm)", type=["xlsm", "xlsx"], key="tmpl_cash")

# ── Step 3: Source Reports ────────────────────────────────────────────────────
st.divider()
st.subheader("③ Source Report Files")

st.markdown("""
| Upload File | → Sheet(s) | Notes |
|---|---|---|
| Activity Statement | GST Summary · GST Detail · BAS Field | 3-tab Xero export |
| Balance Sheet | BS | 1 tab |
| Profit & Loss | PL | 1 tab |
| Payroll Activity Summary | PAYROLL | Optional |
| Accounts Receivable | AR | Optional — Cash Basis only |
| Accounts Payable | AP | Optional — Cash Basis only |
""")

col_a, col_b = st.columns(2)
with col_a:
    file_activity = st.file_uploader("Activity Statement *(required)*", type=["xlsx","xlsm","xls"], key="activity")
    file_balance  = st.file_uploader("Balance Sheet *(required)*",      type=["xlsx","xlsm","xls"], key="balance")
    file_pl       = st.file_uploader("Profit & Loss *(required)*",      type=["xlsx","xlsm","xls"], key="pl")
with col_b:
    file_payroll = st.file_uploader("Payroll Activity Summary *(optional)*", type=["xlsx","xlsm","xls"], key="payroll")
    if accounting_method == "Cash Basis":
        file_ar = st.file_uploader("Accounts Receivable *(optional)*", type=["xlsx","xlsm","xls"], key="ar")
        file_ap = st.file_uploader("Accounts Payable *(optional)*",    type=["xlsx","xlsm","xls"], key="ap")
    else:
        file_ar = None
        file_ap = None

# ── Readiness checks ──────────────────────────────────────────────────────────
st.divider()

all_ok = True
ok_client = bool(client_name_input.strip())
ok_tmpl   = (tmpl_accrual is not None) if accounting_method == "Accrual Basis" else (tmpl_cash is not None)
ok_act    = file_activity is not None
ok_bs     = file_balance is not None
ok_pl_f   = file_pl is not None
if not ok_client or not ok_tmpl or not ok_act or not ok_bs or not ok_pl_f:
    all_ok = False

def ck(ok, label):
    return f"{'✅' if ok else '❌'} {label}"

c1, c2, c3, c4, c5 = st.columns(5)
c1.markdown(ck(ok_client, "Client"))
c2.markdown(ck(ok_tmpl,   "Template"))
c3.markdown(ck(ok_act,    "Activity Stmt"))
c4.markdown(ck(ok_bs,     "Balance Sheet"))
c5.markdown(ck(ok_pl_f,   "P&L"))

# ── Generate ──────────────────────────────────────────────────────────────────
st.write("")
if st.button("⚡  Generate BAS Workbook"):
    if not all_ok:
        st.error("Please fill in all required fields and upload all required files before generating.")
    else:
        with st.spinner("Consolidating files…"):
            try:
                def read_wb(f):
                    return load_workbook(io.BytesIO(f.read()), keep_vba=True) if f else None

                output_bytes = build_output(
                    client_name       = client_name_input.strip().upper(),
                    month             = month_sel,
                    year              = year_sel,
                    accounting_method = accounting_method,
                    payg              = payg,
                    frequency         = frequency,
                    activity_wb       = read_wb(file_activity),
                    balance_wb        = read_wb(file_balance),
                    pl_wb             = read_wb(file_pl),
                    payroll_wb        = read_wb(file_payroll),
                    ar_wb             = read_wb(file_ar) if file_ar else None,
                    ap_wb             = read_wb(file_ap) if file_ap else None,
                    template_accrual_bytes = tmpl_accrual.read() if tmpl_accrual else b"",
                    template_cash_bytes    = tmpl_cash.read()    if tmpl_cash    else b"",
                )

                freq_suffix   = " Qtr" if frequency == "Quarterly" else ""
                yr_short      = str(year_sel)[-2:]
                download_name = f"{month_sel}{yr_short}_BAS{freq_suffix} {client_name_input.strip().upper()}.xlsm"

                st.session_state["output_bytes"] = output_bytes
                st.session_state["download_name"] = download_name
                st.success(f"✅ Workbook generated: **{download_name}**")

            except Exception as e:
                st.error(f"❌ Error: {e}")
                import traceback
                st.code(traceback.format_exc())

if "output_bytes" in st.session_state:
    st.download_button(
        label     = "⬇  Download BAS Workbook",
        data      = st.session_state["output_bytes"],
        file_name = st.session_state["download_name"],
        mime      = "application/vnd.ms-excel.sheet.macroEnabled.12",
    )

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("dexterous · BAS Consolidator — Merge Xero exports into a BAS workbook")
