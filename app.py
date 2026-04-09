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
    layout="wide"
)

st.markdown("""
<style>
  div.stButton > button {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
    color: white !important; font-weight: 600 !important;
    border: none !important; border-radius: 8px !important;
    width: 100%; padding: 0.6rem 1.5rem !important; font-size: 1rem !important;
  }
  div.stDownloadButton > button {
    background: linear-gradient(135deg, #38a169 0%, #2f855a 100%) !important;
    color: white !important; font-weight: 600 !important;
    border: none !important; border-radius: 8px !important;
    width: 100%; padding: 0.6rem 1.5rem !important; font-size: 1rem !important;
  }
  #MainMenu, footer { visibility: hidden; }
  /* Tighten dataframe font */
  [data-testid="stDataFrame"] { font-size: 0.82rem; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown(
    """<div style="background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);
        border-radius:12px;padding:26px 32px 20px;margin-bottom:8px;
        display:flex;align-items:center;gap:18px;">
      <span style="font-size:2.2rem;">&#128203;</span>
      <div>
        <div style="color:#fff;font-size:1.5rem;font-weight:700;margin-bottom:3px;">BAS Consolidator</div>
        <div style="color:#a0aec0;font-size:0.87rem;">dexterous &middot; Bulk upload Xero exports &rarr; single BAS workbook</div>
      </div>
    </div>""",
    unsafe_allow_html=True,
)

# ── Constants ──────────────────────────────────────────────────────────────────
MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
CURRENT_YEAR = datetime.now().year
YEARS = [str(y) for y in range(CURRENT_YEAR - 2, CURRENT_YEAR + 3)]

# Report-type keywords found in Xero filenames → internal key
REPORT_KEYWORDS = {
    "activity_statement":    ["activity_statement", "activity statement"],
    "balance_sheet":         ["balance_sheet", "balance sheet"],
    "profit_and_loss":       ["profit_and_loss", "profit and loss", "profit_&_loss"],
    "payroll_activity":      ["payroll_activity_summary", "payroll activity summary", "payroll_activity", "payroll activity"],
    "aged_receivables":      ["aged_receivables_detail", "aged receivables detail", "aged_receivables", "aged receivables"],
    "aged_payables":         ["aged_payables_detail", "aged payables detail", "aged_payables", "aged payables"],
}

REPORT_LABELS = {
    "activity_statement": ("Activity Statement",      "GST Summary + GST Detail + BAS Field", True),
    "balance_sheet":      ("Balance Sheet",            "BS",                                   True),
    "profit_and_loss":    ("Profit & Loss",            "PL",                                   True),
    "payroll_activity":   ("Payroll Activity Summary", "PAYROLL",                              False),
    "aged_receivables":   ("Aged Receivables Detail",  "AR",                                   False),
    "aged_payables":      ("Aged Payables Detail",     "AP",                                   False),
}

# ── Helpers ────────────────────────────────────────────────────────────────────
def detect_report_type(filename: str) -> str | None:
    """Detect report type from filename."""
    name = filename.lower().replace(" ", "_")
    for rtype, keywords in REPORT_KEYWORDS.items():
        for kw in keywords:
            if kw.replace(" ", "_") in name:
                return rtype
    return None


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
                    dst_cell.font          = copy.copy(cell.font)
                    dst_cell.fill          = copy.copy(cell.fill)
                    dst_cell.border        = copy.copy(cell.border)
                    dst_cell.alignment     = copy.copy(cell.alignment)
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


def get_dst(out_wb, preferred, fallbacks=None):
    for n in [preferred] + (fallbacks or []):
        for sn in out_wb.sheetnames:
            if sn.strip().lower() == n.strip().lower():
                return out_wb[sn]
    return out_wb.create_sheet(preferred)


def build_output(client_name, month, year, accounting_method, payg, frequency, report_wbs):
    """Build the consolidated BAS workbook from detected report workbooks."""
    from openpyxl import Workbook

    # Load correct embedded template
    tmpl_key = "cash" if accounting_method == "Cash Basis" else "accrual"
    template_bytes = st.session_state.get(f"tmpl_{tmpl_key}")
    if template_bytes:
        out_wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    else:
        # Build a blank workbook with required sheets
        out_wb = Workbook()
        out_wb.remove(out_wb.active)
        sheets_needed = ["Queries", "GST Summary", "GST Detail", "BAS field", "BS", "PL", "PAYROLL"]
        if accounting_method == "Cash Basis":
            sheets_needed += ["AR", "AP"]
        for sn in sheets_needed:
            out_wb.create_sheet(sn)

    # ── Activity Statement → 3 sheets
    if "activity_statement" in report_wbs:
        wb = report_wbs["activity_statement"]
        for sn in wb.sheetnames:
            ws = wb[sn]
            h = (ws.cell(1,1).value or "").lower()
            if "activity statement" in h:
                copy_sheet_data(ws, get_dst(out_wb, "GST Summary"))
            elif "transactions by tax rate" in h:
                copy_sheet_data(ws, get_dst(out_wb, "GST Detail"))
            elif "transactions by bas field" in h:
                copy_sheet_data(ws, get_dst(out_wb, "BAS field", ["BAS Field"]))

    # ── Balance Sheet
    if "balance_sheet" in report_wbs:
        copy_sheet_data(report_wbs["balance_sheet"].active, get_dst(out_wb, "BS", ["BS "]))

    # ── Profit & Loss
    if "profit_and_loss" in report_wbs:
        copy_sheet_data(report_wbs["profit_and_loss"].active, get_dst(out_wb, "PL", ["P&L", "P&L "]))

    # ── Payroll
    if "payroll_activity" in report_wbs:
        copy_sheet_data(report_wbs["payroll_activity"].active, get_dst(out_wb, "PAYROLL", ["Payroll Activity Smry"]))

    # ── AR / AP (Cash only)
    if "aged_receivables" in report_wbs and accounting_method == "Cash Basis":
        copy_sheet_data(report_wbs["aged_receivables"].active, get_dst(out_wb, "AR"))
    if "aged_payables" in report_wbs and accounting_method == "Cash Basis":
        copy_sheet_data(report_wbs["aged_payables"].active, get_dst(out_wb, "AP"))

    # ── Queries sheet
    freq_suffix    = " Qtr" if frequency == "Quarterly" else ""
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
    queries_ws["A6"] = "Email sent for queries and confirmation"
    queries_ws["A7"] = "Subject reference"
    queries_ws["A8"] = file_name_val

    buf = io.BytesIO()
    out_wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# LAYOUT  — two columns: left = controls, right = masterlist
# ══════════════════════════════════════════════════════════════════════════════
left, right = st.columns([1.1, 1], gap="large")

# ─────────────────────────────────────────────────────────────────────────────
# LEFT PANEL
# ─────────────────────────────────────────────────────────────────────────────
with left:

    # ── Client Details ────────────────────────────────────────────────────────
    st.subheader("① Client Details")

    col1, col2 = st.columns([2, 1])
    with col1:
        client_name_input = st.text_input("Client Name", placeholder="e.g. TSM THE SERVICE MANAGER PTY LTD")
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

    if client_name_input:
        freq_suffix  = " Qtr" if frequency == "Quarterly" else ""
        yr_short     = str(year_sel)[-2:]
        preview_name = f"{month_sel}{yr_short}_BAS{freq_suffix} {client_name_input.strip().upper()}.xlsm"
        st.info(f"📁 **{preview_name}**")

    st.divider()

    # ── Bulk Upload ───────────────────────────────────────────────────────────
    st.subheader("② Upload Xero Exports")
    st.caption(
        "Drop all files at once — the tool reads the filename to identify each report automatically. "
        "Expected filename format: `ClientName_-_ReportType.xlsx`"
    )

    uploaded_files = st.file_uploader(
        "Select all Xero export files",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=True,
        key="bulk_upload",
        label_visibility="collapsed",
    )

    # ── Detection results ─────────────────────────────────────────────────────
    detected   = {}   # rtype → UploadedFile
    undetected = []

    if uploaded_files:
        for f in uploaded_files:
            rtype = detect_report_type(f.name)
            if rtype:
                detected[rtype] = f
            else:
                undetected.append(f.name)

        st.markdown("**Detected files:**")
        for rtype, f in detected.items():
            label, sheet_dest, required = REPORT_LABELS[rtype]
            st.markdown(f"- ✅ **{label}** → `{sheet_dest}` &nbsp;·&nbsp; `{f.name}`")

        if undetected:
            st.warning("⚠️ Could not identify these files — check filenames match Xero export format:")
            for fn in undetected:
                st.markdown(f"  - `{fn}`")

    st.divider()

    # ── Readiness check & Generate ────────────────────────────────────────────
    st.subheader("③ Generate")

    required_types  = ["activity_statement", "balance_sheet", "profit_and_loss"]
    ok_client = bool(client_name_input.strip())
    ok_files  = all(rt in detected for rt in required_types)
    all_ok    = ok_client and ok_files

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown("✅ Client" if ok_client else "❌ Client name")
    c2.markdown("✅ Activity Stmt"  if "activity_statement" in detected else "❌ Activity Stmt")
    c3.markdown("✅ Balance Sheet"  if "balance_sheet"       in detected else "❌ Balance Sheet")
    c4.markdown("✅ P&L"           if "profit_and_loss"     in detected else "❌ P&L")

    st.write("")
    if st.button("⚡  Generate BAS Workbook"):
        if not all_ok:
            st.error("Fill in client name and upload the 3 required files first.")
        else:
            with st.spinner("Consolidating…"):
                try:
                    report_wbs = {}
                    for rtype, uf in detected.items():
                        report_wbs[rtype] = load_workbook(io.BytesIO(uf.read()), keep_vba=False)

                    output_bytes = build_output(
                        client_name       = client_name_input.strip().upper(),
                        month             = month_sel,
                        year              = year_sel,
                        accounting_method = accounting_method,
                        payg              = payg,
                        frequency         = frequency,
                        report_wbs        = report_wbs,
                    )

                    freq_suffix   = " Qtr" if frequency == "Quarterly" else ""
                    yr_short      = str(year_sel)[-2:]
                    download_name = f"{month_sel}{yr_short}_BAS{freq_suffix} {client_name_input.strip().upper()}.xlsm"

                    st.session_state["output_bytes"]  = output_bytes
                    st.session_state["download_name"] = download_name
                    st.success(f"✅ **{download_name}** ready!")

                except Exception as e:
                    import traceback
                    st.error(f"❌ {e}")
                    st.code(traceback.format_exc())

    if "output_bytes" in st.session_state:
        st.download_button(
            label     = "⬇  Download BAS Workbook",
            data      = st.session_state["output_bytes"],
            file_name = st.session_state["download_name"],
            mime      = "application/vnd.ms-excel.sheet.macroEnabled.12",
        )


# ─────────────────────────────────────────────────────────────────────────────
# RIGHT PANEL — Masterlist
# ─────────────────────────────────────────────────────────────────────────────
with right:
    st.subheader("📋 Client Masterlist")
    st.caption("Upload the BAS Monthly Automation Tracker to load the masterlist.")

    tracker_file = st.file_uploader(
        "BAS Monthly Automation Tracker (.xlsx)",
        type=["xlsx"],
        key="tracker",
        label_visibility="collapsed",
    )

    if tracker_file:
        try:
            tracker_wb = load_workbook(io.BytesIO(tracker_file.read()), data_only=True)
            ws = tracker_wb["BAS"] if "BAS" in tracker_wb.sheetnames else tracker_wb.active

            rows = list(ws.iter_rows(values_only=True))
            headers = [str(h) if h else "" for h in rows[0]]

            def ci(name):
                for i, h in enumerate(headers):
                    if name.lower() in h.lower():
                        return i
                return None

            name_col  = ci("XeroClientName") or ci("ClientName") or 0
            freq_col  = ci("Frequency")
            payg_col  = ci("PAYGI")
            acct_col  = ci("GSTAccountingMethod") or ci("Accounting")
            stat_col  = ci("Status")

            table_rows = []
            for row in rows[1:]:
                name = str(row[name_col]).strip() if row[name_col] else ""
                if not name or name == "None":
                    continue
                table_rows.append({
                    "Client Name":         name,
                    "Status":              str(row[stat_col]).strip()  if stat_col  is not None and row[stat_col]  else "Active",
                    "Frequency":           str(row[freq_col]).strip()  if freq_col  is not None and row[freq_col]  else "",
                    "PAYG Instalment":     str(row[payg_col]).strip()  if payg_col  is not None and row[payg_col]  else "",
                    "Accounting Method":   str(row[acct_col]).strip()  if acct_col  is not None and row[acct_col]  else "",
                })

            df = pd.DataFrame(table_rows)

            # Highlight row matching current client
            if client_name_input.strip():
                search = client_name_input.strip().upper()

                def highlight_match(row):
                    match = search in row["Client Name"].upper() or row["Client Name"].upper() in search
                    return ["background-color: #ebf8ff; font-weight: 600" if match else "" for _ in row]

                st.dataframe(
                    df.style.apply(highlight_match, axis=1),
                    use_container_width=True,
                    hide_index=True,
                    height=600,
                )

                # Auto-fill hint
                matched = df[df["Client Name"].str.upper().str.contains(search, na=False)]
                if not matched.empty:
                    r = matched.iloc[0]
                    parts = []
                    if r["Frequency"]:        parts.append(f"Frequency: **{r['Frequency']}**")
                    if r["PAYG Instalment"]:  parts.append(f"PAYG: **{r['PAYG Instalment']}**")
                    if r["Accounting Method"]:parts.append(f"Method: **{r['Accounting Method']}**")
                    if parts:
                        st.success("✅ " + " · ".join(parts))
            else:
                st.dataframe(df, use_container_width=True, hide_index=True, height=600)

        except Exception as e:
            st.error(f"Could not read tracker: {e}")
    else:
        # Placeholder table showing the structure
        placeholder = pd.DataFrame([
            {"Client Name": "e.g. Livewire Markets Pty Ltd", "Status": "Active",   "Frequency": "Quarterly", "PAYG Instalment": "Monthly",   "Accounting Method": "Accrual Basis"},
            {"Client Name": "e.g. ISH Dental Pty Ltd",       "Status": "Active",   "Frequency": "Quarterly", "PAYG Instalment": "—",         "Accounting Method": "Cash Basis"},
            {"Client Name": "e.g. Harmony Build Pty Ltd",    "Status": "Inactive", "Frequency": "Quarterly", "PAYG Instalment": "Quarterly", "Accounting Method": "Cash Basis"},
        ])
        st.dataframe(placeholder, use_container_width=True, hide_index=True)
        st.caption("⬆ Upload tracker above to see the full client list.")

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("dexterous · BAS Consolidator — bulk Xero export consolidation")
