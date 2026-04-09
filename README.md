# 📋 BAS Consolidator

A Streamlit app built for **Dexterous Group** that merges multiple Xero export files into a single, properly structured BAS workbook — replacing the manual copy-paste process.

---

## 🚀 Getting Started

### Running Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

### Deploying to Streamlit Cloud

1. Push `app.py` and `requirements.txt` to your GitHub repo
2. Go to [share.streamlit.io](https://share.streamlit.io) and connect the repo
3. Set the main file path to `app.py`
4. Deploy — dependencies install automatically via `requirements.txt`

---

## 📁 Repository Structure

```
├── app.py                  # Main Streamlit app
├── requirements.txt        # Python dependencies
└── README.md               # This file
```

---

## 🧭 How to Use the App

### Step 1 — Client Details

Fill in the following fields:

| Field | Description |
|-------|-------------|
| **Client Name** | Full client name e.g. `UPSTAGE WORLD PTY LTD` |
| **Frequency** | `Quarterly` or `Monthly` |
| **Month / Year** | Period being processed e.g. `DEC / 2025` |
| **Accounting Method** | `Cash Basis` or `Accrual Basis` |
| **PAYG Instalment** | `Monthly`, `Quarterly`, or `No Payroll` |

> **Tip:** Upload the **BAS Monthly Automation Tracker** (`BAS_Monthly_Automation_Tracker.xlsx`) to auto-detect the client's Frequency, PAYG, and Accounting Method from the master list.

The output filename is automatically generated in this format:
```
DEC25_BAS Qtr UPSTAGE WORLD PTY LTD.xlsm
```

---

### Step 2 — BAS Templates

Upload both template files (only needs to be done once per session). The correct template is selected automatically based on the Accounting Method chosen.

| Template | Used When |
|----------|-----------|
| `BAS_template_-_Accrual_Basis.xlsm` | Accounting Method = Accrual Basis |
| `BAS_template_-_Cash_Basis.xlsm` | Accounting Method = Cash Basis |

> **Cash Basis** template includes two extra sheets: **AP** and **AR**

---

### Step 3 — Source Report Files

Upload the Xero exports. Each file's data is mapped to the corresponding sheet in the output workbook:

| Xero Export | → Output Sheet(s) | Required? |
|-------------|-------------------|-----------|
| Activity Statement | GST Summary, GST Detail, BAS Field | ✅ Required |
| Balance Sheet | BS | ✅ Required |
| Profit & Loss | PL | ✅ Required |
| Payroll Activity Summary | PAYROLL | ⬜ Optional |
| Accounts Receivable | AR | ⬜ Optional (Cash Basis only) |
| Accounts Payable | AP | ⬜ Optional (Cash Basis only) |

> The **Activity Statement** file contains 3 sheets from Xero — the app automatically identifies and routes each one to the correct destination sheet.

---

### Step 4 — Generate & Download

Click **⚡ Generate BAS Workbook**. The app will:

1. Load the correct template (Cash or Accrual)
2. Copy data from each uploaded report into the matching sheet
3. Populate the **Queries** sheet with client name, period, accounting method, PAYG, and file name
4. Produce a ready-to-download `.xlsm` file

Click **⬇ Download BAS Workbook** to save the file locally.

---

## 🗂 Output Sheet Reference

| Sheet | Source |
|-------|--------|
| Queries | Auto-populated from inputs |
| GST Summary | Activity Statement → "Activity Statement" tab |
| GST Detail | Activity Statement → "Transactions by Tax Rate" tab |
| BAS Field | Activity Statement → "Transactions by BAS Field" tab |
| BS | Balance Sheet file |
| PL | Profit & Loss file |
| PAYROLL | Payroll Activity Summary file |
| AR | Accounts Receivable file *(Cash Basis only)* |
| AP | Accounts Payable file *(Cash Basis only)* |

---

## 📦 Dependencies

| Package | Purpose |
|---------|---------|
| `streamlit` | Web app framework |
| `openpyxl` | Read/write Excel `.xlsx` / `.xlsm` files |
| `pandas` | Data handling |

---

## 📝 Notes

- The app preserves the VBA macros and existing formatting from the template files
- The Queries sheet is always updated with the correct client metadata
- AR and AP upload fields are only shown when **Cash Basis** is selected
- The tracker auto-fill is a read-only lookup — it does not modify the tracker file
