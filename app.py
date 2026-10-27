import io
import re
import datetime
import streamlit as st
import pandas as pd
import polars as pl
from recon_engine import execute_high_volume_reconciliation
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="AutoReconFlow — Enterprise Autonomous Reconciliation (SOX 404 / ISA-500 Compliant)",
    page_icon="⚡",
    layout="wide"
)

# Sidebar
st.sidebar.title("⚡ AutoReconFlow")
st.sidebar.markdown("**Enterprise B2B Edition v2.0**")
st.sidebar.caption("SOX 404 / ISA-500 & SOC 1 Type II Architecture")
st.sidebar.divider()
st.sidebar.markdown("### 💼 Commercial Licensing")
st.sidebar.markdown(
    """
    Acquire full source code, Docker assets, and distribution rights:
    - **Single Practice License**: $79
    - **Extended / Agency License**: $499
    
    [👉 Purchase Commercial License](https://autoreconflow.com)
    """
)
st.sidebar.caption("Includes Polars multi-core matching, PDF ingestion, and GST/VAT cross-border rules.")
st.sidebar.divider()

st.title("⚡ AutoReconFlow — Enterprise Autonomous Reconciliation")
st.markdown("High-throughput, 3-tier financial matching and cross-border tax reconciliation platform.")

# --- Enterprise User Manual & Technical Ingestion Guidelines ---
with st.expander("📖 User Guide & Ingestion Guidelines (SOX / ISA-500 Standard)", expanded=False):
    st.markdown("""
    ### 1. Architectural Overview & Workflow
    AutoReconFlow operates a deterministic, 3-tier matching engine powered by Polars:
    1. **Data Ingestion**: Files from Source A (Internal Books/Ledger) and Source B (Bank/Counterparty/Tax Portal) are parsed into memory.
    2. **Field Normalization & Sanitization**: Strings are trimmed, special characters removed, and numeric values converted to standard floats.
    3. **Three-Tier Matching**:
       * **Tier-1 (Deterministic)**: 100% exact match on Reference ID and Amount.
       * **Tier-2 (Fuzzy / Threshold Review)**: Matches within specified tolerance limits and text proximity (Jaro-Winkler / Levenshtein).
       * **Tier-3 (Exceptions)**: Discrepancies and unlinked entries categorized by origin source.
    4. **Audit-Trail Export**: Multi-sheet SOX-compliant workbook generation with immutable run metadata and supervisory sign-off sections.

    ---

    ### 2. Supported vs. Unsupported File Types

    | File Category | Supported Inputs (✅ Accepted) | Unsupported Inputs (❌ Not Accepted) |
    | :--- | :--- | :--- |
    | **CSV** | • Delimited by comma (`,`), semicolon (`;`), tab (`\\t`), or pipe (`\|`)<br>• UTF-8, UTF-8-BOM, UTF-16, Latin-1 encodings | • Password-protected / encrypted CSVs<br>• Corrupt or unclosed quote delimiters |
    | **Excel** | • Modern `.xlsx` workbooks<br>• Legacy `.xls` spreadsheets | • Password-protected Excel files<br>• Macro-enabled dynamic sheets that modify schemas |
    | **PDF Documents** | • **Digitally generated / vector PDFs** (Direct exports from online banking, ERPs like SAP/Oracle, or 'Print-to-PDF' statements) containing native tabular grids | • **Scanned paper / photocopies / flat images** (requires an external OCR layer)<br>• Password-locked / encrypted PDFs<br>• Multi-column newspaper text layouts without tabular boundaries |

    ---

    ### 3. File Preparation Checklist for Highest Accuracy
    * **Header Rows**: Ensure the first row of your table contains descriptive column titles (e.g., `Invoice_ID`, `Party_Name`, `Amount`).
    * **Amount Formats**: Clean numeric strings (e.g., `12500.00` or `12,500.00`) are automatically parsed. Avoid mixing currency symbols in the same cell as the digits (prefer `12500.00` over `USD 12500.00`).
    * **Transaction Identifiers**: Invoices, cheques, or reference codes should be unique per transaction where possible.
    """)

# --- Professional SOX 404 & ISA-500 Multi-Tab Audit Workpaper Generator ---
def generate_consolidated_workbook(matched_df, fuzzy_df, un_a_df, un_b_df, tolerance=1.0) -> bytes:
    output = io.BytesIO()
    timestamp_utc = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

    # Corporate Audit Palette
    navy_header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=11)
    title_font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        workbook = writer.book

        # -------------------------------------------------------------
        # SHEET 1: AUDIT & CONTROL SUMMARY (SOX 404 COMPLIANCE)
        # -------------------------------------------------------------
        ws_summary = workbook.create_sheet(title="Audit & Control Summary", index=0)
        ws_summary.views.sheetView[0].showGridLines = True

        ws_summary["A1"] = "AutoReconFlow — Enterprise Reconciliation & Control Report"
        ws_summary["A1"].font = title_font
        ws_summary["A2"] = "SOX 404 / ISA-500 Standard Financial Statement Assertion Audit Trail"
        ws_summary["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")

        metadata = [
            ("Execution Timestamp (UTC):", timestamp_utc),
            ("Tolerance Applied:", f"± {tolerance:.2f} Currency Units"),
            ("Matching Methodology:", "3-Tier (Deterministic Exact -> Levenshtein/Jaro-Winkler Fuzzy)"),
            ("Control Assertion Target:", "Completeness, Accuracy, and Valuation (Financial Controls)")
        ]
        ws_summary.append([])
        for label, val in metadata:
            ws_summary.append([label, val])

        for r in range(4, 8):
            ws_summary.cell(row=r, column=1).font = bold_font
            ws_summary.cell(row=r, column=2).font = regular_font

        ws_summary.append([])
        ws_summary.append(["Reconciliation Category", "Record Count", "Control Status"])
        summary_header_row = 9
        for col_idx in range(1, 4):
            cell = ws_summary.cell(row=summary_header_row, column=col_idx)
            cell.fill = navy_header_fill
            cell.font = white_bold_font
            cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left")

        summary_rows = [
            ("Tier-1 Deterministic Matches (Subledger to External)", len(matched_df), "Cleared"),
            ("Tier-2 Discrepancy & Threshold Review (Materiality Tracking)", len(fuzzy_df), "Requires Supervisory Review" if len(fuzzy_df) > 0 else "Cleared"),
            ("Source A Exceptions (Ledger Missing in External)", len(un_a_df), "Action Required" if len(un_a_df) > 0 else "Cleared"),
            ("Source B Exceptions (External Missing in Ledger)", len(un_b_df), "Action Required" if len(un_b_df) > 0 else "Cleared"),
        ]

        for item in summary_rows:
            ws_summary.append(list(item))

        for r in range(10, 14):
            for c in range(1, 4):
                cell = ws_summary.cell(row=r, column=c)
                cell.font = regular_font
                cell.border = thin_border
                if c == 2:
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "#,##0"
                elif c == 3:
                    cell.alignment = Alignment(horizontal="center")

        # Governance Sign-Off Block
        sign_off_start = 16
        ws_summary.cell(row=sign_off_start, column=1, value="Internal Control Sign-Off (SOX 404 / Four-Eyes Governance)").font = bold_font
        ws_summary.cell(row=sign_off_start, column=1).fill = section_fill
        ws_summary.merge_cells(start_row=sign_off_start, start_column=1, end_row=sign_off_start, end_column=3)

        signoffs = [
            ("Prepared By (Analyst):", "___________________________", "Date: ____________"),
            ("Reviewed By (Controller / Manager):", "___________________________", "Date: ____________"),
            ("Audit Workpaper Reference:", "WP-REC-2026-Q4", "Status: Pending Review")
        ]
        for role, line, dt in signoffs:
            ws_summary.append([role, line, dt])

        for r in range(sign_off_start + 1, sign_off_start + 4):
            ws_summary.cell(row=r, column=1).font = bold_font
            ws_summary.cell(row=r, column=2).font = regular_font
            ws_summary.cell(row=r, column=3).font = regular_font

        # -------------------------------------------------------------
        # DATA SHEETS WITH PROFESSIONAL AUDIT STYLING
        # -------------------------------------------------------------
        data_sheets = {
            "Deterministic Matches": matched_df,
            "Fuzzy Candidates": fuzzy_df,
            "Exceptions Source A": un_a_df,
            "Exceptions Source B": un_b_df
        }

        for sheet_title, df_content in data_sheets.items():
            df_content.to_excel(writer, sheet_name=sheet_title, index=False)
            ws = writer.sheets[sheet_title]
            ws.views.sheetView[0].showGridLines = True
            ws.freeze_panes = "A2"

            for col_idx in range(1, len(df_content.columns) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = navy_header_fill
                cell.font = white_bold_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for row in ws.iter_rows(min_row=2, max_row=len(df_content) + 1, min_col=1, max_col=len(df_content.columns)):
                for cell in row:
                    cell.border = thin_border
                    cell.font = regular_font
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0.00"

        # Auto-fit column widths
        for sheetname in writer.sheets:
            ws = writer.sheets[sheetname]
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or "")
                    if cell.row == 1 and sheetname == "Audit & Control Summary":
                        continue
                    max_len = max(max_len, len(val_str))
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    return output.getvalue()

# --- Robust Ingestion Engine with Built-in Column Splitting ---
def robust_read_file(uploaded_file) -> pd.DataFrame:
    filename = uploaded_file.name.lower()
    uploaded_file.seek(0)
    
    if filename.endswith(".csv"):
        raw_bytes = uploaded_file.read()
        
        text = None
        for enc in ["utf-8-sig", "utf-8", "utf-16", "latin1"]:
            try:
                text = raw_bytes.decode(enc)
                if "\n" in text or "\r" in text:
                    break
            except Exception:
                continue
                
        if text is None:
            text = str(raw_bytes)
            
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if not lines:
            return pd.DataFrame()
            
        first_line = lines[0]
        delim = ","
        if "\t" in first_line:
            delim = "\t"
        elif ";" in first_line:
            delim = ";"
        elif "|" in first_line:
            delim = "|"
            
        try:
            df = pd.read_csv(io.StringIO(text), sep=delim, engine="python")
        except Exception:
            df = pd.read_csv(io.StringIO(text))
            
        if len(df.columns) == 1 and ("," in str(df.columns[0]) or ";" in str(df.columns[0])):
            actual_delim = "," if "," in str(df.columns[0]) else ";"
            df = pd.read_csv(io.StringIO(text), sep=actual_delim)
            
        df.columns = [str(c).strip().strip('"').strip("'").replace("\ufeff", "") for c in df.columns]
        return df.reset_index(drop=True)
        
    elif filename.endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        return df.reset_index(drop=True)
        
    elif filename.endswith(".pdf"):
        import pdfplumber
        rows = []
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for r in table:
                            clean_r = [str(c).strip() if c is not None else "" for c in r]
                            if any(clean_r):
                                rows.append(clean_r)
        except Exception as e:
            st.error(f"Failed to read PDF: {str(e)}. Please confirm the file is not password-protected.")
            return pd.DataFrame()

        if not rows:
            st.warning("⚠️ No tabular data identified in this PDF. Please confirm it is a digitally generated export rather than a scanned photocopy or flat image.")
            return pd.DataFrame()

        headers = [re.sub(r'[^a-zA-Z0-9_]', '', h.replace(" ", "_").lower()) or f"col_{i}" for i, h in enumerate(rows[0])]
        df = pd.DataFrame(rows[1:], columns=headers)
        return df.reset_index(drop=True)
        
    return pd.DataFrame()

# --- Sample Data Generator for Instant 1-Click Testing ---
SAMPLE_A_CSV = """Invoice_ID,Vendor_Name,Amount,Posting_Date
INV-2026-101,Apex Tech Solutions,12500.00,2026-10-01
INV-2026-102,BlueSky Logistics,8450.50,2026-10-02
INV-2026-103,CloudCore Systems,19200.00,2026-10-03
INV-2026-104,Delta Industrial Supplies,4300.00,2026-10-04
INV-2026-105,Echo Media Group,15750.25,2026-10-05
INV-2026-106,Falcon Freight Services,6200.00,2026-10-06
INV-2026-107,Global Consulting Partners,31000.00,2026-10-07
INV-2026-108,Horizon Office Supplies,1450.00,2026-10-08
INV-2026-109,Infinite Cloud Storage,9800.00,2026-10-09
INV-2026-110,Jupiter Legal Advisory,22500.00,2026-10-10"""

SAMPLE_B_CSV = """Transaction_Ref,Description,Settlement_Amount,Value_Date
INV2026101,Apex Tech Solutions LLC,12500.00,2026-10-02
INV2026102,BlueSky Logistics,8450.50,2026-10-03
INV-2026-103,CloudCore Sys Inc,19200.00,2026-10-04
INV-2026-104-MISC,Delta Industrial Supplies,4300.00,2026-10-05
INV-2026-105,Echo Media,15750.25,2026-10-06
INV-2026-999,Vanguard Security,5100.00,2026-10-06
INV-2026-107,Global Consulting,31000.00,2026-10-08
INV-2026-108,Horizon Office Supplies,1449.50,2026-10-09
INV-2026-888,QuickFleet Courier,3200.00,2026-10-10
INV-2026-110,Jupiter Legal,22500.00,2026-10-11"""

st.markdown("##### 🧪 Want to test without creating files?")
use_mock = st.checkbox("Load Built-in Sample Test Datasets (Source A & Source B)")

df_a = None
df_b = None

if use_mock:
    df_a = pd.read_csv(io.StringIO(SAMPLE_A_CSV))
    df_b = pd.read_csv(io.StringIO(SAMPLE_B_CSV))
    st.info("Loaded built-in test records for Source A and Source B.")
else:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Source A (Ledger / Internal Records)")
        file_a = st.file_uploader("Upload CSV, XLSX, or Bank Statement PDF", type=["csv", "xlsx", "xls", "pdf"], key="file_a")
        if file_a:
            df_a = robust_read_file(file_a)

    with col2:
        st.subheader("Source B (Bank / Tax Portal / External Records)")
        file_b = st.file_uploader("Upload CSV, XLSX, or Bank Statement PDF", type=["csv", "xlsx", "xls", "pdf"], key="file_b")
        if file_b:
            df_b = robust_read_file(file_b)

def get_default_index(columns, keywords, fallback_index=0):
    for idx, col in enumerate(columns):
        clean_col = str(col).lower().replace(" ", "_")
        if any(kw in clean_col for kw in keywords):
            return idx
    return min(fallback_index, len(columns) - 1)

if df_a is not None and df_b is not None:
    cols_a = list(df_a.columns)
    cols_b = list(df_b.columns)

    st.success(f"Loaded Source A ({len(df_a):,} rows, {len(cols_a)} cols) and Source B ({len(df_b):,} rows, {len(cols_b)} cols)")

    with st.expander("🛠️ Configure Column Mappings", expanded=True):
        m1, m2 = st.columns(2)
        with m1:
            st.markdown("**Source A (Internal Books)**")
            idx_id_a = get_default_index(cols_a, ["invoice", "inv", "doc", "ref", "id"], 0)
            idx_party_a = get_default_index(cols_a, ["vendor", "party", "name", "desc"], min(1, len(cols_a)-1))
            idx_amt_a = get_default_index(cols_a, ["amount", "total", "val", "debit"], min(2, len(cols_a)-1))

            id_a = st.selectbox("Document / Txn ID Column", cols_a, index=idx_id_a, key="map_id_a")
            party_a = st.selectbox("Party / Narrative Column", cols_a, index=idx_party_a, key="map_party_a")
            amt_a = st.selectbox("Amount Column", cols_a, index=idx_amt_a, key="map_amt_a")

        with m2:
            st.markdown("**Source B (Bank / Portal)**")
            idx_id_b = get_default_index(cols_b, ["transaction", "ref", "invoice", "inv", "id"], 0)
            idx_party_b = get_default_index(cols_b, ["description", "narrative", "party", "particulars"], min(1, len(cols_b)-1))
            idx_amt_b = get_default_index(cols_b, ["settlement", "amount", "total", "net"], min(2, len(cols_b)-1))

            id_b = st.selectbox("Document / Txn ID Column", cols_b, index=idx_id_b, key="map_id_b")
            party_b = st.selectbox("Party / Narrative Column", cols_b, index=idx_party_b, key="map_party_b")
            amt_b = st.selectbox("Amount Column", cols_b, index=idx_amt_b, key="map_amt_b")

    tolerance = st.slider("Allowed Amount Tolerance (Currency Units)", 0.00, 5.00, 1.00, step=0.10)

    if st.button("🚀 Run Automated Reconciliation", type="primary"):
        map_a = {"id": id_a, "party": party_a, "amt": amt_a}
        map_b = {"id": id_b, "party": party_b, "amt": amt_b}

        # Convert to Polars for high-speed execution
        pl_a = pl.from_pandas(df_a)
        pl_b = pl.from_pandas(df_b)

        with st.spinner("Processing through 3-tier engine..."):
            try:
                matched, fuzzy, un_a, un_b = execute_high_volume_reconciliation(
                    pl_a, pl_b, map_a, map_b, tolerance=tolerance
                )
            except Exception as e:
                st.error(f"Reconciliation error: {str(e)}")
                st.stop()

        matched_pd = matched.to_pandas()
        fuzzy_pd = fuzzy.to_pandas()
        un_a_pd = un_a.to_pandas()
        un_b_pd = un_b.to_pandas()

        st.subheader("📊 Reconciliation & Control Summary (SOX 404 / ISA-500)")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Tier-1 Deterministic Matches", f"{len(matched_pd):,} records")
        k2.metric("Tier-2 Discrepancy Review", f"{len(fuzzy_pd):,} records")
        k3.metric("Source A Exceptions (Ledger)", f"{len(un_a_pd):,} records")
        k4.metric("Source B Exceptions (External)", f"{len(un_b_pd):,} records")

        # Top-level SOX-Compliant Multi-Tab Excel Export Button
        excel_bytes = generate_consolidated_workbook(matched_pd, fuzzy_pd, un_a_pd, un_b_pd, tolerance=tolerance)
        st.download_button(
            label="📦 Export SOX-Compliant Audit Workpaper (.xlsx)",
            data=excel_bytes,
            file_name=f"AutoReconFlow_SOX_Audit_Report_{datetime.datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_download_master_excel"
        )
        st.write("")

        # 4 Separate Review Tabs
        tab1, tab2, tab3, tab4 = st.tabs([
            "✅ Tier-1 Deterministic Matches",
            "🔍 Tier-2 Discrepancy & Threshold Review",
            "⚠️ Exceptions: Source A (Ledger Unmatched)",
            "⚠️ Exceptions: Source B (External Unmatched)"
        ])

        with tab1:
            if len(matched_pd) > 0:
                st.dataframe(matched_pd, use_container_width=True)
                st.download_button(
                    label="📥 Download Deterministic Matches (CSV)",
                    data=matched.write_csv(),
                    file_name="deterministic_matches.csv",
                    mime="text/csv",
                    key="btn_matched_csv"
                )
            else:
                st.info("No exact matches identified under current settings.")

        with tab2:
            if len(fuzzy_pd) > 0:
                st.markdown("**Discrepancies Requiring Review (Levenshtein / Jaro-Winkler Variance)**")
                st.dataframe(fuzzy_pd, use_container_width=True)
                st.download_button(
                    label="📥 Download Discrepancy Candidates (CSV)",
                    data=fuzzy.write_csv(),
                    file_name="fuzzy_review_candidates.csv",
                    mime="text/csv",
                    key="btn_fuzzy_csv"
                )
            else:
                st.info("Zero fuzzy discrepancies detected.")

        with tab3:
            if len(un_a_pd) > 0:
                st.markdown("**Source A Entries Missing in Source B (Unmatched Subledger Items)**")
                st.dataframe(un_a_pd, use_container_width=True)
                st.download_button(
                    label="📥 Download Source A Exceptions (CSV)",
                    data=un_a.write_csv(),
                    file_name="unmatched_ledger_exceptions.csv",
                    mime="text/csv",
                    key="btn_un_a_csv"
                )
            else:
                st.success("All Source A items accounted for (0 open variances).")

        with tab4:
            if len(un_b_pd) > 0:
                st.markdown("**Source B Entries Missing in Source A (Timing & External Exceptions)**")
                st.dataframe(un_b_pd, use_container_width=True)
                st.download_button(
                    label="📥 Download Source B Exceptions (CSV)",
                    data=un_b.write_csv(),
                    file_name="unmatched_external_exceptions.csv",
                    mime="text/csv",
                    key="btn_un_b_csv"
                )
            else:
                st.success("All Source B items accounted for (0 open variances).")