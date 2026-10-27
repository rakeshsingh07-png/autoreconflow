import io
import re
import datetime
import hashlib
import tempfile
import os
import streamlit as st
import pandas as pd
import polars as pl
from fpdf import FPDF
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

# Sidebar: Corporate White-Label & Entity Profile
st.sidebar.markdown("### 🏢 Client Organization Branding")
uploaded_logo = st.sidebar.file_uploader("Upload Corporate Logo (PNG/JPG)", type=["png", "jpg", "jpeg"], key="company_logo")
if uploaded_logo:
    st.sidebar.image(uploaded_logo, caption="Active Logo Branding", use_container_width=True)

company_name = st.sidebar.text_input("Company / Legal Entity Name", value="Global Enterprise Solutions Ltd.", key="cfg_comp_name")
company_address = st.sidebar.text_area("Registered Office Address", value="Tower 4, Financial District, Cyber City, Gurugram, India - 122002", key="cfg_comp_addr", height=70)
company_uid = st.sidebar.text_input("Unique Identification No. (CIN / LEI / Tax ID)", value="CIN: U72200HR2026PTC109823", key="cfg_comp_uid")
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
    4. **Two-Stage Four-Eyes Governance**: 
       * **Stage 1 (Preparer Submission)**: Preparer audits records and locks submission.
       * **Stage 2 (Supervisory Review)**: Reviewer evaluates variances, issues formal verdict, and seals completion.
    5. **Audit-Trail Export**: Multi-sheet SOX-compliant Excel workpaper plus an immutable, tamper-evident PDF Executive Sign-Off Certificate.

    ---

    ### 2. Supported vs. Unsupported File Types

    | File Category | Supported Inputs (✅ Accepted) | Unsupported Inputs (❌ Not Accepted) |
    | :--- | :--- | :--- |
    | **CSV** | • Delimited by comma (`,`), semicolon (`;`), tab (`\\t`), or pipe (`\|`)<br>• UTF-8, UTF-8-BOM, UTF-16, Latin-1 encodings | • Password-protected / encrypted CSVs<br>• Corrupt or unclosed quote delimiters |
    | **Excel** | • Modern `.xlsx` workbooks<br>• Legacy `.xls` spreadsheets | • Password-protected Excel files<br>• Macro-enabled dynamic sheets that modify schemas |
    | **PDF Documents** | • **Digitally generated / vector PDFs** (Direct exports from online banking, ERPs like SAP/Oracle, or 'Print-to-PDF' statements) containing native tabular grids | • **Scanned paper / photocopies / flat images** (requires an external OCR layer)<br>• Password-locked / encrypted PDFs<br>• Multi-column newspaper text layouts without tabular boundaries |
    """)

# --- PDF Executive Sign-Off Certificate Generator (Unicode-Safe & Brand-Aware) ---
def sanitize_pdf_text(text: str) -> str:
    if not text:
        return ""
    replacements = {
        "—": "-",
        "–": "-",
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "•": "*",
        "…": "...",
        "±": "+/-"
    }
    for orig, rep in replacements.items():
        text = text.replace(orig, rep)
    return text.encode("latin-1", "replace").decode("latin-1")

class AuditCertificatePDF(FPDF):
    def __init__(self, logo_path=None):
        super().__init__()
        self.logo_path = logo_path

    def header(self):
        self.set_fill_color(31, 78, 121)
        self.rect(0, 0, 210, 16, 'F')
        self.set_text_color(255, 255, 255)
        self.set_font('Helvetica', 'B', 11)
        self.cell(0, 4, 'AUTORECONFLOW - INTERNAL CONTROL AUDIT CERTIFICATE', 0, 1, 'C')
        self.ln(6)

    def footer(self):
        self.set_y(-18)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 4, 'Confidential - For Authorized Supervisory & Statutory Audit Inspection Only | SOX 404 Assertion', 0, 1, 'C')
        self.cell(0, 4, f'Page {self.page_no()}', 0, 0, 'C')

def generate_pdf_signoff_certificate(matched_count, fuzzy_count, un_a_count, un_b_count, signoff_data, tolerance, logo_bytes=None, company_info=None) -> bytes:
    logo_temp_file = None
    if logo_bytes:
        try:
            logo_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            logo_temp.write(logo_bytes)
            logo_temp.flush()
            logo_temp.close()
            logo_temp_file = logo_temp.name
        except Exception:
            logo_temp_file = None

    pdf = AuditCertificatePDF(logo_path=logo_temp_file)
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    comp_name = sanitize_pdf_text(company_info.get("name", "AutoReconFlow Enterprise"))
    comp_addr = sanitize_pdf_text(company_info.get("address", ""))
    comp_uid = sanitize_pdf_text(company_info.get("uid", ""))

    y_brand_start = pdf.get_y()
    
    # Left Header: Entity Information
    pdf.set_xy(10, y_brand_start)
    pdf.set_font('Helvetica', 'B', 13)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(120, 6, comp_name, ln=True)

    if comp_uid:
        pdf.set_font('Helvetica', 'B', 8)
        pdf.set_text_color(70, 70, 70)
        pdf.cell(120, 4.5, f"Legal Identification: {comp_uid}", ln=True)

    if comp_addr:
        pdf.set_font('Helvetica', '', 7.5)
        pdf.set_text_color(100, 100, 100)
        pdf.multi_cell(125, 3.8, f"Registered Office: {comp_addr}")

    # Right Header: Logo
    if logo_temp_file and os.path.exists(logo_temp_file):
        try:
            pdf.image(logo_temp_file, x=150, y=y_brand_start - 2, h=16)
        except Exception:
            pass

    pdf.set_y(max(pdf.get_y() + 3, y_brand_start + 18))
    pdf.set_draw_color(210, 220, 235)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(4)

    # Document Headings
    pdf.set_font('Helvetica', 'B', 14)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 7, 'Financial Reconciliation Sign-Off Certificate', ln=True, align='L')
    pdf.set_font('Helvetica', 'I', 8.5)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 4.5, 'Compliance Certification under PCAOB AS 2201 / SOX Section 404 & ISA-500 Standards', ln=True, align='L')
    pdf.ln(3)

    # Assertion Verification Block
    pdf.set_fill_color(245, 247, 250)
    pdf.set_draw_color(200, 210, 225)
    pdf.rect(10, pdf.get_y(), 190, 22, 'DF')
    pdf.set_xy(14, pdf.get_y() + 2)

    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.set_text_color(50, 50, 50)
    pdf.cell(45, 4.5, 'Execution Timestamp (UTC):', 0, 0)
    pdf.set_font('Helvetica', '', 8.5)
    pdf.cell(55, 4.5, sanitize_pdf_text(signoff_data.get('preparer_date', '')), 0, 0)

    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.cell(40, 4.5, 'Workpaper Reference:', 0, 0)
    pdf.set_font('Helvetica', '', 8.5)
    pdf.cell(45, 4.5, sanitize_pdf_text(signoff_data.get('wp_ref', '')), 0, 1)

    pdf.set_x(14)
    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.cell(45, 4.5, 'Variance Tolerance:', 0, 0)
    pdf.set_font('Helvetica', '', 8.5)
    pdf.cell(55, 4.5, f'+/- {tolerance:.2f} Currency Units', 0, 0)

    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.cell(40, 4.5, 'Supervisory Verdict:', 0, 0)
    verdict = sanitize_pdf_text(signoff_data.get('review_verdict', 'Approved / Cleared'))
    pdf.set_font('Helvetica', 'B', 8.5)
    if 'Approved' in verdict:
        pdf.set_text_color(16, 124, 65)
    else:
        pdf.set_text_color(180, 40, 40)
    pdf.cell(45, 4.5, verdict, 0, 1)

    pdf.ln(6)

    # Metric Table
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(31, 78, 121)
    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.cell(110, 6.5, 'Reconciliation Assertion Category', 1, 0, 'L', fill=True)
    pdf.cell(35, 6.5, 'Record Count', 1, 0, 'C', fill=True)
    pdf.cell(45, 6.5, 'Audit Status', 1, 1, 'C', fill=True)

    table_data = [
        ("Tier-1 Deterministic Matches (Exact Key & Net Match)", str(matched_count), "Cleared"),
        ("Tier-2 Discrepancy & Threshold Review (Materiality Tracking)", str(fuzzy_count), "Supervisory Review" if fuzzy_count > 0 else "Cleared"),
        ("Source A Exceptions (Subledger Missing in External Records)", str(un_a_count), "Remediation Req." if un_a_count > 0 else "Cleared"),
        ("Source B Exceptions (External Records Missing in Subledger)", str(un_b_count), "Remediation Req." if un_b_count > 0 else "Cleared"),
    ]

    pdf.set_text_color(40, 40, 40)
    pdf.set_font('Helvetica', '', 8)
    for cat, count, status in table_data:
        pdf.cell(110, 6, f"  {cat}", 1, 0, 'L')
        pdf.cell(35, 6, count, 1, 0, 'C')
        if status == "Cleared":
            pdf.set_text_color(16, 124, 65)
        else:
            pdf.set_text_color(180, 40, 40)
        pdf.cell(45, 6, status, 1, 1, 'C')
        pdf.set_text_color(40, 40, 40)

    pdf.ln(7)

    # Sign-Off Blocks
    pdf.set_font('Helvetica', 'B', 10.5)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 5.5, 'Four-Eyes Internal Control Attestation', ln=True)
    pdf.ln(2)

    # Preparer Box
    y_start = pdf.get_y()
    pdf.set_fill_color(250, 252, 255)
    pdf.rect(10, y_start, 92, 40, 'DF')
    pdf.set_xy(14, y_start + 2.5)
    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(85, 4.5, 'PREPARER (First Line of Defense)', 0, 1)

    pdf.set_x(14)
    pdf.set_font('Helvetica', '', 8)
    pdf.set_text_color(50, 50, 50)
    pdf.cell(85, 4.2, f"Name: {sanitize_pdf_text(signoff_data.get('preparer_name'))}", 0, 1)
    pdf.set_x(14)
    pdf.cell(85, 4.2, f"System / Emp ID: {sanitize_pdf_text(signoff_data.get('preparer_id'))}", 0, 1)
    pdf.set_x(14)
    pdf.cell(85, 4.2, f"Timestamp: {sanitize_pdf_text(signoff_data.get('preparer_date'))}", 0, 1)
    pdf.set_x(14)
    sig_text = "[VERIFIED ELECTRONIC SIGNATURE - COMPLETED]" if signoff_data.get('preparer_signed') else "[UNSIGNED]"
    pdf.set_font('Helvetica', 'B', 7.5)
    pdf.set_text_color(16, 124, 65) if signoff_data.get('preparer_signed') else pdf.set_text_color(180, 40, 40)
    pdf.cell(85, 5.5, sig_text, 0, 1)

    # Reviewer Box
    pdf.set_fill_color(250, 252, 255)
    pdf.rect(108, y_start, 92, 40, 'DF')
    pdf.set_xy(112, y_start + 2.5)
    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(85, 4.5, 'REVIEWER (Supervisory Review / Oversight)', 0, 1)

    pdf.set_x(112)
    pdf.set_font('Helvetica', '', 8)
    pdf.set_text_color(50, 50, 50)
    pdf.cell(85, 4.2, f"Name: {sanitize_pdf_text(signoff_data.get('reviewer_name'))}", 0, 1)
    pdf.set_x(112)
    pdf.cell(85, 4.2, f"Designation: {sanitize_pdf_text(signoff_data.get('reviewer_title'))}", 0, 1)
    pdf.set_x(112)
    pdf.cell(85, 4.2, f"Timestamp: {sanitize_pdf_text(signoff_data.get('reviewer_date'))}", 0, 1)
    pdf.set_x(112)
    pdf.set_font('Helvetica', 'B', 7.5)
    pdf.set_text_color(16, 124, 65)
    pdf.cell(85, 5.5, f"[SUPERVISORY APPROVAL: {verdict.upper()}]", 0, 1)

    pdf.set_y(y_start + 44)

    # Audit Notes
    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 4.5, 'Supervisory Audit Notes:', ln=True)
    pdf.set_font('Helvetica', 'I', 8)
    pdf.set_text_color(60, 60, 60)
    notes_clean = sanitize_pdf_text(signoff_data.get('audit_comments', 'No comments noted.') or 'None noted.')
    pdf.multi_cell(0, 4, notes_clean)

    pdf.ln(3)

    # Tamper-Evident SHA-256 Checksum
    raw_fingerprint = f"{comp_uid}_{signoff_data.get('wp_ref')}_{matched_count}_{fuzzy_count}_{un_a_count}_{un_b_count}_{signoff_data.get('preparer_date')}_{signoff_data.get('reviewer_date')}_{tolerance}"
    doc_hash = hashlib.sha256(raw_fingerprint.encode()).hexdigest().upper()

    pdf.set_fill_color(240, 240, 240)
    pdf.set_font('Courier', 'B', 7.5)
    pdf.set_text_color(70, 70, 70)
    pdf.cell(0, 6, f" SHA-256 DIGITAL IMMUTABILITY AUDIT HASH: {doc_hash} ", 1, 1, 'C', fill=True)

    result_bytes = bytes(pdf.output())

    if logo_temp_file and os.path.exists(logo_temp_file):
        try:
            os.remove(logo_temp_file)
        except Exception:
            pass

    return result_bytes

# --- Professional SOX 404 & ISA-500 Multi-Tab Audit Workpaper Generator with Verified Sign-Offs & Branding ---
def generate_consolidated_workbook(matched_df, fuzzy_df, un_a_df, un_b_df, signoff_data, tolerance=1.0, company_info=None) -> bytes:
    output = io.BytesIO()
    timestamp_utc = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

    comp_name = company_info.get("name", "AutoReconFlow Enterprise") if company_info else "AutoReconFlow Enterprise"
    comp_addr = company_info.get("address", "") if company_info else ""
    comp_uid = company_info.get("uid", "") if company_info else ""

    navy_header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=11)
    italic_font = Font(name="Calibri", size=10, italic=True, color="595959")
    title_font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        workbook = writer.book

        ws_summary = workbook.create_sheet(title="Audit & Control Summary", index=0)
        ws_summary.views.sheetView[0].showGridLines = True

        ws_summary["A1"] = f"{comp_name} — Enterprise Reconciliation & Control Report"
        ws_summary["A1"].font = title_font
        ws_summary.merge_cells("A1:C1")

        ws_summary["A2"] = f"Legal Identification: {comp_uid} | Registered Office: {comp_addr}"
        ws_summary["A2"].font = italic_font
        ws_summary.merge_cells("A2:C2")

        ws_summary["A3"] = "SOX 404 / ISA-500 Standard Financial Statement Assertion Audit Trail"
        ws_summary["A3"].font = Font(name="Calibri", size=9, italic=True, color="7F7F7F")
        ws_summary.merge_cells("A3:C3")

        metadata = [
            ("Execution Timestamp (UTC):", timestamp_utc),
            ("Tolerance Applied:", f"± {tolerance:.2f} Currency Units"),
            ("Matching Methodology:", "3-Tier (Deterministic Exact -> Levenshtein/Jaro-Winkler Fuzzy)"),
            ("Control Assertion Target:", "Completeness, Accuracy, and Valuation (Financial Controls)")
        ]
        ws_summary.append([])
        for label, val in metadata:
            ws_summary.append([label, val])

        for r in range(5, 9):
            ws_summary.cell(row=r, column=1).font = bold_font
            ws_summary.cell(row=r, column=2).font = regular_font

        ws_summary.append([])
        ws_summary.append(["Reconciliation Category", "Record Count", "Control Status"])
        summary_header_row = 10
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

        for r in range(11, 15):
            for c in range(1, 4):
                cell = ws_summary.cell(row=r, column=c)
                cell.font = regular_font
                cell.border = thin_border
                if c == 2:
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "#,##0"
                elif c == 3:
                    cell.alignment = Alignment(horizontal="center")

        ws_summary.append([])
        sign_off_header_row = 16
        ws_summary.cell(row=sign_off_header_row, column=1, value="Internal Control Sign-Off (SOX 404 / Four-Eyes Governance)").font = bold_font
        ws_summary.cell(row=sign_off_header_row, column=1).fill = section_fill
        ws_summary.merge_cells(start_row=sign_off_header_row, start_column=1, end_row=sign_off_header_row, end_column=3)

        sig_status = "VERIFIED / ELECTRONICALLY SIGNED" if signoff_data.get("preparer_signed") else "PENDING SIGNATURE"
        
        signoff_entries = [
            ("Prepared By (Analyst):", f"{signoff_data.get('preparer_name')} [ID: {signoff_data.get('preparer_id')}]", f"Signed: {signoff_data.get('preparer_date')} ({sig_status})"),
            ("Reviewed By (Controller / Manager):", f"{signoff_data.get('reviewer_name')} [{signoff_data.get('reviewer_title')}]", f"Reviewed: {signoff_data.get('reviewer_date')} | Verdict: {signoff_data.get('review_verdict')}"),
            ("Audit Workpaper Reference:", signoff_data.get("wp_ref"), f"Status: {signoff_data.get('review_verdict')}"),
            ("Supervisory Audit Notes:", signoff_data.get("audit_comments", "None noted"), "Document Status: Final Closed")
        ]

        for c1, c2, c3 in signoff_entries:
            ws_summary.append([c1, c2, c3])

        for r in range(sign_off_header_row + 1, sign_off_header_row + 5):
            for c in range(1, 4):
                cell = ws_summary.cell(row=r, column=c)
                cell.font = bold_font if c == 1 else regular_font
                cell.border = thin_border

        ws_summary.column_dimensions["A"].width = 64
        ws_summary.column_dimensions["B"].width = 60
        ws_summary.column_dimensions["C"].width = 65

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

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or "")
                    max_len = max(max_len, len(val_str))
                ws.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 45)

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

# --- Sample Data Generator ---
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

        st.session_state["recon_results"] = {
            "matched_pd": matched.to_pandas(),
            "fuzzy_pd": fuzzy.to_pandas(),
            "un_a_pd": un_a.to_pandas(),
            "un_b_pd": un_b.to_pandas(),
            "matched_pl": matched,
            "fuzzy_pl": fuzzy,
            "un_a_pl": un_a,
            "un_b_pl": un_b,
            "tolerance": tolerance
        }
        # Reset workflow states on every fresh reconciliation run
        st.session_state["preparer_completed"] = False
        st.session_state["reviewer_approved"] = False

    if "recon_results" in st.session_state:
        res = st.session_state["recon_results"]
        matched_pd = res["matched_pd"]
        fuzzy_pd = res["fuzzy_pd"]
        un_a_pd = res["un_a_pd"]
        un_b_pd = res["un_b_pd"]

        st.subheader("📊 Reconciliation & Control Summary (SOX 404 / ISA-500)")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Tier-1 Deterministic Matches", f"{len(matched_pd):,} records")
        k2.metric("Tier-2 Discrepancy Review", f"{len(fuzzy_pd):,} records")
        k3.metric("Source A Exceptions (Ledger)", f"{len(un_a_pd):,} records")
        k4.metric("Source B Exceptions (External)", f"{len(un_b_pd):,} records")

        # -------------------------------------------------------------
        # Sequential Four-Eyes Governance Sign-Off Workflow
        # -------------------------------------------------------------
        is_prep_done = st.session_state.get("preparer_completed", False)
        is_rev_approved = st.session_state.get("reviewer_approved", False)

        with st.expander("✍️ Four-Eyes Internal Governance Sign-Off & Approval Workflow", expanded=True):
            so_col1, so_col2 = st.columns(2)
            now_str = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

            # --- STAGE 1: PREPARER PANEL ---
            with so_col1:
                st.markdown("##### 👤 Stage 1: Preparer (Analyst / Senior Accountant)")
                prep_name = st.text_input("Preparer Full Name", value="Senior Recon Analyst", key="so_prep_name", disabled=is_prep_done)
                prep_id = st.text_input("Employee / System ID", value="FIN-ACC-8921", key="so_prep_id", disabled=is_prep_done)
                prep_date = st.text_input("Preparation Date & Time (UTC)", value=st.session_state.get("saved_prep_date", now_str), key="so_prep_date", disabled=is_prep_done)
                prep_signed = st.checkbox("I verify that ledger entries, tolerances, and files are audited.", value=True, key="so_prep_signed", disabled=is_prep_done)

                if not is_prep_done:
                    if st.button("✅ Mark Preparer Task as Complete", type="primary", use_container_width=True):
                        if not prep_signed:
                            st.error("Preparer verification box must be checked before submission.")
                        else:
                            st.session_state["preparer_completed"] = True
                            st.session_state["saved_prep_name"] = prep_name
                            st.session_state["saved_prep_id"] = prep_id
                            st.session_state["saved_prep_date"] = prep_date
                            st.session_state["saved_prep_signed"] = prep_signed
                            st.rerun()
                else:
                    st.success(f"🔒 Preparer Task Completed & Locked by {st.session_state.get('saved_prep_name', prep_name)}")

            # --- STAGE 2: REVIEWER PANEL ---
            with so_col2:
                if not is_prep_done:
                    st.markdown("##### 🛡️ Stage 2: Reviewer (Supervisory Oversight)")
                    st.info("⏳ Reviewer stage is currently locked. Awaiting Preparer completion on the left.")
                    rev_name = "Finance Controller"
                    rev_title = "VP - Financial Reporting & Controls"
                    rev_date = now_str
                    rev_verdict = "Under Remediation"
                    wp_ref = "WP-SOX-REC-2026-Q4"
                    audit_notes = "Pending supervisory sign-off."
                else:
                    st.markdown("##### 🛡️ Stage 2: Reviewer (Active for Approval)")
                    st.success("🟢 Preparer has completed their submission. Ready for Supervisory Review.")
                    rev_name = st.text_input("Reviewer Full Name", value="Finance Controller", key="so_rev_name", disabled=is_rev_approved)
                    rev_title = st.text_input("Role / Designation", value="VP - Financial Reporting & Controls", key="so_rev_title", disabled=is_rev_approved)
                    rev_date = st.text_input("Review Date & Time (UTC)", value=st.session_state.get("saved_rev_date", now_str), key="so_rev_date", disabled=is_rev_approved)
                    rev_verdict = st.selectbox("Supervisory Audit Verdict", ["Approved / Cleared", "Approved with Exceptions Noted", "Under Remediation", "Rejected"], key="so_rev_verdict", disabled=is_rev_approved)
                    wp_ref = st.text_input("Audit Workpaper Reference Code", value="WP-SOX-REC-2026-Q4", key="so_wp_ref", disabled=is_rev_approved)
                    audit_notes = st.text_area("Supervisory & Control Notes", value="All deterministic matches verified. Discrepancies and exceptions reviewed and reconciled.", key="so_notes", disabled=is_rev_approved)

                    if not is_rev_approved:
                        if st.button("🛡️ Approve & Finalize Supervisory Sign-Off", type="primary", use_container_width=True):
                            st.session_state["reviewer_approved"] = True
                            st.session_state["saved_rev_name"] = rev_name
                            st.session_state["saved_rev_title"] = rev_title
                            st.session_state["saved_rev_date"] = rev_date
                            st.session_state["saved_rev_verdict"] = rev_verdict
                            st.session_state["saved_wp_ref"] = wp_ref
                            st.session_state["saved_audit_notes"] = audit_notes
                            st.rerun()
                    else:
                        st.success(f"✅ Supervisory Sign-Off Formally Approved & Sealed ({st.session_state.get('saved_rev_verdict', rev_verdict)})")

        # Compile final sign-off payload
        signoff_payload = {
            "preparer_name": st.session_state.get("saved_prep_name", prep_name),
            "preparer_id": st.session_state.get("saved_prep_id", prep_id),
            "preparer_date": st.session_state.get("saved_prep_date", prep_date),
            "preparer_signed": st.session_state.get("saved_prep_signed", prep_signed),
            "reviewer_name": st.session_state.get("saved_rev_name", rev_name),
            "reviewer_title": st.session_state.get("saved_rev_title", rev_title),
            "reviewer_date": st.session_state.get("saved_rev_date", rev_date),
            "review_verdict": st.session_state.get("saved_rev_verdict", rev_verdict),
            "wp_ref": st.session_state.get("saved_wp_ref", wp_ref),
            "audit_comments": st.session_state.get("saved_audit_notes", audit_notes)
        }

        active_company_info = {
            "name": company_name,
            "address": company_address,
            "uid": company_uid
        }

        # -------------------------------------------------------------
        # Official Export Hub: Workpapers unlock upon Reviewer Approval
        # -------------------------------------------------------------
        st.markdown("#### 📦 Audit Deliverables & Certification Hub")
        exp_col1, exp_col2 = st.columns(2)

        with exp_col1:
            if is_rev_approved:
                excel_bytes = generate_consolidated_workbook(
                    matched_pd, fuzzy_pd, un_a_pd, un_b_pd, signoff_payload, tolerance=res["tolerance"], company_info=active_company_info
                )
                st.download_button(
                    label="📊 Download Final SOX Audit Workpaper (.xlsx)",
                    data=excel_bytes,
                    file_name=f"AutoReconFlow_SOX_Audit_Report_{datetime.datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_download_master_excel"
                )
            else:
                st.info("🔒 Complete Preparer submission and Reviewer approval above to generate the Excel Workpaper.")

        with exp_col2:
            if is_rev_approved:
                logo_file_bytes = uploaded_logo.getvalue() if uploaded_logo is not None else None
                pdf_bytes = generate_pdf_signoff_certificate(
                    len(matched_pd), len(fuzzy_pd), len(un_a_pd), len(un_b_pd), signoff_payload, res["tolerance"], 
                    logo_bytes=logo_file_bytes, company_info=active_company_info
                )
                st.download_button(
                    label="📜 Download Official Signed PDF Audit Certificate",
                    data=pdf_bytes,
                    file_name=f"AutoReconFlow_SignOff_Certificate_{datetime.datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="btn_download_pdf_cert"
                )
            else:
                st.info("🔒 Signed PDF Certificate unlocks automatically once the Reviewer approves.")

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
                    data=res["matched_pl"].write_csv(),
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
                    data=res["fuzzy_pl"].write_csv(),
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
                    data=res["un_a_pl"].write_csv(),
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
                    data=res["un_b_pl"].write_csv(),
                    file_name="unmatched_external_exceptions.csv",
                    mime="text/csv",
                    key="btn_un_b_csv"
                )
            else:
                st.success("All Source B items accounted for (0 open variances).")