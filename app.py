import io
import re
import datetime
import hashlib
import tempfile
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
import pandas as pd
import polars as pl
from fpdf import FPDF
from recon_engine import execute_high_volume_reconciliation
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- Page Configuration & Professional Executive Styling ---
st.set_page_config(
    page_title="AutoReconFlow — Enterprise Financial Reconciliation, Tax & Audit Engine",
    page_icon="⚖️",
    layout="wide"
)

# Executive Institutional Palette (Wall Street Navy & Slate Theme)
st.markdown("""
<style>
    /* Global Base Typography & Smooth Rendering */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        color: #1e293b;
    }
    
    /* Top Header Branding Container */
    .brand-hero {
        background: linear-gradient(135deg, #0A2540 0%, #1B365D 60%, #204066 100%);
        color: #ffffff;
        padding: 24px 30px;
        border-radius: 10px;
        margin-bottom: 24px;
        box-shadow: 0 4px 14px rgba(10, 37, 64, 0.15);
    }
    .brand-hero h1 {
        color: #ffffff !important;
        font-weight: 700;
        font-size: 26px;
        margin: 0 0 6px 0;
        letter-spacing: -0.5px;
    }
    .brand-hero p {
        color: #94a3b8 !important;
        font-size: 14px;
        margin: 0;
    }

    /* Metric Summary Badges */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 14px 18px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        border-left: 4px solid #1B365D;
    }
    div[data-testid="stMetricLabel"] p {
        font-size: 13px;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    div[data-testid="stMetricValue"] div {
        color: #0f172a;
        font-weight: 700;
        font-size: 22px;
    }

    /* Sidebar Refinement */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    
    /* Primary Action Buttons */
    button[kind="primary"] {
        background: linear-gradient(180deg, #1B365D 0%, #0A2540 100%) !important;
        border: none !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        padding: 8px 18px !important;
        border-radius: 6px !important;
        box-shadow: 0 2px 6px rgba(10, 37, 64, 0.2) !important;
        transition: all 0.2s ease;
    }
    button[kind="primary"]:hover {
        background: linear-gradient(180deg, #23426e 0%, #102d4d 100%) !important;
        box-shadow: 0 4px 12px rgba(10, 37, 64, 0.3) !important;
        transform: translateY(-1px);
    }
    
    /* Professional DataFrames / Tables */
    div[data-testid="stDataFrame"] {
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# --- License Verification Engine (Deterministic SHA-256) ---
def verify_license_key(email: str, license_key: str):
    """Validates the license key against the email, plan prefix, and monthly salt."""
    if not email or not license_key:
        return False, None
        
    email_clean = email.strip().lower()
    key_clean = license_key.strip().upper()
    salt = "AutoReconFlowSecretKey2026"
    
    tier_map = {
        "ARF-PRAC": "Practitioner",
        "ARF-BIZ": "Business & Tax Suite",
        "ARF-ENT": "Enterprise Turnkey"
    }
    
    parts = key_clean.split("-")
    if len(parts) != 5:
        return False, None
        
    prefix = f"{parts[0]}-{parts[1]}"
    if prefix not in tier_map:
        return False, None
        
    tier_name = tier_map[prefix]
    
    # Check current month and previous month for validity
    current_date = datetime.datetime.utcnow()
    prev_month_date = current_date.replace(day=1) - datetime.timedelta(days=1)
    
    months_to_test = [
        current_date.strftime("%Y%m"),
        prev_month_date.strftime("%Y%m")
    ]
    
    for m in months_to_test:
        raw_str = f"{email_clean}:{tier_name}:{salt}:{m}"
        h = hashlib.sha256(raw_str.encode()).hexdigest().upper()
        expected_key = f"{prefix}-{h[:4]}-{h[4:8]}-{h[8:12]}"
        if key_clean == expected_key:
            return True, tier_name
            
    return False, None

# --- Pure Python Token-Free Email Dispatcher ---
def send_workflow_email(smtp_settings, recipient_email, subject, plain_text, html_content=None):
    """Sends notification emails using pure Python smtplib with zero token cost."""
    if not smtp_settings.get("enabled", False):
        return True, "Simulation mode active (SMTP disabled)."
    
    server_host = smtp_settings.get("host", "smtp.gmail.com").strip()
    server_port = int(smtp_settings.get("port", 465))
    sender_user = smtp_settings.get("user", "").strip()
    sender_pass = smtp_settings.get("password", "").strip()

    if not sender_user or not sender_pass:
        return False, "SMTP sender username or password is missing."
    if not recipient_email or "@" not in recipient_email:
        return False, f"Invalid recipient email address: '{recipient_email}'"

    try:
        msg = MIMEMultipart("alternative")
        msg['From'] = f"AutoReconFlow Audit <{sender_user}>"
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # Attach Plain Text fallback
        msg.attach(MIMEText(plain_text, 'plain'))
        
        # Attach Rich HTML layout if provided
        if html_content:
            msg.attach(MIMEText(html_content, 'html'))

        if server_port == 465:
            with smtplib.SMTP_SSL(server_host, server_port, timeout=12) as server:
                server.login(sender_user, sender_pass)
                server.send_message(msg)
        else:
            with smtplib.SMTP(server_host, server_port, timeout=12) as server:
                server.starttls()
                server.login(sender_user, sender_pass)
                server.send_message(msg)
        return True, "Email successfully delivered."
    except smtplib.SMTPAuthenticationError:
        return False, "SMTP Authentication Failed: Incorrect email or Google App Password."
    except Exception as e:
        return False, f"SMTP Dispatch Error: {str(e)}"

def build_html_email(title, header_desc, details_dict, action_button_url="https://www.autoreconflow.com"):
    """Builds a responsive, corporate-styled HTML email."""
    rows_html = "".join([
        f"<tr><td style='padding:9px;font-weight:600;color:#1B365D;width:35%;border-bottom:1px solid #E2E8F0;'>{k}</td>"
        f"<td style='padding:9px;color:#334155;border-bottom:1px solid #E2E8F0;'>{v}</td></tr>"
        for k, v in details_dict.items()
    ])
    
    return f"""
    <html>
      <body style="font-family:'Segoe UI',Helvetica,Arial,sans-serif;margin:0;padding:24px;background-color:#F8FAFC;">
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="600" style="background:#ffffff;border:1px solid #CBD5E1;border-radius:8px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,0.06);">
          <tr>
            <td bgcolor="#1B365D" style="padding:20px 28px;color:#ffffff;">
              <h2 style="margin:0;font-size:19px;font-weight:700;">⚖️ AutoReconFlow | Financial Governance Notice</h2>
            </td>
          </tr>
          <tr>
            <td style="padding:28px;">
              <h3 style="margin-top:0;color:#0F172A;font-size:17px;">{title}</h3>
              <p style="color:#475569;font-size:14px;line-height:1.6;">{header_desc}</p>
              <table width="100%" cellpadding="0" cellspacing="0" style="margin:20px 0;border-collapse:collapse;font-size:13px;">
                {rows_html}
              </table>
              <div style="text-align:center;margin-top:28px;">
                <a href="{action_button_url}" style="background-color:#1B365D;color:#ffffff;text-decoration:none;padding:12px 26px;border-radius:5px;font-weight:600;display:inline-block;box-shadow:0 2px 6px rgba(27,54,93,0.3);">Open Workpaper & Verify Record</a>
              </div>
            </td>
          </tr>
          <tr>
            <td bgcolor="#F1F5F9" style="padding:14px 28px;font-size:11px;color:#64748B;text-align:center;">
              SOX 404 / ISA-500 Four-Eyes Governance Automation | Tamper-Evident SHA-256 Audit Trails
            </td>
          </tr>
        </table>
      </body>
    </html>
    """

# --- Sidebar Configuration & Commercial Hub ---
st.sidebar.markdown("### ⚡ AutoReconFlow")
st.sidebar.markdown("**Enterprise B2B Edition v2.0**")
st.sidebar.caption("SOX 404 / ISA-500 & SOC 1 Type II Governance Architecture")
st.sidebar.divider()

# Organization White-Label Profile
st.sidebar.markdown("#### 🏢 Client Branding Configuration")
uploaded_logo = st.sidebar.file_uploader("Upload Corporate Logo (PNG/JPG)", type=["png", "jpg", "jpeg"], key="company_logo")
if uploaded_logo:
    st.sidebar.image(uploaded_logo, caption="Active Workpaper Branding", use_container_width=True)

company_name = st.sidebar.text_input("Company / Legal Entity Name", value="Global Enterprise Solutions Ltd.", key="cfg_comp_name")
company_address = st.sidebar.text_area("Registered Office Address", value="Tower 4, Financial District, Cyber City, Gurugram, India - 122002", key="cfg_comp_addr", height=70)
company_uid = st.sidebar.text_input("Unique Identification No. (CIN / LEI / Tax ID)", value="CIN: U72200HR2026PTC109823", key="cfg_comp_uid")
st.sidebar.divider()

# Email Notification Settings with Connection Verification
with st.sidebar.expander("📬 Automated Email Alert Settings", expanded=False):
    st.caption("Zero-token SMTP engine. Compatible with Gmail App Passwords, Outlook, and Brevo.")
    enable_smtp = st.checkbox("Enable Live Email Dispatch", value=False, key="smtp_enabled")
    
    default_host = st.secrets.get("SMTP_HOST", "smtp.gmail.com") if hasattr(st, "secrets") else "smtp.gmail.com"
    default_port = str(st.secrets.get("SMTP_PORT", "465")) if hasattr(st, "secrets") else "465"
    default_user = st.secrets.get("SMTP_USER", "") if hasattr(st, "secrets") else ""
    default_pass = st.secrets.get("SMTP_PASSWORD", "") if hasattr(st, "secrets") else ""

    smtp_host = st.text_input("SMTP Server Host", value=default_host, key="smtp_host")
    smtp_port = st.text_input("SMTP Port (465=SSL / 587=TLS)", value=default_port, key="smtp_port")
    smtp_user = st.text_input("Sender Email Address", value=default_user, placeholder="yourname@gmail.com", key="smtp_user")
    smtp_pass = st.text_input("Sender App Password", value=default_pass, type="password", placeholder="16-character app password", key="smtp_pass")

    smtp_cfg = {
        "enabled": enable_smtp,
        "host": smtp_host,
        "port": smtp_port,
        "user": smtp_user,
        "password": smtp_pass
    }

    test_target = st.text_input("Test Target Email", value=smtp_user or "", placeholder="recipient@domain.com", key="smtp_test_target")
    if st.button("🧪 Send Test Ping Email", use_container_width=True):
        if not enable_smtp:
            st.warning("Please check 'Enable Live Email Dispatch' first.")
        else:
            with st.spinner("Connecting to SMTP server..."):
                ok, feedback = send_workflow_email(
                    smtp_cfg, 
                    test_target, 
                    "AutoReconFlow - SMTP Connection Verified", 
                    "This confirms your AutoReconFlow email trigger configuration is working properly.",
                    build_html_email("SMTP Connection Verified", "Your AutoReconFlow email alerts are configured properly and operational.", {"Status": "Operational", "Tested At": datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")})
                )
                if ok:
                    st.success("✅ Test email dispatched successfully! Check your inbox.")
                else:
                    st.error(f"❌ {feedback}")

st.sidebar.divider()

# --- Interactive In-App License Key Verification Widget ---
with st.sidebar.expander("🔑 License Activation & Verification", expanded=True):
    user_email_input = st.text_input("Registered Email", placeholder="buyer@clientdomain.com", key="lic_user_email")
    user_key_input = st.text_input("License Key", placeholder="ARF-PRAC-XXXX-XXXX-XXXX", key="lic_user_key")
    
    if user_email_input and user_key_input:
        is_valid, active_tier = verify_license_key(user_email_input, user_key_input)
        if is_valid:
            st.success(f"✅ Verified: {active_tier}")
            st.session_state["licensed"] = True
            st.session_state["license_tier"] = active_tier
        else:
            st.error("❌ Invalid License Key or Email Mismatch.")
            st.session_state["licensed"] = False
            st.session_state["license_tier"] = None
    else:
        st.session_state["licensed"] = False
        st.session_state["license_tier"] = None
        st.caption("Enter your registered credentials to unlock full audit features.")

st.sidebar.divider()

# High-Converting 3-Tier Commercial Licensing Hub
st.sidebar.markdown("### 💼 Commercial Licensing & Deployment")
st.sidebar.markdown(
    """
    **Enterprise Financial Reconciliation, Tax & Audit Suite**
    
    * **Practitioner License (₹6,999)**: Single-seat desktop deployment, unlimited matching, automated email alerts, and SHA-256 PDF certificates.
    * **Business & Tax Suite (₹16,999)**: Subledger-to-bank & GST/VAT portal modules, custom white-labeling, and full source code access.
    * **Enterprise Turnkey (₹39,999)**: Private cloud/server deployment, multi-user database integration, and compliance onboarding.
    """
)

# Live Razorpay Multi-Tier Checkout Button
st.sidebar.link_button(
    "💳 Purchase Commercial License (UPI / Cards / NetBanking)",
    "https://rzp.io/rzp/G2XPtSpy",
    use_container_width=True,
    type="primary"
)

st.sidebar.markdown(
    """
    *Need custom ERP/SAP integration or an enterprise demo?*  
    📧 **Contact Sales:** [contact@rakeshsingh.tech](mailto:contact@rakeshsingh.tech)  
    📞 **Direct Desk:** +91 7042218226
    """
)
st.sidebar.caption("Pure Python • 0 Token Costs • 100% On-Premise Data Privacy")
st.sidebar.divider()

# --- Main App Title & Subtitle (Unified Hybrid Hero) ---
st.markdown("""
<div class="brand-hero">
    <h1>⚖️ AutoReconFlow — Enterprise Financial Reconciliation, Tax & Audit Engine</h1>
    <p>High-throughput 3-tier financial matching, cross-border tax reconciliation, and SOX 404 / ISA-500 Four-Eyes governance.</p>
</div>
""", unsafe_allow_html=True)

# --- Comprehensive Enterprise Standard Operating Procedure & Technical Manual ---
with st.expander("📖 Comprehensive Procedure & Configuration Manual (SOX 404 / ISA-500 Standard)", expanded=False):
    st.markdown("""
    ### 🏛️ Section 1: Corporate Entity & Workpaper Branding Setup
    AutoReconFlow stamps corporate identity metadata across all audit deliverables (Multi-tab Excel workpapers and signed PDF certificates):
    * **Corporate Logo**: Upload a high-resolution PNG or JPG (transparent background recommended, ideal dimensions: 400x120 px). The logo is automatically rendered on the upper-right corner of the official PDF Sign-Off Certificate.
    * **Legal Entity Name**: Enter the full registered statutory entity name (e.g., *Global Enterprise Solutions Ltd.*). This appears on the title block of both Excel and PDF documents.
    * **Registered Office Address**: Provide the legal registered address. It will be stamped directly beneath the entity title in standard audit report format.
    * **Unique Identification Number (CIN / LEI / GSTIN / Tax ID)**: 
      * For India: Corporate Identification Number (CIN), GSTIN, or Permanent Account Number (PAN).
      * For US/Global: Legal Entity Identifier (LEI), Employer Identification Number (EIN), or VAT Registration Number.
      * *Audit Note*: This unique ID forms part of the immutable SHA-256 digital fingerprint.

    ---

    ### 📬 Section 2: Automated Email Dispatch & SMTP Mailbox Mapping Matrix
    AutoReconFlow features a zero-token, built-in SMTP engine (`smtplib`). No third-party paid API tokens are required. Configure your preferred mailbox using the exact mapping parameters below:

    #### Option A: Gmail / Google Workspace (Recommended for Quick Setup)
    * **SMTP Host**: `smtp.gmail.com`
    * **SMTP Port**: `465` (SSL)
    * **Username**: Your complete Gmail address (e.g., `finance.controller@gmail.com`)
    * **Password Generation (16-Character App Password)**:
      1. Log into your Google Account at **[myaccount.google.com/security](https://myaccount.google.com/security)**.
      2. Ensure **2-Step Verification** is active.
      3. In the top search bar, type **App passwords**.
      4. Enter `AutoReconFlow` as the app name and click **Create**.
      5. Copy the generated 16-character code (e.g., `abcd efgh ijkl mnop`) and paste it into the **Sender App Password** field.
      * *Security Note*: Standard Gmail login passwords will be blocked by Google. You must use an App Password.

    #### Option B: Microsoft 365 / Corporate Outlook
    * **SMTP Host**: `smtp.office365.com`
    * **SMTP Port**: `587` (TLS)
    * **Username**: Your corporate Office 365 email (e.g., `analyst@corporatedomain.com`)
    * **Password**: Microsoft 365 Account Password (or App Password if MFA is enforced by your IT administrator).
    * *IT Note*: Ensure *Authenticated SMTP* is permitted on the mailbox in the Microsoft 365 Admin Center.

    #### Option C: Brevo (Formerly Sendinblue) / Resend (High-Volume Free SaaS Tier)
    * **SMTP Host**: `smtp-relay.brevo.com`
    * **SMTP Port**: `587` (TLS)
    * **Username**: Your Brevo login email
    * **Password**: Master SMTP Key generated under *Transactional -> Settings -> Configuration*.
    * *Advantage*: Up to 300 free corporate emails per day with verified custom domain deliverability (SPF/DKIM).

    #### Option D: Custom Corporate / On-Premises Mail Server
    * **SMTP Host**: `mail.yourcompany.com` or internal IP (e.g., `10.0.0.25`)
    * **SMTP Port**: `465` (SSL) or `587` (TLS / STARTTLS)
    * **Authentication**: Provide standard service-account credentials authorized to relay internal audit notices.

    ---

    ### 🔄 Section 3: Four-Eyes Governance & Line-Item Audit Query SOP
    AutoReconFlow enforces strict segregation of duties (SoD) under PCAOB AS 2201 and SOX 404:

    ```
    [ Stage 1: Preparer Audit & Submission ]
          │
          ▼
    Preparer uploads Source A & Source B -> Runs matching engine -> Reviews metrics -> 
    Inputs Emp ID & checks verification box -> Clicks "Mark Preparer Task as Complete & Notify Reviewer"
          │
          ├──────────────────────────► [ Automated HTML Email sent to Reviewer ]
          │
    [ Stage 2: Supervisory Review & Line-Item Query ]
          │
          ▼
    Reviewer inspects exceptions. Two pathways available:
          ├── Pathway 1 (Direct Clearance): Selects verdict -> Clicks "Approve & Finalize Sign-Off"
          └── Pathway 2 (Line-Item Clarification Loop):
                1. Reviewer selects disputed transaction from dropdown.
                2. Enters audit observation (e.g., "Variance of 50.00 requires swift advice copy").
                3. Clicks "Send Query Back to Preparer".
                4. [ Automated Email dispatched to Preparer ].
                5. Preparer logs in, enters line-item clarification, and resubmits.
                6. [ Automated Email dispatched to Reviewer ].
                7. Reviewer accepts explanation and clicks "Accept Explanation & Formally Approve".
          │
    [ Stage 3: Audit Deliverable Sealing ]
          │
          ▼
    Excel Workpaper & Signed PDF Certificate unlock for download. 
    Full query dialogue is permanently stamped into the document body under the SHA-256 seal.
    ```

    ---

    ### 📁 Section 4: Data Ingestion & File Compatibility Guidelines

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
        self.set_fill_color(27, 54, 93)  # Professional Oxford Navy
        self.rect(0, 0, 210, 16, 'F')
        self.set_text_color(255, 255, 255)
        self.set_font('Helvetica', 'B', 11)
        self.cell(0, 4, 'AUTORECONFLOW - FINANCIAL RECONCILIATION, TAX & AUDIT CERTIFICATE', 0, 1, 'C')
        self.ln(6)

    def footer(self):
        self.set_y(-18)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 4, 'Confidential - For Authorized Supervisory & Statutory Audit Inspection Only | SOX 404 Assertion', 0, 1, 'C')
        self.cell(0, 4, f'Page {self.page_no()}', 0, 0, 'C')

def generate_pdf_signoff_certificate(matched_count, fuzzy_count, un_a_count, un_b_count, signoff_data, tolerance, logo_bytes=None, company_info=None, query_trail="") -> bytes:
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
    pdf.set_text_color(27, 54, 93)
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
    pdf.set_draw_color(203, 213, 225)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(4)

    # Document Headings
    pdf.set_font('Helvetica', 'B', 14)
    pdf.set_text_color(27, 54, 93)
    pdf.cell(0, 7, 'Financial Reconciliation & Tax Audit Certificate', ln=True, align='L')
    pdf.set_font('Helvetica', 'I', 8.5)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 4.5, 'Compliance Certification under PCAOB AS 2201 / SOX Section 404 & ISA-500 Standards', ln=True, align='L')
    pdf.ln(3)

    # Assertion Verification Block
    pdf.set_fill_color(248, 250, 252)
    pdf.set_draw_color(203, 213, 225)
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
    pdf.set_fill_color(27, 54, 93)
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
    pdf.set_text_color(27, 54, 93)
    pdf.cell(0, 5.5, 'Four-Eyes Internal Control Attestation', ln=True)
    pdf.ln(2)

    # Preparer Box
    y_start = pdf.get_y()
    pdf.set_fill_color(248, 250, 252)
    pdf.rect(10, y_start, 92, 40, 'DF')
    pdf.set_xy(14, y_start + 2.5)
    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.set_text_color(27, 54, 93)
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
    pdf.set_fill_color(248, 250, 252)
    pdf.rect(108, y_start, 92, 40, 'DF')
    pdf.set_xy(112, y_start + 2.5)
    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.set_text_color(27, 54, 93)
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

    # Audit Notes & Clarification Thread
    pdf.set_font('Helvetica', 'B', 8.5)
    pdf.set_text_color(27, 54, 93)
    pdf.cell(0, 4.5, 'Supervisory Audit Notes & Resolution Trail:', ln=True)
    pdf.set_font('Helvetica', 'I', 8)
    pdf.set_text_color(60, 60, 60)
    
    notes_clean = sanitize_pdf_text(signoff_data.get('audit_comments', 'No comments noted.') or 'None noted.')
    if query_trail:
        notes_clean += f" | {sanitize_pdf_text(query_trail)}"
    pdf.multi_cell(0, 4, notes_clean)

    pdf.ln(3)

    # Tamper-Evident SHA-256 Checksum
    raw_fingerprint = f"{comp_uid}_{signoff_data.get('wp_ref')}_{matched_count}_{fuzzy_count}_{un_a_count}_{un_b_count}_{signoff_data.get('preparer_date')}_{signoff_data.get('reviewer_date')}_{tolerance}_{query_trail}"
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

# --- Professional SOX 404 Multi-Tab Audit Workpaper Generator with Query Trail Column ---
def generate_consolidated_workbook(matched_df, fuzzy_df, un_a_df, un_b_df, signoff_data, tolerance=1.0, company_info=None, query_obj=None) -> bytes:
    output = io.BytesIO()
    timestamp_utc = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

    comp_name = company_info.get("name", "AutoReconFlow Enterprise") if company_info else "AutoReconFlow Enterprise"
    comp_addr = company_info.get("address", "") if company_info else ""
    comp_uid = company_info.get("uid", "") if company_info else ""

    navy_header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    section_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=11)
    italic_font = Font(name="Calibri", size=10, italic=True, color="595959")
    title_font = Font(name="Calibri", size=14, bold=True, color="1B365D")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    m_copy = matched_df.copy()
    f_copy = fuzzy_df.copy()
    a_copy = un_a_df.copy()
    b_copy = un_b_df.copy()

    if query_obj and query_obj.get("status") == "APPROVED" and query_obj.get("item_id"):
        flag_id = str(query_obj["item_id"])
        trail_entry = f"[REVIEWER {query_obj['reviewer_query_time']}]: {query_obj['reviewer_query']} | [PREPARER {query_obj['preparer_reply_time']}]: {query_obj['preparer_reply']}"
        
        for df_target in [f_copy, a_copy, b_copy]:
            match_col = None
            for c in df_target.columns:
                if any(k in str(c).lower() for k in ["id", "invoice", "ref", "trans"]):
                    match_col = c
                    break
            if match_col:
                df_target["Audit_Clarification_Trail"] = df_target[match_col].apply(
                    lambda x: trail_entry if str(x) == flag_id else ""
                )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        workbook = writer.book

        ws_summary = workbook.create_sheet(title="Audit & Control Summary", index=0)
        ws_summary.views.sheetView[0].showGridLines = True

        ws_summary["A1"] = f"{comp_name} — Financial Reconciliation, Tax & Control Workpaper"
        ws_summary["A1"].font = title_font
        ws_summary.merge_cells("A1:C1")

        ws_summary["A2"] = f"Legal Identification: {comp_uid} | Registered Office: {comp_addr}"
        ws_summary["A2"].font = italic_font
        ws_summary.merge_cells("A2:C2")

        ws_summary["A3"] = "SOX 404 / ISA-500 Standard Financial Statement & Tax Assertion Audit Trail"
        ws_summary["A3"].font = Font(name="Calibri", size=9, italic=True, color="64748B")
        ws_summary.merge_cells("A3:C3")

        metadata = [
            ("Execution Timestamp (UTC):", timestamp_utc),
            ("Tolerance Applied:", f"± {tolerance:.2f} Currency Units"),
            ("Matching Methodology:", "3-Tier (Deterministic Exact -> Levenshtein/Jaro-Winkler Fuzzy)"),
            ("Control Assertion Target:", "Completeness, Accuracy, and Valuation (Financial & Tax Controls)")
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
            ("Tier-1 Deterministic Matches (Subledger to External)", len(m_copy), "Cleared"),
            ("Tier-2 Discrepancy & Threshold Review (Materiality Tracking)", len(f_copy), "Requires Supervisory Review" if len(f_copy) > 0 else "Cleared"),
            ("Source A Exceptions (Ledger Missing in External)", len(a_copy), "Action Required" if len(a_copy) > 0 else "Cleared"),
            ("Source B Exceptions (External Missing in Ledger)", len(b_copy), "Action Required" if len(b_copy) > 0 else "Cleared"),
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
            "Deterministic Matches": m_copy,
            "Fuzzy Candidates": f_copy,
            "Exceptions Source A": a_copy,
            "Exceptions Source B": b_copy
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
                ws.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 55)

    return output.getvalue()

# --- Ingestion Engine ---
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
            
        delim = ","
        if "\t" in lines[0]:
            delim = "\t"
        elif ";" in lines[0]:
            delim = ";"
        elif "|" in lines[0]:
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
            st.error(f"Failed to read PDF: {str(e)}.")
            return pd.DataFrame()

        if not rows:
            st.warning("⚠️ No tabular data identified in this PDF.")
            return pd.DataFrame()

        headers = [re.sub(r'[^a-zA-Z0-9_]', '', h.replace(" ", "_").lower()) or f"col_{i}" for i, h in enumerate(rows[0])]
        df = pd.DataFrame(rows[1:], columns=headers)
        return df.reset_index(drop=True)
        
    return pd.DataFrame()

# --- Sample Test Data ---
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
        # Reset entire workflow state on fresh run
        st.session_state["preparer_completed"] = False
        st.session_state["reviewer_approved"] = False
        st.session_state["audit_query_thread"] = {
            "status": "NONE",
            "item_id": "",
            "reviewer_query": "",
            "reviewer_query_time": "",
            "preparer_reply": "",
            "preparer_reply_time": ""
        }

    if "recon_results" in st.session_state:
        res = st.session_state["recon_results"]
        matched_pd = res["matched_pd"]
        fuzzy_pd = res["fuzzy_pd"]
        un_a_pd = res["un_a_pd"]
        un_b_pd = res["un_b_pd"]

        st.markdown("#### 📊 Internal Control Summary & Transaction Metrics")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Tier-1 Exact Matches", f"{len(matched_pd):,} records")
        k2.metric("Tier-2 Variance Review", f"{len(fuzzy_pd):,} records")
        k3.metric("Source A Ledger Exceptions", f"{len(un_a_pd):,} records")
        k4.metric("Source B External Exceptions", f"{len(un_b_pd):,} records")

        # Session tracking state
        is_prep_done = st.session_state.get("preparer_completed", False)
        is_rev_approved = st.session_state.get("reviewer_approved", False)
        if "audit_query_thread" not in st.session_state:
            st.session_state["audit_query_thread"] = {
                "status": "NONE",
                "item_id": "",
                "reviewer_query": "",
                "reviewer_query_time": "",
                "preparer_reply": "",
                "preparer_reply_time": ""
            }
        query_state = st.session_state["audit_query_thread"]

        # -------------------------------------------------------------
        # Sequential Governance & Sign-Off Workflow
        # -------------------------------------------------------------
        with st.expander("✍️ Four-Eyes Internal Governance Sign-Off & Approval Workflow", expanded=True):
            so_col1, so_col2 = st.columns(2)
            now_str = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

            # --- STAGE 1: PREPARER PANEL ---
            with so_col1:
                st.markdown("##### 👤 Stage 1: Preparer (Analyst / Senior Accountant)")
                prep_name = st.text_input("Preparer Full Name", value="Senior Recon Analyst", key="so_prep_name", disabled=is_prep_done)
                prep_id = st.text_input("Employee / System ID", value="FIN-ACC-8921", key="so_prep_id", disabled=is_prep_done)
                prep_email = st.text_input("Preparer Email Address", value="analyst@company.com", key="so_prep_email", disabled=is_prep_done)
                rev_target_email = st.text_input("Reviewer Notification Email", value="controller@company.com", key="so_rev_target_email", disabled=is_prep_done)
                prep_date = st.text_input("Preparation Date & Time (UTC)", value=st.session_state.get("saved_prep_date", now_str), key="so_prep_date", disabled=is_prep_done)
                prep_signed = st.checkbox("I verify that ledger entries, tolerances, and files are audited.", value=True, key="so_prep_signed", disabled=is_prep_done)

                if not is_prep_done:
                    if st.button("✅ Mark Preparer Task as Complete & Notify Reviewer", type="primary", use_container_width=True):
                        if not prep_signed:
                            st.error("Preparer verification box must be checked before submission.")
                        else:
                            st.session_state["preparer_completed"] = True
                            st.session_state["saved_prep_name"] = prep_name
                            st.session_state["saved_prep_id"] = prep_id
                            st.session_state["saved_prep_email"] = prep_email
                            st.session_state["saved_rev_target_email"] = rev_target_email
                            st.session_state["saved_prep_date"] = prep_date
                            st.session_state["saved_prep_signed"] = prep_signed

                            # Dispatch automated notification to Reviewer
                            plain_body = f"Hello,\n\nReconciliation workpaper has been prepared and submitted by {prep_name} ({prep_id}) at {prep_date}.\n\nPlease review discrepancies and issue supervisory sign-off: https://www.autoreconflow.com\n\nAutoReconFlow Governance Engine"
                            html_body = build_html_email(
                                "Action Required: Reconciliation Awaiting Approval",
                                f"A reconciliation workpaper has been prepared and submitted for your supervisory sign-off under SOX 404 controls.",
                                {
                                    "Preparer": f"{prep_name} ({prep_id})",
                                    "Submitted At": prep_date,
                                    "Matched Items": f"{len(matched_pd):,} records",
                                    "Exceptions to Review": f"{len(un_a_pd) + len(un_b_pd) + len(fuzzy_pd):,} items"
                                }
                            )
                            ok, info = send_workflow_email(smtp_cfg, rev_target_email, "ACTION REQUIRED: Reconciliation Sign-Off Pending - AutoReconFlow", plain_body, html_body)
                            if not ok:
                                st.warning(f"Note: Workflow updated locally, but email could not be sent: {info}")
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
                    st.markdown("##### 🛡️ Stage 2: Reviewer (Active for Evaluation)")
                    st.success("🟢 Preparer has completed their submission. Ready for Supervisory Evaluation.")
                    rev_name = st.text_input("Reviewer Full Name", value="Finance Controller", key="so_rev_name", disabled=is_rev_approved)
                    rev_title = st.text_input("Role / Designation", value="VP - Financial Reporting & Controls", key="so_rev_title", disabled=is_rev_approved)
                    rev_date = st.text_input("Review Date & Time (UTC)", value=st.session_state.get("saved_rev_date", now_str), key="so_rev_date", disabled=is_rev_approved)
                    rev_verdict = st.selectbox("Supervisory Audit Verdict", ["Approved / Cleared", "Approved with Exceptions Noted", "Under Remediation", "Rejected"], key="so_rev_verdict", disabled=is_rev_approved)
                    wp_ref = st.text_input("Audit Workpaper Reference Code", value="WP-SOX-REC-2026-Q4", key="so_wp_ref", disabled=is_rev_approved)
                    audit_notes = st.text_area("General Supervisory Notes", value="All deterministic matches verified. Exceptions evaluated under tolerable limit.", key="so_notes", disabled=is_rev_approved)

                    if not is_rev_approved and query_state["status"] not in ["PENDING_CLARIFICATION"]:
                        if st.button("🛡️ Approve & Finalize Supervisory Sign-Off (No Queries)", type="primary", use_container_width=True):
                            st.session_state["reviewer_approved"] = True
                            st.session_state["saved_rev_name"] = rev_name
                            st.session_state["saved_rev_title"] = rev_title
                            st.session_state["saved_rev_date"] = rev_date
                            st.session_state["saved_rev_verdict"] = rev_verdict
                            st.session_state["saved_wp_ref"] = wp_ref
                            st.session_state["saved_audit_notes"] = audit_notes

                            # Notify Preparer of final approval
                            prep_mail = st.session_state.get("saved_prep_email", "analyst@company.com")
                            html_approved = build_html_email(
                                "Reconciliation Workpaper Approved & Sealed",
                                "The supervisory review cycle has been formally cleared. Final audit artifacts are unlocked.",
                                {
                                    "Reviewer": f"{rev_name} ({rev_title})",
                                    "Verdict": rev_verdict,
                                    "Approved Timestamp": rev_date,
                                    "Workpaper Ref": wp_ref
                                }
                            )
                            send_workflow_email(smtp_cfg, prep_mail, f"COMPLETED: Reconciliation Workpaper [{wp_ref}] Approved", "Workpaper formally cleared.", html_approved)
                            st.rerun()
                    elif is_rev_approved:
                        st.success(f"✅ Supervisory Sign-Off Formally Approved & Sealed ({st.session_state.get('saved_rev_verdict', rev_verdict)})")

        # -------------------------------------------------------------
        # Line-Item Query, Clarification & Reverse Action Engine
        # -------------------------------------------------------------
        if is_prep_done and not is_rev_approved:
            st.markdown("#### 💬 Supervisory Exception Query & Clarification Thread")

            disputed_items = []
            for col_id_candidate in ["Invoice_ID", "Transaction_Ref", "col_0"]:
                if col_id_candidate in un_b_pd.columns:
                    disputed_items.extend(un_b_pd[col_id_candidate].astype(str).tolist())
                if col_id_candidate in fuzzy_pd.columns:
                    disputed_items.extend(fuzzy_pd[col_id_candidate].astype(str).tolist())
                if col_id_candidate in un_a_pd.columns:
                    disputed_items.extend(un_a_pd[col_id_candidate].astype(str).tolist())
            disputed_items = sorted(list(set([x for x in disputed_items if x])))

            # Scenario A: Reviewer raises a query
            if query_state["status"] in ["NONE", "RESUBMITTED"]:
                with st.container():
                    st.info("💡 Reviewer Option: Raise a specific transaction query to send back to the Preparer.")
                    sel_item = st.selectbox("Select Line Item Requiring Clarification:", options=disputed_items, key="sel_query_item")
                    q_text = st.text_area("Reviewer Query / Required Explanation:", placeholder="e.g., Variance of 50.00 observed. Please provide bank swift advice or explain timing delay.", key="rev_query_input")
                    
                    if st.button("❓ Send Query Back to Preparer for Clarification", type="secondary"):
                        if not q_text.strip():
                            st.error("Please enter a query description before sending.")
                        else:
                            query_state["status"] = "PENDING_CLARIFICATION"
                            query_state["item_id"] = sel_item
                            query_state["reviewer_query"] = q_text
                            query_state["reviewer_query_time"] = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

                            target_prep_mail = st.session_state.get("saved_prep_email", "analyst@company.com")
                            html_query = build_html_email(
                                f"Clarification Required: Line Item [{sel_item}]",
                                "The Reviewer has raised an audit observation on this line item. Please provide your explanation.",
                                {
                                    "Item Queried": sel_item,
                                    "Reviewer Note": q_text,
                                    "Queried At": query_state["reviewer_query_time"]
                                }
                            )
                            send_workflow_email(
                                smtp_cfg,
                                target_prep_mail,
                                f"ACTION REQUIRED: Reviewer Clarification Query [{sel_item}] - AutoReconFlow",
                                f"Reviewer query on {sel_item}: {q_text}",
                                html_query
                            )
                            st.rerun()

            # Scenario B: Preparer responds to the query
            elif query_state["status"] == "PENDING_CLARIFICATION":
                st.warning(f"⚠️ Action Required by Preparer: Reviewer queried item [{query_state['item_id']}]")
                st.markdown(f"**Reviewer Observation ({query_state['reviewer_query_time']}):**")
                st.markdown(f"> *\"{query_state['reviewer_query']}\"*")

                prep_reply = st.text_area("Preparer Explanation / Response:", placeholder="e.g., Confirmed with Treasury team. Difference represents swift processing fee ref #8810; approved by Controller.", key="prep_reply_input")

                if st.button("📤 Submit Explanation Back to Reviewer", type="primary"):
                    if not prep_reply.strip():
                        st.error("Please enter an explanation before resubmitting.")
                    else:
                        query_state["status"] = "RESUBMITTED"
                        query_state["preparer_reply"] = prep_reply
                        query_state["preparer_reply_time"] = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

                        target_rev_mail = st.session_state.get("saved_rev_target_email", "controller@company.com")
                        html_reply = build_html_email(
                            f"Clarification Submitted for [{query_state['item_id']}]",
                            "The Preparer has responded to your observation. The item is ready for supervisory evaluation.",
                            {
                                "Item ID": query_state["item_id"],
                                "Original Query": query_state["reviewer_query"],
                                "Preparer Explanation": prep_reply,
                                "Resubmitted At": query_state["preparer_reply_time"]
                            }
                        )
                        send_workflow_email(
                            smtp_cfg,
                            target_rev_mail,
                            f"EXPLANATION SUBMITTED: Item [{query_state['item_id']}] Ready for Approval - AutoReconFlow",
                            f"Preparer responded to query on {query_state['item_id']}: {prep_reply}",
                            html_reply
                        )
                        st.rerun()

            # Scenario C: Reviewer evaluates response and approves
            if query_state["status"] == "RESUBMITTED":
                st.success(f"📥 Preparer Explanation Received for [{query_state['item_id']}].")
                st.markdown(f"**Reviewer Query:** *\"{query_state['reviewer_query']}\"*")
                st.markdown(f"**Preparer Explanation ({query_state['preparer_reply_time']}):**")
                st.markdown(f"> **\"{query_state['preparer_reply']}\"**")

                if st.button("🛡️ Accept Explanation & Formally Approve Workpaper", type="primary"):
                    st.session_state["reviewer_approved"] = True
                    query_state["status"] = "APPROVED"
                    st.session_state["saved_rev_name"] = rev_name
                    st.session_state["saved_rev_title"] = rev_title
                    st.session_state["saved_rev_date"] = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
                    st.session_state["saved_rev_verdict"] = "Approved with Clarifications Noted"
                    st.session_state["saved_wp_ref"] = wp_ref
                    st.session_state["saved_audit_notes"] = audit_notes

                    # Notify Preparer
                    prep_mail = st.session_state.get("saved_prep_email", "analyst@company.com")
                    html_resolved = build_html_email(
                        "Clarification Cleared & Workpaper Approved",
                        "Your explanation was accepted by the Reviewer. The workpaper has been permanently approved.",
                        {
                            "Item Resolved": query_state["item_id"],
                            "Verdict": "Approved with Clarifications Noted",
                            "Reviewer": rev_name,
                            "Timestamp": st.session_state["saved_rev_date"]
                        }
                    )
                    send_workflow_email(smtp_cfg, prep_mail, f"RESOLVED: Workpaper [{wp_ref}] Formally Approved", "Clarification accepted.", html_resolved)
                    st.rerun()

        # Build payload
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

        query_trail_summary = ""
        if query_state.get("status") == "APPROVED" and query_state.get("item_id"):
            query_trail_summary = f"Item {query_state['item_id']}: Queried ({query_state['reviewer_query']}) -> Clarified ({query_state['preparer_reply']}) -> Cleared."

        # -------------------------------------------------------------
        # Official Export Hub: Workpapers Unlock Upon Reviewer Approval
        # -------------------------------------------------------------
        st.markdown("#### 📦 Audit Deliverables & Certification Hub")
        exp_col1, exp_col2 = st.columns(2)

        with exp_col1:
            if is_rev_approved:
                excel_bytes = generate_consolidated_workbook(
                    matched_pd, fuzzy_pd, un_a_pd, un_b_pd, signoff_payload, 
                    tolerance=res["tolerance"], company_info=active_company_info, query_obj=query_state
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
                st.info("🔒 Excel Workpaper unlocks automatically upon complete Reviewer approval.")

        with exp_col2:
            if is_rev_approved:
                logo_file_bytes = uploaded_logo.getvalue() if uploaded_logo is not None else None
                pdf_bytes = generate_pdf_signoff_certificate(
                    len(matched_pd), len(fuzzy_pd), len(un_a_pd), len(un_b_pd), signoff_payload, res["tolerance"], 
                    logo_bytes=logo_file_bytes, company_info=active_company_info, query_trail=query_trail_summary
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
