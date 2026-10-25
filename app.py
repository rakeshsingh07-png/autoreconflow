import streamlit as st
import pandas as pd
from recon_engine import CommercialReconEngine

# --- Page Configuration ---
st.set_page_config(
    page_title="AutoReconFlow - Financial Reconciliation Platform",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- License & Tier State Management ---
if "user_tier" not in st.session_state:
    st.session_state["user_tier"] = "Tier 1: Free Starter"

VALID_LICENSE_KEYS = {
    "PRO-RECON-2026": "Tier 2: Pro Commercial ($49/mo)",
    "RECON-CFO-ACCESS": "Tier 2: Pro Commercial ($49/mo)",
    "ENTERPRISE-LIFETIME": "Tier 3: Enterprise / Lifetime ($299)"
}

# --- Sidebar: Tier Licensing & Controls ---
st.sidebar.title("⚖️ AutoReconFlow")

st.sidebar.markdown("### 🔑 License Activation")
entered_key = st.sidebar.text_input("Enter License Key", type="password", placeholder="Paste license key here")

if entered_key:
    if entered_key.strip() in VALID_LICENSE_KEYS:
        st.session_state["user_tier"] = VALID_LICENSE_KEYS[entered_key.strip()]
        st.sidebar.success(f"Active: {st.session_state['user_tier']}")
    else:
        st.session_state["user_tier"] = "Tier 1: Free Starter"
        st.sidebar.error("Invalid license key. Running on Free Tier.")
else:
    st.session_state["user_tier"] = "Tier 1: Free Starter"
    st.sidebar.info("Plan: Tier 1 Free Starter (Limited to 50 rows)")

is_pro = st.session_state["user_tier"] != "Tier 1: Free Starter"

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Matching Engine Controls")

allowed_variance = st.sidebar.slider("Allowed Variance ($)", min_value=0.0, max_value=10.0, value=0.0, step=0.01)

if is_pro:
    fuzzy_sensitivity = st.sidebar.slider("Fuzzy Match Sensitivity (%)", min_value=50, max_value=100, value=80, step=1)
    solve_batch = st.sidebar.checkbox("Solve 1-to-Many Batch Payments (Subset-Sum)", value=True)
else:
    st.sidebar.slider("Fuzzy Match Sensitivity (%)", min_value=50, max_value=100, value=80, disabled=True)
    st.sidebar.checkbox("Solve 1-to-Many Batch Payments", value=False, disabled=True)
    st.sidebar.caption("🔒 *Fuzzy tuning and 1-to-many subset solver require Pro.*")

st.sidebar.markdown("---")
st.sidebar.markdown("### 💳 Upgrade Commercial Plan")
col_btn1, col_btn2 = st.sidebar.columns(2)
with col_btn1:
    st.link_button("Pro ($49/mo)", "https://buy.stripe.com/your_pro_link_here")
with col_btn2:
    st.link_button("Lifetime ($299)", "https://buy.stripe.com/your_lifetime_link_here")

st.sidebar.markdown("---")
st.sidebar.caption("Enterprise inquiries:")
st.sidebar.markdown("[contact@rakeshsingh.tech](mailto:contact@rakeshsingh.tech)")

# --- Main App Header ---
col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.title("⚖️ AutoReconFlow")
    st.caption("Deterministic Multi-Pass Financial & Bank Reconciliation Platform")
with col_head2:
    st.markdown("<br>", unsafe_allow_html=True)
    st.success("🟢 Production Live")

# --- Tabs ---
tab_demo, tab_upload, tab_pricing = st.tabs([
    "🚀 1-Click Interactive Demo",
    "📁 Reconcile Your Files",
    "🏷️ Plans & Licensing"
])

# --- Tab 1: Interactive Demo ---
with tab_demo:
    st.markdown("#### Test the 3-Tier engine on messy corporate transactions:")
    
    demo_bank = pd.DataFrame([
        {"Bank_Date": "2026-10-01", "Narration": "INV-1001-VENDOR-ACME", "Amount": 1500.00},
        {"Bank_Date": "2026-10-02", "Narration": "ACH PMT MICROSOFT CORP REF 8821", "Amount": 3200.00},
        {"Bank_Date": "2026-10-03", "Narration": "WIRE INFLOW 9901 GLOBAL LTD", "Amount": 4750.50},
        {"Bank_Date": "2026-10-04", "Narration": "CHECK 4421 DEPOSIT", "Amount": 890.00},
        {"Bank_Date": "2026-10-05", "Narration": "BULK VENDOR PAYOUT COMBINED", "Amount": 4000.00}
    ])
    
    demo_ledger = pd.DataFrame([
        {"Posting_Date": "2026-10-01", "Document_No": "INV-1001-VENDOR-ACME", "Ledger_Amt": 1500.00},
        {"Posting_Date": "2026-10-02", "Document_No": "MSFT CORP REF 8821", "Ledger_Amt": 3200.00},
        {"Posting_Date": "2026-10-03", "Document_No": "WIRE INFLOW 9901 GLOBAL LTD", "Ledger_Amt": 4750.50},
        {"Posting_Date": "2026-10-05", "Document_No": "BATCH-PART-A", "Ledger_Amt": 2500.00},
        {"Posting_Date": "2026-10-05", "Document_No": "BATCH-PART-B", "Ledger_Amt": 1500.00}
    ])
    
    c_demo1, c_demo2 = st.columns(2)
    with c_demo1:
        st.caption("Source A: Bank Statement (Messy narration)")
        st.dataframe(demo_bank)
    with c_demo2:
        st.caption("Source B: General Ledger (ERP Bookings)")
        st.dataframe(demo_ledger)
        
    if st.button("⚡ Run Interactive Demo Reconciliation", type="primary"):
        with st.spinner("Processing through 3-tier cascade..."):
            engine = CommercialReconEngine(demo_bank, demo_ledger)
            matched_df, unmatched_bank, unmatched_ledger = engine.run_reconciliation(
                col_ref_a="Narration",
                col_amt_a="Amount",
                col_ref_b="Document_No",
                col_amt_b="Ledger_Amt",
                tolerance=allowed_variance,
                fuzzy_threshold=80.0 if is_pro else 100.0,
                enable_split_matching=True if is_pro else False
            )
            
            st.success("Reconciliation Completed Successfully!")
            st.markdown("##### ✅ Matched Transactions")
            st.dataframe(matched_df)
            
            c_un1, c_un2 = st.columns(2)
            with c_un1:
                st.markdown("##### ⚠️ Unmatched Bank Items")
                st.dataframe(unmatched_bank)
            with c_un2:
                st.markdown("##### ⚠️ Unmatched Ledger Items")
                st.dataframe(unmatched_ledger)

# --- Tab 2: Reconcile Your Files ---
with tab_upload:
    st.markdown("#### Upload Your Reconciliation Files (CSV or Excel)")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        file_a = st.file_uploader("Upload Bank Statement / Source A", type=["csv", "xlsx"])
    with col_up2:
        file_b = st.file_uploader("Upload General Ledger / Source B", type=["csv", "xlsx"])
        
    if file_a and file_b:
        try:
            df_a = pd.read_csv(file_a) if file_a.name.endswith(".csv") else pd.read_excel(file_a)
            df_b = pd.read_csv(file_b) if file_b.name.endswith(".csv") else pd.read_excel(file_b)
            
            if not is_pro and (len(df_a) > 50 or len(df_b) > 50):
                st.warning("⚠️ Free Tier limit: Only processing first 50 rows. Upgrade to Pro for unlimited volume.")
                df_a = df_a.head(50)
                df_b = df_b.head(50)
                
            cols_a = df_a.columns.tolist()
            cols_b = df_b.columns.tolist()
            
            c_sel1, c_sel2, c_sel3, c_sel4 = st.columns(4)
            with c_sel1:
                amt_col_a = st.selectbox("Source A Amount Column", cols_a)
            with c_sel2:
                desc_col_a = st.selectbox("Source A Reference/Text Column", cols_a)
            with c_sel3:
                amt_col_b = st.selectbox("Source B Amount Column", cols_b)
            with c_sel4:
                desc_col_b = st.selectbox("Source B Reference/Text Column", cols_b)
                
            if st.button("🚀 Process File Reconciliation", type="primary"):
                with st.spinner("Executing reconciliation cascade..."):
                    engine = CommercialReconEngine(df_a, df_b)
                    m_df, um_a, um_b = engine.run_reconciliation(
                        col_ref_a=desc_col_a,
                        col_amt_a=amt_col_a,
                        col_ref_b=desc_col_b,
                        col_amt_b=amt_col_b,
                        tolerance=allowed_variance,
                        fuzzy_threshold=float(fuzzy_sensitivity) if is_pro else 100.0,
                        enable_split_matching=solve_batch if is_pro else False
                    )
                    
                    st.success("Reconciliation Run Finished!")
                    st.dataframe(m_df)
                    
                    total_matched_volume = m_df["Amount_A"].sum() if not m_df.empty and "Amount_A" in m_df.columns else 0.0
                    match_rate = round((len(m_df) / max(len(df_a), 1)) * 100, 1)
                    
                    metadata = {
                        "count_a": len(df_a),
                        "count_b": len(df_b),
                        "match_rate": match_rate,
                        "reconciled_volume": total_matched_volume,
                        "hours_saved": round(len(m_df) * 0.08, 1)
                    }
                    
                    excel_bytes = CommercialReconEngine.generate_executive_workpaper(m_df, um_a, um_b, metadata)
                    
                    st.download_button(
                        label="📥 Download CFO-Ready Reconciliation Workbook (.xlsx)",
                        data=excel_bytes,
                        file_name="reconciliation_workpapers.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Error reading uploaded files: {str(e)}")

# --- Tab 3: Plans & Licensing ---
with tab_pricing:
    st.markdown("### Choose the Right Plan for Your Organization")
    st.markdown("Upgrade your license to unlock full 3-tier algorithmic power and unlimited file volume.")
    
    card1, card2, card3 = st.columns(3)
    
    with card1:
        st.info("### Tier 1: Free Starter")
        st.markdown("""
        **$0 / month**
        * Ideal for initial testing & basic batches
        
        - Up to **50 rows** per reconciliation run
        - Pass 1: Deterministic matching
        - Standard table results preview
        - Community web access
        """)
        st.caption("Active by default on open access.")
        
    with card2:
        st.success("### Tier 2: Pro Commercial")
        st.markdown("""
        **$49 / month**
        * For freelance accountants, bookkeepers & CPAs
        
        - **Unlimited** monthly rows & file sizes
        - Pass 2: **RapidFuzz** fuzzy matching controls
        - Pass 3: **1-to-Many** split payment solver
        - CFO-Ready formatted Excel export
        - Commercial license activation key
        """)
        st.link_button("Subscribe Pro ($49/mo)", "https://buy.stripe.com/your_pro_link_here", type="primary")
        
    with card3:
        st.warning("### Tier 3: Enterprise / Lifetime")
        st.markdown("""
        **$299 one-time**
        * For high-volume businesses & micro-PE buyers
        
        - All Pro Commercial features included
        - **Lifetime license** (no recurring monthly fees)
        - Standalone **Docker deployment** package
        - Custom columns & enterprise SLA support
        - Transferable IP rights for acquirers
        """)
        st.link_button("Buy Lifetime Pass ($299)", "https://buy.stripe.com/your_lifetime_link_here")
