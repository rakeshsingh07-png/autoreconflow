import pandas as pd
from recon_engine import CommercialReconEngine
import streamlit as st

st.set_page_config(
    page_title="AutoReconFlow | Financial Reconciliation Platform",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Live Razorpay Payment Link and Direct Support
RAZORPAY_PRO_LINK = "https://rzp.io/rzp/jFe4aJWI"
CONTACT_EMAIL = "contact@rakeshsingh.tech"

# Commercial Enterprise Styling
st.markdown(
    """
    <style>
    .metric-card {
        background-color: #f8f9fa;
        border-left: 5px solid #0052cc;
        padding: 16px;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .pricing-box {
        border: 2px solid #0052cc;
        border-radius: 10px;
        padding: 22px;
        background-color: #ffffff;
        text-align: center;
    }
    .stButton>button {
        background-color: #0052cc;
        color: white;
        font-weight: 600;
        border-radius: 6px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Header Navigation
col_logo, col_badge = st.columns([4, 1])
with col_logo:
    st.title("⚖️ AutoReconFlow")
    st.caption("Deterministic Multi-Pass Financial & Bank Reconciliation Platform")
with col_badge:
    st.markdown(
        "<div style='text-align: right; margin-top: 15px;'><span style='background-color: #e3fcef; color: #006644; padding: 6px 12px; border-radius: 12px; font-weight: bold;'>Production Live</span></div>",
        unsafe_allow_html=True,
    )

st.divider()

# Sidebar: Controls & Instant Checkout
with st.sidebar:
    st.header("⚙️ Matching Engine Controls")
    tolerance = st.slider(
        "Allowed Variance ($)",
        0.00,
        1.00,
        0.00,
        0.01,
        help="Tolerance threshold for rounding discrepancies.",
    )
    fuzzy_ratio = st.slider(
        "Fuzzy Match Sensitivity (%)",
        60,
        100,
        80,
        help="Catches noisy payment references (e.g., ACH/Wire prefixes).",
    )
    enable_split = st.checkbox(
        "Solve 1-to-Many Batch Payments",
        value=True,
        help="Matches multiple invoice entries to a single lump-sum settlement.",
    )

    st.markdown("---")
    st.header("💳 Commercial Upgrade")
    st.markdown(
        """
        **Pro Commercial License (₹3,999)**
        - Unlimited monthly row volume
        - Full 3-Tier cascade & split solver
        - CFO-Ready formatted Excel workpapers
        - Priority enterprise support
        """
    )
    st.link_button(
        "💳 Upgrade to Pro (₹3,999)",
        RAZORPAY_PRO_LINK,
        use_container_width=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.caption(f"Enterprise inquiries: [{CONTACT_EMAIL}](mailto:{CONTACT_EMAIL})")

# Sample Datasets for Instant Demo
demo_df_bank = pd.DataFrame({
    "Bank_Date": [
        "2026-10-01",
        "2026-10-02",
        "2026-10-03",
        "2026-10-04",
        "2026-10-05",
        "2026-10-06",
    ],
    "Narration": [
        "INV-1001-VENDOR-ACME",
        "ACH PMT MICROSOFT CORP REF 8821",
        "WIRE INFLOW 9901 GLOBAL LTD",
        "CHECK 4421 DEPOSIT",
        "LUMP SUM VENDOR BATCH #502",
        "UNKNOWN BANK FEE",
    ],
    "Amount": [1500.00, 3200.00, 4750.50, 890.00, 6000.00, 45.00],
})

demo_df_gl = pd.DataFrame({
    "Posting_Date": [
        "2026-10-01",
        "2026-10-02",
        "2026-10-03",
        "2026-10-05",
        "2026-10-05",
        "2026-10-07",
    ],
    "Document_No": [
        "INV-1001-VENDOR-ACME",
        "MSFT CORP REF 8821",
        "WIRE INFLOW 9901 GLOBAL LTD",
        "BATCH-PART-A",
        "BATCH-PART-B",
        "UNRECORDED ACCRUAL",
    ],
    "Ledger_Amt": [1500.00, 3200.00, 4750.50, 4000.00, 2000.00, 1250.00],
})

tab_demo, tab_upload, tab_pricing = st.tabs([
    "🚀 1-Click Interactive Demo",
    "📁 Reconcile Your Files",
    "🏷️ Plans & Licensing",
])

# Tab 1: Interactive Demo
with tab_demo:
    st.markdown("##### Test the 3-Tier engine on messy corporate transactions:")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.caption("Source A: Bank Statement (Messy narration)")
        st.dataframe(demo_df_bank, use_container_width=True, height=180)
    with col_d2:
        st.caption("Source B: General Ledger (ERP Bookings)")
        st.dataframe(demo_df_gl, use_container_width=True, height=180)

    if st.button("⚡ Run Instant Engine Demo", use_container_width=True):
        engine = CommercialReconEngine(demo_df_bank, demo_df_gl)
        with st.spinner("Processing multi-pass cascade..."):
            matched_df, exc_a, exc_b = engine.run_reconciliation(
                col_ref_a="Narration",
                col_amt_a="Amount",
                col_ref_b="Document_No",
                col_amt_b="Ledger_Amt",
                tolerance=tolerance,
                fuzzy_threshold=fuzzy_ratio,
                enable_split_matching=enable_split,
            )

        total_rec = len(demo_df_bank) + len(demo_df_gl)
        match_rate = (
            round((len(matched_df) * 2 / total_rec) * 100, 1)
            if total_rec > 0
            else 0
        )
        volume = matched_df["Amount_A"].sum() if not matched_df.empty else 0
        hours_saved = round((len(matched_df) * 4) / 60, 1)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Cascade Match Rate", f"{match_rate}%", "+35% vs Manual")
        m2.metric("Matched Pairs", len(matched_df))
        m3.metric("Reconciled Balance", f"${volume:,.2f}")
        m4.metric("Est. Manual Time Saved", f"{hours_saved} Hours")

        r1, r2, r3 = st.tabs([
            "✅ Matched Audit Ledger",
            "⚠️ Source A Exceptions",
            "⚠️ Source B Exceptions",
        ])
        with r1:
            st.dataframe(matched_df, use_container_width=True)
        with r2:
            st.dataframe(exc_a, use_container_width=True)
        with r3:
            st.dataframe(exc_b, use_container_width=True)

        meta = {
            "count_a": len(demo_df_bank),
            "count_b": len(demo_df_gl),
            "match_rate": match_rate,
            "reconciled_volume": volume,
            "hours_saved": hours_saved,
        }
        excel_bytes = CommercialReconEngine.generate_executive_workpaper(
            matched_df, exc_a, exc_b, meta
        )

        st.download_button(
            label="📥 Download CFO-Ready Workpaper (.xlsx)",
            data=excel_bytes,
            file_name="AutoReconFlow_Demo_Workpaper.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# Tab 2: User File Upload
with tab_upload:
    u1, u2 = st.columns(2)
    with u1:
        f_a = st.file_uploader(
            "Upload Bank / Subledger (CSV or Excel)",
            type=["csv", "xlsx"],
            key="file_a",
        )
    with u2:
        f_b = st.file_uploader(
            "Upload ERP / General Ledger (CSV or Excel)",
            type=["csv", "xlsx"],
            key="file_b",
        )

    if f_a and f_b:
        df_a = pd.read_csv(f_a) if f_a.name.endswith(".csv") else pd.read_excel(f_a)
        df_b = pd.read_csv(f_b) if f_b.name.endswith(".csv") else pd.read_excel(f_b)

        st.markdown("##### Dynamic Schema Mapping")
        c1, c2 = st.columns(2)
        with c1:
            col_ref_a = st.selectbox(
                "Source A Reference / Narration", df_a.columns
            )
            col_amt_a = st.selectbox("Source A Amount", df_a.columns)
        with c2:
            col_ref_b = st.selectbox("Source B Document / Key", df_b.columns)
            col_amt_b = st.selectbox("Source B Amount", df_b.columns)

        if st.button("🚀 Reconcile Uploaded Files", use_container_width=True):
            engine = CommercialReconEngine(df_a, df_b)
            matched_df, exc_a, exc_b = engine.run_reconciliation(
                col_ref_a,
                col_amt_a,
                col_ref_b,
                col_amt_b,
                tolerance=tolerance,
                fuzzy_threshold=fuzzy_ratio,
                enable_split_matching=enable_split,
            )
            st.success(
                f"Matched {len(matched_df)} transaction pairs across the cascade!"
            )
            st.dataframe(matched_df, use_container_width=True)

            meta = {
                "count_a": len(df_a),
                "count_b": len(df_b),
                "match_rate": round(
                    (len(matched_df) * 2 / (len(df_a) + len(df_b))) * 100, 1
                ),
                "reconciled_volume": (
                    matched_df["Amount_A"].sum() if not matched_df.empty else 0
                ),
                "hours_saved": round((len(matched_df) * 4) / 60, 1),
            }
            excel_bytes = CommercialReconEngine.generate_executive_workpaper(
                matched_df, exc_a, exc_b, meta
            )
            st.download_button(
                label="📥 Download Audit Workpaper (.xlsx)",
                data=excel_bytes,
                file_name="AutoReconFlow_Production_Workpaper.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# Tab 3: Pricing & Checkout
with tab_pricing:
    st.subheader("Select a Plan for Your Finance Operations")
    p1, p2, p3 = st.columns(3)

    with p1:
        st.markdown(
            """
            <div class='pricing-box'>
                <h4>Starter</h4>
                <h2>Free</h2>
                <p>For independent accountants</p>
                <hr>
                <p>✔ Up to 1,000 rows/month</p>
                <p>✔ Exact match pass</p>
                <p>✖ No split-sum solving</p>
                <p>✖ Standard export</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with p2:
        st.markdown(
            """
            <div class='pricing-box' style='border-color: #0052cc; box-shadow: 0 4px 12px rgba(0,82,204,0.15);'>
                <span style='background-color: #0052cc; color: white; padding: 3px 8px; border-radius: 4px; font-size: 12px;'>POPULAR</span>
                <h4>Professional</h4>
                <h2>₹3,999 <span style='font-size:16px;'>one-time / monthly</span></h2>
                <p>For high-volume finance teams</p>
                <hr>
                <p>✔ Unlimited transaction rows</p>
                <p>✔ 3-Tier Fuzzy & Split Solver</p>
                <p>✔ CFO-Ready Audit Workpapers</p>
                <p>✔ Email support</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.link_button(
            "💳 Instant Checkout (Razorpay)",
            RAZORPAY_PRO_LINK,
            use_container_width=True,
        )

    with p3:
        st.markdown(
            f"""
            <div class='pricing-box'>
                <h4>Enterprise</h4>
                <h2>Custom</h2>
                <p>For SSCs, Banks & Enterprises</p>
                <hr>
                <p>✔ Direct SAP / Oracle connectors</p>
                <p>✔ SOC-2 & On-premise VPC deployment</p>
                <p>✔ Multi-currency FX revaluation</p>
                <p>✔ Dedicated Support SLA</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.link_button(
            "Contact Sales",
            f"mailto:{CONTACT_EMAIL}?subject=Enterprise%20Recon%20Inquiry",
            use_container_width=True,
        )
