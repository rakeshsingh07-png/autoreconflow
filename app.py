import io
import pandas as pd
from recon_engine import CommercialReconEngine
import streamlit as st

st.set_page_config(
    page_title="AutoReconFlow | Enterprise Financial Reconciliation",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling for SaaS Appeal
st.markdown(
    """
    <style>
    .metric-card {
        background-color: #f8f9fa;
        border-left: 5px solid #0052cc;
        padding: 18px;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    .stButton>button {
        background-color: #0052cc;
        color: white;
        font-weight: 600;
        border-radius: 6px;
        height: 48px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Top Brand Bar
col_logo, col_badge = st.columns([4, 1])
with col_logo:
    st.title("⚖️ AutoReconFlow")
    st.caption(
        "Deterministic Multi-Source Financial & Bank Reconciliation Platform"
    )
with col_badge:
    st.markdown(
        "<div style='text-align: right; margin-top: 15px;'><span style='background-color: #e3fcef; color: #006644; padding: 6px 12px; border-radius: 12px; font-weight: bold;'>Enterprise Edition</span></div>",
        unsafe_allow_html=True,
    )

st.divider()

# Sidebar: Configuration & Pricing CTA
with st.sidebar:
    st.header("⚙️ Matching Controls")
    tolerance = st.slider(
        "Variance Tolerance ($)",
        0.00,
        1.00,
        0.00,
        0.01,
        help="Allowed penny differences for rounding discrepancies.",
    )
    fuzzy_ratio = st.slider(
        "Fuzzy Match Threshold (%)",
        60,
        100,
        80,
        help="Catches noisy payment references (e.g. ACH prefixes).",
    )
    enable_split = st.checkbox(
        "Solve 1-to-Many Batch Payments",
        value=True,
        help="Knapsack solver for lump-sum vendor settlements.",
    )

    st.markdown("---")
    st.markdown("### 💼 Commercial License")
    st.info(
        "**Unlimited Transactions Tier**\n- Multi-ERP connectors (SAP, Oracle, NetSuite)\n- SOC-2 compliant audit trails\n- Dedicated AWS VPC deployment"
    )
    st.link_button(
        "Schedule Enterprise Demo",
        "mailto:contact@rakeshsingh.tech?subject=AutoReconFlow%20Enterprise%20Inquiry",
    )

# Demo Mode Setup Function
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

# Workflow Selection
tab_upload, tab_demo = st.tabs(
    ["📁 Upload Your Datasets", "🚀 Instant 1-Click Interactive Demo"]
)

with tab_demo:
    st.write(
        "Experience the 3-Tier engine instantly on sample messy corporate data (Exact match, noisy reference match, and split settlement):"
    )
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.caption("Source A: Bank Statement (Messy ACH / Wire Narration)")
        st.dataframe(demo_df_bank, use_container_width=True, height=180)
    with col_d2:
        st.caption("Source B: General Ledger (ERP Bookings)")
        st.dataframe(demo_df_gl, use_container_width=True, height=180)

    if st.button("⚡ Run Instant Interactive Recon"):
        engine = CommercialReconEngine(demo_df_bank, demo_df_gl)
        with st.spinner("Processing deterministic multi-pass cascade..."):
            matched_df, exc_a, exc_b = engine.run_reconciliation(
                col_ref_a="Narration",
                col_amt_a="Amount",
                col_ref_b="Document_No",
                col_amt_b="Ledger_Amt",
                tolerance=tolerance,
                fuzzy_threshold=fuzzy_ratio,
                enable_split_matching=enable_split,
            )

        # ROI & Performance Metrics
        total_rec = len(demo_df_bank) + len(demo_df_gl)
        match_rate = round(
            (len(matched_df) * 2 / total_rec) * 100, 1
        ) if total_rec > 0 else 0
        volume = matched_df["Amount_A"].sum() if not matched_df.empty else 0
        hours_saved = round(
            (len(matched_df) * 4) / 60, 1
        )  # Avg 4 mins per complex exception manually

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Overall Match Rate", f"{match_rate}%", "+32% vs Manual")
        m2.metric("Matched Pairs", len(matched_df))
        m3.metric("Reconciled Volume", f"${volume:,.2f}")
        m4.metric("Est. Time Saved", f"{hours_saved} Hours", "High Efficiency")

        # Results Tabs
        r1, r2, r3 = st.tabs([
            "✅ Matched Records (Traceable)",
            "⚠️ Source A Exceptions",
            "⚠️ Source B Exceptions",
        ])
        with r1:
            st.dataframe(matched_df, use_container_width=True)
        with r2:
            st.dataframe(exc_a, use_container_width=True)
        with r3:
            st.dataframe(exc_b, use_container_width=True)

        # Export Package
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
            label="📥 Download CFO-Ready Audit Workpaper (.xlsx)",
            data=excel_bytes,
            file_name=f"AutoReconFlow_Audit_Package_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab_upload:
    u1, u2 = st.columns(2)
    with u1:
        f_a = st.file_uploader(
            "Source A (Bank/Subledger CSV/Excel)",
            type=["csv", "xlsx"],
            key="up_a",
        )
    with u2:
        f_b = st.file_uploader(
            "Source B (General Ledger CSV/Excel)",
            type=["csv", "xlsx"],
            key="up_b",
        )

    if f_a and f_b:
        df_a_custom = (
            pd.read_csv(f_a) if f_a.name.endswith(".csv") else pd.read_excel(f_a)
        )
        df_b_custom = (
            pd.read_csv(f_b) if f_b.name.endswith(".csv") else pd.read_excel(f_b)
        )

        st.markdown("#### Dynamic Column Mapping")
        c1, c2 = st.columns(2)
        with c1:
            col_ref_a = st.selectbox(
                "Source A Reference / Narration Column", df_a_custom.columns
            )
            col_amt_a = st.selectbox(
                "Source A Amount Column", df_a_custom.columns
            )
        with c2:
            col_ref_b = st.selectbox(
                "Source B Document / Key Column", df_b_custom.columns
            )
            col_amt_b = st.selectbox(
                "Source B Ledger Amount Column", df_b_custom.columns
            )

        if st.button("🚀 Reconcile Uploaded Files"):
            engine = CommercialReconEngine(df_a_custom, df_b_custom)
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
                f"Completed! {len(matched_df)} transaction pairs resolved across 3 passes."
            )
            st.dataframe(matched_df, use_container_width=True)

            meta = {
                "count_a": len(df_a_custom),
                "count_b": len(df_b_custom),
                "match_rate": round(
                    (len(matched_df) * 2 / (len(df_a_custom) + len(df_b_custom)))
                    * 100,
                    1,
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
                file_name="AutoReconFlow_Reconciliation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )