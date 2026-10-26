import streamlit as st
import polars as pl
from parser_engine import parse_to_polars
from recon_engine import execute_high_volume_reconciliation

# 1. Page Configuration
st.set_page_config(
    page_title="AutoReconFlow — Enterprise Autonomous Reconciliation",
    page_icon="⚡",
    layout="wide"
)

# 2. Commercial Monetization Sidebar
st.sidebar.title("⚡ AutoReconFlow")
st.sidebar.markdown("**Enterprise B2B Edition v2.0**")
st.sidebar.divider()

st.sidebar.markdown("### 💼 Commercial Licensing")
st.sidebar.markdown(
    """
    Acquire full source code, Docker assets, and distribution rights:
    - **Single Practice License**: \$79
    - **Extended / Agency License**: \$499
    
    [👉 Purchase Commercial License](https://autoreconflow.com)
    """
)
st.sidebar.caption("Includes Polars multi-core matching, PDF ingestion, and GST/VAT cross-border rules.")
st.sidebar.divider()

# 3. Main Header
st.title("⚡ AutoReconFlow — Enterprise Autonomous Reconciliation")
st.markdown("High-throughput, 3-tier financial matching and cross-border tax reconciliation platform.")

# 4. Ingestion Dropzones
col1, col2 = st.columns(2)

with col1:
    st.subheader("Source A (Ledger / Internal Records)")
    file_a = st.file_uploader(
        "Upload CSV, XLSX, or Bank Statement PDF",
        type=["csv", "xlsx", "xls", "pdf"],
        key="file_a"
    )

with col2:
    st.subheader("Source B (Bank / Tax Portal / External Records)")
    file_b = st.file_uploader(
        "Upload CSV, XLSX, or Bank Statement PDF",
        type=["csv", "xlsx", "xls", "pdf"],
        key="file_b"
    )

# 5. Helper function for Smart Column Detection
def get_default_index(columns, keywords, fallback_index=0):
    """Finds the most likely column matching a set of keywords."""
    for idx, col in enumerate(columns):
        clean_col = str(col).lower().replace(" ", "_")
        if any(kw in clean_col for kw in keywords):
            return idx
    return min(fallback_index, len(columns) - 1)

# 6. Processing Pipeline
if file_a and file_b:
    with st.spinner("Parsing documents via Polars multi-core parser..."):
        try:
            df_a = parse_to_polars(file_a)
            df_b = parse_to_polars(file_b)
        except Exception as e:
            st.error(f"Error reading uploaded files: {str(e)}")
            st.stop()

    if len(df_a.columns) == 0 or len(df_b.columns) == 0:
        st.error("Could not extract tabular columns from one of the files. Please check the file formatting.")
        st.stop()

    st.success(f"Successfully loaded Source A ({len(df_a):,} rows, {len(df_a.columns)} cols) and Source B ({len(df_b):,} rows, {len(df_b.columns)} cols)")

    # 7. Safe Dynamic Column Mappings
    cols_a = list(df_a.columns)
    cols_b = list(df_b.columns)

    with st.expander("🛠️ Configure Column Mappings", expanded=True):
        m1, m2 = st.columns(2)

        with m1:
            st.markdown("**Source A (Internal Books)**")
            idx_id_a = get_default_index(cols_a, ["invoice", "inv", "bill", "doc", "ref", "id"], 0)
            idx_party_a = get_default_index(cols_a, ["vendor", "party", "name", "customer", "desc"], min(1, len(cols_a) - 1))
            idx_amt_a = get_default_index(cols_a, ["amount", "total", "val", "debit", "credit"], min(2, len(cols_a) - 1))

            id_a = st.selectbox("Document / Txn ID Column", cols_a, index=idx_id_a, key="map_id_a")
            party_a = st.selectbox("Party / Narrative Column", cols_a, index=idx_party_a, key="map_party_a")
            amt_a = st.selectbox("Amount Column", cols_a, index=idx_amt_a, key="map_amt_a")

        with m2:
            st.markdown("**Source B (Bank / Portal)**")
            idx_id_b = get_default_index(cols_b, ["transaction", "ref", "invoice", "inv", "cheque", "id"], 0)
            idx_party_b = get_default_index(cols_b, ["description", "narrative", "party", "name", "particulars"], min(1, len(cols_b) - 1))
            idx_amt_b = get_default_index(cols_b, ["settlement", "amount", "total", "balance", "net"], min(2, len(cols_b) - 1))

            id_b = st.selectbox("Document / Txn ID Column", cols_b, index=idx_id_b, key="map_id_b")
            party_b = st.selectbox("Party / Narrative Column", cols_b, index=idx_party_b, key="map_party_b")
            amt_b = st.selectbox("Amount Column", cols_b, index=idx_amt_b, key="map_amt_b")

    # 8. Tolerance Controls
    tolerance = st.slider("Allowed Amount Tolerance (Currency Units)", 0.00, 5.00, 1.00, step=0.10)

    # 9. Execution Button
    if st.button("🚀 Run Automated Reconciliation", type="primary", width="stretch"):
        map_a = {"id": id_a, "party": party_a, "amt": amt_a}
        map_b = {"id": id_b, "party": party_b, "amt": amt_b}

        with st.spinner("Processing through 3-tier engine (Deterministic ➔ Heuristic Fuzzy ➔ Unmatched Exceptions)..."):
            try:
                matched, fuzzy, un_a, un_b = execute_high_volume_reconciliation(
                    df_a, df_b, map_a, map_b, tolerance=tolerance
                )
            except Exception as e:
                st.error(f"Reconciliation error: {str(e)}")
                st.stop()

        # 10. Results & KPI Summary
        st.subheader("📊 Reconciliation Summary")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("100% Deterministic Matched", f"{len(matched):,} records")
        k2.metric("Fuzzy Candidates", f"{len(fuzzy):,} records")
        k3.metric("Unmatched (Source A)", f"{len(un_a):,} records")
        k4.metric("Unmatched (Source B)", f"{len(un_b):,} records")

        # 11. Detailed Output Tabs
        tab1, tab2, tab3 = st.tabs(["✅ Matched Records", "🔍 Fuzzy Review", "⚠️ Exceptions"])

        with tab1:
            if len(matched) > 0:
                st.dataframe(matched.head(1000).to_pandas(), width="stretch")
                csv_matched = matched.write_csv()
                st.download_button(
                    label="📥 Download Matched CSV",
                    data=csv_matched,
                    file_name="reconciled_matched_records.csv",
                    mime="text/csv"
                )
            else:
                st.info("No exact matches found under the current tolerance settings.")

        with tab2:
            if len(fuzzy) > 0:
                st.dataframe(fuzzy.to_pandas(), width="stretch")
                csv_fuzzy = fuzzy.write_csv()
                st.download_button(
                    label="📥 Download Fuzzy Review CSV",
                    data=csv_fuzzy,
                    file_name="fuzzy_review_candidates.csv",
                    mime="text/csv"
                )
            else:
                st.info("No fuzzy discrepancies found.")

        with tab3:
            ex_col1, ex_col2 = st.columns(2)
            with ex_col1:
                st.markdown("**Source A Items Missing in Source B**")
                if len(un_a) > 0:
                    st.dataframe(un_a.head(1000).to_pandas(), width="stretch")
                    st.download_button(
                        label="📥 Download Source A Exceptions",
                        data=un_a.write_csv(),
                        file_name="unmatched_source_a.csv",
                        mime="text/csv"
                    )
                else:
                    st.success("All Source A items accounted for.")

            with ex_col2:
                st.markdown("**Source B Items Missing in Source A**")
                if len(un_b) > 0:
                    st.dataframe(un_b.head(1000).to_pandas(), width="stretch")
                    st.download_button(
                        label="📥 Download Source B Exceptions",
                        data=un_b.write_csv(),
                        file_name="unmatched_source_b.csv",
                        mime="text/csv"
                    )
                else:
                    st.success("All Source B items accounted for.")
