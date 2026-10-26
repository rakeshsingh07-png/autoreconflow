import streamlit as st
import polars as pl
from parser_engine import parse_to_polars
from recon_engine import execute_high_volume_reconciliation

st.set_page_config(page_title="AutoReconFlow | Enterprise Suite", layout="wide")
st.title("⚡ AutoReconFlow — Enterprise Autonomous Reconciliation")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Source A (Ledger / Internal Records)")
    file_a = st.file_uploader("Upload CSV, XLSX, or Bank Statement PDF", type=["csv", "xlsx", "xls", "pdf"], key="file_a")

with col2:
    st.subheader("Source B (Bank / Tax Portal / External Records)")
    file_b = st.file_uploader("Upload CSV, XLSX, or Bank Statement PDF", type=["csv", "xlsx", "xls", "pdf"], key="file_b")

if file_a and file_b:
    with st.spinner("Loading documents via Polars multi-core parser..."):
        df_a = parse_to_polars(file_a)
        df_b = parse_to_polars(file_b)

    st.success(f"Loaded Source A ({len(df_a):,} rows) and Source B ({len(df_b):,} rows)")

    # Mapping selectors
    with st.expander("Configure Column Mappings", expanded=True):
        m1, m2 = st.columns(2)
        with m1:
            st.markdown("**Source A Columns**")
            id_a = st.selectbox("Document / Txn ID", df_a.columns, index=0)
            party_a = st.selectbox("Party / Narrative", df_a.columns, index=min(1, len(df_a.columns)-1))
            amt_a = st.selectbox("Amount Column", df_a.columns, index=min(2, len(df_a.columns)-1))
        with m2:
            st.markdown("**Source B Columns**")
            id_b = st.selectbox("Document / Txn ID", df_b.columns, index=0, key="id_b")
            party_b = st.selectbox("Party / Narrative", df_b.columns, index=min(1, len(df_b.columns)-1), key="party_b")
            amt_b = st.selectbox("Amount Column", df_b.columns, index=min(2, len(df_b.columns)-1), key="amt_b")

    tolerance = st.slider("Allowed Amount Tolerance", 0.00, 5.00, 1.00, step=0.10)

    if st.button("🚀 Run Automated Reconciliation", type="primary", width="stretch"):
        map_a = {"id": id_a, "party": party_a, "amt": amt_a}
        map_b = {"id": id_b, "party": party_b, "amt": amt_b}

        with st.spinner("Processing through 3-Tier engine..."):
            matched, fuzzy, un_a, un_b = execute_high_volume_reconciliation(
                df_a, df_b, map_a, map_b, tolerance=tolerance
            )

        st.subheader("📊 Reconciliation Summary")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("100% Matched", f"{len(matched):,} records")
        kpi2.metric("Fuzzy Matches", f"{len(fuzzy):,} records")
        kpi3.metric("Unmatched (Source A)", f"{len(un_a):,} records")
        kpi4.metric("Unmatched (Source B)", f"{len(un_b):,} records")

        tab1, tab2, tab3 = st.tabs(["Matched Records", "Fuzzy Review", "Exceptions"])
        
        with tab1:
            st.dataframe(matched.head(1000).to_pandas(), width="stretch")
            st.download_button("📥 Download Matched CSV", matched.write_csv(), "matched_ledger.csv")
            
        with tab2:
            if len(fuzzy) > 0:
                st.dataframe(fuzzy.to_pandas(), width="stretch")
            else:
                st.info("No fuzzy discrepancies found.")
                
        with tab3:
            st.write("Source A Unmatched Records:")
            st.dataframe(un_a.head(1000).to_pandas(), width="stretch")
