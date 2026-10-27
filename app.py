import io
import re
import streamlit as st
import pandas as pd
import polars as pl
from recon_engine import execute_high_volume_reconciliation

st.set_page_config(
    page_title="AutoReconFlow — Enterprise Autonomous Reconciliation",
    page_icon="⚡",
    layout="wide"
)

# Sidebar
st.sidebar.title("⚡ AutoReconFlow")
st.sidebar.markdown("**Enterprise B2B Edition v2.0**")
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

# --- Multi-Tab Consolidated Excel Generator ---
def generate_consolidated_workbook(matched_df, fuzzy_df, un_a_df, un_b_df) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df = pd.DataFrame({
            "Reconciliation Category": [
                "100% Deterministic Matched",
                "Fuzzy Review Candidates",
                "Source A Exceptions (Ledger Missing in Bank)",
                "Source B Exceptions (Bank Missing in Ledger)"
            ],
            "Record Count": [
                len(matched_df),
                len(fuzzy_df),
                len(un_a_df),
                len(un_b_df)
            ]
        })
        summary_df.to_excel(writer, sheet_name="Recon Summary", index=False)
        matched_df.to_excel(writer, sheet_name="Matched Records", index=False)
        fuzzy_df.to_excel(writer, sheet_name="Fuzzy Candidates", index=False)
        un_a_df.to_excel(writer, sheet_name="Exceptions Source A", index=False)
        un_b_df.to_excel(writer, sheet_name="Exceptions Source B", index=False)
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
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for r in table:
                        clean_r = [str(c).strip() if c is not None else "" for c in r]
                        if any(clean_r):
                            rows.append(clean_r)
        if not rows:
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

        # Convert Polars DataFrames to Pandas for clean export
        matched_pd = matched.to_pandas()
        fuzzy_pd = fuzzy.to_pandas()
        un_a_pd = un_a.to_pandas()
        un_b_pd = un_b.to_pandas()

        st.subheader("📊 Reconciliation Summary")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("100% Deterministic Matched", f"{len(matched_pd):,} records")
        k2.metric("Fuzzy Candidates", f"{len(fuzzy_pd):,} records")
        k3.metric("Unmatched (Source A)", f"{len(un_a_pd):,} records")
        k4.metric("Unmatched (Source B)", f"{len(un_b_pd):,} records")

        # Top-level Consolidated Multi-Tab Excel Export Button
        excel_bytes = generate_consolidated_workbook(matched_pd, fuzzy_pd, un_a_pd, un_b_pd)
        st.download_button(
            label="📦 Download Complete Consolidated Workbook (Excel .xlsx)",
            data=excel_bytes,
            file_name="AutoReconFlow_Consolidated_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_download_master_excel"
        )
        st.write("")

        # Distinct 4 tabs: Matched, Fuzzy, and Separate Exceptions for Source A & B
        tab1, tab2, tab3, tab4 = st.tabs([
            "✅ Matched Records",
            "🔍 Fuzzy Review",
            "⚠️ Exceptions: Source A (Ledger)",
            "⚠️ Exceptions: Source B (External / Bank)"
        ])

        with tab1:
            if len(matched_pd) > 0:
                st.dataframe(matched_pd, use_container_width=True)
                st.download_button(
                    label="📥 Download Matched CSV",
                    data=matched.write_csv(),
                    file_name="matched_records.csv",
                    mime="text/csv",
                    key="btn_matched_csv"
                )
            else:
                st.info("No exact matches found under current settings.")

        with tab2:
            if len(fuzzy_pd) > 0:
                st.dataframe(fuzzy_pd, use_container_width=True)
                st.download_button(
                    label="📥 Download Fuzzy Review CSV",
                    data=fuzzy.write_csv(),
                    file_name="fuzzy_candidates.csv",
                    mime="text/csv",
                    key="btn_fuzzy_csv"
                )
            else:
                st.info("No fuzzy discrepancies found.")

        with tab3:
            if len(un_a_pd) > 0:
                st.markdown("**Source A Entries Missing in Source B (Unmatched Ledger)**")
                st.dataframe(un_a_pd, use_container_width=True)
                st.download_button(
                    label="📥 Download Source A Exceptions CSV",
                    data=un_a.write_csv(),
                    file_name="exceptions_source_a_ledger.csv",
                    mime="text/csv",
                    key="btn_un_a_csv"
                )
            else:
                st.success("All Source A items accounted for (0 exceptions).")

        with tab4:
            if len(un_b_pd) > 0:
                st.markdown("**Source B Entries Missing in Source A (Unmatched Bank/Portal)**")
                st.dataframe(un_b_pd, use_container_width=True)
                st.download_button(
                    label="📥 Download Source B Exceptions CSV",
                    data=un_b.write_csv(),
                    file_name="exceptions_source_b_external.csv",
                    mime="text/csv",
                    key="btn_un_b_csv"
                )
            else:
                st.success("All Source B items accounted for (0 exceptions).")
