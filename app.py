import io
import streamlit as st
import pandas as pd
from recon_engine import run_reconciliation

st.set_page_config(
    page_title="AutoReconFlow | Financial Reconciliation Engine",
    page_icon="⚖️",
    layout="wide"
)

st.title("⚖️ AutoReconFlow")
st.caption("Deterministic Multi-Source Financial & Bank Reconciliation Platform")

with st.sidebar:
    st.header("1. Upload Datasets")
    file_bank = st.file_uploader("Upload Bank / Source Statement (CSV/XLSX)", type=["csv", "xlsx"])
    file_gl = st.file_uploader("Upload General Ledger / ERP Extract (CSV/XLSX)", type=["csv", "xlsx"])
    
    st.divider()
    st.header("2. Matching Rules")
    tolerance = st.number_input("Amount Tolerance ($ / ₹)", min_value=0.00, value=0.01, step=0.01)
    use_fuzzy = st.checkbox("Enable Fuzzy Text Matching", value=True)
    fuzzy_cutoff = st.slider("Similarity Threshold (%)", 50, 100, 85) if use_fuzzy else 100

def load_file(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)

if file_bank and file_gl:
    df_bank = load_file(file_bank)
    df_gl = load_file(file_gl)
    
    st.subheader("Field Mapping")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Source A (Bank / Claims)**")
        bank_ref = st.selectbox("Reference / Key Column", df_bank.columns, key="b_ref")
        bank_amt = st.selectbox("Amount Column", df_bank.columns, key="b_amt")
        
    with col2:
        st.markdown("**Source B (General Ledger / Core ERP)**")
        gl_ref = st.selectbox("Reference / Key Column", df_gl.columns, key="g_ref")
        gl_amt = st.selectbox("Amount Column", df_gl.columns, key="g_amt")
        
    if st.button("Run Reconciliation Engine", type="primary"):
        with st.spinner("Processing ledger transactions..."):
            matches, bank_res, gl_res = run_reconciliation(
                df_bank, df_gl,
                bank_ref, bank_amt,
                gl_ref, gl_amt,
                amt_tolerance=tolerance,
                enable_fuzzy=use_fuzzy,
                fuzzy_threshold=fuzzy_cutoff
            )
            
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Bank Records", len(bank_res))
        m2.metric("Total GL Records", len(gl_res))
        m3.metric("Matched Pairs", len(matches))
        unmatched_total = len(bank_res[bank_res['Recon_Status'] == 'Unmatched']) + len(gl_res[gl_res['Recon_Status'] == 'Unmatched'])
        m4.metric("Unmatched Exceptions", unmatched_total)
        
        tab1, tab2, tab3 = st.tabs(["Matched Records", "Bank Exceptions", "GL Exceptions"])
        
        with tab1:
            st.dataframe(matches, use_container_width=True)
        with tab2:
            st.dataframe(bank_res[bank_res['Recon_Status'] == 'Unmatched'], use_container_width=True)
        with tab3:
            st.dataframe(gl_res[gl_res['Recon_Status'] == 'Unmatched'], use_container_width=True)
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            matches.to_excel(writer, sheet_name='Matched_Summary', index=False)
            bank_res.to_excel(writer, sheet_name='Bank_Full_Audit', index=False)
            gl_res.to_excel(writer, sheet_name='GL_Full_Audit', index=False)
            
        st.download_button(
            label="📥 Download Full Reconciliation Audit Package (.xlsx)",
            data=output.getvalue(),
            file_name="AutoReconFlow_Reconciliation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload both statement and ledger datasets from the sidebar to begin reconciliation.")