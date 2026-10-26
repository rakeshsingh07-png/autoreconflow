# Mapping selectors (Safe indexing)
    with st.expander("Configure Column Mappings", expanded=True):
        m1, m2 = st.columns(2)
        cols_a = list(df_a.columns)
        cols_b = list(df_b.columns)
        
        with m1:
            st.markdown("**Source A Columns**")
            id_a = st.selectbox("Document / Txn ID", cols_a, index=0)
            party_a = st.selectbox("Party / Narrative", cols_a, index=min(1, len(cols_a)-1))
            amt_a = st.selectbox("Amount Column", cols_a, index=min(2, len(cols_a)-1))
        with m2:
            st.markdown("**Source B Columns**")
            id_b = st.selectbox("Document / Txn ID", cols_b, index=0, key="id_b")
            party_b = st.selectbox("Party / Narrative", cols_b, index=min(1, len(cols_b)-1), key="party_b")
            amt_b = st.selectbox("Amount Column", cols_b, index=min(2, len(cols_b)-1), key="amt_b")
