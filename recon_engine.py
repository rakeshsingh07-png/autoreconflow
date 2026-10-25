import pandas as pd
import numpy as np
from rapidfuzz import fuzz

def normalize_text(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.upper()

def run_reconciliation(
    df_bank: pd.DataFrame,
    df_gl: pd.DataFrame,
    bank_ref_col: str,
    bank_amt_col: str,
    gl_ref_col: str,
    gl_amt_col: str,
    amt_tolerance: float = 0.01,
    enable_fuzzy: bool = False,
    fuzzy_threshold: int = 85
):
    b = df_bank.copy()
    g = df_gl.copy()
    
    b['_b_idx'] = b.index
    g['_g_idx'] = g.index
    
    b['_ref'] = normalize_text(b[bank_ref_col])
    g['_ref'] = normalize_text(g[gl_ref_col])
    
    b['_amt'] = pd.to_numeric(b[bank_amt_col], errors='coerce').fillna(0.0).round(2)
    g['_amt'] = pd.to_numeric(g[gl_amt_col], errors='coerce').fillna(0.0).round(2)
    
    b['Recon_Status'] = 'Unmatched'
    g['Recon_Status'] = 'Unmatched'
    
    matched_pairs = []
    match_counter = 1
    
    # Pass 1: Exact Key + Exact Amount
    merged_exact = pd.merge(
        b[b['Recon_Status'] == 'Unmatched'],
        g[g['Recon_Status'] == 'Unmatched'],
        on=['_ref'],
        suffixes=('_bank', '_gl')
    )
    
    merged_exact['amt_diff'] = (merged_exact['_amt_bank'] - merged_exact['_amt_gl']).abs()
    exact_matches = merged_exact[merged_exact['amt_diff'] <= amt_tolerance]
    
    exact_used_b = set()
    exact_used_g = set()
    
    for _, row in exact_matches.iterrows():
        b_idx = int(row['_b_idx'])
        g_idx = int(row['_g_idx'])
        if b_idx not in exact_used_b and g_idx not in exact_used_g:
            exact_used_b.add(b_idx)
            exact_used_g.add(g_idx)
            matched_pairs.append({
                'Match_ID': f"M-{match_counter:05d}",
                'Match_Type': 'Exact Key & Amount',
                'Bank_Ref': row['_ref'],
                'GL_Ref': row['_ref'],
                'Bank_Amount': row['_amt_bank'],
                'GL_Amount': row['_amt_gl'],
                'Variance': round(row['_amt_bank'] - row['_amt_gl'], 2),
                '_b_idx': b_idx,
                '_g_idx': g_idx
            })
            match_counter += 1
            
    b.loc[list(exact_used_b), 'Recon_Status'] = 'Matched'
    g.loc[list(exact_used_g), 'Recon_Status'] = 'Matched'

    # Pass 2: Optional Fuzzy Match
    if enable_fuzzy:
        unmatched_b = b[b['Recon_Status'] == 'Unmatched']
        unmatched_g = g[g['Recon_Status'] == 'Unmatched']
        
        fuzzy_used_b = set()
        fuzzy_used_g = set()
        
        for _, b_row in unmatched_b.iterrows():
            b_idx = int(b_row['_b_idx'])
            candidate_g = unmatched_g[
                (~unmatched_g['_g_idx'].isin(fuzzy_used_g)) &
                ((unmatched_g['_amt'] - b_row['_amt']).abs() <= amt_tolerance)
            ]
            
            for _, g_row in candidate_g.iterrows():
                g_idx = int(g_row['_g_idx'])
                similarity = fuzz.token_sort_ratio(str(b_row['_ref']), str(g_row['_ref']))
                if similarity >= fuzzy_threshold:
                    fuzzy_used_b.add(b_idx)
                    fuzzy_used_g.add(g_idx)
                    matched_pairs.append({
                        'Match_ID': f"M-{match_counter:05d}",
                        'Match_Type': f'Fuzzy ({similarity}% Match)',
                        'Bank_Ref': b_row['_ref'],
                        'GL_Ref': g_row['_ref'],
                        'Bank_Amount': b_row['_amt'],
                        'GL_Amount': g_row['_amt'],
                        'Variance': round(b_row['_amt'] - g_row['_amt'], 2),
                        '_b_idx': b_idx,
                        '_g_idx': g_idx
                    })
                    match_counter += 1
                    break
                    
        b.loc[list(fuzzy_used_b), 'Recon_Status'] = 'Matched (Fuzzy)'
        g.loc[list(fuzzy_used_g), 'Recon_Status'] = 'Matched (Fuzzy)'

    df_matches = pd.DataFrame(matched_pairs) if matched_pairs else pd.DataFrame(
        columns=['Match_ID', 'Match_Type', 'Bank_Ref', 'GL_Ref', 'Bank_Amount', 'GL_Amount', 'Variance']
    )
    
    b_out = b.drop(columns=['_b_idx', '_ref', '_amt'], errors='ignore')
    g_out = g.drop(columns=['_g_idx', '_ref', '_amt'], errors='ignore')
    
    return df_matches, b_out, g_out