import polars as pl
import pandas as pd
from rapidfuzz import fuzz

def clean_series_id(s):
    """Universal invoice/doc cleaner: strips punctuation, spaces, and leading zeros."""
    return (
        s.astype(str)
        .str.upper()
        .str.replace(r"[^A-Z0-9]", "", regex=True)
        .str.lstrip("0")
    )

def clean_series_amount(s):
    """Universal numeric cleaner: strips currency symbols, commas, and handles accounting negatives."""
    cleaned = (
        s.astype(str)
        .str.replace(r"[₹$€£,\s]", "", regex=True)
        .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0).abs()

def execute_high_volume_reconciliation(df_a_pl: pl.DataFrame, df_b_pl: pl.DataFrame, 
                                      map_a: dict, map_b: dict, 
                                      tolerance: float = 1.0):
    """
    3-Tier Reconciliation Engine:
    Pass 1: Deterministic Exact Match (Normalized ID + Amount within Tolerance)
    Pass 2: Partitioned Fuzzy Matching (Amount match + Identifier typo/suffix similarity)
    Pass 3: Unmatched Exceptions
    """
    # Convert to pandas for robust hybrid join & string matching
    df_a = df_a_pl.to_pandas()
    df_b = df_b_pl.to_pandas()

    # 1. Normalize Keys & Amounts
    df_a["_clean_id"] = clean_series_id(df_a[map_a["id"]])
    df_a["_clean_party"] = df_a[map_a["party"]].astype(str).str.strip()
    df_a["_clean_amount"] = clean_series_amount(df_a[map_a["amt"]])
    df_a["_row_id_a"] = df_a.index

    df_b["_clean_id"] = clean_series_id(df_b[map_b["id"]])
    df_b["_clean_party"] = df_b[map_b["party"]].astype(str).str.strip()
    df_b["_clean_amount"] = clean_series_amount(df_b[map_b["amt"]])
    df_b["_row_id_b"] = df_b.index

    # ---------------------------------------------------------
    # PASS 1: Exact Deterministic Match
    # ---------------------------------------------------------
    matched_rows = []
    matched_a_indices = set()
    matched_b_indices = set()

    # Create fast lookup dictionary for Source B by cleaned ID
    b_lookup_by_id = {}
    for idx_b, row_b in df_b.iterrows():
        b_lookup_by_id.setdefault(row_b["_clean_id"], []).append((idx_b, row_b))

    for idx_a, row_a in df_a.iterrows():
        c_id = row_a["_clean_id"]
        if c_id in b_lookup_by_id:
            for idx_b, row_b in b_lookup_by_id[c_id]:
                if idx_b not in matched_b_indices:
                    diff = round(abs(row_a["_clean_amount"] - row_b["_clean_amount"]), 2)
                    if diff <= tolerance:
                        matched_a_indices.add(idx_a)
                        matched_b_indices.add(idx_b)
                        matched_rows.append({
                            "Matched_ID": c_id,
                            "Doc_ID_A": str(row_a[map_a["id"]]),
                            "Doc_ID_B": str(row_b[map_b["id"]]),
                            "Party_A": row_a["_clean_party"],
                            "Party_B": row_b["_clean_party"],
                            "Amount_A": row_a["_clean_amount"],
                            "Amount_B": row_b["_clean_amount"],
                            "Variance": diff,
                            "Status": "Exact Match"
                        })
                        break

    # ---------------------------------------------------------
    # PASS 2: Heuristic & Fuzzy Matching on Unmatched Items
    # ---------------------------------------------------------
    unmatched_a_idx = [i for i in df_a.index if i not in matched_a_indices]
    unmatched_b_idx = [i for i in df_b.index if i not in matched_b_indices]

    fuzzy_rows = []
    for idx_a in unmatched_a_idx:
        row_a = df_a.loc[idx_a]
        for idx_b in unmatched_b_idx:
            if idx_b in matched_b_indices:
                continue
            row_b = df_b.loc[idx_b]
            
            diff = round(abs(row_a["_clean_amount"] - row_b["_clean_amount"]), 2)
            if diff <= tolerance:
                # Compare fuzzy similarity between identifiers
                id_score = fuzz.token_sort_ratio(str(row_a["_clean_id"]), str(row_b["_clean_id"]))
                party_score = fuzz.partial_ratio(str(row_a["_clean_party"]).lower(), str(row_b["_clean_party"]).lower())
                
                # Check if either ID has typos/extra tags OR parties strongly match
                if id_score >= 70 or party_score >= 80:
                    matched_a_indices.add(idx_a)
                    matched_b_indices.add(idx_b)
                    fuzzy_rows.append({
                        "Doc_ID_A": str(row_a[map_a["id"]]),
                        "Doc_ID_B": str(row_b[map_b["id"]]),
                        "Party_A": row_a["_clean_party"],
                        "Party_B": row_b["_clean_party"],
                        "Amount_A": row_a["_clean_amount"],
                        "Amount_B": row_b["_clean_amount"],
                        "ID_Confidence": f"{id_score}%",
                        "Party_Confidence": f"{party_score}%",
                        "Status": "Fuzzy Match Review"
                    })
                    break

    # ---------------------------------------------------------
    # PASS 3: Remaining Exceptions
    # ---------------------------------------------------------
    final_unmatched_a = df_a[~df_a.index.isin(matched_a_indices)][[map_a["id"], map_a["party"], map_a["amt"]]].copy()
    final_unmatched_b = df_b[~df_b.index.isin(matched_b_indices)][[map_b["id"], map_b["party"], map_b["amt"]]].copy()

    # Convert back to Polars for clean table rendering and streaming export
    matched_df = pl.from_pandas(pd.DataFrame(matched_rows)) if matched_rows else pl.DataFrame()
    fuzzy_df = pl.from_pandas(pd.DataFrame(fuzzy_rows)) if fuzzy_rows else pl.DataFrame()
    unmatched_a_df = pl.from_pandas(final_unmatched_a) if len(final_unmatched_a) > 0 else pl.DataFrame()
    unmatched_b_df = pl.from_pandas(final_unmatched_b) if len(final_unmatched_b) > 0 else pl.DataFrame()

    return matched_df, fuzzy_df, unmatched_a_df, unmatched_b_df
