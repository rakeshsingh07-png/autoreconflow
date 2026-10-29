import re
import polars as pl
from rapidfuzz import fuzz

def clean_alphanumeric_id(val: str) -> str:
    """Normalizes document and transaction identifiers by stripping noise."""
    if val is None:
        return ""
    s = str(val).strip().upper()
    s = re.sub(r'[^A-Z0-9]', '', s)
    s = re.sub(r'^(INV|BILL|DOC|TXN|SETTLE|TRANS)', '', s)
    s = re.sub(r'^(202[4-9]|2[4-9])', '', s)
    s = s.lstrip('0')
    return s if s else "0"

def clean_party_string(val: str) -> str:
    """Normalizes vendor and corporate entity names for fuzzy matching."""
    if val is None:
        return ""
    s = str(val).strip().upper()
    s = re.sub(r'[^A-Z0-9\s]', ' ', s)
    legal_stopwords = ['LTD', 'LIMITED', 'PVT', 'PRIVATE', 'INC', 'INCORPORATED', 'LLC', 'CORP', 'CORPORATION', 'CO', 'COMPANY', 'SERVICES', 'SOLUTIONS']
    tokens = [t for t in s.split() if t not in legal_stopwords]
    return " ".join(tokens) if tokens else s

def execute_high_volume_reconciliation(
    df_a: pl.DataFrame,
    df_b: pl.DataFrame,
    map_a: dict,
    map_b: dict,
    tolerance: float = 1.0,
    fuzzy_threshold: float = 85.0
):
    """
    High-Throughput Vectorized 3-Tier Financial Reconciliation Engine.
    Tier 1: Deterministic Exact Match (Normalized ID + Amount within Tolerance)
    Tier 2: Discrepancy & Threshold Variance Review (Materiality Tracking)
    Tier 3: Open Exceptions (Source A Open vs Source B Open)
    """
    id_a, party_a, amt_a = map_a["id"], map_a["party"], map_a["amt"]
    id_b, party_b, amt_b = map_b["id"], map_b["party"], map_b["amt"]

    # Pre-process Source A in Polars
    a_proc = df_a.clone().with_columns([
        pl.col(id_a).cast(pl.Utf8).map_elements(clean_alphanumeric_id, return_dtype=pl.Utf8).alias("_match_id"),
        pl.col(party_a).cast(pl.Utf8).map_elements(clean_party_string, return_dtype=pl.Utf8).alias("_match_party"),
        pl.col(amt_a).cast(pl.Float64, strict=False).fill_null(0.0).alias("_match_amt"),
        pl.int_range(0, pl.len()).alias("_row_idx_a")
    ])

    # Pre-process Source B in Polars
    b_proc = df_b.clone().with_columns([
        pl.col(id_b).cast(pl.Utf8).map_elements(clean_alphanumeric_id, return_dtype=pl.Utf8).alias("_match_id"),
        pl.col(party_b).cast(pl.Utf8).map_elements(clean_party_string, return_dtype=pl.Utf8).alias("_match_party"),
        pl.col(amt_b).cast(pl.Float64, strict=False).fill_null(0.0).alias("_match_amt"),
        pl.int_range(0, pl.len()).alias("_row_idx_b")
    ])

    # --- TIER 1: DETERMINISTIC MATCHING (Vectorized Join) ---
    joined_id = a_proc.join(b_proc, on="_match_id", how="inner", suffix="_B")
    
    exact_matches = joined_id.filter(
        (pl.col("_match_amt") - pl.col("_match_amt_B")).abs() <= tolerance
    )

    matched_a_indices = set(exact_matches["_row_idx_a"].to_list())
    matched_b_indices = set(exact_matches["_row_idx_b"].to_list())

    # --- TIER 2: DISCREPANCY & THRESHOLD CANDIDATES ---
    value_discrepancies = joined_id.filter(
        (pl.col("_match_amt") - pl.col("_match_amt_B")).abs() > tolerance
    ).filter(
        ~pl.col("_row_idx_a").is_in(list(matched_a_indices)) &
        ~pl.col("_row_idx_b").is_in(list(matched_b_indices))
    )

    fuzzy_a_indices = set(value_discrepancies["_row_idx_a"].to_list())
    fuzzy_b_indices = set(value_discrepancies["_row_idx_b"].to_list())

    # --- TIER 3: OPEN EXCEPTIONS ---
    excluded_a = list(matched_a_indices.union(fuzzy_a_indices))
    excluded_b = list(matched_b_indices.union(fuzzy_b_indices))

    unmatched_a = a_proc.filter(~pl.col("_row_idx_a").is_in(excluded_a))
    unmatched_b = b_proc.filter(~pl.col("_row_idx_b").is_in(excluded_b))

    drop_helpers = ["_match_id", "_match_party", "_match_amt", "_row_idx_a", "_row_idx_b", "_match_id_B", "_match_party_B", "_match_amt_B"]

    matched_out = exact_matches.drop([c for c in drop_helpers if c in exact_matches.columns])
    fuzzy_out = value_discrepancies.drop([c for c in drop_helpers if c in value_discrepancies.columns])
    un_a_out = unmatched_a.drop([c for c in drop_helpers if c in unmatched_a.columns])
    un_b_out = unmatched_b.drop([c for c in drop_helpers if c in unmatched_b.columns])

    return matched_out, fuzzy_out, un_a_out, un_b_out
