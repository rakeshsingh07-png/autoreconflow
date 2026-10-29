import re
import polars as pl
from rapidfuzz import fuzz

def clean_alphanumeric_id(val: str) -> str:
    """
    Big 4 Standard Document ID Cleaner:
    Strips noise, standard business prefixes, and formatting while preserving true transaction serials.
    """
    if val is None:
        return ""
    s = str(val).strip().upper()
    s = re.sub(r'[^A-Z0-9]', '', s)
    s = re.sub(r'^(INV|BILL|DOC|TXN|SETTLE|TRANS)', '', s)
    s = re.sub(r'^(202[4-9]|2[4-9])', '', s)
    s = s.lstrip('0')
    return s if s else "0"

def clean_party_string(val: str) -> str:
    """
    Normalizes legal corporate suffixes to enable objective entity matching across heterogeneous ERPs.
    """
    if val is None:
        return ""
    s = str(val).strip().upper()
    s = re.sub(r'[^A-Z0-9\s]', ' ', s)
    legal_stopwords = [
        'LTD', 'LIMITED', 'PVT', 'PRIVATE', 'INC', 'INCORPORATED', 
        'LLC', 'CORP', 'CORPORATION', 'CO', 'COMPANY', 'SERVICES', 'SOLUTIONS'
    ]
    tokens = [t for t in s.split() if t not in legal_stopwords]
    return " ".join(tokens) if tokens else s

def execute_high_volume_reconciliation(
    df_a: pl.DataFrame,
    df_b: pl.DataFrame,
    map_a: dict,
    map_b: dict,
    tolerance: float = 1.0,
    fuzzy_threshold: float = 70.0
):
    """
    Vectorized 3-Tier Reconciliation Engine adhering to PCAOB AS 2201 / ISA 500 standards.
    
    Tier 1: Deterministic Exact Match (Normalized Key + Amount within Tolerance)
    Tier 2: Discrepancy & Fuzzy Review (Amount within Tolerance + ID/Party Heuristic >= 70%)
    Tier 3: Open Exceptions (Source A Ledger Accruals vs Source B External Bank/Portal items)
    """
    id_a, party_a, amt_a = map_a["id"], map_a["party"], map_a["amt"]
    id_b, party_b, amt_b = map_b["id"], map_b["party"], map_b["amt"]

    # Pre-process Source A
    a_proc = df_a.clone().with_columns([
        pl.col(id_a).cast(pl.Utf8).map_elements(clean_alphanumeric_id, return_dtype=pl.Utf8).alias("_match_id"),
        pl.col(party_a).cast(pl.Utf8).map_elements(clean_party_string, return_dtype=pl.Utf8).alias("_match_party"),
        pl.col(amt_a).cast(pl.Float64, strict=False).fill_null(0.0).alias("_match_amt"),
        pl.int_range(0, pl.len()).alias("_row_idx_a")
    ])

    # Pre-process Source B
    b_proc = df_b.clone().with_columns([
        pl.col(id_b).cast(pl.Utf8).map_elements(clean_alphanumeric_id, return_dtype=pl.Utf8).alias("_match_id"),
        pl.col(party_b).cast(pl.Utf8).map_elements(clean_party_string, return_dtype=pl.Utf8).alias("_match_party"),
        pl.col(amt_b).cast(pl.Float64, strict=False).fill_null(0.0).alias("_match_amt"),
        pl.int_range(0, pl.len()).alias("_row_idx_b")
    ])

    # --- TIER 1: DETERMINISTIC EXACT MATCH ---
    joined_id = a_proc.join(b_proc, on="_match_id", how="inner", suffix="_B")
    exact_matches = joined_id.filter(
        (pl.col("_match_amt") - pl.col("_match_amt_B")).abs() <= tolerance
    ).with_columns([
        pl.lit("Cleared - Deterministic Key & Value Match").alias("Audit_Assertion_Status"),
        (pl.col("_match_amt") - pl.col("_match_amt_B")).round(2).alias("Variance")
    ])

    matched_a_indices = set(exact_matches["_row_idx_a"].to_list())
    matched_b_indices = set(exact_matches["_row_idx_b"].to_list())

    # --- TIER 2: HEURISTIC / FUZZY REVIEW (Amount Match + Text Similarity) ---
    unmatched_a_p1 = a_proc.filter(~pl.col("_row_idx_a").is_in(list(matched_a_indices)))
    unmatched_b_p1 = b_proc.filter(~pl.col("_row_idx_b").is_in(list(matched_b_indices)))

    fuzzy_matches_list = []
    used_a_idx = set()
    used_b_idx = set()

    if len(unmatched_a_p1) > 0 and len(unmatched_b_p1) > 0:
        a_records = unmatched_a_p1.to_dicts()
        b_records = unmatched_b_p1.to_dicts()

        for r_a in a_records:
            idx_a = r_a["_row_idx_a"]
            if idx_a in used_a_idx:
                continue

            for r_b in b_records:
                idx_b = r_b["_row_idx_b"]
                if idx_b in used_b_idx:
                    continue

                diff = abs(r_a["_match_amt"] - r_b["_match_amt"])
                if diff <= tolerance:
                    id_score = fuzz.token_sort_ratio(r_a["_match_id"], r_b["_match_id"])
                    party_score = fuzz.partial_ratio(r_a["_match_party"], r_b["_match_party"])

                    # Catches variations like INV-2026-104 vs INV-2026-104-MISC
                    if (
                        r_a["_match_id"] in r_b["_match_id"]
                        or r_b["_match_id"] in r_a["_match_id"]
                        or id_score >= fuzzy_threshold
                        or party_score >= 80.0
                    ):
                        used_a_idx.add(idx_a)
                        used_b_idx.add(idx_b)

                        merged_row = {**r_a}
                        for k, v in r_b.items():
                            merged_row[f"{k}_B"] = v
                        merged_row["Fuzzy_Confidence"] = f"{max(id_score, party_score):.1f}%"
                        merged_row["Variance"] = round(r_a["_match_amt"] - r_b["_match_amt"], 2)
                        merged_row["Audit_Assertion_Status"] = "Materiality Review - Value Matched with Suffix/Name Variance"
                        fuzzy_matches_list.append(merged_row)
                        break

    fuzzy_out_df = pl.from_dicts(fuzzy_matches_list) if fuzzy_matches_list else pl.DataFrame()

    # --- TIER 3: OPEN EXCEPTIONS ---
    all_excluded_a = list(matched_a_indices.union(used_a_idx))
    all_excluded_b = list(matched_b_indices.union(used_b_idx))

    unmatched_a = a_proc.filter(~pl.col("_row_idx_a").is_in(all_excluded_a)).with_columns(
        pl.lit("Open Exception: Subledger Entry Missing in External Records").alias("Audit_Assertion_Status")
    )
    unmatched_b = b_proc.filter(~pl.col("_row_idx_b").is_in(all_excluded_b)).with_columns(
        pl.lit("Open Exception: External Record Missing in Subledger").alias("Audit_Assertion_Status")
    )

    drop_helpers = [
        "_match_id", "_match_party", "_match_amt", "_row_idx_a", "_row_idx_b",
        "_match_id_B", "_match_party_B", "_match_amt_B"
    ]

    matched_out = exact_matches.drop([c for c in drop_helpers if c in exact_matches.columns])
    fuzzy_out = fuzzy_out_df.drop([c for c in drop_helpers if c in fuzzy_out_df.columns]) if len(fuzzy_out_df) > 0 else pl.DataFrame()
    un_a_out = unmatched_a.drop([c for c in drop_helpers if c in unmatched_a.columns])
    un_b_out = unmatched_b.drop([c for c in drop_helpers if c in unmatched_b.columns])

    return matched_out, fuzzy_out, un_a_out, un_b_out
