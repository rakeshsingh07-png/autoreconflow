import re
import polars as pl
from rapidfuzz import fuzz

def clean_alphanumeric_id(val: str) -> str:
    """Normalizes document and transaction identifiers by stripping punctuation and prefixes."""
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
    fuzzy_threshold: float = 70.0
):
    """
    High-Throughput Vectorized 3-Tier Financial Reconciliation Engine.
    Tier 1: Deterministic Exact Match (Normalized ID + Amount within Tolerance)
    Tier 2: Fuzzy / Discrepancy Match (Side-by-Side Amounts & Variance)
    Tier 3: True Exceptions (Unmatched A & Unmatched B)
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
    )

    matched_a_indices = set(exact_matches["_row_idx_a"].to_list())
    matched_b_indices = set(exact_matches["_row_idx_b"].to_list())

    # --- TIER 2: FUZZY CANDIDATES (Side-by-Side Amount Layout) ---
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

                diff = round(abs(r_a["_match_amt"] - r_b["_match_amt"]), 2)
                if diff <= tolerance:
                    id_score = fuzz.token_sort_ratio(r_a["_match_id"], r_b["_match_id"])
                    party_score = fuzz.partial_ratio(r_a["_match_party"], r_b["_match_party"])

                    if (
                        r_a["_match_id"] in r_b["_match_id"]
                        or r_b["_match_id"] in r_a["_match_id"]
                        or id_score >= fuzzy_threshold
                        or party_score >= 80.0
                    ):
                        used_a_idx.add(idx_a)
                        used_b_idx.add(idx_b)

                        # Clean row structure with side-by-side comparison columns
                        row_entry = {
                            "Doc_ID_A": str(r_a.get(id_a, "")),
                            "Doc_ID_B": str(r_b.get(id_b, "")),
                            "Party_A": str(r_a.get(party_a, "")),
                            "Party_B": str(r_b.get(party_b, "")),
                            "Amount_Source_A": r_a["_match_amt"],
                            "Amount_Source_B": r_b["_match_amt"],
                            "Variance": diff,
                            "Confidence_Score": f"{max(id_score, party_score):.1f}%",
                            "Match_Status": "Fuzzy Discrepancy Review"
                        }

                        # Append remaining original metadata fields without internal keys
                        for k, v in r_a.items():
                            if k not in row_entry and not k.startswith("_"):
                                row_entry[f"{k}_A"] = v
                        for k, v in r_b.items():
                            if k not in row_entry and not k.startswith("_"):
                                row_entry[f"{k}_B"] = v

                        fuzzy_matches_list.append(row_entry)
                        break

    fuzzy_out_df = pl.from_dicts(fuzzy_matches_list) if fuzzy_matches_list else pl.DataFrame()

    # --- TIER 3: OPEN EXCEPTIONS ---
    all_excluded_a = list(matched_a_indices.union(used_a_idx))
    all_excluded_b = list(matched_b_indices.union(used_b_idx))

    unmatched_a = a_proc.filter(~pl.col("_row_idx_a").is_in(all_excluded_a))
    unmatched_b = b_proc.filter(~pl.col("_row_idx_b").is_in(all_excluded_b))

    # Strip internal processing columns
    def clean_output_df(df: pl.DataFrame) -> pl.DataFrame:
        cols_to_drop = [c for c in df.columns if c.startswith("_") or c.endswith("_B_B") or "_row_idx" in c or "_match_" in c]
        return df.drop(cols_to_drop) if cols_to_drop else df

    matched_out = clean_output_df(exact_matches)
    fuzzy_out = clean_output_df(fuzzy_out_df) if len(fuzzy_out_df) > 0 else pl.DataFrame()
    un_a_out = clean_output_df(unmatched_a)
    un_b_out = clean_output_df(unmatched_b)

    return matched_out, fuzzy_out, un_a_out, un_b_out
