import polars as pl
from rapidfuzz import fuzz

def normalize_dataset(df: pl.DataFrame, id_col: str, party_col: str, amt_col: str) -> pl.DataFrame:
    """Normalizes identifiers and amounts for rapid vector joins."""
    return df.select([
        pl.col(id_col).cast(pl.Utf8).str.to_uppercase().str.replace_all(r"[^A-Z0-9]", "").alias("clean_id"),
        pl.col(party_col).cast(pl.Utf8).alias("party"),
        pl.col(amt_col).cast(pl.Utf8)
          .str.replace_all(r"[₹$€£,\s]", "")
          .str.replace(r"^\((.*)\)$", "-$1")
          .cast(pl.Float64, strict=False)
          .fill_null(0.0)
          .abs()
          .alias("clean_amount")
    ])

def execute_high_volume_reconciliation(df_a: pl.DataFrame, df_b: pl.DataFrame, 
                                      map_a: dict, map_b: dict, 
                                      tolerance: float = 1.0):
    """
    Executes high-speed 3-tier reconciliation across large datasets with safe null guards.
    """
    clean_a = normalize_dataset(df_a, map_a["id"], map_a["party"], map_a["amt"])
    clean_b = normalize_dataset(df_b, map_b["id"], map_b["party"], map_b["amt"])

    # PASS 1: Parallel Outer Hash Join
    joined = clean_a.join(
        clean_b,
        on="clean_id",
        how="full",
        suffix="_b"
    ).with_columns([
        (pl.col("clean_amount") - pl.col("clean_amount_b")).abs().alias("amount_diff")
    ])

    matched = joined.filter(
        (pl.col("clean_id").is_not_null()) &
        (pl.col("clean_amount_b").is_not_null()) &
        (pl.col("amount_diff") <= tolerance)
    )

    unmatched_a = joined.filter(
        pl.col("clean_amount_b").is_null() | (pl.col("amount_diff") > tolerance)
    ).select([
        pl.col("clean_id"),
        pl.col("party"),
        pl.col("clean_amount").alias("amount")
    ])

    unmatched_b = joined.filter(
        pl.col("clean_id_b").is_null()
    ).select([
        pl.col("clean_id_b").alias("clean_id"),
        pl.col("party_b").alias("party"),
        pl.col("clean_amount_b").alias("amount")
    ])

    # PASS 2: Safe Partitioned Fuzzy Matching
    fuzzy_matches = []
    if len(unmatched_a) > 0 and len(unmatched_b) > 0:
        a_records = unmatched_a.to_dicts()
        b_records = unmatched_b.to_dicts()
        
        # Partition with safe float/int conversion
        b_buckets = {}
        for r in b_records:
            val = r.get("amount")
            amt_num = int(float(val)) if val is not None else 0
            b_buckets.setdefault(amt_num, []).append(r)

        for item_a in a_records:
            val_a = item_a.get("amount")
            target = int(float(val_a)) if val_a is not None else 0
            
            candidates = b_buckets.get(target, []) + b_buckets.get(target - 1, []) + b_buckets.get(target + 1, [])
            
            for cand in candidates:
                score = fuzz.ratio(str(item_a.get("clean_id", "")), str(cand.get("clean_id", "")))
                if score >= 80:
                    fuzzy_matches.append({
                        "id_internal": item_a.get("clean_id"),
                        "id_external": cand.get("clean_id"),
                        "party_internal": item_a.get("party"),
                        "party_external": cand.get("party"),
                        "amount": item_a.get("amount"),
                        "confidence_score": score
                    })
                    break

    fuzzy_df = pl.DataFrame(fuzzy_matches) if fuzzy_matches else pl.DataFrame()

    return matched, fuzzy_df, unmatched_a, unmatched_b
