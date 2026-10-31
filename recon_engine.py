import re
import datetime
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

def calculate_aging_days(date_val, as_at_date=None) -> tuple:
    """Calculates aging in days against an explicit 'As At' cut-off date."""
    if not date_val or date_val is None:
        return 0, "Unspecified Date"
    
    if as_at_date is None:
        as_at_date = datetime.date.today()
    elif isinstance(as_at_date, datetime.datetime):
        as_at_date = as_at_date.date()
    
    dt = None
    date_str = str(date_val).strip()
    
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            dt = datetime.datetime.strptime(date_str[:10], fmt).date()
            break
        except ValueError:
            continue
            
    if not dt:
        return 0, "Unspecified Date"
        
    delta_days = (as_at_date - dt).days
    delta_days = max(delta_days, 0)
    
    if delta_days <= 30:
        bucket = "0-30 Days (Current Transit)"
    elif delta_days <= 60:
        bucket = "31-60 Days (Follow-up Queue)"
    elif delta_days <= 90:
        bucket = "61-90 Days (Escalated Discrepancy)"
    else:
        bucket = "> 90 Days (Stale / Provision Risk)"
        
    return delta_days, bucket

def execute_high_volume_reconciliation(
    df_a: pl.DataFrame,
    df_b: pl.DataFrame,
    map_a: dict,
    map_b: dict,
    tolerance: float = 1.0,
    fuzzy_threshold: float = 70.0,
    as_at_date: datetime.date = None
):
    """
    Vectorized 3-Tier Financial Reconciliation Engine with Dynamic 'As At' Aging.
    Tier 1: Deterministic Exact Match (Normalized ID + Amount within Tolerance)
    Tier 2: Heuristic Fuzzy Review (Side-by-Side Amounts & Variance)
    Tier 3: Open Exceptions with Dynamic Aging & Risk Buckets against 'as_at_date'
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

                        for k, v in r_a.items():
                            if k not in row_entry and not k.startswith("_"):
                                row_entry[f"{k}_A"] = v
                        for k, v in r_b.items():
                            if k not in row_entry and not k.startswith("_"):
                                row_entry[f"{k}_B"] = v

                        fuzzy_matches_list.append(row_entry)
                        break

    fuzzy_out_df = pl.from_dicts(fuzzy_matches_list) if fuzzy_matches_list else pl.DataFrame()

    # --- TIER 3: OPEN EXCEPTIONS & AGING ENRICHMENT ---
    all_excluded_a = list(matched_a_indices.union(used_a_idx))
    all_excluded_b = list(matched_b_indices.union(used_b_idx))

    unmatched_a = a_proc.filter(~pl.col("_row_idx_a").is_in(all_excluded_a))
    unmatched_b = b_proc.filter(~pl.col("_row_idx_b").is_in(all_excluded_b))

    def clean_output_df(df: pl.DataFrame) -> pl.DataFrame:
        cols_to_drop = [c for c in df.columns if c.startswith("_") or c.endswith("_B_B") or "_row_idx" in c or "_match_" in c]
        return df.drop(cols_to_drop) if cols_to_drop else df

    matched_out = clean_output_df(exact_matches)
    fuzzy_out = clean_output_df(fuzzy_out_df) if len(fuzzy_out_df) > 0 else pl.DataFrame()
    un_a_out = clean_output_df(unmatched_a)
    un_b_out = clean_output_df(unmatched_b)

    # Calculate Aging for Source A Exceptions using the 'as_at_date'
    date_col_a = None
    for c in df_a.columns:
        if any(k in str(c).lower() for k in ["date", "dt", "posting", "value"]):
            date_col_a = c
            break

    if len(un_a_out) > 0 and date_col_a and date_col_a in un_a_out.columns:
        aging_res_a = [calculate_aging_days(d, as_at_date=as_at_date) for d in un_a_out[date_col_a].to_list()]
        un_a_out = un_a_out.with_columns([
            pl.Series("Aging_Days", [r[0] for r in aging_res_a]),
            pl.Series("Aging_Bucket", [r[1] for r in aging_res_a])
        ])

    # Calculate Aging for Source B Exceptions using the 'as_at_date'
    date_col_b = None
    for c in df_b.columns:
        if any(k in str(c).lower() for k in ["date", "dt", "posting", "value"]):
            date_col_b = c
            break

    if len(un_b_out) > 0 and date_col_b and date_col_b in un_b_out.columns:
        aging_res_b = [calculate_aging_days(d, as_at_date=as_at_date) for d in un_b_out[date_col_b].to_list()]
        un_b_out = un_b_out.with_columns([
            pl.Series("Aging_Days", [r[0] for r in aging_res_b]),
            pl.Series("Aging_Bucket", [r[1] for r in aging_res_b])
        ])

    return matched_out, fuzzy_out, un_a_out, un_b_out
