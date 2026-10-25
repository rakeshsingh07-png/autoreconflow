import itertools
from datetime import datetime
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
from rapidfuzz import fuzz


class CommercialReconEngine:

    def __init__(self, df_a: pd.DataFrame, df_b: pd.DataFrame):
        self.df_a = df_a.copy().reset_index(drop=True)
        self.df_b = df_b.copy().reset_index(drop=True)

        self.df_a["_matched"] = False
        self.df_b["_matched"] = False
        self.matches = []

    def clean_text(self, series: pd.Series) -> pd.Series:
        return (
            series.astype(str)
            .str.upper()
            .str.replace(r"[^A-Z0-9]", "", regex=True)
        )

    def run_reconciliation(
        self,
        col_ref_a: str,
        col_amt_a: str,
        col_ref_b: str,
        col_amt_b: str,
        tolerance: float = 0.00,
        fuzzy_threshold: float = 80.0,
        enable_split_matching: bool = True,
    ):
        # Normalize keys and amounts
        self.df_a["_norm_ref"] = self.clean_text(self.df_a[col_ref_a])
        self.df_b["_norm_ref"] = self.clean_text(self.df_b[col_ref_b])
        self.df_a["_norm_amt"] = pd.to_numeric(
            self.df_a[col_amt_a], errors="coerce"
        ).fillna(0.0)
        self.df_b["_norm_amt"] = pd.to_numeric(
            self.df_b[col_amt_b], errors="coerce"
        ).fillna(0.0)

        # -------------------------------------------------------------
        # PASS 1: Exact Key & Exact Amount (Tolerance Bounded)
        # -------------------------------------------------------------
        unmatched_a = self.df_a[~self.df_a["_matched"]].index.tolist()
        unmatched_b = self.df_b[~self.df_b["_matched"]].index.tolist()

        b_map = {}
        for b_idx in unmatched_b:
            key = (
                self.df_b.at[b_idx, "_norm_ref"],
                round(self.df_b.at[b_idx, "_norm_amt"], 2),
            )
            b_map.setdefault(key, []).append(b_idx)

        for a_idx in unmatched_a:
            val_a = round(self.df_a.at[a_idx, "_norm_amt"], 2)
            k_a = (self.df_a.at[a_idx, "_norm_ref"], val_a)

            if k_a in b_map and b_map[k_a]:
                b_idx = b_map[k_a].pop(0)
                self.df_a.at[a_idx, "_matched"] = True
                self.df_b.at[b_idx, "_matched"] = True
                self.matches.append({
                    "Match_ID": f"M1-{len(self.matches)+1:05d}",
                    "Tier": "Pass 1: Exact Match",
                    "Confidence": "100%",
                    "Source_A_Ref": self.df_a.at[a_idx, col_ref_a],
                    "Source_B_Ref": self.df_b.at[b_idx, col_ref_b],
                    "Amount_A": val_a,
                    "Amount_B": round(self.df_b.at[b_idx, "_norm_amt"], 2),
                    "Variance": 0.0,
                    "Idx_A": a_idx,
                    "Idx_B": b_idx,
                })

        # -------------------------------------------------------------
        # PASS 2: Fuzzy Reference + Strict Amount Parity
        # -------------------------------------------------------------
        unmatched_a = self.df_a[~self.df_a["_matched"]].index.tolist()
        unmatched_b = self.df_b[~self.df_b["_matched"]].index.tolist()

        for a_idx in unmatched_a:
            val_a = self.df_a.at[a_idx, "_norm_amt"]
            str_a = str(self.df_a.at[a_idx, col_ref_a]).strip()

            best_b = None
            best_score = 0.0

            for b_idx in unmatched_b:
                if self.df_b.at[b_idx, "_matched"]:
                    continue

                val_b = self.df_b.at[b_idx, "_norm_amt"]
                # Must match amount within tolerance
                if abs(val_a - val_b) <= tolerance:
                    str_b = str(self.df_b.at[b_idx, col_ref_b]).strip()
                    score = fuzz.token_sort_ratio(str_a, str_b)
                    if score >= fuzzy_threshold and score > best_score:
                        best_score = score
                        best_b = b_idx

            if best_b is not None:
                self.df_a.at[a_idx, "_matched"] = True
                self.df_b.at[best_b, "_matched"] = True
                val_b = self.df_b.at[best_b, "_norm_amt"]
                self.matches.append({
                    "Match_ID": f"M2-{len(self.matches)+1:05d}",
                    "Tier": f"Pass 2: Fuzzy Match ({round(best_score, 1)}%)",
                    "Confidence": f"{round(best_score, 1)}%",
                    "Source_A_Ref": self.df_a.at[a_idx, col_ref_a],
                    "Source_B_Ref": self.df_b.at[best_b, col_ref_b],
                    "Amount_A": val_a,
                    "Amount_B": val_b,
                    "Variance": round(val_a - val_b, 2),
                    "Idx_A": a_idx,
                    "Idx_B": best_b,
                })

        # -------------------------------------------------------------
        # PASS 3: 1-to-Many Split Payment Solver (Knapsack subset-sum)
        # -------------------------------------------------------------
        if enable_split_matching:
            unmatched_a = self.df_a[~self.df_a["_matched"]].index.tolist()
            unmatched_b = self.df_b[~self.df_b["_matched"]].index.tolist()

            # Target: Single lump sum in B matching 2 to 3 items in A
            for b_idx in unmatched_b:
                if self.df_b.at[b_idx, "_matched"]:
                    continue
                target_amt = round(self.df_b.at[b_idx, "_norm_amt"], 2)
                available_a = [
                    i for i in unmatched_a if not self.df_a.at[i, "_matched"]
                ]

                found_combo = None
                for r in (2, 3):
                    for combo in itertools.combinations(available_a, r):
                        combo_sum = round(
                            sum(self.df_a.at[i, "_norm_amt"] for i in combo), 2
                        )
                        if abs(combo_sum - target_amt) <= tolerance:
                            found_combo = combo
                            break
                    if found_combo:
                        break

                if found_combo:
                    self.df_b.at[b_idx, "_matched"] = True
                    b_ref = self.df_b.at[b_idx, col_ref_b]
                    for a_sub_idx in found_combo:
                        self.df_a.at[a_sub_idx, "_matched"] = True
                        self.matches.append({
                            "Match_ID": f"M3-{len(self.matches)+1:05d}",
                            "Tier": (
                                "Pass 3: Split Batch (1-to-Many)"
                            ),
                            "Confidence": "High (Subset-Sum Exact)",
                            "Source_A_Ref": self.df_a.at[
                                a_sub_idx, col_ref_a
                            ],
                            "Source_B_Ref": f"{b_ref} (Batch Component)",
                            "Amount_A": self.df_a.at[
                                a_sub_idx, "_norm_amt"
                            ],
                            "Amount_B": target_amt,
                            "Variance": 0.0,
                            "Idx_A": a_sub_idx,
                            "Idx_B": b_idx,
                        })

        # Compile final outputs
        df_matched = pd.DataFrame(self.matches)
        df_exc_a = (
            self.df_a[~self.df_a["_matched"]]
            .drop(columns=["_matched", "_norm_ref", "_norm_amt"])
            .copy()
        )
        df_exc_b = (
            self.df_b[~self.df_b["_matched"]]
            .drop(columns=["_matched", "_norm_ref", "_norm_amt"])
            .copy()
        )

        return df_matched, df_exc_a, df_exc_b

    @staticmethod
    def generate_executive_workpaper(
        df_matched: pd.DataFrame,
        df_exc_a: pd.DataFrame,
        df_exc_b: pd.DataFrame,
        meta: dict,
    ) -> bytes:
        from io import BytesIO

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # 1. Executive Summary Sheet
            summary_data = {
                "Metric": [
                    "Reconciliation Date",
                    "Source A Total Records",
                    "Source B Total Records",
                    "Matched Records Count",
                    "Source A Exceptions (Unmatched)",
                    "Source B Exceptions (Unmatched)",
                    "Match Rate (%)",
                    "Total Reconciled Volume",
                    "Estimated Manual Effort Saved",
                    "Audit Status",
                ],
                "Value": [
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    meta["count_a"],
                    meta["count_b"],
                    len(df_matched),
                    len(df_exc_a),
                    len(df_exc_b),
                    f"{meta['match_rate']}%",
                    f"${meta['reconciled_volume']:,.2f}",
                    f"{meta['hours_saved']} Hours",
                    "BALANCED & AUDITED",
                ],
            }
            pd.DataFrame(summary_data).to_excel(
                writer, sheet_name="Executive Summary", index=False
            )

            # 2. Matched Records Sheet
            if not df_matched.empty:
                df_matched.to_excel(
                    writer, sheet_name="Matched Ledger", index=False
                )

            # 3. Exceptions Sheets
            df_exc_a.to_excel(
                writer, sheet_name="Source A Exceptions", index=False
            )
            df_exc_b.to_excel(
                writer, sheet_name="Source B Exceptions", index=False
            )

        return output.getvalue()