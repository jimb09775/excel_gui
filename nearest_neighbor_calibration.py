import pandas as pd
from tkinter import Tk, filedialog
import os

# ---------------- CONFIG ----------------
# Column positions (0-based): A=0, C=2, D=3
TARGET_VALUE_COL = 0          # Target: column A (value to correct)
INDEX_COL = 2                 # Both files: column C (index)
CORRECTION_VALUE_COL = 3      # Correction file: column D (correction value)

# Apply correction: corrected = value - correction
TOLERANCE = 0.5               # max allowed |target_idx - corr_idx|
# ----------------------------------------

def pick_file(title: str) -> str:
    Tk().withdraw()
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not path:
        raise SystemExit("No file selected")
    return path

def coerce_numeric(series, name: str) -> pd.Series:
    out = pd.to_numeric(series, errors="coerce")
    if out.isna().all():
        raise ValueError(f"{name} has no numeric values after parsing.")
    return out

def main():
    correction_file = pick_file("Select CORRECTION Excel file (C=index, D=correction)")
    target_file = pick_file("Select TARGET Excel file (C=index, A=value-to-correct)")

    # Read
    corr_df_raw = pd.read_excel(correction_file)
    tgt_df = pd.read_excel(target_file)

    # Build correction table
    corr = pd.DataFrame({
        "idx": coerce_numeric(corr_df_raw.iloc[:, INDEX_COL], "Correction C (index)"),
        "corr_val": coerce_numeric(corr_df_raw.iloc[:, CORRECTION_VALUE_COL], "Correction D (correction)"),
    }).dropna()

    if len(corr) < 1:
        raise ValueError("Correction file contains no valid (index, correction) pairs.")

    # Extract target index and value
    tgt_idx = coerce_numeric(tgt_df.iloc[:, INDEX_COL], "Target C (index)")
    tgt_val = coerce_numeric(tgt_df.iloc[:, TARGET_VALUE_COL], "Target A (value)")

    # Prepare frames for merge_asof (must be sorted by key)
    corr = corr.sort_values("idx").reset_index(drop=True)

    tgt_work = pd.DataFrame({
        "_row": range(len(tgt_df)),   # preserve original order
        "_idx": tgt_idx,
        "_val": tgt_val
    }).dropna(subset=["_idx", "_val"])

    # Sort for asof merge
    tgt_sorted = tgt_work.sort_values("_idx").reset_index(drop=True)

    # Nearest match within tolerance
    merged = pd.merge_asof(
        tgt_sorted,
        corr,
        left_on="_idx",
        right_on="idx",
        direction="nearest",
        tolerance=TOLERANCE
    )

    # Apply correction: subtract
    merged["corrected"] = merged["_val"] - merged["corr_val"]

    # Put results back in original target row order
    merged = merged.sort_values("_row")

    # Create output columns (aligned to original target rows)
    # Start with all-NaN columns, then fill only rows that had valid idx/val in tgt_work
    tgt_df["matched_correction_index"] = pd.NA
    tgt_df["applied_correction_value"] = pd.NA
    tgt_df["corrected_value"] = pd.NA

    # Map merged results back into the correct row indices
    rows = merged["_row"].to_numpy()
    tgt_df.loc[rows, "matched_correction_index"] = merged["idx"].to_numpy()
    tgt_df.loc[rows, "applied_correction_value"] = merged["corr_val"].to_numpy()
    tgt_df.loc[rows, "corrected_value"] = merged["corrected"].to_numpy()

    # Diagnostics: count unmatched within tolerance among rows we attempted to correct
    attempted = len(tgt_work)
    unmatched = merged["corr_val"].isna().sum()
    if unmatched:
        print(f"WARNING: {unmatched} of {attempted} target rows had no correction match within tolerance={TOLERANCE}.")

    # Save next to target file
    out_dir = os.path.dirname(target_file)
    base = os.path.splitext(os.path.basename(target_file))[0]
    out_path = os.path.join(out_dir, f"{base}_corrected.xlsx")

    tgt_df.to_excel(out_path, index=False)
    print(f"Saved corrected file to:\n{out_path}")

if __name__ == "__main__":
    main()
