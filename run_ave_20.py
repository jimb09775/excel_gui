import pandas as pd
from tkinter import Tk, filedialog
import os

# Hide Tk window
Tk().withdraw()

# Select Excel file
input_file = filedialog.askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not input_file:
    raise SystemExit("No file selected")

# Read Excel
df = pd.read_excel(input_file)

# Columns C and D (0-based indexing: C=2, D=3)
col_c = pd.to_numeric(df.iloc[:, 2], errors="coerce")
col_d = pd.to_numeric(df.iloc[:, 3], errors="coerce")

# Drop rows with NaNs in either column
valid = col_c.notna() & col_d.notna()
col_c = col_c[valid]
col_d = col_d[valid]

window = 20

if len(col_c) < window:
    raise ValueError(f"Not enough rows for a {window}-point running average.")

# 20-point running (sliding) average
c_avg_20 = col_c.rolling(window=window).mean().dropna().reset_index(drop=True)
d_avg_20 = col_d.rolling(window=window).mean().dropna().reset_index(drop=True)

# Output DataFrame
out_df = pd.DataFrame({
    "C_avg_20": c_avg_20,
    "D_avg_20": d_avg_20
})

# Output filename
base = os.path.splitext(os.path.basename(input_file))[0]
output_file = os.path.join(
    os.path.dirname(input_file),
    f"{base}_running_avg20.xlsx"
)

# Write Excel
out_df.to_excel(output_file, index=False)

print(f"Saved 20-point running-average file to:\n{output_file}")
