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

# Columns C and D (0-based indexing)
col_c = pd.to_numeric(df.iloc[:, 2], errors="coerce")
col_d = pd.to_numeric(df.iloc[:, 3], errors="coerce")

# Drop rows with NaNs in either column
valid = col_c.notna() & col_d.notna()
col_c = col_c[valid]
col_d = col_d[valid]

if len(col_c) < 5:
    raise ValueError("Not enough rows for a 5-point running average.")

# 5-point running (sliding) average
c_avg_5 = col_c.rolling(window=5).mean().dropna().reset_index(drop=True)
d_avg_5 = col_d.rolling(window=5).mean().dropna().reset_index(drop=True)

# Output DataFrame
out_df = pd.DataFrame({
    "C_avg_5": c_avg_5,
    "D_avg_5": d_avg_5
})

# Output filename
base = os.path.splitext(os.path.basename(input_file))[0]
output_file = os.path.join(
    os.path.dirname(input_file),
    f"{base}_running_avg5.xlsx"
)

# Write Excel
out_df.to_excel(output_file, index=False)

print(f"Saved running-average file to:\n{output_file}")
