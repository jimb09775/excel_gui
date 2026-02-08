import pandas as pd
import numpy as np
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

# Columns C and D (0-based index: C=2, D=3)
col_c = pd.to_numeric(df.iloc[:, 2], errors="coerce")
col_d = pd.to_numeric(df.iloc[:, 3], errors="coerce")

# Drop rows with NaNs in either column
valid = col_c.notna() & col_d.notna()
col_c = col_c[valid].to_numpy()
col_d = col_d[valid].to_numpy()

# Number of complete 3-row blocks
n_blocks = len(col_c) // 3

if n_blocks == 0:
    raise ValueError("Not enough rows to compute a 3-point average.")

# Reshape and average
c_avg = col_c[:n_blocks * 3].reshape(n_blocks, 3).mean(axis=1)
d_avg = col_d[:n_blocks * 3].reshape(n_blocks, 3).mean(axis=1)

# Create output DataFrame
out_df = pd.DataFrame({
    "C_avg_3": c_avg,
    "D_avg_3": d_avg
})

# Output filename
base = os.path.splitext(os.path.basename(input_file))[0]
output_file = os.path.join(
    os.path.dirname(input_file),
    f"{base}_avg3.xlsx"
)

# Write Excel
out_df.to_excel(output_file, index=False)

print(f"Saved averaged file to:\n{output_file}")
