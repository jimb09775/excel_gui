import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog
import os

# ---------------- CONFIG ----------------
x_column = "linear_encoder_position"
y_column = "measurement_delta"
sheet_name = 0

output_dir = "/Users/jamesmckee/Dropbox/SierraWave Systems/measured data"
# ----------------------------------------

# Hide Tk root window
Tk().withdraw()

# -------- Select files --------
file1 = filedialog.askopenfilename(
    title="Select FIRST Excel file (reference X)",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not file1:
    raise SystemExit("No first file selected")

file2 = filedialog.askopenfilename(
    title="Select SECOND Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not file2:
    raise SystemExit("No second file selected")

name1 = os.path.basename(file1)
name2 = os.path.basename(file2)

# -------- Read Excel --------
df1 = pd.read_excel(file1, sheet_name=sheet_name)
df2 = pd.read_excel(file2, sheet_name=sheet_name)

df1.columns = df1.columns.astype(str).str.strip()
df2.columns = df2.columns.astype(str).str.strip()

# -------- Validate columns --------
for df, label in [(df1, "File 1"), (df2, "File 2")]:
    missing = [c for c in (x_column, y_column) if c not in df.columns]
    if missing:
        raise ValueError(f"{label} missing columns {missing}")

# -------- Extract numeric data --------
x1 = pd.to_numeric(df1[x_column], errors="coerce")
y1 = pd.to_numeric(df1[y_column], errors="coerce")
mask1 = x1.notna() & y1.notna()
x1, y1 = x1[mask1].to_numpy(), y1[mask1].to_numpy()

x2 = pd.to_numeric(df2[x_column], errors="coerce")
y2 = pd.to_numeric(df2[y_column], errors="coerce")
mask2 = x2.notna() & y2.notna()
x2, y2 = x2[mask2].to_numpy(), y2[mask2].to_numpy()

# -------- Sort by X --------
o1 = np.argsort(x1)
x1, y1 = x1[o1], y1[o1]

o2 = np.argsort(x2)
x2, y2 = x2[o2], y2[o2]

# -------- Align File 2 to File 1 X (overlap only) --------
in_range = (x1 >= x2.min()) & (x1 <= x2.max())
x = x1[in_range]
y1_aligned = y1[in_range]
y2_aligned = np.interp(x, x2, y2)

if len(x) < 2:
    raise ValueError("Not enough overlapping X values to plot.")

# -------- Plot --------
plt.figure()
plt.plot(x, y1_aligned, label=name1)
plt.plot(x, y2_aligned, label=name2)

plt.xlabel("Linear Encoder Position")
plt.ylabel("Measurement Delta")
plt.title("Measurement Delta Comparison")
plt.grid(True)
plt.legend()

# Overlap annotation
plt.text(
    0.02, 0.02,
    f"Overlap X range: [{x.min():.4g}, {x.max():.4g}]",
    transform=plt.gca().transAxes,
    ha="left",
    va="bottom",
    fontsize=9,
    bbox=dict(boxstyle="round", facecolor="white", alpha=0.7)
)

plt.tight_layout()

# -------- Auto-increment PNG save --------
os.makedirs(output_dir, exist_ok=True)

base = os.path.splitext(name1)[0]
ext = ".png"
png_path = os.path.join(output_dir, base + ext)

if os.path.exists(png_path):
    i = 1
    while True:
        png_path = os.path.join(output_dir, f"{base}_{i:02d}{ext}")
        if not os.path.exists(png_path):
            break
        i += 1

plt.savefig(png_path, dpi=300)
print(f"Saved plot to: {png_path}")

plt.show()
