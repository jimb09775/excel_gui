import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog
import os

# ---------------- CONFIG ----------------
x_column = "linear_encoder_position"
y_column = "measurement_delta"
sheet_name = 0
# ----------------------------------------

# Hide Tk root window
Tk().withdraw()

# -------- File 1 --------
file1 = filedialog.askopenfilename(
    title="Select FIRST Excel file (reference X)",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not file1:
    raise SystemExit("No first file selected")

# -------- File 2 --------
file2 = filedialog.askopenfilename(
    title="Select SECOND Excel file (same X positions)",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not file2:
    raise SystemExit("No second file selected")

name1 = os.path.basename(file1)
name2 = os.path.basename(file2)

# -------- Read files --------
df1 = pd.read_excel(file1, sheet_name=sheet_name)
df2 = pd.read_excel(file2, sheet_name=sheet_name)

df1.columns = df1.columns.astype(str).str.strip()
df2.columns = df2.columns.astype(str).str.strip()

for df, label in [(df1, "File 1"), (df2, "File 2")]:
    missing = [c for c in (x_column, y_column) if c not in df.columns]
    if missing:
        raise ValueError(f"{label} missing columns {missing}")

# -------- Extract numeric data --------
x1 = pd.to_numeric(df1[x_column], errors="coerce")
y1 = pd.to_numeric(df1[y_column], errors="coerce")

x2 = pd.to_numeric(df2[x_column], errors="coerce")
y2 = pd.to_numeric(df2[y_column], errors="coerce")

mask1 = x1.notna() & y1.notna()
mask2 = x2.notna() & y2.notna()

x1, y1 = x1[mask1].to_numpy(), y1[mask1].to_numpy()
x2, y2 = x2[mask2].to_numpy(), y2[mask2].to_numpy()

# -------- Sanity check: same X --------
if len(x1) != len(x2) or not np.allclose(x1, x2, rtol=0, atol=1e-9):
    raise ValueError(
        "X values do not match between files.\n"
        "This script assumes identical linear_encoder_position arrays."
    )

x = x1  # shared X

# -------- Fits --------
m1, b1 = np.polyfit(x, y1, 1)
m2, b2 = np.polyfit(x, y2, 1)

y1_fit = m1 * x + b1
y2_fit = m2 * x + b2

res1 = y1 - y1_fit
res2 = y2 - y2_fit

stats1 = {
    "std": res1.std(ddof=1),
    "rms": np.sqrt(np.mean(res1**2)),
    "max": np.abs(res1).max(),
}

stats2 = {
    "std": res2.std(ddof=1),
    "rms": np.sqrt(np.mean(res2**2)),
    "max": np.abs(res2).max(),
}

# -------- Sort for plotting --------
order = np.argsort(x)
xs = x[order]

# -------- Plot --------
plt.figure()
plt.plot(xs, y1[order], label=f"Data 1: {name1}")
plt.plot(xs, y1_fit[order], linestyle="--", label="Fit 1")

plt.plot(xs, y2[order], label=f"Data 2: {name2}")
plt.plot(xs, y2_fit[order], linestyle="--", label="Fit 2")

plt.xlabel("Linear Encoder Position")
plt.ylabel("Measurement Delta")
plt.title("Measurement Delta Comparison")
plt.grid(True)
plt.legend()

# -------- Stats boxes --------
box1 = (
    f"{name1}\n"
    f"m = {m1:.4g}\n"
    f"σ = {stats1['std']:.4g}\n"
    f"RMS = {stats1['rms']:.4g}\n"
    f"Max |dev| = {stats1['max']:.4g}"
)

box2 = (
    f"{name2}\n"
    f"m = {m2:.4g}\n"
    f"σ = {stats2['std']:.4g}\n"
    f"RMS = {stats2['rms']:.4g}\n"
    f"Max |dev| = {stats2['max']:.4g}"
)

ax = plt.gca()

ax.text(
    0.98, 0.02, box1,
    transform=ax.transAxes,
    ha="right", va="bottom",
    fontsize=8,
    bbox=dict(boxstyle="round", facecolor="white", alpha=0.85)
)

ax.text(
    0.98, 0.32, box2,
    transform=ax.transAxes,
    ha="right", va="bottom",
    fontsize=8,
    bbox=dict(boxstyle="round", facecolor="white", alpha=0.85)
)

plt.tight_layout()
plt.show()
