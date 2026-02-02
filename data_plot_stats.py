import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog
import os

# Hide Tk root window
Tk().withdraw()

# File picker
excel_file = filedialog.askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not excel_file:
    raise SystemExit("No file selected")

# Extract filename (no path)
filename = os.path.basename(excel_file)

# Fixed column names
x_column = "linear_encoder_position"
y_column = "measurement_delta"
sheet_name = 0

# Read Excel
df = pd.read_excel(excel_file, sheet_name=sheet_name)
df.columns = df.columns.astype(str).str.strip()

# Validate columns
missing = [c for c in (x_column, y_column) if c not in df.columns]
if missing:
    raise ValueError(
        f"Missing columns: {missing}\n"
        f"Available columns: {list(df.columns)}"
    )

# Numeric coercion + cleanup
x = pd.to_numeric(df[x_column], errors="coerce")
y = pd.to_numeric(df[y_column], errors="coerce")
mask = x.notna() & y.notna()
x = x[mask].to_numpy()
y = y[mask].to_numpy()

if len(x) < 2:
    raise ValueError("Not enough valid data points to fit a line.")

# Best-fit line
m, b = np.polyfit(x, y, 1)
y_fit = m * x + b

# Residuals and statistics
residuals = y - y_fit
resid_std = residuals.std(ddof=1)
max_abs_dev = np.abs(residuals).max()
rms = np.sqrt(np.mean(residuals**2))

# Sort for plotting
order = np.argsort(x)
xs = x[order]
ys = y[order]
yfs = m * xs + b

# Plot
plt.figure()
plt.plot(xs, ys, label="Data")
plt.plot(xs, yfs, label="Best-fit line")

plt.xlabel("Linear Encoder Position")
plt.ylabel("Measurement Delta")
plt.title("Measurement Delta vs Linear Encoder Position")
plt.suptitle(filename, fontsize=9, y=0.94)

plt.grid(True)
plt.legend()

# ---- Stats box (lower right) ----
stats_text = (
    f"Fit: y = m·x + b\n"
    f"m = {m:.4g}\n"
    f"b = {b:.4g}\n"
    f"\n"
    f"Residual σ = {resid_std:.4g}\n"
    f"RMS error  = {rms:.4g}\n"
    f"Max |dev| = {max_abs_dev:.4g}\n"
    f"N = {len(x)}"
)

plt.text(
    0.98, 0.02, stats_text,
    transform=plt.gca().transAxes,
    ha="right",
    va="bottom",
    fontsize=9,
    bbox=dict(boxstyle="round", facecolor="white", alpha=0.85)
)

plt.tight_layout()
plt.show()
