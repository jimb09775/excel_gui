import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog

# Hide Tk root window
Tk().withdraw()

# File picker
excel_file = filedialog.askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not excel_file:
    raise SystemExit("No file selected")

# Fixed column names
x_column = "linear_encoder_position"
y_column = "measurement_delta"
sheet_name = 0

# Read Excel
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Validate columns
missing = [c for c in (x_column, y_column) if c not in df.columns]
if missing:
    raise ValueError(
        f"Missing columns: {missing}\n"
        f"Available columns: {list(df.columns)}"
    )

# Plot
plt.figure()
plt.plot(df[x_column], df[y_column])
plt.xlabel("Linear Encoder Position")
plt.ylabel("Measurement Delta")
plt.title("Measurement Delta vs Linear Encoder Position")
plt.grid(True)
plt.tight_layout()
plt.show()

