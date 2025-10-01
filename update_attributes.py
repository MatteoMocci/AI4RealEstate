# save as transform_disponibilita_and_posti_auto.py
import pandas as pd
import re
from pathlib import Path
from typing import Optional

INPUT_XLSX = "Cagliari-Def.updated.xlsx"
SHEET = 0
OUTPUT_XLSX = "Cagliari-Def.updated.xlsx"
OUTPUT_CSV  = "Cagliari-Def.updated.csv"

def to_binary_disponibilita(v):
    """Only exact 'Libero' stays Libero; everything else (including NaN) becomes Non Libero."""
    if pd.isna(v):
        return "Non Libero"
    s = str(v).strip()
    return "Libero" if s.lower() == "libero" else "Non Libero"

def normalize_parking_value(v) -> str:
    """
    Convert Posti Auto to '3+' when value >= 3.
    Tries to parse numbers from strings like '3', '3 posti', '2-3', '>=3', '3 o più'.
    If no number is found, returns the original trimmed string.
    NaN stays NaN.
    """
    if pd.isna(v):
        return v
    s = str(v).strip().lower()

    # quick checks for textual 3+ hints
    if any(tok in s for tok in ["3+", "≥3", ">=3", "=>3", "3 o piu", "3 o più", "almeno 3", "min 3"]):
        return "3+"

    # extract all integers present (handles '2-3', '2 / 3', '3 posti', etc.)
    nums = [int(x) for x in re.findall(r"\d+", s)]
    if nums:
        # choose the max mentioned number, e.g., '2-3' -> 3
        m = max(nums)
        if m >= 3:
            return "3+"
        else:
            return str(m)

    # fallback: if the cell is a plain numeric type (e.g., 2.0)
    try:
        f = float(s.replace(",", "."))
        if f >= 3:
            return "3+"
        # keep integer-like formatting for small numbers
        return str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        # nothing numeric detected, return original as-is
        return s
    
import re
import numpy as np

# --- Piano cleanup ---

# strategy for ranges: "lower" | "upper" | "mid"
PIANO_RANGE_STRATEGY = "lower"

import re
import numpy as np
import pandas as pd

def parse_piano_to_int(v) -> object:
    """
    Normalize Piano to an integer.
    S -> -1, T -> 0, R -> 0
    '3' -> 3
    '3-4' or '3 to 4' -> 3  (lower bound)
    Any numeric like '4.0' -> 4  (rounded)
    Unparseable -> NaN
    """
    if pd.isna(v):
        return np.nan
    s = str(v).strip().upper()

    # letter codes
    if s == "S":
        return -1
    if s in {"T", "R"}:
        return 0

    # plain integer
    if s.isdigit():
        return int(s)

    # range like 3-4 or 3 to 4
    m = re.match(r"^\s*(\d+)\s*(?:-|TO)\s*(\d+)\s*$", s)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        return min(a, b)  # lower bound

    # numeric with decimals or commas
    try:
        f = float(s.replace(",", "."))
        return int(round(f))
    except ValueError:
        return np.nan







def main():
    in_path = Path(INPUT_XLSX)
    if not in_path.exists():
        raise FileNotFoundError(f"Input file not found: {in_path.resolve()}")

    df = pd.read_excel(in_path, sheet_name=SHEET, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # 1) Disponibilità -> Libero / Non Libero
    DISP_COL = "Disponibilità"
    if DISP_COL in df.columns:
        df[DISP_COL] = df[DISP_COL].map(to_binary_disponibilita)
    else:
        print(f'Warning: column "{DISP_COL}" not found. Skipping its transformation.')

    # 2) Posti Auto -> '3+' when >=3
    POSTI_COL = "Posti Auto"
    if POSTI_COL in df.columns:
        df[POSTI_COL] = df[POSTI_COL].map(normalize_parking_value)
    else:
        print(f'Warning: column "{POSTI_COL}" not found. Skipping its transformation.')

    # Apply to your dataframe df
    if "Piano" in df.columns:
        df["Piano"] = df["Piano"].map(parse_piano_to_int).astype("Int64")
    # write outputs
    df.to_excel(OUTPUT_XLSX, index=False)
    df.to_csv(OUTPUT_CSV, index=False)
    print(f"Saved:\n - {OUTPUT_XLSX}\n - {OUTPUT_CSV}")

if __name__ == "__main__":
    main()
