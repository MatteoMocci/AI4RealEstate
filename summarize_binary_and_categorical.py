# save as summarize_binary_and_categorical.py
import pandas as pd
import numpy as np
import re
from pathlib import Path
from typing import Any, List, Tuple

INPUT_XLSX = "Cagliari-Def.updated.xlsx"
SHEET = 0  # change to sheet name or index if needed

BINARY_CSV = "binary_summary.csv"
CATEGORICAL_CSV = "categorical_summary.csv"

CATEGORICAL_MIN_UNIQUE = 3
CATEGORICAL_MAX_UNIQUE = 30
CATEGORICAL_MAX_UNIQUE_FRAC = 0.05
MULTISELECT_DELIMS = [",", ";", "|"]

def load_df() -> pd.DataFrame:
    try:
        df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET, engine="openpyxl")
    except ImportError as e:
        raise SystemExit("openpyxl is required. Install with: pip install openpyxl") from e
    df.columns = [str(c).strip() for c in df.columns]
    return df

def normalize_binary_series(s: pd.Series) -> pd.Series:
    mapping = {
        "yes": 1, "y": 1, "true": 1, "t": 1, "si": 1, "sì": 1, "on": 1, "x": 1, "1": 1, 1: 1, True: 1,
        "no": 0,  "n": 0, "false": 0, "f": 0, "off": 0, "0": 0, 0: 0, False: 0
    }
    def _map(v: Any):
        if pd.isna(v):
            return np.nan
        if isinstance(v, str):
            key = v.strip().lower()
            return mapping.get(key, v.strip())
        return mapping.get(v, v)
    return s.map(_map)

def is_binary(col: pd.Series) -> Tuple[bool, pd.Series]:
    norm = normalize_binary_series(col)
    uniques = pd.Series(norm.dropna().unique())
    return len(uniques) == 2, norm

def looks_categorical(col: pd.Series, n_rows: int) -> bool:
    if col.dtype.name in ("object", "category", "bool"):
        pass
    else:
        non_null = col.dropna()
        if non_null.empty:
            return False
        nunique = non_null.nunique(dropna=True)
        if not (nunique <= CATEGORICAL_MAX_UNIQUE or nunique / len(non_null) <= CATEGORICAL_MAX_UNIQUE_FRAC):
            return False
    return True

def expand_multiselect(series: pd.Series, delims: List[str]) -> pd.Series:
    if not delims:
        return series
    delim_regex = "|".join(map(re.escape, delims))
    def to_list_or_scalar(v):
        if pd.isna(v):
            return np.nan
        if isinstance(v, str):
            if any(d in v for d in delims):
                parts = [p.strip() for p in re.split(delim_regex, v) if p.strip() != ""]
                return parts
            return v.strip()
        return v
    tmp = series.astype("object").map(to_list_or_scalar)
    return tmp.explode(ignore_index=False) if tmp.apply(lambda x: isinstance(x, list)).any() else tmp

def main():
    df = load_df()
    n_rows = len(df)

    bin_rows = []
    cat_rows = []

    for colname in df.columns:
        col = df[colname]
        missing = int(col.isna().sum())
        missing_rate = missing / n_rows if n_rows else 0.0
        missing_over_50pct = missing_rate > 0.5

        # Binary with arbitrary two values
        detected_binary, norm = is_binary(col)
        if detected_binary:
            # Determine the two values from the normalized series, ordered by frequency desc then value asc
            counts = pd.Series(norm.dropna()).value_counts()
            values = counts.index.tolist()[:2]
            # Guarantee stable order
            if len(values) == 2:
                v1, v2 = values[0], values[1]
            else:
                # extremely rare edge case where only one non missing value exists
                v1, v2 = (values[0] if values else ""), ""

            c1 = int((norm == v1).sum()) if v1 != "" else 0
            c2 = int((norm == v2).sum()) if v2 != "" else 0

            bin_rows.append({
                "column": colname,
                "first_value": "NaN" if pd.isna(v1) else str(v1),
                "second_value": "NaN" if pd.isna(v2) else str(v2),
                "first_count": c1,
                "second_count": c2,
                "missing": missing,
                "missing_rate": round(missing_rate, 4),
                "missing_over_50pct": missing_over_50pct
            })
            continue

        # Categorical with 3..30 distinct values after multiselect split
        if looks_categorical(col, n_rows):
            exploded = expand_multiselect(col, MULTISELECT_DELIMS)
            vc = exploded.value_counts(dropna=False)

            vc_no_nan = vc.drop(labels=[np.nan], errors="ignore")
            k = int(vc_no_nan.shape[0])

            if CATEGORICAL_MIN_UNIQUE <= k <= CATEGORICAL_MAX_UNIQUE:
                vc_no_nan = vc_no_nan.sort_values(ascending=False)
                values_list = [str(v) for v in vc_no_nan.index.tolist()]
                counts_list = [int(c) for c in vc_no_nan.tolist()]

                cat_rows.append({
                    "column": colname,
                    "n_values": k,
                    "values": " | ".join(values_list),
                    "counts": " | ".join(map(str, counts_list)),
                    "missing": missing,
                    "missing_rate": round(missing_rate, 4),
                    "missing_over_50pct": missing_over_50pct
                })

    # Write outputs
    pd.DataFrame(
        bin_rows,
        columns=[
            "column", "first_value", "second_value",
            "first_count", "second_count",
            "missing", "missing_rate", "missing_over_50pct"
        ]
    ).sort_values("column").to_csv(BINARY_CSV, index=False)

    pd.DataFrame(
        cat_rows,
        columns=[
            "column", "n_values", "values", "counts",
            "missing", "missing_rate", "missing_over_50pct"
        ]
    ).sort_values("column").to_csv(CATEGORICAL_CSV, index=False)

    print(f"Wrote {BINARY_CSV} with {len(bin_rows)} rows.")
    print(f"Wrote {CATEGORICAL_CSV} with {len(cat_rows)} rows.")

if __name__ == "__main__":
    main()
