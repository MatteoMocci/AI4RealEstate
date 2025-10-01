# save as describe_tipologia_and_anno.py
import pandas as pd
from pathlib import Path

INPUT_XLSX = "Cagliari-Def.updated.xlsx"
SHEET = 0  # change if needed

SAVE_CSV = True # set True to also save CSVs next to the script

def value_counts_with_missing(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """
    Return a table with: value, count, pct_over_total
    Adds an explicit row for missing with value='<MISSING>'
    Percentages are over total rows (including missing).
    """
    if col not in df.columns:
        raise SystemExit(f'Column "{col}" not found. Available columns: {", ".join(df.columns)}')

    total = len(df)

    # Trim whitespace but keep original casing
    s = df[col].astype("object").map(lambda x: x.strip() if isinstance(x, str) else x)

    # Counts including NaN
    vc = s.value_counts(dropna=False)

    # Build result rows
    rows = []
    for val, cnt in vc.items():
        label = "<MISSING>" if pd.isna(val) else str(val)
        rows.append({"value": label, "count": int(cnt), "pct_over_total": round(cnt / total * 100, 2)})
    out = pd.DataFrame(rows)

    # Sort with non missing first by count desc, then the <MISSING> row at the end
    non_missing = out[out["value"] != "<MISSING>"].sort_values(["count", "value"], ascending=[False, True])
    missing = out[out["value"] == "<MISSING>"]
    out = pd.concat([non_missing, missing], ignore_index=True)

    return out

def main():
    path = Path(INPUT_XLSX)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path.resolve()}")

    df = pd.read_excel(path, sheet_name=SHEET, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # Tipologia
    tip_table = value_counts_with_missing(df, "Tipologia")
    print("\n=== Tipologia ===")
    print(tip_table.to_string(index=False))

    # Anno di costruzione
    anno_table = value_counts_with_missing(df, "Anno di costruzione")
    print("\n=== Anno di costruzione ===")
    print(anno_table.to_string(index=False))

    if SAVE_CSV:
        tip_table.to_csv("Tipologia_summary.csv", index=False)
        anno_table.to_csv("Anno_di_costruzione_summary.csv", index=False)
        print('\nSaved CSVs: Tipologia_summary.csv, Anno_di_costruzione_summary.csv')

if __name__ == "__main__":
    main()
