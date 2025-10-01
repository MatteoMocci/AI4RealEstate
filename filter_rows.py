# save as filter_rows.py
import pandas as pd
from pathlib import Path

INPUT_XLSX = "Cagliari-Def.updated.xlsx"
SHEET = 0
OUTPUT_XLSX = "Cagliari-Def.updated.xlsx"

# Rows to remove
TIPOLOGIA_REMOVE = {
    "casale",
    "villa plurifamiliare",
    "rustico",
    "casa indipendente",
    "villa",
}
ALIM_REMOVE = {
    "alimentato a legna",
    "alimentato a teleriscaldamento",
}

def norm(v):
    if pd.isna(v):
        return None
    return str(v).strip().lower()

def main():
    in_path = Path(INPUT_XLSX)
    if not in_path.exists():
        raise FileNotFoundError(f"Input file not found: {in_path.resolve()}")

    df = pd.read_excel(in_path, sheet_name=SHEET, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # Required columns
    col_tip = "Tipologia"
    col_alim = "Alimentazione riscaldamento"
    missing_cols = [c for c in [col_tip, col_alim] if c not in df.columns]
    if missing_cols:
        raise SystemExit(f"Missing column(s): {', '.join(missing_cols)}. Available: {', '.join(df.columns)}")

    tip_norm = df[col_tip].map(norm)
    alim_norm = df[col_alim].map(norm)

    mask_tip = tip_norm.isin(TIPOLOGIA_REMOVE)
    mask_alim = alim_norm.isin(ALIM_REMOVE)
    mask_drop = mask_tip | mask_alim

    # Stats
    total_rows = len(df)
    remove_tip = int(mask_tip.sum())
    remove_alim = int(mask_alim.sum())
    remove_both = int((mask_tip & mask_alim).sum())
    remove_total = int(mask_drop.sum())

    # Filter and save
    df_filtered = df.loc[~mask_drop].copy()
    df_filtered.to_excel(OUTPUT_XLSX, index=False)

    print(f"Rows in:  {total_rows}")
    print(f"Removed by Tipologia: {remove_tip}")
    print(f"Removed by Alimentazione riscaldamento: {remove_alim}")
    print(f"Overlap removed by both: {remove_both}")
    print(f"Total removed: {remove_total}")
    print(f"Rows out: {len(df_filtered)}")
    print(f"Saved {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
