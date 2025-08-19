import pandas as pd
import re
import math

# --- File paths ---
first_file = "extract.xlsx"     # DESCRIPTION | INVENTORY
second_file = "template.xlsx"   # DESCRIPTION | ALLOCATION | INVENTORY | ORDERING
output_file = "reconciliation.xlsx"

# --- Load excels ---
df_first = pd.read_excel(first_file)
df_second = pd.read_excel(second_file)

# --- Normalize descriptions ---
def normalize_desc(s: str) -> str:
    if pd.isna(s): return ""
    s = str(s).upper().strip()
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

df_first["DESC_KEY"] = df_first["DESCRIPTION"].map(normalize_desc)
df_second["DESC_KEY"] = df_second["DESCRIPTION"].map(normalize_desc)

# --- Merge inventory from first into second ---
df_merged = df_second.merge(
    df_first[["DESC_KEY", "INVENTORY"]].rename(columns={"INVENTORY": "INVENTORY_FIRST"}),
    on="DESC_KEY",
    how="left"
)

# Replace INVENTORY with first file’s if available
df_merged["INVENTORY"] = df_merged["INVENTORY_FIRST"].fillna(df_merged["INVENTORY"])

# --- Allocation rule (box = 100 pcs, syrup = per piece) ---
def compute_order(desc, alloc, inv):
    if pd.isna(alloc) or pd.isna(inv):
        return 0

    desc_upper = str(desc).upper()

    if "SYRUP" in desc_upper or "ML" in desc_upper:  
        pieces_per_alloc = 1   # syrup/ml by piece
    else:
        pieces_per_alloc = 100 # default: box = 100 pcs

    needed_pcs = alloc * pieces_per_alloc
    gap = needed_pcs - inv

    if gap <= 0:
        return 0

    return math.ceil(gap / pieces_per_alloc)

df_merged["ORDERING"] = df_merged.apply(
    lambda x: compute_order(x["DESCRIPTION"], x["ALLOCATION"], x["INVENTORY"]), axis=1
)

# --- Keep only the same format as second file ---
df_result = df_merged[["DESCRIPTION", "ALLOCATION", "INVENTORY", "ORDERING"]]

# --- Save final result ---
df_result.to_excel(output_file, index=False)

print(f"✅ Done! Output saved as {output_file}")
