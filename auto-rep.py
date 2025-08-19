import re
import pandas as pd
from rapidfuzz import process, fuzz
import math

# File paths
template_file = "template.xlsx"
extract_file = "extract.xlsx"
output_file = "order_result.xlsx"

# Load excels
df_template = pd.read_excel(template_file)  # Has Description, Allocation (boxes/pieces)
df_extract = pd.read_excel(extract_file)    # Has Description, Current_Stock (pcs)

print("Template Columns:", df_template.columns.tolist())
print("Extract Columns:", df_extract.columns.tolist())

# Normalize text for comparison
df_template['Description_norm'] = df_template['DESCRIPTION'].str.lower().str.strip()
df_extract['Description_norm'] = df_extract['DESCRIPTION'].str.lower().str.strip()

# Fuzzy match function
def best_match(desc, choices):
    match, score, idx = process.extractOne(desc, choices, scorer=fuzz.token_sort_ratio)
    return match, score

matches = []
for i, row in df_extract.iterrows():
    match, score = best_match(row['Description_norm'], df_template['Description_norm'].tolist())
    matches.append((row['DESCRIPTION'], match, score, row['INVENTORY']))

df_map = pd.DataFrame(matches, columns=['NEW DESCRIPTION', 'Template_Desc', 'Score', 'CURRENT STOCK'])
df_merge = df_map.merge(df_template, left_on="Template_Desc", right_on="Description_norm", how="left")

# Function: detect pcs per box
def get_pieces_per_box(desc):
    desc_lower = desc.lower()
    # Syrup or ml means count = 1 piece
    if "syrup" in desc_lower or "ml" in desc_lower or "syr" in desc_lower:
        return 1
    # Try to detect number of tablets/caps
    m = re.search(r'(\d+)\s*(tabs?|caps?|s)', desc_lower)
    if m:
        return int(m.group(1))
    return 1  # fallback if nothing found

df_merge['PIECES FOR BOX'] = df_merge['NEW DESCRIPTION'].apply(get_pieces_per_box)

# Business logic for order calculation
def compute_order(current_stock, allocation_boxes, pieces_per_box):
    if pd.isna(allocation_boxes) or allocation_boxes == 0 or pieces_per_box == 0:
        return None  # needs manual audit

    # Convert stock into boxes (or pieces if syrup)
    boxes_available = current_stock / pieces_per_box

    # If below allocation, check tolerance
    if boxes_available < allocation_boxes:
        needed_pieces = allocation_boxes * pieces_per_box
        if current_stock <= (needed_pieces - 25):  # tolerance 25 pcs
            return allocation_boxes - math.floor(boxes_available)
        else:
            return 0

    return 0  # already enough

df_merge['ORDER_QUANTITY'] = df_merge.apply(
    lambda x: compute_order(x['CURRENT STOCK'], x['ALLOCATION'], x['PIECES FOR BOX']), axis=1
)

# Flag mixed medicines
df_merge['MANUAL AUDIT'] = df_merge['NEW DESCRIPTION'].apply(
    lambda d: "Yes" if "+" in d or "/" in d else "No"
)

# Save final result
df_merge[['NEW DESCRIPTION', 'ALLOCATION', 'CURRENT STOCK', 'ORDER_QUANTITY', 'MANUAL AUDIT']].to_excel(output_file, index=False)

'''
df_merge[['NEW DESCRIPTION', 'ALLOCATION', 'CURRENT STOCK', 'PIECES FOR BOX', 'ORDER_QUANTITY', 'MANUAL AUDIT']].to_excel(output_file, index=False)
'''
print("✅ Processing complete! Check:", output_file)