import pandas as pd
import re
import math

# ==== CONFIG ====
FILE_ALLOCATION = "template.xlsx"  # first Excel: Col A=Description, Col B=Allocation (boxes/bottles)
FILE_INVENTORY  = "extract.xlsx"   # second Excel: Col A=Description, Col B=Inventory (pieces/bottles)
FILE_OUTPUT     = "output.xlsx"

# ==== HELPERS ====
SOLID_FORMS = {
    "tab", "tabs", "tablet", "tablets",
    "cap", "caps", "capsule", "capsules",
    "softgel", "softgels", "sachet", "pwdr", "powder", "lozenge"
}
LIQUID_TOPICAL_FORMS = {
    "syr", "syrup", "susp", "suspension", "drops", "drop",
    "cream", "ointment", "spray", "nebule", "inhaler", "solution", "soln", "gel", "lotion"
}

def round_half_up(x: float) -> int:
    if x <= 0:
        return 0
    return int(math.floor(x + 0.5))

def safe_int(x):
    try:
        v = float(x)
        if v < 0 or pd.isna(v):
            return 0
        return int(v) if float(v).is_integer() else float(v)
    except:
        return 0

def norm_text(s: str) -> str:
    if pd.isna(s): return ""
    s = str(s).lower()
    s = s.replace("’", "'")
    s = re.sub(r"[\(\)\[\]{}]", " ", s)
    s = re.sub(r"[^a-z0-9\+\s\.%/x']", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    # unify synonyms
    s = s.replace("tablets", "tablet").replace("tabs", "tab")
    s = s.replace("capsules", "capsule").replace("caps", "cap")
    s = s.replace("suspension", "susp").replace("syr.", "syr")
    s = s.replace("powder", "pwdr").replace("sol'n", "soln").replace("solution", "soln")
    return s

def detect_form_family(s: str) -> str:
    tokens = set(norm_text(s).split())
    if tokens & SOLID_FORMS:
        return "solid"
    if tokens & LIQUID_TOPICAL_FORMS:
        return "liquid_topical"
    if re.search(r"\b\d+\s*'?s\b", s.lower()):
        return "solid"
    return "liquid_topical"

def extract_strength(s: str):
    s = norm_text(s)
    m = re.search(r"(\d+(?:\.\d+)?)\s*(mcg|mg|g)\b", s)
    if m:
        num, unit = m.group(1), m.group(2)
        if unit == "g":
            try:
                num = str(float(num) * 1000.0)
                unit = "mg"
            except:
                pass
        return f"{num}{unit}"
    m2 = re.search(r"(\d+(?:\.\d+)?)\s*mg\/?(?:ml|mL)", s)
    if m2:
        return f"{m2.group(1)}mg/ml"
    return ""

def extract_pack(s: str):
    s0 = s
    s = norm_text(s)
    if re.search(r"\b\d+\s*[x×]\s*\d+\b", s) or re.search(r"\bbxs?\b.*[x×].*\d", s):
        return None, True
    m = re.search(r"\b(\d+)\s*'?s\b", s)
    if m:
        try:
            return int(m.group(1)), False
        except:
            return None, False
    return None, False

def is_combo_product(s: str) -> bool:
    return "+" in s

def primary_generic_root(s: str) -> str:
    s = norm_text(s)
    cut = re.split(r"\d", s, maxsplit=1)[0]
    toks = [t for t in cut.split() if t not in {"film", "coated"}]
    return " ".join(toks[:2]) if toks else s

def same_combo_pattern(a: str, b: str) -> bool:
    return is_combo_product(a) == is_combo_product(b)

def cap_tab_mismatch(a: str, b: str) -> bool:
    at = set(norm_text(a).split())
    bt = set(norm_text(b).split())
    a_form = ("tablet" in at or "tab" in at) or ("capsule" in at or "cap" in at)
    b_form = ("tablet" in bt or "tab" in bt) or ("capsule" in bt or "cap" in bt)
    if not (a_form and b_form):
        return False
    a_is_tab = ("tablet" in at or "tab" in at)
    b_is_tab = ("tablet" in bt or "tab" in bt)
    return a_is_tab != b_is_tab

def is_unclear_pack(s: str) -> bool:
    """Mark products like Pwdr or Soln as unclear pack size."""
    s = norm_text(s)
    return any(kw in s for kw in ["pwdr", "soln", "solution", "powder"])

# ==== READ FILES ====
dfa = pd.read_excel(FILE_ALLOCATION, usecols=[0,1])
dfa.columns = ["Description", "Allocation"]

dfb = pd.read_excel(FILE_INVENTORY, usecols=[0,1])
dfb.columns = ["Description", "Inventory"]

dfb["Inventory"] = dfb["Inventory"].apply(safe_int)

# Pre-compute parsed features
dfa["norm"] = dfa["Description"].apply(norm_text)
dfa["strength"] = dfa["Description"].apply(extract_strength)
dfa["form_family"] = dfa["Description"].apply(detect_form_family)
dfa[["pack", "pack_complex"]] = dfa["Description"].apply(lambda s: pd.Series(extract_pack(s)))
dfa["root"] = dfa["Description"].apply(primary_generic_root)

dfb["norm"] = dfb["Description"].apply(norm_text)
dfb["strength"] = dfb["Description"].apply(extract_strength)
dfb["form_family"] = dfb["Description"].apply(detect_form_family)
dfb[["pack", "pack_complex"]] = dfb["Description"].apply(lambda s: pd.Series(extract_pack(s)))
dfb["root"] = dfb["Description"].apply(primary_generic_root)

rows = []

for idx, A in dfa.iterrows():
    alloc = safe_int(A["Allocation"])
    a_desc = A["Description"]
    a_norm = A["norm"]
    a_root = A["root"]
    a_strength = A["strength"]
    a_formfam = A["form_family"]
    a_pack, a_complex = A["pack"], A["pack_complex"]

    remarks = []

    # --- Candidate selection ---
    cand = dfb[
        (dfb["Inventory"] > 0) &
        (dfb["form_family"] == a_formfam) &
        (dfb["strength"] == a_strength) &
        (dfb["root"].str.contains(a_root.split()[0] if a_root else "", na=False)) &
        (dfb["norm"].str.contains(a_root.split()[0] if a_root else "", na=False)) &
        (dfb["norm"].apply(lambda s: same_combo_pattern(a_norm, s)))
    ].copy()

    if cand.empty:
        fw = (a_root.split()[0] if a_root else "")
        if fw:
            cand = dfb[
                (dfb["Inventory"] > 0) &
                (dfb["form_family"] == a_formfam) &
                (dfb["strength"] == a_strength) &
                (dfb["norm"].str.contains(fw))
            ].copy()

    if cand.empty:
        rows.append([a_desc, alloc, 0, alloc, "No match found"])
        continue

    total_inventory_pieces = int(sum([safe_int(x) for x in cand["Inventory"]]))

    # --- Exact vs Fuzzy match detection ---
    if any(cand["norm"] == a_norm):
        remarks.append("Exact Match")
    else:
        remarks.append("Alternatives found")

    # --- Extra remarks from audit rules ---
    if any(cand["pack_complex"]) or a_complex:
        remarks.append("Audit pack: complex count")
    if any(cand["pack"].isna()) and a_formfam == "solid":
        remarks.append("Check pack size")
    if cand.shape[0] > 1:
        remarks.append("Combined multiple brands")
    if any(cand["norm"].apply(lambda s: cap_tab_mismatch(a_norm, s))):
        remarks.append("Capsule vs Tablet; review")
    if is_unclear_pack(a_desc):
        remarks.append("Unclear pack (Pwdr/Soln); needs checking")

    # --- Remaining computation ---
    remaining = 0
    if a_formfam == "solid":
        if a_pack is None:
            remaining = alloc
            remarks.append("Check pack size")
        else:
            need_float = alloc - (float(total_inventory_pieces) / float(a_pack))
            need_float = max(0.0, need_float)
            remaining = round_half_up(need_float)
    else:
        need_float = alloc - float(total_inventory_pieces)
        need_float = max(0.0, need_float)
        remaining = round_half_up(need_float)

    if remaining < 0:
        remaining = 0

    remarks_str = "; ".join(sorted(set(remarks))) if remarks else ""
    rows.append([a_desc, alloc, total_inventory_pieces, remaining, remarks_str])

out = pd.DataFrame(rows, columns=["Description", "Allocation", "Inventory", "Remaining", "Remarks"])
out.to_excel(FILE_OUTPUT, index=False)
print(f"✅ Done. Wrote {FILE_OUTPUT}")
