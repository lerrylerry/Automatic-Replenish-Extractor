import pandas as pd
import re
import math
import streamlit as st
import io

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

# Round half up
def round_half_up(x: float) -> int:
    if x <= 0: return 0
    return int(math.floor(x + 0.5))

def safe_int(x):
    try:
        v = float(x)
        if v < 0 or pd.isna(v): return 0
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
            try: num = str(float(num)*1000.0); unit="mg"
            except: pass
        return f"{num}{unit}"
    m2 = re.search(r"(\d+(?:\.\d+)?)\s*mg\/?(?:ml|mL)", s)
    if m2: return f"{m2.group(1)}mg/ml"
    return ""

def extract_pack(s: str):
    s = norm_text(s)
    if re.search(r"\b\d+\s*[x×]\s*\d+\b", s) or re.search(r"\bbxs?\b.*[x×].*\d", s):
        return None, True
    m = re.search(r"\b(\d+)\s*'?s\b", s)
    if m:
        try: return int(m.group(1)), False
        except: return None, False
    return None, False

def is_combo_product(s: str) -> bool:
    return any(x in s for x in ["+", "&", "with"])

def primary_generic_root(s: str) -> str:
    s = norm_text(s)
    cut = re.split(r"\d", s, maxsplit=1)[0]
    toks = [t for t in cut.split() if t not in {"film", "coated"}]
    return " ".join(toks[:2]) if toks else s

def same_combo_exact(a: str, b: str) -> bool:
    a_set = set(a.lower().split("+"))
    b_set = set(b.lower().split("+"))
    return a_set == b_set

def cap_tab_mismatch(a: str, b: str) -> bool:
    at = set(norm_text(a).split())
    bt = set(norm_text(b).split())
    a_form = ("tablet" in at or "tab" in at) or ("capsule" in at or "cap" in at)
    b_form = ("tablet" in bt or "tab" in bt) or ("capsule" in bt or "cap" in bt)
    if not (a_form and b_form): return False
    a_is_tab = ("tablet" in at or "tab" in at)
    b_is_tab = ("tablet" in bt or "tab" in bt)
    return a_is_tab != b_is_tab

def is_unclear_pack(s: str) -> bool:
    s = norm_text(s)
    return any(kw in s for kw in ["pwdr", "soln", "solution", "powder"])

# ==== STREAMLIT UI ====
st.title("📦 Automatic Replenish Extractor")

file_allocation = st.file_uploader("Upload Template Excel", type=["xlsx"])
file_inventory = st.file_uploader("Upload Extract Excel", type=["xlsx"])

if file_allocation and file_inventory:
    dfa = pd.read_excel(file_allocation)
    dfb = pd.read_excel(file_inventory)

    # Auto-detect columns
    def detect_col(df, key):
        for c in df.columns:
            if key in c.lower(): return c
        raise ValueError(f"No column with '{key}' found")
    
    desc_a = detect_col(dfa, "desc")
    alloc_col = detect_col(dfa, "alloc")
    desc_b = detect_col(dfb, "desc")
    inv_col = detect_col(dfb, "invent")
    
    # Precompute features
    for df, desc_col in [(dfa, desc_a), (dfb, desc_b)]:
        df["norm"] = df[desc_col].apply(norm_text)
        df["strength"] = df[desc_col].apply(extract_strength)
        df["form_family"] = df[desc_col].apply(detect_form_family)
        df[["pack", "pack_complex"]] = df[desc_col].apply(lambda s: pd.Series(extract_pack(s)))
        df["root"] = df["norm"].apply(primary_generic_root)
    
    # Apply safe int to numeric columns
    dfa[alloc_col] = dfa[alloc_col].apply(safe_int)
    dfb[inv_col] = dfb[inv_col].apply(safe_int)

    used_indices = set()
    rows = []

    for idx, A in dfa.iterrows():
        alloc = safe_int(A[alloc_col])
        a_desc = A[desc_a]
        a_norm = A["norm"]
        a_root = A["root"]
        a_strength = A["strength"]
        a_formfam = A["form_family"]
        a_pack, a_complex = A["pack"], A["pack_complex"]

        remarks = []

        # Candidate selection
        cand = dfb[
            (dfb[inv_col] > 0) &
            (dfb["form_family"] == a_formfam) &
            (dfb["strength"] == a_strength) &
            (dfb["root"] == a_root) &
            (dfb.index.map(lambda i: i not in used_indices)) &
            ((dfb["pack"] == a_pack) | (dfb["pack"].isna() & pd.isna(a_pack)))
        ].copy()

        if is_combo_product(a_norm):
            rows.append([a_desc, alloc, 0, alloc, "Combo product; audit"])
            continue

        if cand.shape[0] > 1:
            rows.append([a_desc, alloc, 0, alloc, "Multiple inventory matches; audit"])
            continue

        if cand.empty:
            rows.append([a_desc, alloc, 0, alloc, "No match found"])
            continue

        total_inventory = int(sum([safe_int(x) for x in cand[inv_col]]))

        # Exact vs fuzzy
        if any(cand["norm"] == a_norm):
            remarks.append("Exact Match")
        else:
            remarks.append("Alternatives found")

        if any(cand["pack_complex"]) or a_complex:
            remarks.append("Audit pack: complex count")
        if any(cand["pack"].isna()) and a_formfam == "solid":
            remarks.append("Check pack size")
        if any(cand["norm"].apply(lambda s: cap_tab_mismatch(a_norm, s))):
            remarks.append("Capsule vs Tablet; review")
        if is_unclear_pack(a_desc):
            remarks.append("Unclear pack (Pwdr/Soln); needs checking")

        remaining = 0
        if a_formfam == "solid":
            if a_pack is None:
                remaining = alloc
                remarks.append("Check pack size")
            else:
                need_float = alloc - (float(total_inventory) / float(a_pack))
                need_float = max(0.0, need_float)
                remaining = round_half_up(need_float)
        else:
            need_float = alloc - float(total_inventory)
            need_float = max(0.0, need_float)
            remaining = round_half_up(need_float)

        if remaining < 0: remaining = 0

        # Mark used indices so inventory is not double-counted
        for i in cand.index: used_indices.add(i)

        remarks_str = "; ".join(sorted(set(remarks))) if remarks else ""
        rows.append([a_desc, alloc, total_inventory, remaining, remarks_str])

    out = pd.DataFrame(rows, columns=["Description", "Allocation", "Inventory", "Ordering", "Remarks"])
    st.success("✅ Processing complete!")
    st.dataframe(out)

    # Download
    buffer = io.BytesIO()
    out.to_excel(buffer, index=False)
    buffer.seek(0)
    st.download_button(
        label="📥 Download Output Excel",
        data=buffer,
        file_name="output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    