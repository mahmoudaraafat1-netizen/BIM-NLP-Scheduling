# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import re
import subprocess
import tkinter as tk
from tkinter import filedialog, simpledialog
from sentence_transformers import SentenceTransformer, util

# ========= Ensure packages (no-op if already installed) =========
for pkg in ["pandas", "openpyxl", "sentence-transformers", "torch"]:
    try:
        __import__(pkg.split("-")[0])
    except ImportError:
        print(f"Installing {pkg} ...")
        subprocess.run(["pip", "install", pkg], check=False)

# ========= Load SBERT model =========
model = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")

# ========= UI (English) =========
root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)

similarity_threshold = simpledialog.askfloat(
    "Similarity Threshold",
    "Enter similarity threshold (0.0 - 1.0), e.g., 0.40:",
    minvalue=0.0, maxvalue=1.0, initialvalue=0.40
)
if similarity_threshold is None:
    raise SystemExit

print("Please select the Items (Elements) file...")
items_path = filedialog.askopenfilename(
    title="Select Items File", filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
)
if not items_path:
    raise SystemExit("No Items file selected.")

print("Please select the Pricing Dictionary file...")
pricing_path = filedialog.askopenfilename(
    title="Select Pricing Dictionary File", filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
)
if not pricing_path:
    raise SystemExit("No Pricing Dictionary file selected.")

# ========= Load data =========
items_df = pd.read_excel(items_path)             # first sheet
pricing_df = pd.read_excel(pricing_path, sheet_name=0)

# ========= Normalize headers (preserve originals) =========
items_df.columns   = [str(c).strip() for c in items_df.columns]
pricing_df.columns = [str(c).strip() for c in pricing_df.columns]
items_cols_l   = [c.lower() for c in items_df.columns]
pricing_cols_l = [c.lower() for c in pricing_df.columns]

def find_col(df_cols_lower, candidates):
    for cand in candidates:
        lc = cand.lower()
        if lc in df_cols_lower:
            return df_cols_lower.index(lc)
    return None

# Items columns (must exist)
idx_type = find_col(items_cols_l, ["Type"])
idx_name = find_col(items_cols_l, ["Element Name"])
idx_area = find_col(items_cols_l, ["Area"])
idx_vol  = find_col(items_cols_l, ["Volume"])
if any(x is None for x in [idx_type, idx_name, idx_area, idx_vol]):
    raise ValueError("Items file must contain columns: Type, Element Name, Area, Volume.")

col_type = items_df.columns[idx_type]
col_name = items_df.columns[idx_name]
col_area = items_df.columns[idx_area]
col_vol  = items_df.columns[idx_vol]

# Pricing columns (must exist)
idx_desc = find_col(pricing_cols_l, ["BOQ Description", "Description", "Item Name"])
idx_unit = find_col(pricing_cols_l, ["Unit of Measure", "Unit"])
idx_rate = find_col(pricing_cols_l, ["Selling Price Rate", "Unit Price", "Rate"])
if any(x is None for x in [idx_desc, idx_unit, idx_rate]):
    raise ValueError("Pricing file must contain: BOQ Description, Unit of Measure, Selling Price Rate.")

col_desc = pricing_df.columns[idx_desc]
col_unit = pricing_df.columns[idx_unit]
col_rate = pricing_df.columns[idx_rate]

# ========= Clean numerics =========
def to_float(x):
    if pd.isna(x): return np.nan
    if isinstance(x, (int, float, np.number)): return float(x)
    s = str(x).replace(",", "").strip()
    try: return float(s)
    except: return np.nan

items_df[col_area] = pd.to_numeric(items_df[col_area], errors="coerce")
items_df[col_vol]  = pd.to_numeric(items_df[col_vol],  errors="coerce")
pricing_df[col_rate] = pricing_df[col_rate].apply(to_float)

# ========= UOM helpers =========
def norm_uom(u):
    if pd.isna(u): return ""
    s = str(u).strip().lower().replace(" ", "")
    s = s.replace("meter", "m").replace("square", "m2").replace("cubic", "m3")
    if s in {"m²", "sqm", "m2"}: return "m2"
    if s in {"m³", "cum", "m3"}: return "m3"
    if s in {"pcs", "pc", "piece"}: return "pcs"
    return s

def choose_qty_by_uom(uom, area, vol):
    u = norm_uom(uom)
    if u == "m3": return (area if False else (vol if pd.notna(vol) else np.nan), "volume(m3)")
    if u == "m2": return (area if pd.notna(area) else np.nan, "area(m2)")
    if pd.notna(area): return (area, "area(fallback)")
    if pd.notna(vol):  return (vol,  "volume(fallback)")
    return (np.nan, "no-qty")

# ========= Compose texts & embeddings =========
def compose_item_text(t, n):
    t = "" if pd.isna(t) else str(t).strip().lower()
    n = "" if pd.isna(n) else str(n).strip().lower()
    return f"{t} | {n}" if t and n else (t or n)

items_texts = [compose_item_text(items_df.loc[i, col_type], items_df.loc[i, col_name]) for i in range(len(items_df))]
desc_texts  = [str(pricing_df.loc[i, col_desc]).strip().lower() for i in range(len(pricing_df))]

print("Computing embeddings...")
items_emb = model.encode(items_texts, convert_to_tensor=True, normalize_embeddings=True)
desc_emb  = model.encode(desc_texts,  convert_to_tensor=True, normalize_embeddings=True)

# ========= Level-aware patterns (SOG / Floors / Basement) =========
SOG_PATTERNS = [
    r"\bslab\s*on\s*grade\b", r"\bon-?grade\b", r"\bground\s*floor\b", r"\bsog\b", r"\bgf\b"
]
FLOOR_PATTERNS = [
    r"\b(first|1st|second|2nd|third|3rd|fourth|4th|fifth|5th)\s*floor\b",
    r"\btypical\s*floor\b", r"\bupper\b", r"\bpent\b", r"\broof\b", r"\bmezz?\b", r"\bl\d+\b"
]
BASEMENT_PATTERNS = [r"\bbasement\b", r"\bb\d+\b"]

def any_match(text, patterns):
    t = (text or "").lower()
    return any(re.search(p, t) for p in patterns)

def adjust_scores_for_level(item_text, desc_text, base_score):
    bonus = 0.0
    item_is_sog = any_match(item_text, SOG_PATTERNS)
    item_is_floor = any_match(item_text, FLOOR_PATTERNS)
    item_is_basement = any_match(item_text, BASEMENT_PATTERNS)

    desc_has_sog = any_match(desc_text, SOG_PATTERNS)
    desc_has_floor = any_match(desc_text, FLOOR_PATTERNS)
    desc_has_basement = any_match(desc_text, BASEMENT_PATTERNS)

    if item_is_sog:
        if desc_has_sog: bonus += 0.15
        if desc_has_floor or desc_has_basement: bonus -= 0.15
    if item_is_floor:
        if desc_has_floor: bonus += 0.15
        if desc_has_sog:   bonus -= 0.15
    if item_is_basement:
        if desc_has_basement: bonus += 0.15
        if desc_has_sog:      bonus -= 0.10

    if re.search(r"\bon\s*grade\b", (item_text or "").lower()) and re.search(r"\bon\s*grade\b", (desc_text or "").lower()):
        bonus += 0.05

    return max(0.0, min(1.0, base_score + bonus))

# ========= Type-aware patterns (Columns / Slab / Foundation) =========
COLUMN_PATTERNS = [r"\bcolumn\b", r"\bcolumns\b"]
SLAB_PATTERNS   = [r"\bslab\b", r"\broof\s*slab\b", r"\bfloor\s*slab\b", r"\bon\s*grade\b"]
FOUND_PATTERNS  = [r"\bfooting\b", r"\bfoundation\b"]

def adjust_scores_by_type(item_type, desc_text, base_score):
    bonus = 0.0
    t = (item_type or "").lower()
    d = (desc_text or "").lower()

    if "column" in t:
        if any(re.search(p, d) for p in COLUMN_PATTERNS): bonus += 0.20
        if any(re.search(p, d) for p in SLAB_PATTERNS):   bonus -= 0.20

    if "slab" in t:
        if any(re.search(p, d) for p in SLAB_PATTERNS):   bonus += 0.20
        if any(re.search(p, d) for p in COLUMN_PATTERNS): bonus -= 0.20

    if "foundation" in t or "footing" in t:
        if any(re.search(p, d) for p in FOUND_PATTERNS):  bonus += 0.20

    return max(0.0, min(1.0, base_score + bonus))

# ========= Matching & pricing =========
selling_rate_col = "Selling Price rate"
selling_cost_col = "Selling Price Cost"
unit_note_col    = "Unit Note"
matched_boq_col  = "Matched BOQ Description"
matched_unit_col = "Matched Unit"
score_col        = "Score"

rates_out, costs_out = [], []
unit_notes, matched_desc, matched_unit, scores = [], [], [], []

pricing_uoms_norm = pricing_df[col_unit].apply(norm_uom).tolist()

for i in range(len(items_df)):
    # raw sims (numpy array)
    sims = util.pytorch_cos_sim(items_emb[i], desc_emb)[0].cpu().numpy()

    # adjust by level (SOG/floor/basement) then by type (slab/column/foundation)
    item_text_i = items_texts[i]
    item_type_i = items_df.loc[i, col_type]
    adj_scores = sims.copy()
    for j in range(len(desc_texts)):
        s1 = adjust_scores_for_level(item_text_i, desc_texts[j], float(sims[j]))
        adj_scores[j] = adjust_scores_by_type(item_type_i, desc_texts[j], s1)

    # Top-3 by adjusted scores
    top_idx = np.argsort(-adj_scores)[:3]
    chosen_j = top_idx[0]
    chosen_score = float(adj_scores[chosen_j])

    # prefer suitable unit (m2 -> Area, m3 -> Volume)
    area_i = items_df.loc[i, col_area]
    vol_i  = items_df.loc[i, col_vol]

    def unit_is_suitable(j):
        u = pricing_uoms_norm[j]
        return (u == "m2" and pd.notna(area_i)) or (u == "m3" and pd.notna(vol_i))

    if not unit_is_suitable(chosen_j):
        for j2 in top_idx[1:]:
            if unit_is_suitable(j2) and adj_scores[j2] >= similarity_threshold * 0.95:
                chosen_j = j2
                chosen_score = float(adj_scores[j2])
                break

    # write outputs
    if chosen_score >= similarity_threshold and pd.notna(pricing_df.loc[chosen_j, col_rate]):
        rate = float(pricing_df.loc[chosen_j, col_rate])
        uom  = pricing_df.loc[chosen_j, col_unit]
        qty, _ = choose_qty_by_uom(uom, area_i, vol_i)
        cost = (rate * float(qty)) if pd.notna(qty) else np.nan

        note = ""
        if norm_uom(uom) == "m2" and pd.isna(area_i): note = "⚠️ Expected area but missing"
        if norm_uom(uom) == "m3" and pd.isna(vol_i):  note = "⚠️ Expected volume but missing"

        rates_out.append(rate)
        costs_out.append(cost)
        unit_notes.append(note)
        matched_desc.append(pricing_df.loc[chosen_j, col_desc])
        matched_unit.append(uom)
        scores.append(round(chosen_score, 3))
    else:
        rates_out.append(np.nan)
        costs_out.append(np.nan)
        unit_notes.append("Manual Review")
        matched_desc.append("No Match")
        matched_unit.append(np.nan)
        scores.append(round(chosen_score, 3))

# ========= Write output columns =========
items_df[selling_rate_col] = pd.to_numeric(rates_out, errors="coerce").round(4)
items_df[selling_cost_col] = pd.to_numeric(costs_out, errors="coerce").round(4)
items_df[unit_note_col]    = unit_notes
items_df[matched_boq_col]  = matched_desc
items_df[matched_unit_col] = matched_unit
items_df[score_col]        = scores

# ========= Save =========
save_path = filedialog.asksaveasfilename(
    title="Save Priced Items", defaultextension=".xlsx",
    initialfile="Priced_Items.xlsx",
    filetypes=[("Excel files","*.xlsx")]
)
if save_path:
    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
        items_df.to_excel(writer, index=False, sheet_name="Priced Items")
    print(f"Saved: {save_path}")
else:
    print("Save cancelled.")
