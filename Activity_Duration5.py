# -*- coding: utf-8 -*-
import math, re, subprocess
from collections import Counter

import numpy as np
import pandas as pd
from sentence_transformers import SentenceTransformer, util
import tkinter as tk
from tkinter import simpledialog, filedialog

# --------- Package check & install ----------
required_packages = ["pandas", "openpyxl", "sentence-transformers", "torch", "numpy"]
for package in required_packages:
    try:
        __import__(package.split("-")[0])
    except ImportError:
        print(f"Installing {package} ...")
        subprocess.run(["pip", "install", package], check=False)

print("All required packages are installed.")

# --------- Load SBERT model ----------
model = SentenceTransformer('sentence-transformers/all-MiniLM-L6-v2')

# --------- Initialize tkinter ----------
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)

# --------- User inputs via dialogs ----------
max_duration_days = simpledialog.askinteger(
    "Max Duration (days)", "Enter the maximum allowed duration (days):",
    minvalue=1, initialvalue=25
)
if max_duration_days is None: raise SystemExit

similarity_threshold = simpledialog.askfloat(
    "Similarity Threshold", "Enter similarity threshold (0.0 - 1.0), e.g., 0.40:",
    minvalue=0.0, maxvalue=1.0, initialvalue=0.40
)
if similarity_threshold is None: raise SystemExit

default_crews = simpledialog.askinteger(
    "Number of Crews", "How many crews?", minvalue=1, initialvalue=1
)
if default_crews is None: raise SystemExit

baseline_area = simpledialog.askfloat(
    "Baseline Qty (default m²)",
    "Enter baseline quantity for reference-duration scaling (e.g., 1500):",
    minvalue=1.0, initialvalue=1500.0
)
if baseline_area is None: raise SystemExit

# ---- Steel intensity defaults (editable in dialog) ----
steel_col_kgm3 = simpledialog.askfloat(
    "Steel Intensity - Columns", "Default steel intensity for Columns (kg/m³):",
    minvalue=1.0, initialvalue=120.0
)
steel_slab_kgm3 = simpledialog.askfloat(
    "Steel Intensity - Slabs", "Default steel intensity for Slabs (kg/m³):",
    minvalue=1.0, initialvalue=100.0
)
steel_found_kgm3 = simpledialog.askfloat(
    "Steel Intensity - Foundations", "Default steel intensity for Foundations (kg/m³):",
    minvalue=1.0, initialvalue=90.0
)

print(f"Max duration: {max_duration_days} days")
print(f"Threshold: {similarity_threshold}")
print(f"Crews: {default_crews}")
print(f"Baseline Qty: {baseline_area}")
print(f"Steel factors (kg/m³): Columns={steel_col_kgm3}, Slabs={steel_slab_kgm3}, Foundations={steel_found_kgm3}")

# --------- File selection ----------
print("Please select the Activity List file...")
activity_list_file = filedialog.askopenfilename(
    title="Select Activity List file", filetypes=[("Excel files", "*.xlsx")]
)
if not activity_list_file: raise SystemExit

print("Please select the Dictionary file...")
dictionary_file = filedialog.askopenfilename(
    title="Select Dictionary file", filetypes=[("Excel files", "*.xlsx")]
)
if not dictionary_file: raise SystemExit

# --------- Load data ----------
activity_list_df = pd.read_excel(activity_list_file)
dictionary_df = pd.read_excel(dictionary_file, sheet_name="Duration")

# normalize headers
activity_list_df.columns = [str(c).strip().lower() for c in activity_list_df.columns]
dictionary_df.columns = [str(c).strip().lower() for c in dictionary_df.columns]

# activity columns
col_activity_name = "activity name"
col_type   = "type"
col_element = "element"
col_area   = "area"
col_volume = "volume"     # optional
col_weight = "weight"     # optional actual weight (kg) if available

# dictionary columns
dict_activity_name = "activity name"
dict_prod_rate     = "production rate"
dict_ref_duration  = "reference duration"
# detect unit column (e.g., "Unit /day")
dict_unit_raw = None
for cand in ["unit /day","unit/day","unit","production unit","rate unit","uom","unit of measure"]:
    if cand in dictionary_df.columns:
        dict_unit_raw = cand
        break
if dict_unit_raw is None:
    raise ValueError("Dictionary must include a unit column (e.g., 'Unit /day').")

# --------- Ensure numeric ----------
if col_area in activity_list_df.columns:
    activity_list_df[col_area] = pd.to_numeric(activity_list_df[col_area], errors="coerce").fillna(0)
else:
    activity_list_df[col_area] = 0

if col_volume in activity_list_df.columns:
    activity_list_df[col_volume] = pd.to_numeric(activity_list_df[col_volume], errors="coerce").fillna(0)
else:
    activity_list_df[col_volume] = 0

if col_weight in activity_list_df.columns:
    activity_list_df[col_weight] = pd.to_numeric(activity_list_df[col_weight], errors="coerce").fillna(0)
else:
    activity_list_df[col_weight] = 0

# ---- Clean production rate: handle '250 m2/day' or '250,0' etc. ----
prod_raw = dictionary_df[dict_prod_rate].astype(str).str.replace(",", ".", regex=False)
dictionary_df[dict_prod_rate] = pd.to_numeric(
    prod_raw.str.extract(r"([\d.]+)")[0],
    errors="coerce"
)
dictionary_df[dict_ref_duration] = pd.to_numeric(dictionary_df[dict_ref_duration], errors="coerce")

activity_list_df["number of crews"] = default_crews

# =========================
# SMART MATCH HELPERS (NEW)
# =========================
ACTION_SYNS = {
    "shuttering": {"shuttering", "formwork", "shutter", "form"},
    "steelfixing": {"steelfixing", "rebar", "steel fixing", "reinforcement"},
    "pouring": {"pouring", "cast", "casting", "concreting"},
    "deshuttering": {"deshuttering", "striking", "deform", "de-shuttering", "strip"},
}

TYPE_ANCHORS = {
    "slab": {"slab", "slabs", "rc slab"},
    "columns": {"column", "columns", "rc column"},
    "foundation": {"foundation", "foundations", "raft", "footing", "pile cap", "pc", "p.c"},
}

ELEMENT_GROUP_PATTERNS = [
    ("sog", r"\bslab(?:\s+)?on(?:\s+)?grade\b|\bsog\b"),
    ("floor", r"\b(first|second|third|forth|fourth|fifth|sixth)\s+floor\b|\bfloor\b"),
    ("pc", r"\b(p\.?c)\b|\bprecast\b"),
    ("foundation", r"\bfoundation\b|\bfooting\b|\bpile\s*cap\b|\braft\b"),
]

STOPWORDS = {"rc","concrete","of","on","the"}

def normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"[-_/]", " ", s)
    s = re.sub(r"\d+(\.\d+)?", "", s)   # remove numbers like 103.30, 106.00
    s = re.sub(r"\s+", " ", s).strip()
    return s

def detect_action(s: str):
    s = normalize_text(s)
    best, score = None, 0
    for act, syns in ACTION_SYNS.items():
        hits = [x for x in syns if x in s]
        if hits:
            local = max(len(x) for x in hits)
            if local > score:
                score, best = local, act
    return best

def detect_type_anchor(s: str):
    s = normalize_text(s)
    for t, keys in TYPE_ANCHORS.items():
        if any(k in s for k in keys):
            return t
    if "column" in s: return "columns"
    if "slab" in s: return "slab"
    if any(k in s for k in ["foundation","footing","pile cap","raft"]): return "foundation"
    return None

def detect_element_group(s: str):
    s = normalize_text(s)
    for name, pat in ELEMENT_GROUP_PATTERNS:
        if re.search(pat, s):
            return name
    if "floor" in s: return "floor"
    if "slab on grade" in s or "sog" in s: return "sog"
    return None

def token_overlap(a: str, b: str) -> float:
    toksA = [t for t in re.findall(r"[a-z]+", normalize_text(a)) if t not in STOPWORDS]
    toksB = [t for t in re.findall(r"[a-z]+", normalize_text(b)) if t not in STOPWORDS]
    if not toksA or not toksB: return 0.0
    ca, cb = Counter(toksA), Counter(toksB)
    inter = sum((ca & cb).values())
    uni   = sum((ca | cb).values())
    return inter/uni if uni else 0.0

def feature_match_score(query_txt: str, cand_txt: str, emb_sim: float):
    q_action  = detect_action(query_txt)
    c_action  = detect_action(cand_txt)
    q_type    = detect_type_anchor(query_txt)
    c_type    = detect_type_anchor(cand_txt)
    q_elem    = detect_element_group(query_txt)
    c_elem    = detect_element_group(cand_txt)

    flags = {
        "action_eq": q_action and c_action and (q_action == c_action),
        "type_eq":   q_type and c_type and (q_type == c_type),
        "elem_eq":   q_elem and c_elem and (q_elem == c_elem),
        "q_action": q_action, "c_action": c_action,
        "q_type": q_type, "c_type": c_type,
        "q_elem": q_elem, "c_elem": c_elem,
    }

    # HARD FILTERS
    if q_action and c_action and q_action != c_action:
        return -1.0, flags, "reject: action mismatch"
    if q_elem and c_elem and q_elem != c_elem:
        return -1.0, flags, "reject: element mismatch"
    if q_type and c_type and q_type != c_type:
        return -1.0, flags, "reject: type mismatch"

    overlap = token_overlap(query_txt, cand_txt)
    bonus = 0.0
    if flags["action_eq"]: bonus += 0.06
    if flags["elem_eq"]:   bonus += 0.06
    if flags["type_eq"]:   bonus += 0.04

    final = 0.7*emb_sim + 0.2*overlap + 0.1*bonus
    return final, flags, f"emb={emb_sim:.3f}, ovl={overlap:.3f}, bonus={bonus:.2f}"

# --------- Units: robust extractor (PATCHED) ----------
def norm_uom(u: str) -> str:
    """Robust unit extractor from messy strings like 'Area @ m2/day'."""
    if pd.isna(u):
        return ""
    s = str(u).lower().strip()

    # normalize visuals/aliases
    s = (s.replace("²", "2").replace("³", "3")
           .replace("sqm", "m2").replace("m^2", "m2").replace("m^3", "m3"))

    # detect anywhere in the string
    if "m3" in s: return "m3"
    if "m2" in s: return "m2"
    if "kg" in s: return "kg"
    if "ton" in s or re.search(r"\b(t|tonne)\b", s): return "ton"
    if re.search(r"\bpcs?\b|\bno\b|\bnr\b", s): return "pcs"
    if re.search(r"\blm\b|\bm\b", s): return "m"

    # fallback: left of '/'
    if "/" in s:
        left = s.split("/", 1)[0]
        if "m3" in left: return "m3"
        if "m2" in left: return "m2"
        if "kg" in left: return "kg"
        if "ton" in left or left.strip() == "t": return "ton"

    return ""

def steel_factor_for_activity(activity_text: str) -> float:
    t = (activity_text or "").lower()
    if "column" in t: return steel_col_kgm3
    if "slab" in t: return steel_slab_kgm3
    if "foundation" in t or "footing" in t or "raft" in t: return steel_found_kgm3
    return steel_slab_kgm3

def choose_quantity_strict_by_unit(row, unit_value, activity_text):
    """
    m3 -> Volume
    m2 -> Area
    kg/ton -> Weight; if not present, estimate from Volume × steel_factor(activity) (kg/m3)
    """
    u = norm_uom(unit_value)
    a = float(row.get(col_area, 0) or 0)
    v = float(row.get(col_volume, 0) or 0)
    w = float(row.get(col_weight, 0) or 0)

    if u == "m3":
        return (v if v > 0 else 0.0), "Area @ m3/day (Volume)"
    if u == "m2":
        return (a if a > 0 else 0.0), "Area @ m2/day"
    if u in {"kg","ton"}:
        if w > 0:
            if u == "kg":  return w, "Weight(kg) @ kg/day"
            else:          return w/1000.0, "Weight(ton from kg) @ ton/day"
        if v > 0:
            factor = steel_factor_for_activity(activity_text)  # kg/m3
            est_kg = v * factor
            if u == "kg":  return est_kg, f"Weight(est: {factor} kg/m3 × Volume) @ kg/day"
            else:          return est_kg/1000.0, f"Weight(est: {factor} kg/m3 × Volume) @ ton/day"
        return 0.0, "Weight missing"
    return 0.0, "None"

# --------- Prepare embeddings ----------
activity_names = activity_list_df[col_activity_name].astype(str).str.lower().tolist()
dict_names = dictionary_df[dict_activity_name].astype(str).str.lower().tolist()

print("Computing embeddings...")
activity_emb = model.encode(activity_names, convert_to_tensor=True, normalize_embeddings=True)
dict_emb = model.encode(dict_names, convert_to_tensor=True, normalize_embeddings=True)

# --------- Matching & calculation loop ----------
matched_names, matched_scores = [], []
durations, basis_list = [], []
suggested_crews_list, suggested_durations = [], []
origin_pre_cap_list = []
estimated_weight_kg_list = []  # for final output after Volume
match_flags_list, match_reason_list = [], []
embedding_sim_list = []
parsed_unit_list, qty_used_list = [], []  # DEBUG (PATCHED)

for idx, act in enumerate(activity_names):
    # Build query text with (type/element) if available
    row = activity_list_df.iloc[idx]
    qtxt = f"{row.get(col_activity_name,'')} {row.get(col_type,'')} {row.get(col_element,'')}"
    qtxt_norm = normalize_text(qtxt)

    sims_t = util.pytorch_cos_sim(activity_emb[idx], dict_emb)[0]  # tensor of sims
    sims = sims_t.cpu().numpy().ravel()

    # Rank candidates with smart score
    scored = []
    for i, cand in enumerate(dict_names):
        final, flags, reason = feature_match_score(qtxt_norm, cand, float(sims[i]))
        if final >= 0:
            scored.append((final, i, flags, reason))

    # Fallback if all rejected
    if not scored:
        best_idx = int(np.argmax(sims))
        final_score = float(sims[best_idx])
        best_reason = "no candidate passed hard filters; used highest emb_sim"
        best_flags = {"fallback": True}
    else:
        scored.sort(reverse=True, key=lambda x: x[0])
        final_score, best_idx, best_flags, best_reason = scored[0]

    embedding_sim_list.append(round(float(sims[best_idx]), 3))
    match_flags_list.append(str(best_flags))
    match_reason_list.append(best_reason)

    # Pull dict values
    matched_name = dictionary_df.iloc[best_idx][dict_activity_name]
    prod_rate    = dictionary_df.iloc[best_idx][dict_prod_rate]
    ref_dur      = dictionary_df.iloc[best_idx][dict_ref_duration]
    unit_val     = dictionary_df.iloc[best_idx][dict_unit_raw]  # e.g., "Area @ m2/day"

    crews = activity_list_df.iloc[idx]["number of crews"]

    # Quantity by unit (with steel estimation if needed)
    qty, qty_basis = choose_quantity_strict_by_unit(
        activity_list_df.iloc[idx], unit_val, activity_list_df.iloc[idx][col_activity_name]
    )
    parsed_unit_list.append(norm_uom(unit_val))
    qty_used_list.append(qty)

    # compute Estimated Weight (kg) for output column
    u_norm = norm_uom(unit_val)
    vol_i = float(activity_list_df.iloc[idx].get(col_volume, 0) or 0)
    w_i = float(activity_list_df.iloc[idx].get(col_weight, 0) or 0)
    if w_i > 0:
        est_weight_kg = w_i
    elif u_norm in {"kg","ton"} and vol_i > 0:
        est_weight_kg = vol_i * steel_factor_for_activity(activity_list_df.iloc[idx][col_activity_name])
    else:
        est_weight_kg = 0.0
    estimated_weight_kg_list.append(round(est_weight_kg, 4))

    duration, basis = None, "Manual Review"
    suggested_crews, suggested_dur = "N/A", "N/A"
    pre_cap_dur = None

    # استخدمنا threshold على final_score (الأذكى)
    if final_score >= similarity_threshold:
        if pd.notna(prod_rate) and prod_rate > 0 and qty > 0:
            raw = qty / (prod_rate * max(crews, 1))
            pre_cap_dur = max(1, math.ceil(raw))
            duration = min(pre_cap_dur, max_duration_days)
            basis = f"{qty_basis}"
        elif pd.notna(ref_dur) and ref_dur > 0:
            if qty > 0:
                scaled = ref_dur * (qty / max(baseline_area, 1e-6))
                raw = scaled / max(crews, 1)
                pre_cap_dur = max(1, math.ceil(raw))
                duration = min(pre_cap_dur, max_duration_days)
                basis = f"Reference-Scaled ({qty_basis})"
            else:
                pre_cap_dur = int(ref_dur)
                duration = min(pre_cap_dur, max_duration_days)
                basis = "Reference"
        else:
            duration = "Manual Review"

        # Suggested crews if capped
        if isinstance(duration, (int, float)) and pre_cap_dur and pre_cap_dur > max_duration_days:
            if qty > 0 and prod_rate and prod_rate > 0:
                crews_needed = math.ceil(qty / (prod_rate * max_duration_days))
                suggested_crews = crews_needed
                sug_raw = qty / (prod_rate * crews_needed)
                suggested_dur = max(1, math.ceil(sug_raw))
            elif qty > 0 and ref_dur and ref_dur > 0:
                crews_needed = math.ceil((ref_dur * (qty / max(baseline_area, 1e-6))) / max_duration_days)
                suggested_crews = crews_needed
                sug_raw = (ref_dur * (qty / max(baseline_area, 1e-6))) / max(crews_needed, 1)
                suggested_dur = max(1, math.ceil(sug_raw))
    else:
        duration = "Manual Review"

    matched_names.append(matched_name if final_score >= similarity_threshold else "No Match")
    matched_scores.append(round(final_score, 3))
    durations.append(duration)
    basis_list.append(basis)
    suggested_crews_list.append(suggested_crews)
    suggested_durations.append(suggested_dur)
    origin_pre_cap_list.append(pre_cap_dur if pre_cap_dur is not None else "N/A")

# --------- Add results ----------
activity_list_df["matched activity"] = matched_names
activity_list_df["similarity score"] = matched_scores              # final (smart) score
activity_list_df["embedding similarity"] = embedding_sim_list      # raw cosine for مراجعة
activity_list_df["match_flags"] = match_flags_list
activity_list_df["match_reason"] = match_reason_list
activity_list_df["unit (parsed)"] = parsed_unit_list               # (PATCHED) تشخيص
activity_list_df["qty (used)"] = qty_used_list                     # (PATCHED) تشخيص
activity_list_df["basis"] = basis_list
activity_list_df["origin duration (pre-cap)"] = origin_pre_cap_list
activity_list_df["duration (days)"] = durations
activity_list_df["suggested crews (to meet max)"] = suggested_crews_list
activity_list_df["suggested duration (days)"] = suggested_durations
activity_list_df["activity duration (final)"] = activity_list_df["duration (days)"]

# --- Insert 'Estimated Weight (kg)' after 'Volume'
insert_after = col_volume
new_col_name = "Estimated Weight (kg)"
estimated_weight_series = pd.to_numeric(pd.Series(estimated_weight_kg_list), errors="coerce").fillna(0).round(4)

if insert_after in activity_list_df.columns:
    cols = activity_list_df.columns.tolist()
    if new_col_name in cols:
        cols.remove(new_col_name)
        if new_col_name in activity_list_df.columns:
            activity_list_df.drop(columns=[new_col_name], inplace=True)
    activity_list_df[new_col_name] = estimated_weight_series
    cols = activity_list_df.columns.tolist()
    cols.remove(new_col_name)
    insert_pos = cols.index(insert_after) + 1
    cols = cols[:insert_pos] + [new_col_name] + cols[insert_pos:]
    activity_list_df = activity_list_df[cols]
else:
    activity_list_df[new_col_name] = estimated_weight_series

# keep tail order for duration-related columns
desired_tail = ["origin duration (pre-cap)", "duration (days)", "suggested crews (to meet max)",
                "suggested duration (days)", "activity duration (final)"]
cols_tail_first = [c for c in activity_list_df.columns if c not in desired_tail] + desired_tail
activity_list_df = activity_list_df[cols_tail_first]

# --------- Save output ----------
print("Select location to save output...")
output_filename = filedialog.asksaveasfilename(
    title="Save Output File",
    defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx")],
    initialfile="Updated_Activity_List.xlsx"
)
if output_filename:
    activity_list_df.to_excel(output_filename, index=False)
    print(f"Results saved in: {output_filename}")
else:
    print("Save cancelled.")
