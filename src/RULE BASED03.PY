import pandas as pd
import numpy as np
import re
from tkinter import filedialog, Tk

# ===================== UI: Select input file =====================
Tk().withdraw()
input_file = filedialog.askopenfilename(
    title="Select Activity List File",
    filetypes=[("Excel Files", "*.xlsx")]
)
if not input_file:
    print("❌ No file selected. Exiting...")
    raise SystemExit

print(f"✅ File selected: {input_file}")

# ===================== Load Excel =====================
df = pd.read_excel(input_file)
print(f"📄 Loaded {len(df)} rows from Excel.")

# ===================== Basic cleaning =====================
required_cols = ['Activity ID', 'Activity Name']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise ValueError(f"Input must contain columns: {missing}")

df['Activity Name'] = (
    df['Activity Name']
    .astype(str)
    .str.strip()
    .str.replace("-", " ", regex=False)
    .str.replace(r"\s+", " ", regex=True)
    .str.lower()
)

# ===================== Floor extraction (robust) =====================
def normalize_floor_text(text: str) -> str:
    return (text or '').replace('lvl', 'level').replace('flr', 'floor')

FLOOR_PATTERNS = [
    ('LEVEL_NUM',     r'\blevel\s*(\d+)\b'),
    ('L_NUM',         r'\bl\s*(\d+)\b'),
    ('FLOOR_NUM',     r'\bfloor\s*(\d+)\b'),
    ('BASEMENT_WORD', r'\bbasement\s*(\d+)\b'),
    ('BASEMENT_B',    r'\bb(\d+)\b'),
    ('GF',            r'\b(?:gf|g\.?f\.?|ground\s*floor)\b'),
    ('ROOF',          r'\broof\b'),
    ('PODIUM',        r'\bpodium\b'),
    ('MEZZ',          r'\bmezzanine\b'),
    ('ORDINAL_FLOOR', r'\b(first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)\s+floor\b')
]
ORDINAL_MAP = {
    'first': 'L1', 'second': 'L2', 'third': 'L3', 'fourth': 'L4', 'fifth': 'L5',
    'sixth': 'L6', 'seventh': 'L7', 'eighth': 'L8', 'ninth': 'L9', 'tenth': 'L10'
}

def extract_floor_token(text: str) -> str:
    t = (text or '').lower()
    t = normalize_floor_text(t)
    for kind, pat in FLOOR_PATTERNS:
        m = re.search(pat, t, flags=re.IGNORECASE)
        if not m:
            continue
        if kind in ('LEVEL_NUM', 'L_NUM', 'FLOOR_NUM'):
            return f"L{int(m.group(1))}"
        if kind in ('BASEMENT_WORD', 'BASEMENT_B'):
            return f"B{int(m.group(1))}"
        if kind == 'GF':
            return 'GF'
        if kind == 'ROOF':
            return 'ROOF'
        if kind == 'PODIUM':
            return 'PODIUM'
        if kind == 'MEZZ':
            return 'MEZZ'
        if kind == 'ORDINAL_FLOOR':
            return ORDINAL_MAP.get(m.group(1).lower(), '')
    return ''

def floor_sort_key(token: str):
    if not token or not isinstance(token, str):
        return (5, 0)
    if token.startswith('B') and token[1:].isdigit():
        return (0, -int(token[1:]))    # B3 < B2 < B1
    if token == 'GF':
        return (1, 0)                  # GF ~ 0
    if token == 'MEZZ':
        return (1, 0.5)
    if token == 'PODIUM':
        return (1, 0.75)
    if token.startswith('L') and token[1:].isdigit():
        return (1, int(token[1:]))     # L1, L2, ...
    if token == 'ROOF':
        return (2, 10**6)
    return (5, 0)

# ===================== Component & Action detection =====================
def extract_component(text: str) -> str:
    t = (text or '').lower()
    # foundation family
    if 'foundation' in t or 'footing' in t or 'raft' in t:
        return 'foundation'
    # column (include "col" as common abbreviation)
    if re.search(r'\bcolumn(s)?\b', t) or re.search(r'\bcol\b', t):
        return 'column'
    # slab family (will also catch slab on grade)
    if 'slab' in t:
        return 'slab'
    return 'other'

def detect_action(text: str) -> str:
    t = (text or '').lower()
    # التسلسل: steelfixing -> shuttering -> concrete -> deshuttering
    if re.search(r'\bdeshutter|deformwork|remove formwork|strik', t):
        return 'deshuttering'
    if re.search(r'\bshutter|formwork', t):
        return 'shuttering'
    if re.search(r'\bconcret|pouring|casting', t):
        return 'concrete'
    if re.search(r'\bsteel|reinforc|rebar|fixing', t):
        return 'steelfixing'
    return 'other'

ACTION_ORDER = {'steelfixing':1, 'shuttering':2, 'concrete':3, 'deshuttering':4, 'other':9}

df['FloorToken'] = df['Activity Name'].apply(extract_floor_token)
df['Component']  = df['Activity Name'].apply(extract_component)
df['Action']     = df['Activity Name'].apply(detect_action)
df['_orig_idx']  = np.arange(len(df))

# ======== Detect SOG and assign GF if floor is missing ========
def is_sog(name: str) -> bool:
    t = str(name).lower()
    return bool(re.search(r'\b(slab\s*on\s*grade|sog|ground\s*slab)\b', t))

sog_mask = df['Activity Name'].apply(is_sog)
# لو النشاط SOG ومفيش FloorToken، خليه GF
df.loc[sog_mask & ((df['FloorToken'] == '') | (df['FloorToken'].isna())), 'FloorToken'] = 'GF'

# ===================== Optional dates =====================
DATE_CANDIDATES_START = ['Start', 'Planned Start', 'Baseline Start', 'Data Date', 'Early Start']
DATE_CANDIDATES_FIN   = ['Finish', 'Planned Finish', 'Baseline Finish', 'Early Finish']

def find_first_existing(colnames):
    for c in colnames:
        if c in df.columns:
            return c
    return None

col_start = find_first_existing(DATE_CANDIDATES_START)
col_finish = find_first_existing(DATE_CANDIDATES_FIN)

if col_start:
    df['_Start'] = pd.to_datetime(df[col_start], errors='coerce')
if col_finish:
    df['_Finish'] = pd.to_datetime(df[col_finish], errors='coerce')

# ======== Stable sorting helpers ========
def add_action_order(sub: pd.DataFrame) -> pd.DataFrame:
    sub = sub.copy()
    sub['__ActionOrder'] = sub['Action'].map(ACTION_ORDER).fillna(9).astype(int)
    return sub

def sort_within_group(sub: pd.DataFrame) -> pd.DataFrame:
    sub = add_action_order(sub)
    if '_Start' in sub.columns and sub['_Start'].notna().any():
        return sub.sort_values(by=['_Start', '__ActionOrder', '_orig_idx'], kind='mergesort')
    return sub.sort_values(by=['__ActionOrder', '_orig_idx'], kind='mergesort')

# ===================== Build relations =====================
relations = []

# --- Floors ordered ---
unique_floors = sorted([t for t in df['FloorToken'].unique() if isinstance(t, str) and t != ''], key=floor_sort_key)
print(f"🏢 Detected Floors (ordered): {unique_floors}")

# Prefer GF as "first floor" if present (SOG lives here), otherwise L1, otherwise the first token
def choose_first_floor_token(tokens):
    if 'GF' in tokens:
        return 'GF'
    if 'L1' in tokens:
        return 'L1'
    levels = [t for t in tokens if t.startswith('L') and t[1:].isdigit()]
    if levels:
        return sorted(levels, key=lambda t: int(t[1:]))[0]
    return tokens[0] if tokens else ''

first_floor_token = choose_first_floor_token(unique_floors)
print(f"🔰 First floor token used for foundation→columns: {first_floor_token or 'N/A'}")

# ---- Part 1: داخل نفس الدور (Columns → Slabs) ----
for floor in unique_floors:
    sub_floor = df[df['FloorToken'] == floor]
    cols = sub_floor[sub_floor['Component'] == 'column'].copy()
    slbs = sub_floor[sub_floor['Component'] == 'slab'  ].copy()

    if cols.empty or slbs.empty:
        print(f"⚠️ Floor {floor}: No columns or slabs found.")
        continue

    cols_sorted = sort_within_group(cols)
    slbs_sorted = sort_within_group(slbs)

    # آخر أعمدة (يفضّل deshuttering)
    dcols = cols_sorted[cols_sorted['Action'] == 'deshuttering']
    last_column = dcols.iloc[-1] if not dcols.empty else cols_sorted.iloc[-1]

    # أول بلاطات (يفضّل steelfixing كأول خطوة منطقية للبلاطة)
    sfix = slbs_sorted[slbs_sorted['Action'] == 'steelfixing']
    first_slab = sfix.iloc[0] if not sfix.empty else slbs_sorted.iloc[0]

    relations.append({
        "Predecessor ID": last_column['Activity ID'],
        "Predecessor Name": last_column['Activity Name'],
        "Successor ID": first_slab['Activity ID'],
        "Successor Name": first_slab['Activity Name'],
        "Relation": "FS",
        "Lag": 0
    })
    print(f"🔗 {floor}: COL(last='{last_column['Activity Name']}') → SLAB(first='{first_slab['Activity Name']}')")

# ---- Part 2: بين الأدوار (Slabs current → Columns next) ----
for i in range(len(unique_floors) - 1):
    current_floor = unique_floors[i]
    next_floor    = unique_floors[i + 1]

    slbs_curr = df[(df['FloorToken'] == current_floor) & (df['Component'] == 'slab')].copy()
    cols_next = df[(df['FloorToken'] == next_floor)    & (df['Component'] == 'column')].copy()

    if slbs_curr.empty or cols_next.empty:
        continue

    slbs_curr_sorted = sort_within_group(slbs_curr)
    cols_next_sorted = sort_within_group(cols_next)

    # آخر بلاطات (يفضّل deshuttering)
    dslab = slbs_curr_sorted[slbs_curr_sorted['Action'] == 'deshuttering']
    last_slab = dslab.iloc[-1] if not dslab.empty else slbs_curr_sorted.iloc[-1]

    # ✅ تفضيل SHUTTERING كأول أعمدة في الدور التالي (حسب طلبك السابق)
    scol_shut = cols_next_sorted[cols_next_sorted['Action'] == 'shuttering']
    first_col = scol_shut.iloc[0] if not scol_shut.empty else cols_next_sorted.iloc[0]

    relations.append({
        "Predecessor ID": last_slab['Activity ID'],
        "Predecessor Name": last_slab['Activity Name'],
        "Successor ID": first_col['Activity ID'],
        "Successor Name": first_col['Activity Name'],
        "Relation": "FS",
        "Lag": 0
    })
    print(f"↗️  {current_floor} SLAB(last='{last_slab['Activity Name']}') → {next_floor} COL(first='{first_col['Activity Name']}') [pref=shuttering]")

# ---- Part 3: RC FOUNDATION (last) → First Columns in FIRST floor ----
foundations = df[(df['Component'] == 'foundation') | (df['Activity Name'].str.contains(r'\brc\s*foundation\b', na=False))].copy()

if not foundations.empty and first_floor_token:
    f_sorted = sort_within_group(foundations)
    dfound = f_sorted[f_sorted['Action'] == 'deshuttering']
    last_foundation = dfound.iloc[-1] if not dfound.empty else f_sorted.iloc[-1]

    cols_first = df[(df['FloorToken'] == first_floor_token) & (df['Component'] == 'column')].copy()
    if not cols_first.empty:
        cols_first_sorted = sort_within_group(cols_first)
        # هنا فضلنا steelfixing كبداية أعمدة في أول دور (تقدر تغيّرها لـ shuttering لو حابب)
        scol = cols_first_sorted[cols_first_sorted['Action'] == 'steelfixing']
        first_col = scol.iloc[0] if not scol.empty else cols_first_sorted.iloc[0]

        relations.append({
            "Predecessor ID": last_foundation['Activity ID'],
            "Predecessor Name": last_foundation['Activity Name'],
            "Successor ID": first_col['Activity ID'],
            "Successor Name": first_col['Activity Name'],
            "Relation": "FS",
            "Lag": 0
        })
        print(f"🏗️  FOUNDATION(last='{last_foundation['Activity Name']}') → {first_floor_token} COL(first='{first_col['Activity Name']}')")
    else:
        print(f"ℹ️ No columns found in first floor '{first_floor_token}' — skipping foundation→columns link.")
else:
    print("ℹ️ No RC FOUNDATION group detected or no first floor — skipping foundation→columns link.")

# ===================== Save output =====================
output_file = filedialog.asksaveasfilename(
    title="Save Relationships As",
    defaultextension=".xlsx",
    filetypes=[("Excel Files", "*.xlsx")]
)
if not output_file:
    print("❌ No output file selected. Exiting...")
    raise SystemExit

relations_df = pd.DataFrame(relations)
relations_df.to_excel(output_file, index=False)
print(f"✅ Created {len(relations)} relationships.")
print(f"💾 Saved to: {output_file}")

# ===================== Diagnostics =====================
if relations_df.empty:
    print("ℹ️ No relationships were created. Check floor parsing or component/action detection.")
else:
    per_floor = {}
    for f in unique_floors:
        per_floor[f] = {
            "col_in_floor": int(((df['FloorToken'] == f) & (df['Component'] == 'column')).sum()),
            "slab_in_floor": int(((df['FloorToken'] == f) & (df['Component'] == 'slab')).sum()),
        }
    print("📊 Summary per floor:", per_floor)
    print(f"🔎 SOG detected rows: {int(sog_mask.sum())}")
