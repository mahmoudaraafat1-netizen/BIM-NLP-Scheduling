import subprocess, sys

# 1) Ensure required packages
required_packages = {
    'pandas': 'pandas',
    'openpyxl': 'openpyxl',
    'sentence_transformers': 'sentence-transformers',
    'torch': 'torch',
    'tqdm': 'tqdm',
    'nltk': 'nltk'
}
for module_name, pip_name in required_packages.items():
    try:
        __import__(module_name)
    except ImportError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', pip_name])

# 2) Imports
import pandas as pd
import numpy as np
from sentence_transformers import SentenceTransformer, util
import re
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from tqdm import tqdm
from datetime import datetime
import nltk
from nltk.stem import WordNetLemmatizer
from collections import defaultdict

# 3) NLTK resources (tolerant)
try:
    nltk.download('wordnet', quiet=True)
    nltk.download('omw-1.4', quiet=True)
except Exception:
    pass

# 4) Model & NLP utils
model = SentenceTransformer('sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2')
lemmatizer = WordNetLemmatizer()

# 5) UI: threshold
root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
sim_threshold = simpledialog.askfloat(
    'Similarity Threshold', 'Enter similarity threshold (0–1):',
    minvalue=0.0, maxvalue=1.0, initialvalue=0.4
)
if sim_threshold is None:
    messagebox.showwarning('Cancelled', 'Similarity threshold not set.')
    sys.exit(1)
start_time = datetime.now()

# 6) Synonyms (tokens-level)
synonym_map = {
    'casting': 'pouring',
    'pouring': 'pouring',
    'formwork': 'shuttering',
    'shuttering': 'shuttering',
    'deformwork': 'deshuttering',
    'deshuttering': 'deshuttering',
    'steel-fixing': 'steelfixing',
    'steel': 'steelfixing',
    'rebar': 'steelfixing',
    'reinforcement': 'steelfixing',
    'steel fixing': 'steelfixing'
}

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

def extract_floor(text: str) -> str:
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
        if kind == 'GF': return 'GF'
        if kind == 'ROOF': return 'ROOF'
        if kind == 'PODIUM': return 'PODIUM'
        if kind == 'MEZZ': return 'MEZZ'
        if kind == 'ORDINAL_FLOOR':
            return ORDINAL_MAP.get(m.group(1).lower(), '')
    return ''

def extract_comp(text: str) -> str:
    t = (text or '').lower()
    if 'column' in t: return 'columns'
    if 'slab' in t: return 'slab'
    if 'foundation' in t: return 'foundation'
    return 'other'

def extract_action(text: str) -> str:
    t = (text or '').lower()
    if 'deshuttering' in t or 'deformwork' in t or 'strik' in t: return 'deshuttering'
    if 'shuttering' in t or 'formwork' in t: return 'shuttering'
    if 'steel' in t or 'rebar' in t or 'reinforc' in t or 'fixing' in t: return 'steelfixing'
    if 'pouring' in t or 'casting' in t or 'concret' in t: return 'concrete'
    return 'other'

def clean(text: str) -> str:
    text = (text or '').lower()
    text = re.sub(r'[^a-z0-9\s]', ' ', text)
    tokens = text.split()
    normalized_tokens = [synonym_map.get(token, token) for token in tokens]
    lemmatized_tokens = [lemmatizer.lemmatize(token) for token in normalized_tokens]
    return ' '.join(lemmatized_tokens)

key_terms = list(set(synonym_map.values()))
def has_key_term_match(text1: str, text2: str) -> bool:
    for term in key_terms:
        if re.search(rf'\b{re.escape(term)}\b', text1) and re.search(rf'\b{re.escape(term)}\b', text2):
            return True
    return False

# ----------------------------
# OPTIONAL: blocking rule (set True only if you intentionally exclude some templates)
BLOCK_DESHUTTERING_TEMPLATES = False

def is_blocked_template(pred_clean: str) -> bool:
    if BLOCK_DESHUTTERING_TEMPLATES and re.search(r'\bdeshuttering\b', pred_clean):
        return True
    return False
# ----------------------------

# 7) File selection
act_file = filedialog.askopenfilename(title='Select activity file', filetypes=[('Excel','*.xlsx')])
if not act_file:
    messagebox.showerror('Error', 'No activity file selected.')
    sys.exit(1)

dict_file = filedialog.askopenfilename(title='Select dictionary file', filetypes=[('Excel','*.xlsx')])
if not dict_file:
    messagebox.showerror('Error', 'No dictionary file selected.')
    sys.exit(1)

# Read activities
df_acts = pd.read_excel(act_file)

# Read dictionary from sheet "Relationships" robustly
try:
    df_dict = pd.read_excel(dict_file, sheet_name='Relationships')
except Exception:
    try:
        xl = pd.ExcelFile(dict_file)
        target = None
        for name in xl.sheet_names:
            if name.strip().lower() == 'relationships':
                target = name
                break
        if target is None:
            raise ValueError("Sheet 'Relationships' not found.")
        df_dict = xl.parse(sheet_name=target)
    except Exception as e:
        messagebox.showerror('Missing Sheet', f"Cannot find sheet 'Relationships' in dictionary file.\n{e}")
        sys.exit(1)

# Clean headers (handles hidden BOM too)
df_acts.columns = df_acts.columns.astype(str).str.replace('\ufeff','').str.strip()
df_dict.columns = df_dict.columns.astype(str).str.replace('\ufeff','').str.strip()

# Required columns
for c in ['Activity ID', 'Activity Name']:
    if c not in df_acts.columns:
        messagebox.showerror('Missing Column', f'Missing column: {c}')
        sys.exit(1)

for c in ['Pred Name', 'Succ Name', 'Rel Type', 'Lag']:
    if c not in df_dict.columns:
        messagebox.showerror('Missing Column', f"Missing column in 'Relationships' sheet: {c}")
        sys.exit(1)

# 8) Data prep
acts = df_acts.copy()
acts['Floor'] = acts['Activity Name'].apply(extract_floor)
acts['Component'] = acts['Activity Name'].apply(extract_comp)
acts['Action'] = acts['Activity Name'].apply(extract_action)
acts['Cleaned'] = acts['Activity Name'].apply(clean)

dict_df = df_dict.copy()
dict_df['Pred Comp'] = dict_df['Pred Name'].apply(extract_comp)
dict_df['Succ Comp'] = dict_df['Succ Name'].apply(extract_comp)
dict_df['Pred Clean'] = dict_df['Pred Name'].apply(clean)
dict_df['Succ Clean'] = dict_df['Succ Name'].apply(clean)

# same-component only
dict_df = dict_df[dict_df['Pred Comp'] == dict_df['Succ Comp']].reset_index(drop=True)

print("Encoding activities...")
act_emb = model.encode(acts['Cleaned'].tolist(), convert_to_tensor=True, show_progress_bar=True)
print("Encoding dictionary (pred)...")
pred_emb = model.encode(dict_df['Pred Clean'].tolist(), convert_to_tensor=True, show_progress_bar=True)
print("Caching dictionary (succ)...")
succ_emb_cache = {name: model.encode(clean(name), convert_to_tensor=True) for name in dict_df['Succ Name'].unique()}

results, unmatched, prim = [], [], []
visited_pairs = set()

# Cycle prevention
graph_adj = defaultdict(set)
def _would_create_cycle(adj, u, v):
    stack, seen = [v], set()
    while stack:
        node = stack.pop()
        if node == u:
            return True
        if node in seen:
            continue
        seen.add(node)
        stack.extend(adj[node])
    return False

# Cache embeddings for (floor, succ_component)
group_emb_cache = {}

print("Matching activities...")
n = len(acts)
for pos in tqdm(range(n), total=n):
    row = acts.iloc[pos]
    comp = row['Component']
    floor = row['Floor']
    act_id = row['Activity ID']

    sub = dict_df[dict_df['Pred Comp'] == comp]
    if sub.empty:
        unmatched.append({
            'Activity ID': act_id,
            'Activity': row['Activity Name'],
            'Decision': 'NOT_MATCH',
            'Reason': 'No matching component in dictionary',
            'SBERT_BaseSim_Max': None,
            'Score_Final_Max': None,
            'BlockedByRule': 0,
            'Threshold': sim_threshold
        })
        continue

    emb = pred_emb[sub.index]
    sims = util.pytorch_cos_sim(act_emb[pos], emb)[0].cpu().numpy()
    base_max = float(np.max(sims)) if len(sims) else 0.0

    matches = [(j, float(sims[j])) for j in range(len(sims)) if sims[j] >= sim_threshold]
    matches.sort(key=lambda x: -x[1])

    matched = False
    best_final_seen = -1.0
    any_blocked = False

    for idx, base_score in matches:
        pred_row = sub.iloc[idx]

        # block rule (optional)
        if is_blocked_template(pred_row['Pred Clean']):
            any_blocked = True
            best_final_seen = max(best_final_seen, base_score)
            continue

        key = (floor, pred_row['Succ Comp'])
        if key not in group_emb_cache:
            mask = (acts['Floor'] == floor) & (acts['Component'] == pred_row['Succ Comp'])
            acts_masked = acts[mask].copy()
            group_emb_cache[key] = (
                acts_masked,
                model.encode(acts_masked['Cleaned'].tolist(), convert_to_tensor=True) if not acts_masked.empty else None
            )
        acts_masked, masked_emb = group_emb_cache[key]
        if acts_masked.empty or masked_emb is None:
            continue

        succ_encoded = succ_emb_cache.get(pred_row['Succ Name'])
        sims_succ = util.pytorch_cos_sim(succ_encoded, masked_emb)[0].cpu().numpy()
        best_idx = int(sims_succ.argmax())
        succ_best_sim = float(sims_succ.max()) if len(sims_succ) else 0.0

        sid = acts_masked.iloc[best_idx].name
        suc_id = acts.loc[sid, 'Activity ID']

        final_score = base_score
        boosted = 0
        if has_key_term_match(row['Cleaned'], pred_row['Pred Clean']):
            final_score = 1.0
            boosted = 1

        best_final_seen = max(best_final_seen, final_score)

        if act_id == suc_id or (suc_id, act_id) in visited_pairs:
            continue
        if _would_create_cycle(graph_adj, act_id, suc_id):
            continue

        pair = (act_id, suc_id)
        if pair in visited_pairs:
            continue
        visited_pairs.add(pair)

        results.append({
            'Activity ID': act_id,
            'Activity': row['Activity Name'],
            'Decision': 'MATCH',
            'Matched Pred': pred_row['Pred Name'],
            'SBERT_BaseSim': round(float(base_score), 4),
            'Score_Final': round(float(final_score), 4),
            'Boosted': boosted,
            'SuccSim': round(float(succ_best_sim), 4),
            'Activity ID next activity': suc_id,
            'Next Activity': acts.loc[sid, 'Activity Name'],
            'Relation': pred_row['Rel Type'],
            'Lag': pred_row['Lag'],
            'Threshold': sim_threshold
        })

        prim.append({
            'Activity Predecessor ID': act_id,
            'Activity Predecessor Name': row['Activity Name'],
            'Activity Successor ID': suc_id,
            'Activity Successor Name': acts.loc[sid, 'Activity Name'],
            'Relation': pred_row['Rel Type'],
            'Lag': pred_row['Lag']
        })

        graph_adj[act_id].add(suc_id)
        matched = True
        break

    if not matched:
        unmatched.append({
            'Activity ID': act_id,
            'Activity': row['Activity Name'],
            'Decision': 'NOT_MATCH',
            'Reason': 'No suitable match found',
            'SBERT_BaseSim_Max': round(float(base_max), 4),
            'Score_Final_Max': round(float(best_final_seen), 4) if best_final_seen >= 0 else None,
            'BlockedByRule': int(any_blocked),
            'Threshold': sim_threshold
        })

# 9) Output
res_df = pd.DataFrame(results)
un_df = pd.DataFrame(unmatched)
pm_df = pd.DataFrame(prim)

out = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')])
if not out:
    messagebox.showinfo('Cancelled', 'Save cancelled by user.')
    sys.exit(0)

try:
    with pd.ExcelWriter(out) as w:
        res_df.to_excel(w, 'Matches', index=False)
        un_df.to_excel(w, 'Unmatched', index=False)
        pm_df.to_excel(w, 'ForPrimavera', index=False)
        meta = pd.DataFrame([{
            'Model': 'paraphrase-multilingual-MiniLM-L12-v2',
            'Threshold': sim_threshold,
            'Activities': len(acts),
            'Dict Rows (after filter)': len(dict_df),
            'Started At': start_time.strftime('%Y-%m-%d %H:%M:%S'),
            'Finished At': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'Activity File': act_file,
            'Dictionary File': dict_file,
            'BLOCK_DESHUTTERING_TEMPLATES': BLOCK_DESHUTTERING_TEMPLATES
        }])
        meta.to_excel(w, 'RunInfo', index=False)

    print(f"✅ Done! File saved to: {out}")
    messagebox.showinfo('Done', f'File saved to:\n{out}')
except Exception as e:
    messagebox.showerror('Save Error', f'Failed to save file:\n{e}')
    sys.exit(1)
