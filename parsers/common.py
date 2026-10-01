import pandas as pd

ALIASES = {
    'task': ['task','task name','name','activity','activity name','description','title'],
    'start': ['start','start date','planned start','baseline start'],
    'finish': ['finish','end','end date','finish date','planned finish','baseline finish'],
    'owner': ['owner','resource','resource names','assigned to','assignee'],
    'progress': ['% complete','percent complete','progress','complete'],
    'milestone': ['milestone','is milestone'],
    'workstream': ['workstream','phase','summary task','stream','category'],
    'dependency': ['predecessors','dependency','dependencies']
}

def normalize(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    lower = {c.lower(): c for c in df.columns}
    out = pd.DataFrame()
    for key, names in ALIASES.items():
        col = next((lower[n] for n in names if n in lower), None)
        out[key] = df[col] if col else None
    if out['task'].isna().all() and len(df.columns): out['task'] = df.iloc[:,0]
    for c in ['start','finish']:
        out[c] = pd.to_datetime(out[c], errors='coerce')
    out['progress'] = pd.to_numeric(out['progress'].astype(str).str.replace('%','',regex=False), errors='coerce').fillna(0)
    out['milestone'] = out['milestone'].astype(str).str.lower().isin(['true','yes','1','y']) | (out['start'].notna() & (out['start'] == out['finish']))
    out['task'] = out['task'].fillna('').astype(str).str.strip()
    return out[out['task'] != ''].reset_index(drop=True)