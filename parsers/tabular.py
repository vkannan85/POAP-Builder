from pathlib import Path
import pandas as pd
from .common import normalize

def parse(path: Path):
    ext = path.suffix.lower()
    if ext in ['.xlsx','.xlsm']:
        df = pd.read_excel(path)
    elif ext == '.csv':
        df = pd.read_csv(path)
    elif ext == '.txt':
        try: df = pd.read_csv(path, sep=None, engine='python')
        except Exception:
            lines = [x.strip() for x in path.read_text(encoding='utf-8', errors='ignore').splitlines() if x.strip()]
            df = pd.DataFrame({'Task': lines})
    else: raise ValueError('Unsupported tabular format')
    return normalize(df)