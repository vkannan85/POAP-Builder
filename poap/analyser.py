import pandas as pd

def _rag(row, today):
    progress = float(row.get('progress', 0) or 0)
    finish = row.get('finish')
    if progress >= 100: return 'GREEN'
    if pd.notna(finish) and finish < today: return 'RED'
    if pd.notna(finish) and finish <= today + pd.Timedelta(days=14): return 'AMBER'
    return 'GREEN'

def analyse(df: pd.DataFrame, project_name: str):
    today = pd.Timestamp.today().normalize(); df = df.copy()
    df['rag'] = df.apply(lambda r: _rag(r, today), axis=1)
    valid = df[df.start.notna() | df.finish.notna()].copy()
    start = valid.start.min() if len(valid) else None; finish = valid.finish.max() if len(valid) else None
    milestones = df[df.milestone].copy().sort_values(['finish','start']).head(12)
    if len(milestones) < 4: milestones = df[df.finish.notna()].sort_values('finish').drop_duplicates('finish').head(12)
    ws = [x for x in df.workstream.dropna().astype(str).unique().tolist() if x and x.lower() not in ('none','nan')][:8]
    if not ws: ws=['Project Delivery']
    overdue=df[(df.progress<100)&df.finish.notna()&(df.finish<today)]
    no_owner=df[df.owner.isna()|(df.owner.astype(str).str.strip().isin(['','None','nan']))]
    dependencies=df[df.dependency.notna()&~df.dependency.astype(str).str.strip().isin(['','None','nan'])]
    risks=[]
    if len(overdue): risks.append(f'{len(overdue)} incomplete task(s) are past planned finish date')
    if len(no_owner): risks.append(f'{len(no_owner)} task(s) have no owner/resource recorded')
    if len(dependencies): risks.append(f'{len(dependencies)} task(s) have recorded dependencies to monitor')
    rag_counts={k:int((df.rag==k).sum()) for k in ['GREEN','AMBER','RED']}
    overall='RED' if rag_counts['RED'] else ('AMBER' if rag_counts['AMBER'] else 'GREEN')
    complete=round(float(df.progress.clip(0,100).mean()),0) if len(df) else 0
    return {'project':project_name,'start':start,'finish':finish,'workstreams':ws,'milestones':milestones,'risks':risks[:5],'task_count':len(df),'rag_counts':rag_counts,'overall_rag':overall,'average_complete':int(complete),'dependencies':len(dependencies)}