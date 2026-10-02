import pandas as pd

def _rag(row,today):
    progress=float(row.get('progress',0) or 0); finish=row.get('finish')
    if progress>=100:return 'GREEN'
    if pd.notna(finish) and finish<today:return 'RED'
    if pd.notna(finish) and finish<=today+pd.Timedelta(days=14):return 'AMBER'
    return 'GREEN'

def analyse(df: pd.DataFrame, project_name: str):
    today=pd.Timestamp.today().normalize();df=df.copy();df['rag']=df.apply(lambda r:_rag(r,today),axis=1)
    # Summary rows are useful for hierarchy but should not inflate task KPIs or roadmap activities.
    summary_mask=df['summary'].fillna(False) if 'summary' in df else pd.Series(False,index=df.index)
    activities=df[~summary_mask].copy(); valid=activities[activities.start.notna()|activities.finish.notna()].copy()
    start=valid.start.min() if len(valid) else None;finish=valid.finish.max() if len(valid) else None
    milestones=activities[activities.milestone].copy().sort_values(['finish','start']).head(12)
    if len(milestones)<4:milestones=activities[activities.finish.notna()].sort_values('finish').drop_duplicates('finish').head(12)
    ws=[x for x in activities.workstream.dropna().astype(str).unique().tolist() if x and x.lower() not in ('none','nan')][:8]
    if not ws:ws=['Project Delivery']
    activities['workstream']=activities['workstream'].where(activities['workstream'].notna(),'Project Delivery')
    roadmap=[]
    candidates=activities[(activities.start.notna()|activities.finish.notna()) & (activities.workstream.isin(ws)| (activities.workstream=='Project Delivery'))].copy()
    # Prioritise milestones, incomplete work and near-term items while preserving source dates/workstreams.
    candidates['priority']=candidates.milestone.astype(int)*100+(candidates.progress<100).astype(int)*20
    candidates=candidates.sort_values(['priority','finish','start'],ascending=[False,True,True]).head(48)
    for _,r in candidates.iterrows():
        roadmap.append({'task':str(r.task),'start':r.start,'finish':r.finish,'workstream':str(r.workstream),'owner':'' if pd.isna(r.owner) else str(r.owner),'progress':int(max(0,min(100,float(r.progress or 0)))),'rag':str(r.rag),'milestone':bool(r.milestone)})
    overdue=activities[(activities.progress<100)&activities.finish.notna()&(activities.finish<today)]
    no_owner=activities[activities.owner.isna()|(activities.owner.astype(str).str.strip().isin(['','None','nan']))]
    dependencies=activities[activities.dependency.notna()&~activities.dependency.astype(str).str.strip().isin(['','None','nan'])]
    risks=[]
    if len(overdue):risks.append(f'{len(overdue)} incomplete task(s) are past planned finish date')
    if len(no_owner):risks.append(f'{len(no_owner)} task(s) have no owner/resource recorded')
    if len(dependencies):risks.append(f'{len(dependencies)} task(s) have recorded dependencies to monitor')
    rag_counts={k:int((activities.rag==k).sum()) for k in ['GREEN','AMBER','RED']};overall='RED' if rag_counts['RED'] else ('AMBER' if rag_counts['AMBER'] else 'GREEN')
    complete=round(float(activities.progress.clip(0,100).mean()),0) if len(activities) else 0
    return {'project':project_name,'start':start,'finish':finish,'workstreams':ws,'milestones':milestones,'roadmap_items':roadmap,'risks':risks[:5],'task_count':len(activities),'rag_counts':rag_counts,'overall_rag':overall,'average_complete':int(complete),'dependencies':len(dependencies)}