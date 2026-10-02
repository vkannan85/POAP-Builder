from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel, Field
from pathlib import Path
import shutil, uuid, pandas as pd
from parsers.tabular import parse as parse_tabular
from parsers.mpp import parse as parse_mpp
from poap.analyser import analyse
from poap.ppt import generate
ROOT=Path(__file__).parent;UP=ROOT/'uploads';OUT=ROOT/'outputs';UP.mkdir(exist_ok=True);OUT.mkdir(exist_ok=True);app=FastAPI(title='POAP Generator')
class MilestoneEdit(BaseModel):
 task:str;date:str|None=None;owner:str|None=None;rag:str='GREEN'
class RoadmapItem(BaseModel):
 task:str;start:str|None=None;finish:str|None=None;workstream:str='Project Delivery';owner:str|None=None;progress:int=0;rag:str='GREEN';milestone:bool=False
class POAPEdit(BaseModel):
 project:str;start:str|None=None;finish:str|None=None;tasks:int=0;completion:int=Field(0,ge=0,le=100);dependencies:int=0;status:str='GREEN';rag_counts:dict=Field(default_factory=dict);workstreams:list[str]=Field(default_factory=list);milestones:list[MilestoneEdit]=Field(default_factory=list);roadmap_items:list[RoadmapItem]=Field(default_factory=list);risks:list[str]=Field(default_factory=list);executive_summary:str=''
@app.get('/',response_class=HTMLResponse)
def home():return (ROOT/'templates'/'index.html').read_text(encoding='utf-8')
def _ds(v):return v.strftime('%Y-%m-%d') if pd.notna(v) and getattr(v,'strftime',None) else None
def _preview(summary):
 milestones=[]
 for _,r in summary['milestones'].iterrows():
  d=r.get('finish') if pd.notna(r.get('finish')) else r.get('start');milestones.append({'task':str(r.get('task','')),'date':_ds(d),'owner':'' if pd.isna(r.get('owner')) else str(r.get('owner')),'rag':str(r.get('rag','GREEN'))})
 roadmap=[{**r,'start':_ds(r.get('start')),'finish':_ds(r.get('finish'))} for r in summary.get('roadmap_items',[])]
 return {'project':summary['project'],'tasks':summary['task_count'],'start':_ds(summary['start']),'finish':_ds(summary['finish']),'workstreams':summary['workstreams'],'milestones':milestones,'roadmap_items':roadmap,'risks':summary['risks'],'status':summary['overall_rag'],'rag_counts':summary['rag_counts'],'completion':summary['average_complete'],'dependencies':summary['dependencies'],'executive_summary':f"{summary['task_count']} activities across {len(summary['workstreams'])} detected workstream(s). Average recorded completion is {summary['average_complete']}%. Review red/amber items and dependencies before publishing the POAP."}
@app.post('/api/analyse')
async def analyse_file(file:UploadFile=File(...)):
 ext=Path(file.filename or '').suffix.lower()
 if ext not in ['.mpp','.xlsx','.xlsm','.csv','.txt']:raise HTTPException(400,'Use MPP, XLSX, XLSM, CSV or TXT.')
 token=uuid.uuid4().hex[:8];src=UP/f'{token}{ext}'
 with src.open('wb') as f:shutil.copyfileobj(file.file,f)
 try:
  df=parse_mpp(src) if ext=='.mpp' else parse_tabular(src);return _preview(analyse(df,Path(file.filename).stem))
 except Exception as e:raise HTTPException(422,str(e))
@app.post('/api/export')
def export_poap(data:POAPEdit):
 try:
  payload=data.model_dump();token=uuid.uuid4().hex[:8];safe=''.join(c for c in data.project if c.isalnum() or c in (' ','-','_')).strip() or 'Project';out=OUT/f'{safe}_POAP_{token}.pptx';generate(payload,out);return {'download':f'/download/{out.name}'}
 except Exception as e:raise HTTPException(422,str(e))
@app.post('/api/generate')
async def create(file:UploadFile=File(...)):
 preview=await analyse_file(file);payload=POAPEdit(**preview);return {**preview,**export_poap(payload)}
@app.get('/download/{name}')
def download(name:str):
 p=OUT/Path(name).name
 if not p.exists():raise HTTPException(404,'Output not found')
 return FileResponse(p,filename=p.name)