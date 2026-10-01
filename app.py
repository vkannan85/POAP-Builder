from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from pathlib import Path
import shutil, uuid
from parsers.tabular import parse as parse_tabular
from parsers.mpp import parse as parse_mpp
from poap.analyser import analyse
from poap.ppt import generate

ROOT=Path(__file__).parent; UP=ROOT/'uploads'; OUT=ROOT/'outputs'; UP.mkdir(exist_ok=True); OUT.mkdir(exist_ok=True)
app=FastAPI(title='POAP Generator')

@app.get('/', response_class=HTMLResponse)
def home(): return (ROOT/'templates'/'index.html').read_text(encoding='utf-8')

@app.post('/api/generate')
async def create(file: UploadFile=File(...)):
    ext=Path(file.filename).suffix.lower()
    if ext not in ['.mpp','.xlsx','.xlsm','.csv','.txt']: raise HTTPException(400,'Use MPP, XLSX, CSV or TXT.')
    token=uuid.uuid4().hex[:8]; src=UP/f'{token}{ext}'
    with src.open('wb') as f: shutil.copyfileobj(file.file,f)
    try:
        df=parse_mpp(src) if ext=='.mpp' else parse_tabular(src)
        summary=analyse(df, Path(file.filename).stem)
        out=OUT/f'{Path(file.filename).stem}_POAP_{token}.pptx'; generate(summary,out)
        preview={'project':summary['project'],'tasks':summary['task_count'],'start':str(summary['start'].date()) if summary['start'] is not None else None,'finish':str(summary['finish'].date()) if summary['finish'] is not None else None,'workstreams':summary['workstreams'],'risks':summary['risks'],'status':summary['overall_rag'],'completion':summary['average_complete'],'dependencies':summary['dependencies'],'download':f'/download/{out.name}'}
        return preview
    except Exception as e: raise HTTPException(422,str(e))

@app.get('/download/{name}')
def download(name:str):
    p=OUT/Path(name).name
    if not p.exists(): raise HTTPException(404,'Output not found')
    return FileResponse(p, filename=p.name)