from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
import pandas as pd

NAVY=RGBColor(20,48,73); BLUE=RGBColor(28,100,160); PALE=RGBColor(245,248,251); LINE=RGBColor(210,218,226); GREY=RGBColor(105,115,125); WHITE=RGBColor(255,255,255); LIGHTBLUE=RGBColor(222,238,249)
RAG={'GREEN':RGBColor(54,142,74),'AMBER':RGBColor(224,155,32),'RED':RGBColor(196,62,62)}
def _text(slide,x,y,w,h,text,size=10,bold=False,color=NAVY,align=None):
 s=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));p=s.text_frame.paragraphs[0];p.text=str(text);p.font.size=Pt(size);p.font.bold=bold;p.font.color.rgb=color
 if align is not None:p.alignment=align
 return s
def _shape(slide,kind,x,y,w,h,fill,line=LINE):
 s=slide.shapes.add_shape(kind,Inches(x),Inches(y),Inches(w),Inches(h));s.fill.solid();s.fill.fore_color.rgb=fill;s.line.color.rgb=line;return s
def _box(slide,x,y,w,h,title,body=''):
 _shape(slide,MSO_SHAPE.ROUNDED_RECTANGLE,x,y,w,h,PALE);_text(slide,x+.16,y+.10,w-.32,.28,title,11,True,BLUE)
 if body:_text(slide,x+.16,y+.44,w-.32,h-.52,body,9)
def _metric(slide,x,y,w,label,value):
 _shape(slide,MSO_SHAPE.ROUNDED_RECTANGLE,x,y,w,.72,PALE);_text(slide,x+.12,y+.08,w-.24,.22,label.upper(),7,True,BLUE);_text(slide,x+.12,y+.30,w-.24,.30,value,15,True)
def _dt(v):
 if v is None or str(v) in ('','None','NaT'):return None
 try:return pd.to_datetime(v)
 except:return None
def _date(v):
 d=_dt(v);return d.strftime('%d %b %Y') if d is not None else 'N/A'
def _milestone_data(summary):
 raw=summary.get('milestones',[]);result=[]
 if hasattr(raw,'iterrows'):
  for _,r in raw.iterrows():
   d=r.get('finish') if pd.notna(r.get('finish')) else r.get('start');result.append({'task':str(r.get('task','')),'date':_dt(d),'rag':str(r.get('rag','GREEN')).upper(),'owner':str(r.get('owner','') or '')})
 else:
  for r in raw:result.append({'task':str(r.get('task','')),'date':_dt(r.get('date')),'rag':str(r.get('rag','GREEN')).upper(),'owner':str(r.get('owner','') or '')})
 return [m for m in result if m['date'] is not None]
def _timeline(slide,summary,x=.45,y=6.02,w=12.43,h=.88):
 _shape(slide,MSO_SHAPE.ROUNDED_RECTANGLE,x,y,w,h,PALE);_text(slide,x+.16,y+.08,1,.2,'TIMELINE',9,True,BLUE);start=_dt(summary.get('start'));finish=_dt(summary.get('finish'));ms=_milestone_data(summary)
 if start is None or finish is None or finish<=start:return
 left=x+.35;right=x+w-.35;ly=y+.49;span=(finish-start).total_seconds();_shape(slide,MSO_SHAPE.RECTANGLE,left,ly,right-left,.025,LINE,LINE)
 months=pd.date_range(start=start.replace(day=1),end=finish,freq='MS');step=max(1,int(len(months)/7)+1)
 for d in months[::step]:
  p=max(0,min(1,(d-start).total_seconds()/span));px=left+(right-left)*p;_shape(slide,MSO_SHAPE.RECTANGLE,px,ly-.05,.012,.12,GREY,GREY);_text(slide,px-.25,y+.64,.55,.15,d.strftime('%b %y'),6,False,GREY,PP_ALIGN.CENTER)
 today=pd.Timestamp.today().normalize()
 if start<=today<=finish:
  px=left+(right-left)*(today-start).total_seconds()/span;_shape(slide,MSO_SHAPE.RECTANGLE,px,y+.27,.018,.39,BLUE,BLUE);_text(slide,px-.22,y+.15,.46,.12,'TODAY',5,True,BLUE,PP_ALIGN.CENTER)
 for i,m in enumerate(ms[:10]):
  px=left+(right-left)*max(0,min(1,(m['date']-start).total_seconds()/span));_shape(slide,MSO_SHAPE.DIAMOND,px-.055,ly-.055,.11,.11,RAG.get(m['rag'],BLUE),WHITE);label=m['task'][:22]+('…' if len(m['task'])>22 else '');_text(slide,max(x+.1,min(px-.48,x+w-1.05)),y+(.24 if i%2==0 else .52),.96,.18,label,5,False,NAVY,PP_ALIGN.CENTER)
def _roadmap(slide,summary):
 start=_dt(summary.get('start'));finish=_dt(summary.get('finish'));streams=(summary.get('workstreams') or ['Project Delivery'])[:8];ms=_milestone_data(summary)
 _text(slide,.42,.18,8.5,.42,summary['project']+' — Delivery Roadmap',22,True,NAVY);_text(slide,9.2,.25,.85,.2,'RAG',8,True,GREY)
 for i,(lab,col) in enumerate([('R',RAG['RED']),('A',RAG['AMBER']),('G',RAG['GREEN']),('TBC',GREY)]):
  _shape(slide,MSO_SHAPE.CHEVRON,9.75+i*.72,.16,.72,.42,col,col);_text(slide,9.75+i*.72,.28,.65,.14,lab,7,True,WHITE,PP_ALIGN.CENTER)
 if start is None or finish is None or finish<=start:return
 left=2.55;right=12.92;top=1.15;rowh=.68;span=(finish-start).total_seconds();months=pd.date_range(start=start.replace(day=1),end=finish,freq='MS');step=max(1,int(len(months)/12)+1);marks=list(months[::step]);
 _shape(slide,MSO_SHAPE.RECTANGLE,left,.72,right-left,.36,NAVY,NAVY);_text(slide,left,.80,right-left,.16,f'{start.year}' if start.year==finish.year else f'{start.year} — {finish.year}',10,True,WHITE,PP_ALIGN.CENTER)
 for d in marks:
  p=max(0,min(1,(d-start).total_seconds()/span));x=left+(right-left)*p;_text(slide,x-.3,1.00,.6,.16,d.strftime('%b'),7,True,GREY,PP_ALIGN.CENTER)
 for i,ws in enumerate(streams):
  y=top+i*rowh;_shape(slide,MSO_SHAPE.RECTANGLE,left,y,right-left,rowh,LIGHTBLUE,WHITE);_shape(slide,MSO_SHAPE.OVAL,.48,y+.16,.34,.34,RAG['GREEN'],RAG['GREEN']);_text(slide,.48,y+.23,.34,.13,str(i+1),8,True,WHITE,PP_ALIGN.CENTER);_shape(slide,MSO_SHAPE.ROUNDED_RECTANGLE,.9,y+.10,1.48,.46,WHITE,GREY);_text(slide,1.02,y+.24,1.25,.18,ws[:24],8,True,NAVY)
  for d in marks:
   p=max(0,min(1,(d-start).total_seconds()/span));x=left+(right-left)*p;_shape(slide,MSO_SHAPE.RECTANGLE,x,y,.008,rowh,WHITE,WHITE)
 # Distribute source milestones across workstream rows to create a compact POAP roadmap. No source dates are changed.
 for i,m in enumerate(ms[:24]):
  row=i%len(streams);y=top+row*rowh+.18;p=max(0,min(1,(m['date']-start).total_seconds()/span));x=left+(right-left)*p;w=.85; x=max(left,min(x-w/2,right-w));col=RAG.get(m['rag'],BLUE);_shape(slide,MSO_SHAPE.CHEVRON,x,y,w,.28,col,col);_text(slide,x+.04,y+.09,w-.12,.11,m['task'][:17],5,True,WHITE,PP_ALIGN.CENTER)
 today=pd.Timestamp.today().normalize()
 if start<=today<=finish:
  x=left+(right-left)*(today-start).total_seconds()/span;_shape(slide,MSO_SHAPE.RECTANGLE,x,1.05,.018,rowh*len(streams)+.15,BLUE,BLUE);_text(slide,x-.22,.88,.46,.14,'TODAY',6,True,BLUE,PP_ALIGN.CENTER)
 _text(slide,.45,6.83,12.4,.2,'Roadmap generated from source milestone dates • Source schedule remains authoritative',7,False,GREY,PP_ALIGN.RIGHT)
def generate(summary,out_path):
 prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5);slide=prs.slides.add_slide(prs.slide_layouts[6]);_text(slide,.45,.25,10.6,.45,summary['project'],24,True);_text(slide,.45,.73,8,.25,'PLAN ON A PAGE • EXECUTIVE VIEW',9,True,BLUE);rag=str(summary.get('overall_rag',summary.get('status','GREEN'))).upper();rag=rag if rag in RAG else 'GREEN';_shape(slide,MSO_SHAPE.ROUNDED_RECTANGLE,11.55,.28,1.3,.55,RAG[rag],RAG[rag]);_text(slide,11.55,.41,1.3,.22,rag,10,True,WHITE,PP_ALIGN.CENTER);sd=_date(summary.get('start'));ed=_date(summary.get('finish'));tasks=summary.get('task_count',summary.get('tasks',0));complete=summary.get('average_complete',summary.get('completion',0));deps=summary.get('dependencies',0);rc=summary.get('rag_counts',{});_metric(slide,.45,1.15,2,'Tasks',str(tasks));_metric(slide,2.58,1.15,2,'Average complete',f'{complete}%');_metric(slide,4.71,1.15,2.25,'Plan start',sd);_metric(slide,7.09,1.15,2.25,'Plan finish',ed);_metric(slide,9.47,1.15,1.65,'Dependencies',str(deps));_metric(slide,11.25,1.15,1.6,'R / A / G',f"{rc.get('RED',0)} / {rc.get('AMBER',0)} / {rc.get('GREEN',0)}");_box(slide,.45,2.10,3.15,2.08,'WORKSTREAMS','\n'.join('• '+x for x in summary.get('workstreams',[])));md=_milestone_data(summary);_box(slide,3.75,2.10,5.65,2.08,'KEY MILESTONES','\n'.join(f"• {m['date'].strftime('%d %b %y')}   {m['task']}" for m in md[:8]) or 'No dated milestones detected');_box(slide,9.55,2.10,3.33,2.08,'STATUS','\n'.join([f"• Green: {rc.get('GREEN',0)}",f"• Amber: {rc.get('AMBER',0)}",f"• Red: {rc.get('RED',0)}",f"• Overall: {rag}"]));_box(slide,.45,4.36,6.15,1.48,'RISKS / ATTENTION','\n'.join('• '+x for x in summary.get('risks',[])) or '• No rule-based risks detected');executive=summary.get('executive_summary') or f"{tasks} activities across {len(summary.get('workstreams',[]))} detected workstream(s). Average recorded completion is {complete}%.";_box(slide,6.75,4.36,6.13,1.48,'EXECUTIVE SUMMARY',executive);_timeline(slide,summary);_text(slide,.45,7.08,12.4,.18,'Generated by POAP Builder • All PowerPoint shapes and text remain editable',7,False,GREY,PP_ALIGN.RIGHT)
 slide2=prs.slides.add_slide(prs.slide_layouts[6]);_roadmap(slide2,summary);prs.save(out_path)