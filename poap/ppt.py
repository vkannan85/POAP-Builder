from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
import pandas as pd
NAVY=RGBColor(20,48,73);BLUE=RGBColor(28,100,160);PALE=RGBColor(245,248,251);LINE=RGBColor(210,218,226);GREY=RGBColor(105,115,125);WHITE=RGBColor(255,255,255);LIGHT=RGBColor(247,249,251);MUTED=RGBColor(220,226,231)
RAG={'GREEN':RGBColor(54,142,74),'AMBER':RGBColor(224,155,32),'RED':RGBColor(196,62,62)}
def _text(s,x,y,w,h,t,z=10,b=False,c=NAVY,a=None):
 q=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));f=q.text_frame;f.clear();f.margin_left=f.margin_right=Inches(.02);p=f.paragraphs[0];p.text=str(t);p.font.size=Pt(z);p.font.bold=b;p.font.color.rgb=c
 if a is not None:p.alignment=a
 return q
def _shape(s,k,x,y,w,h,fill,line=LINE):
 q=s.shapes.add_shape(k,Inches(x),Inches(y),Inches(w),Inches(h));q.fill.solid();q.fill.fore_color.rgb=fill;q.line.color.rgb=line;return q
def _dt(v):
 if v is None or str(v) in ('','None','NaT'):return None
 try:return pd.to_datetime(v)
 except:return None
def _date(v):
 d=_dt(v);return d.strftime('%d %b %Y') if d is not None else 'N/A'
def _ms(d):
 raw=d.get('milestones',[]);out=[]
 if hasattr(raw,'iterrows'):
  for _,r in raw.iterrows():
   x=r.get('finish') if pd.notna(r.get('finish')) else r.get('start');out.append({'task':str(r.get('task','')),'date':_dt(x),'rag':str(r.get('rag','GREEN')).upper()})
 else:
  for r in raw:out.append({'task':str(r.get('task','')),'date':_dt(r.get('date')),'rag':str(r.get('rag','GREEN')).upper()})
 return [x for x in out if x['date'] is not None]
def _box(s,x,y,w,h,title,body=''):
 _shape(s,MSO_SHAPE.ROUNDED_RECTANGLE,x,y,w,h,PALE);_text(s,x+.16,y+.10,w-.32,.28,title,11,True,BLUE)
 if body:_text(s,x+.16,y+.44,w-.32,h-.52,body,9)
def _metric(s,x,y,w,l,v):
 _shape(s,MSO_SHAPE.ROUNDED_RECTANGLE,x,y,w,.72,PALE);_text(s,x+.12,y+.08,w-.24,.22,l.upper(),7,True,BLUE);_text(s,x+.12,y+.30,w-.24,.30,v,15,True)
def _timeline(s,d,x=.45,y=6.02,w=12.43):
 _shape(s,MSO_SHAPE.ROUNDED_RECTANGLE,x,y,w,.88,PALE);_text(s,x+.16,y+.08,1,.2,'TIMELINE',9,True,BLUE);st=_dt(d.get('start'));en=_dt(d.get('finish'));ms=_ms(d)
 if st is None or en is None or en<=st:return
 L=x+.35;R=x+w-.35;ly=y+.49;span=(en-st).total_seconds();_shape(s,MSO_SHAPE.RECTANGLE,L,ly,R-L,.025,LINE,LINE)
 for m in ms[:10]:
  px=L+(R-L)*max(0,min(1,(m['date']-st).total_seconds()/span));_shape(s,MSO_SHAPE.DIAMOND,px-.055,ly-.055,.11,.11,RAG.get(m['rag'],BLUE),WHITE)
def _roadmap(s,d):
 st=_dt(d.get('start'));en=_dt(d.get('finish'));streams=(d.get('workstreams') or ['Project Delivery'])[:8];ms=_ms(d);title=str(d['project']).replace('_',' ');rag=str(d.get('status',d.get('overall_rag','GREEN'))).upper();complete=d.get('completion',d.get('average_complete',0))
 _text(s,.55,.28,8.0,.32,title,20,True);_text(s,.55,.66,5,.18,'EXECUTIVE DELIVERY ROADMAP',8,True,BLUE)
 _shape(s,MSO_SHAPE.ROUNDED_RECTANGLE,9.15,.25,1.35,.52,PALE,LINE);_text(s,9.28,.35,1.1,.12,f'{complete}% COMPLETE',8,True,NAVY,PP_ALIGN.CENTER)
 _shape(s,MSO_SHAPE.ROUNDED_RECTANGLE,10.68,.25,1.15,.52,RAG.get(rag,GREY),RAG.get(rag,GREY));_text(s,10.78,.35,.95,.12,rag,8,True,WHITE,PP_ALIGN.CENTER)
 _text(s,11.95,.37,.8,.12,f'{_date(st)} → {_date(en)}',5.5,False,GREY,PP_ALIGN.RIGHT)
 if st is None or en is None or en<=st:return
 L=2.35;R=10.65;top=1.48;bottom=6.55;rh=(bottom-top)/len(streams);span=(en-st).total_seconds();months=list(pd.date_range(st.replace(day=1),en,freq='MS'));step=max(1,int(len(months)/8)+1);marks=months[::step]
 _text(s,.55,1.15,1.55,.18,'WORKSTREAM',7,True,GREY);_text(s,L,1.15,R-L,.18,'DELIVERY WINDOW',7,True,GREY)
 for m in marks:
  p=max(0,min(1,(m-st).total_seconds()/span));x=L+(R-L)*p;_text(s,x-.25,1.15,.5,.16,m.strftime('%b'),6,True,GREY,PP_ALIGN.CENTER)
 # right insight panel
 _shape(s,MSO_SHAPE.ROUNDED_RECTANGLE,10.88,1.18,1.9,5.38,PALE,LINE);_text(s,11.08,1.38,1.5,.18,'NEXT / ATTENTION',9,True,BLUE)
 future=[m for m in ms if m['date']>=pd.Timestamp.today().normalize()][:4]
 yy=1.78
 for m in future:
  _shape(s,MSO_SHAPE.DIAMOND,11.08,yy+.03,.12,.12,RAG.get(m['rag'],BLUE),RAG.get(m['rag'],BLUE));_text(s,11.30,yy,1.22,.28,m['task'][:28],6.5,True,NAVY);_text(s,11.30,yy+.25,1.1,.13,m['date'].strftime('%d %b'),5.5,False,GREY);yy+=.72
 risks=d.get('risks',[])[:2]
 if risks:_text(s,11.08,4.85,1.45,.16,'KEY EXCEPTIONS',7,True,RAG['RED']);_text(s,11.08,5.12,1.48,1.0,'\n'.join('• '+str(r)[:48] for r in risks),5.8,False,NAVY)
 # clean swimlanes and milestone bars
 for i,ws in enumerate(streams):
  y=top+i*rh;cy=y+rh/2;_text(s,.55,cy-.08,1.55,.16,ws[:25],7.5,True,NAVY);_shape(s,MSO_SHAPE.RECTANGLE,L,cy-.015,R-L,.03,MUTED,MUTED)
  rowms=[m for j,m in enumerate(ms) if j%len(streams)==i][:3]
  for m in rowms:
   p=max(0,min(1,(m['date']-st).total_seconds()/span));x=L+(R-L)*p;col=RAG.get(m['rag'],BLUE);_shape(s,MSO_SHAPE.DIAMOND,x-.07,cy-.07,.14,.14,col,WHITE);_text(s,max(L,min(x-.38,R-.8)),cy-.29,.76,.14,m['task'][:18],5.2,False,NAVY,PP_ALIGN.CENTER)
 today=pd.Timestamp.today().normalize()
 if st<=today<=en:
  x=L+(R-L)*(today-st).total_seconds()/span;_shape(s,MSO_SHAPE.RECTANGLE,x,top-.12,.014,bottom-top+.18,BLUE,BLUE);_text(s,x-.2,top-.30,.4,.12,'TODAY',5.5,True,BLUE,PP_ALIGN.CENTER)
 _text(s,.55,6.90,12.0,.14,'Source-controlled dates • Executive roadmap view • PowerPoint elements remain editable',6,False,GREY,PP_ALIGN.RIGHT)
def generate(d,out):
 prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5);s=prs.slides.add_slide(prs.slide_layouts[6]);_text(s,.45,.25,10.6,.45,d['project'],24,True);_text(s,.45,.73,8,.25,'PLAN ON A PAGE • EXECUTIVE VIEW',9,True,BLUE);rag=str(d.get('overall_rag',d.get('status','GREEN'))).upper();rag=rag if rag in RAG else 'GREEN';_shape(s,MSO_SHAPE.ROUNDED_RECTANGLE,11.55,.28,1.3,.55,RAG[rag],RAG[rag]);_text(s,11.55,.41,1.3,.22,rag,10,True,WHITE,PP_ALIGN.CENTER);tasks=d.get('task_count',d.get('tasks',0));comp=d.get('average_complete',d.get('completion',0));deps=d.get('dependencies',0);rc=d.get('rag_counts',{});_metric(s,.45,1.15,2,'Tasks',str(tasks));_metric(s,2.58,1.15,2,'Average complete',f'{comp}%');_metric(s,4.71,1.15,2.25,'Plan start',_date(d.get('start')));_metric(s,7.09,1.15,2.25,'Plan finish',_date(d.get('finish')));_metric(s,9.47,1.15,1.65,'Dependencies',str(deps));_metric(s,11.25,1.15,1.6,'R / A / G',f"{rc.get('RED',0)} / {rc.get('AMBER',0)} / {rc.get('GREEN',0)}");_box(s,.45,2.10,3.15,2.08,'WORKSTREAMS','\n'.join('• '+x for x in d.get('workstreams',[])));md=_ms(d);_box(s,3.75,2.10,5.65,2.08,'KEY MILESTONES','\n'.join(f"• {m['date'].strftime('%d %b %y')}   {m['task']}" for m in md[:8]) or 'No dated milestones detected');_box(s,9.55,2.10,3.33,2.08,'STATUS',f"• Green: {rc.get('GREEN',0)}\n• Amber: {rc.get('AMBER',0)}\n• Red: {rc.get('RED',0)}\n• Overall: {rag}");_box(s,.45,4.36,6.15,1.48,'RISKS / ATTENTION','\n'.join('• '+x for x in d.get('risks',[])) or '• No rule-based risks detected');_box(s,6.75,4.36,6.13,1.48,'EXECUTIVE SUMMARY',d.get('executive_summary') or f'{tasks} activities. Average recorded completion is {comp}%.');_timeline(s,d);_text(s,.45,7.08,12.4,.18,'Generated by POAP Builder • All PowerPoint shapes and text remain editable',7,False,GREY,PP_ALIGN.RIGHT);s2=prs.slides.add_slide(prs.slide_layouts[6]);_roadmap(s2,d);prs.save(out)