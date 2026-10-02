from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
import pandas as pd
NAVY=RGBColor(20,48,73);BLUE=RGBColor(28,100,160);PALE=RGBColor(245,248,251);LINE=RGBColor(210,218,226);GREY=RGBColor(105,115,125);WHITE=RGBColor(255,255,255);LIGHTBLUE=RGBColor(235,243,248)
RAG={'GREEN':RGBColor(54,142,74),'AMBER':RGBColor(224,155,32),'RED':RGBColor(196,62,62)}
def _text(slide,x,y,w,h,text,size=10,bold=False,color=NAVY,align=None):
 s=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));tf=s.text_frame;tf.clear();tf.margin_left=tf.margin_right=Inches(.02);tf.margin_top=tf.margin_bottom=Inches(.01);p=tf.paragraphs[0];p.text=str(text);p.font.size=Pt(size);p.font.bold=bold;p.font.color.rgb=color
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
 left=x+.35;right=x+w-.35;ly=y+.49;span=(finish-start).total_seconds();_shape(slide,MSO_SHAPE.RECTANGLE,left,ly,right-left,.025,LINE,LINE);months=pd.date_range(start=start.replace(day=1),end=finish,freq='MS');step=max(1,int(len(months)/7)+1)
 for d in months[::step]:
  p=max(0,min(1,(d-start).total_seconds()/span));px=left+(right-left)*p;_shape(slide,MSO_SHAPE.RECTANGLE,px,ly-.05,.012,.12,GREY,GREY);_text(slide,px-.25,y+.64,.55,.15,d.strftime('%b %y'),6,False,GREY,PP_ALIGN.CENTER)
 today=pd.Timestamp.today().normalize()
 if start<=today<=finish:
  px=left+(right-left)*(today-start).total_seconds()/span;_shape(slide,MSO_SHAPE.RECTANGLE,px,y+.27,.018,.39,BLUE,BLUE);_text(slide,px-.22,y+.15,.46,.12,'TODAY',5,True,BLUE,PP_ALIGN.CENTER)
 for i,m in enumerate(ms[:10]):
  px=left+(right-left)*max(0,min(1,(m['date']-start).total_seconds()/span));_shape(slide,MSO_SHAPE.DIAMOND,px-.055,ly-.055,.11,.11,RAG.get(m['rag'],BLUE),WHITE);_text(slide,max(x+.1,min(px-.48,x+w-1.05)),y+(.24 if i%2==0 else .52),.96,.18,m['task'][:20],5,False,NAVY,PP_ALIGN.CENTER)
def _roadmap(slide,summary):
 start=_dt(summary.get('start'));finish=_dt(summary.get('finish'));streams=(summary.get('workstreams') or ['Project Delivery'])[:8];ms=_milestone_data(summary);title=str(summary['project']).replace('_',' ')
 _text(slide,.48,.22,8.25,.34,title+' — Delivery Roadmap',18,True,NAVY);_text(slide,9.08,.30,.45,.14,'RAG',7,True,GREY)
 for i,(lab,col) in enumerate([('R',RAG['RED']),('A',RAG['AMBER']),('G',RAG['GREEN']),('TBC',GREY)]):
  x=9.55+i*.73;_shape(slide,MSO_SHAPE.CHEVRON,x,.20,.68,.34,col,col);_text(slide,x+.03,.30,.58,.10,lab,6,True,WHITE,PP_ALIGN.CENTER)
 if start is None or finish is None or finish<=start:return
 left=2.72;right=12.78;header_y=.80;months_y=1.14;top=1.42;bottom=6.62;rowh=(bottom-top)/len(streams);span=(finish-start).total_seconds();months=pd.date_range(start=start.replace(day=1),end=finish,freq='MS');step=max(1,int(len(months)/10)+1);marks=list(months[::step])
 _shape(slide,MSO_SHAPE.RECTANGLE,left,header_y,right-left,.30,NAVY,NAVY);_text(slide,left,header_y+.08,right-left,.11,f'{start.year}' if start.year==finish.year else f'{start.year} — {finish.year}',8,True,WHITE,PP_ALIGN.CENTER)
 for d in marks:
  p=max(0,min(1,(d-start).total_seconds()/span));x=left+(right-left)*p;_text(slide,x-.25,months_y,.5,.13,d.strftime('%b'),6,True,GREY,PP_ALIGN.CENTER)
 for i,ws in enumerate(streams):
  y=top+i*rowh;_shape(slide,MSO_SHAPE.RECTANGLE,left,y,right-left,rowh,LIGHTBLUE,WHITE)
  cy=y+rowh/2;_shape(slide,MSO_SHAPE.OVAL,.55,cy-.13,.26,.26,RAG['GREEN'],RAG['GREEN']);_text(slide,.55,cy-.045,.26,.09,str(i+1),6,True,WHITE,PP_ALIGN.CENTER)
  _shape(slide,MSO_SHAPE.ROUNDED_RECTANGLE,.92,cy-.18,1.58,.36,WHITE,GREY);_text(slide,1.00,cy-.055,1.42,.12,ws[:26],7,True,NAVY,PP_ALIGN.CENTER)
  for d in marks:
   p=max(0,min(1,(d-start).total_seconds()/span));x=left+(right-left)*p;_shape(slide,MSO_SHAPE.RECTANGLE,x,y,.006,rowh,WHITE,WHITE)
 # compact source milestones, aligned to exact source date positions
 for i,m in enumerate(ms[:24]):
  row=i%len(streams);y=top+row*rowh;cy=y+rowh/2;p=max(0,min(1,(m['date']-start).total_seconds()/span));x=left+(right-left)*p;w=.66;h=.22;x=max(left+.02,min(x-w/2,right-w-.02));col=RAG.get(m['rag'],BLUE);_shape(slide,MSO_SHAPE.CHEVRON,x,cy-h/2,w,h,col,col);_text(slide,x+.025,cy-.035,w-.08,.08,m['task'][:14],4.5,True,WHITE,PP_ALIGN.CENTER)
 today=pd.Timestamp.today().normalize()
 if start<=today<=finish:
  x=left+(right-left)*(today-start).total_seconds()/span;_shape(slide,MSO_SHAPE.RECTANGLE,x,top-.12,.014,bottom-top+.12,BLUE,BLUE);_text(slide,x-.20,top-.28,.40,.11,'TODAY',5.5,True,BLUE,PP_ALIGN.CENTER)
 _text(slide,.55,6.88,12.1,.15,'Roadmap generated from source milestone dates • Source schedule remains authoritative',6,False,GREY,PP_ALIGN.RIGHT)
def generate(summary,out_path):
 prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5);slide=prs.slides.add_slide(prs.slide_layouts[6]);_text(slide,.45,.25,10.6,.45,summary['project'],24,True);_text(slide,.45,.73,8,.25,'PLAN ON A PAGE • EXECUTIVE VIEW',9,True,BLUE);rag=str(summary.get('overall_rag',summary.get('status','GREEN'))).upper();rag=rag if rag in RAG else 'GREEN';_shape(slide,MSO_SHAPE.ROUNDED_RECTANGLE,11.55,.28,1.3,.55,RAG[rag],RAG[rag]);_text(slide,11.55,.41,1.3,.22,rag,10,True,WHITE,PP_ALIGN.CENTER);sd=_date(summary.get('start'));ed=_date(summary.get('finish'));tasks=summary.get('task_count',summary.get('tasks',0));complete=summary.get('average_complete',summary.get('completion',0));deps=summary.get('dependencies',0);rc=summary.get('rag_counts',{});_metric(slide,.45,1.15,2,'Tasks',str(tasks));_metric(slide,2.58,1.15,2,'Average complete',f'{complete}%');_metric(slide,4.71,1.15,2.25,'Plan start',sd);_metric(slide,7.09,1.15,2.25,'Plan finish',ed);_metric(slide,9.47,1.15,1.65,'Dependencies',str(deps));_metric(slide,11.25,1.15,1.6,'R / A / G',f"{rc.get('RED',0)} / {rc.get('AMBER',0)} / {rc.get('GREEN',0)}");_box(slide,.45,2.10,3.15,2.08,'WORKSTREAMS','\n'.join('• '+x for x in summary.get('workstreams',[])));md=_milestone_data(summary);_box(slide,3.75,2.10,5.65,2.08,'KEY MILESTONES','\n'.join(f"• {m['date'].strftime('%d %b %y')}   {m['task']}" for m in md[:8]) or 'No dated milestones detected');_box(slide,9.55,2.10,3.33,2.08,'STATUS','\n'.join([f"• Green: {rc.get('GREEN',0)}",f"• Amber: {rc.get('AMBER',0)}",f"• Red: {rc.get('RED',0)}",f"• Overall: {rag}"]));_box(slide,.45,4.36,6.15,1.48,'RISKS / ATTENTION','\n'.join('• '+x for x in summary.get('risks',[])) or '• No rule-based risks detected');executive=summary.get('executive_summary') or f"{tasks} activities across {len(summary.get('workstreams',[]))} detected workstream(s). Average recorded completion is {complete}%.";_box(slide,6.75,4.36,6.13,1.48,'EXECUTIVE SUMMARY',executive);_timeline(slide,summary);_text(slide,.45,7.08,12.4,.18,'Generated by POAP Builder • All PowerPoint shapes and text remain editable',7,False,GREY,PP_ALIGN.RIGHT)
 slide2=prs.slides.add_slide(prs.slide_layouts[6]);_roadmap(slide2,summary);prs.save(out_path)