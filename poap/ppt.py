from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

NAVY=RGBColor(20,48,73); BLUE=RGBColor(28,100,160); PALE=RGBColor(245,248,251); LINE=RGBColor(210,218,226)
RAG={'GREEN':RGBColor(54,142,74),'AMBER':RGBColor(224,155,32),'RED':RGBColor(196,62,62)}

def _text(slide,x,y,w,h,text,size=10,bold=False,color=NAVY,align=None):
    s=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=s.text_frame; tf.clear(); tf.word_wrap=True
    p=tf.paragraphs[0]; p.text=str(text); p.font.size=Pt(size); p.font.bold=bold; p.font.color.rgb=color
    if align is not None: p.alignment=align
    return s

def _box(slide,x,y,w,h,title,body=''):
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h)); s.fill.solid(); s.fill.fore_color.rgb=PALE; s.line.color.rgb=LINE
    _text(slide,x+.16,y+.10,w-.32,.28,title,11,True,BLUE)
    if body: _text(slide,x+.16,y+.44,w-.32,h-.52,body,9,False,NAVY)
    return s

def _metric(slide,x,y,w,label,value):
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(.72)); s.fill.solid(); s.fill.fore_color.rgb=PALE; s.line.color.rgb=LINE
    _text(slide,x+.12,y+.08,w-.24,.22,label.upper(),7,True,BLUE); _text(slide,x+.12,y+.30,w-.24,.30,value,15,True,NAVY)

def generate(summary,out_path):
    prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5); slide=prs.slides.add_slide(prs.slide_layouts[6])
    _text(slide,.45,.25,10.6,.45,summary['project'],24,True,NAVY); _text(slide,.45,.73,8,.25,'PLAN ON A PAGE • EXECUTIVE VIEW',9,True,BLUE)
    rag=summary['overall_rag']; badge=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(11.55),Inches(.28),Inches(1.3),Inches(.55)); badge.fill.solid(); badge.fill.fore_color.rgb=RAG[rag]; badge.line.fill.background()
    _text(slide,11.55,.41,1.3,.22,rag,10,True,RGBColor(255,255,255),PP_ALIGN.CENTER)
    sd=summary['start'].strftime('%d %b %Y') if summary['start'] is not None else 'N/A'; ed=summary['finish'].strftime('%d %b %Y') if summary['finish'] is not None else 'N/A'
    _metric(slide,.45,1.15,2.0,'Tasks',str(summary['task_count'])); _metric(slide,2.58,1.15,2.0,'Average complete',f"{summary['average_complete']}%")
    _metric(slide,4.71,1.15,2.25,'Plan start',sd); _metric(slide,7.09,1.15,2.25,'Plan finish',ed); _metric(slide,9.47,1.15,1.65,'Dependencies',str(summary['dependencies']))
    _metric(slide,11.25,1.15,1.6,'R / A / G',f"{summary['rag_counts']['RED']} / {summary['rag_counts']['AMBER']} / {summary['rag_counts']['GREEN']}")
    _box(slide,.45,2.10,3.15,2.08,'WORKSTREAMS','\n'.join('• '+x for x in summary['workstreams']))
    ms=[]
    for _,r in summary['milestones'].iterrows():
        d=r['finish'] if r['finish'] is not None else r['start']; ds=d.strftime('%d %b %y') if getattr(d,'strftime',None) else '—'; ms.append(f"• {ds}   {r['task']}")
    _box(slide,3.75,2.10,5.65,2.08,'KEY MILESTONES','\n'.join(ms[:8]) or 'No dated milestones detected')
    _box(slide,9.55,2.10,3.33,2.08,'STATUS','\n'.join([f"• Green: {summary['rag_counts']['GREEN']}",f"• Amber: {summary['rag_counts']['AMBER']}",f"• Red: {summary['rag_counts']['RED']}",f"• Overall: {summary['overall_rag']}"]))
    _box(slide,.45,4.36,6.15,1.48,'RISKS / ATTENTION','\n'.join('• '+x for x in summary['risks']) or '• No rule-based risks detected')
    _box(slide,6.75,4.36,6.13,1.48,'EXECUTIVE SUMMARY',f"{summary['task_count']} activities across {len(summary['workstreams'])} detected workstream(s). Average recorded completion is {summary['average_complete']}%. Review red/amber items and dependencies before publishing the POAP.")
    _box(slide,.45,6.02,12.43,.88,'TIMELINE',f"{sd}  ─────────────────────────────────────────────────────────  {ed}")
    _text(slide,.45,7.08,12.4,.18,'Generated locally • All PowerPoint shapes and text remain editable',7,False,RGBColor(100,110,120),PP_ALIGN.RIGHT); prs.save(out_path)