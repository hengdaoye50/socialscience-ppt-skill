"""Deterministic structural and rendering QA, with limitations documented."""
from pathlib import Path
from pptx import Presentation
import re, json, subprocess
from zipfile import ZipFile

DANGEROUS=['lorem ipsum','点击添加','双击此处','在此输入','[请填写]','PLACEHOLDER','XXX','待补充']


def inspect(pptx):
    p=Presentation(str(pptx));slides=[]
    for i,s in enumerate(p.slides,1):
        info={'slide':i,'shapes':len(s.shapes),'charts':[],'tables':[],'text':[]}
        for sh in s.shapes:
            if getattr(sh,'has_chart',False):
                info['charts'].append({'shape_id':sh.shape_id,'type':str(sh.chart.chart_type),'series':len(sh.chart.series)})
            if getattr(sh,'has_table',False):info['tables'].append({'shape_id':sh.shape_id,'rows':len(sh.table.rows),'cols':len(sh.table.columns)})
            if sh.has_text_frame and sh.text.strip():info['text'].append({'shape_id':sh.shape_id,'text':sh.text[:170]})
        slides.append(info)
    return {'slides':slides,'slide_count':len(slides),'chart_count':sum(len(z['charts']) for z in slides),'table_count':sum(len(z['tables']) for z in slides)}

def render(pptx,outdir,dpi=125):
    outdir=Path(outdir);outdir.mkdir(parents=True,exist_ok=True)
    pptx=Path(pptx).resolve();pdf=outdir/(pptx.stem+'.pdf')
    cmd=['soffice','-env:UserInstallation=file:///tmp/academic-ppt-lo-'+str(__import__('os').getpid()),'--headless','--convert-to','pdf','--outdir',str(outdir),str(pptx)]
    cp=subprocess.run(cmd,capture_output=True,text=True,timeout=120)
    if cp.returncode or not pdf.exists():raise RuntimeError(f'LibreOffice render failed: {cp.stdout}\n{cp.stderr}')
    import fitz
    doc=fitz.open(pdf);paths=[]
    for i,p in enumerate(doc,1):
        pix=p.get_pixmap(matrix=fitz.Matrix(dpi/72,dpi/72),alpha=False)
        path=outdir/f'slide-{i:02d}.png';pix.save(path);paths.append(str(path))
    return {'pdf':str(pdf),'images':paths,'slide_count':len(paths)}

def audit(pptx,render_pdf=None,expected=None):
    p=Presentation(str(pptx));issues=[];warning=[]
    w,h=p.slide_width,p.slide_height
    headings=[]
    for i,s in enumerate(p.slides,1):
        for sh in s.shapes:
            if sh.left < -1500 or sh.top < -1500 or sh.left+sh.width > w+1500 or sh.top+sh.height > h+1500:
                issues.append(f'Page {i}: shape {sh.shape_id} out of bounds')
            if sh.has_text_frame:
                txt=sh.text.strip()
                for term in DANGEROUS:
                    if term.lower() in txt.lower():warning.append(f'Page {i}: possible placeholder "{term}"')
                if txt and not sh.text_frame.paragraphs[0].runs:warning.append(f'Page {i}: empty formatted run {sh.shape_id}')
                for pp in sh.text_frame.paragraphs:
                    for run in pp.runs:
                        if run.font.size and run.font.size.pt<8.9 and txt:
                            warning.append(f'Page {i}: tiny text {run.font.size.pt:.1f} pt in shape {sh.shape_id}')
            if getattr(sh,'has_chart',False):
                chart=sh.chart
                if not len(chart.series):issues.append(f'Page {i}: chart has no series')
    if expected:
        for i,sd in enumerate(expected.get('slides',[]),1):
            if sd.get('type') in ('chart','table','estimate','image','evidence') and not sd.get('source'):
                if not (sd.get('type')=='evidence' and sd.get('source')):
                    warning.append(f'Spec slide {i}: missing source/illustrative disclaimer')
            if sd.get('type')=='mechanism':
                for e in sd.get('edges',[]):
                    if 'kind' not in e:warning.append(f'Spec slide {i}: relation {e.get("source")}→{e.get("target")} defaults to causal; mark explicitly')
    if render_pdf:
        import fitz
        doc=fitz.open(str(render_pdf))
        if len(doc)!=len(p.slides):issues.append(f'Render page count {len(doc)} differs from PPTX {len(p.slides)}')
        for i,page in enumerate(doc,1):
            words=page.get_text('words'); rect=page.rect
            for x0,y0,x1,y1,*rest in words:
                # Text that spills beyond the printable canvas is an objective rendering error.
                if x0 < -1 or y0 < -1 or x1 > rect.width+1 or y1>rect.height+1:
                    issues.append(f'Page {i}: PDF text outside page edge')
                    break
            if not words:warning.append(f'Page {i}: no selectable text in PDF render')
    return {'status':'fail' if issues else ('warning' if warning else 'pass'),'error_count':len(issues),'warning_count':len(warning),'errors':issues,'warnings':warning,'slide_count':len(p.slides)}

def slide_xml_hashes(pptx,upto=None):
    import hashlib
    with ZipFile(pptx) as z:
        files=[f for f in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',f)]
        result={f:hashlib.sha256(z.read(f)).hexdigest() for f in files}
    if upto is not None:result={k:v for k,v in result.items() if int(re.search(r'slide(\d+)',k).group(1))<=upto}
    return result
