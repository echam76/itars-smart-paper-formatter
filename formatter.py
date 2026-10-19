"""Conservative Word formatter: edits styles in place, never reconstructs the body.
This helps preserve drawings, tables, field codes, hyperlinks and equations.
"""
import io, re, zipfile
from collections import Counter
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docx.enum.style import WD_STYLE_TYPE

SECTIONS = ['Introduction','Literature Review','Methodology','Discussion of Analysis and Findings','Conclusion and Future Research','References']
HEADING_RE = re.compile(r'^\s*(\d+(?:\.\d+)*)[.\s]+(.+?)\s*$')
FIGURE_RE = re.compile(r'^\s*(Figure|Fig\.|Table)\s+\d+[\w.]*\s*[:.\-]?\s*.+',re.I)

def normalize(s): return re.sub(r'\s+',' ',s or '').strip()

def is_section(text):
    m=HEADING_RE.match(text)
    return (text.strip().casefold() in [x.casefold() for x in SECTIONS]) or bool(m and m.group(2).strip().casefold() in [x.casefold() for x in SECTIONS])

def classify(paragraphs):
    nonempty=[i for i,p in enumerate(paragraphs) if normalize(p.text)]
    start=next((i for i in nonempty if normalize(paragraphs[i].text).upper()=='ABSTRACT'),None)
    end=next((i for i in nonempty if i>(start if start is not None else -1) and normalize(paragraphs[i].text).lower().startswith('keywords:')),None)
    first_sec=next((i for i in nonempty if is_section(normalize(paragraphs[i].text))),None)
    kinds={}
    if start is not None:
        kinds[start]='abstract_heading'
        for i in nonempty:
            if start < i < (end if end is not None else (first_sec if first_sec is not None else len(paragraphs))):kinds[i]='abstract_body'
    if end is not None:kinds[end]='keywords'
    meta=[i for i in nonempty if i < (start if start is not None else (first_sec if first_sec is not None else len(paragraphs)))]
    for idx,i in enumerate(meta):kinds[i]= 'title' if idx==0 else 'author'
    for i in nonempty:
        txt=normalize(paragraphs[i].text)
        if is_section(txt):kinds[i]='heading'
        elif HEADING_RE.match(txt) and i >= (first_sec if first_sec is not None else 10**9):
            m=HEADING_RE.match(txt)
            if m and len(m.group(2))<110 and len(txt)<140 and not txt[-1] in '.;': kinds[i]='subheading'
        elif FIGURE_RE.match(txt):kinds[i]='caption'
    return kinds,start,end

def set_font(p, size=11, bold=None, italic=None):
    for run in p.runs:
        # Preserve non-text runs and existing emphasis in body
        run.font.name='Times New Roman'
        run.font.size=Pt(size)
        if bold is not None:run.bold=bold
        if italic is not None:run.italic=italic

def replace_text_runs(p, old, new):
    # Replace only within individual text runs: preserve any inline drawing/field relationships.
    for r in p.runs:
        if old in r.text:r.text=r.text.replace(old,new)

def format_document(raw, label='ITARS 6.0', year='2027'):
    data=io.BytesIO(raw)
    if not zipfile.is_zipfile(data):raise ValueError('Fail bukan DOCX yang sah.')
    if len(raw)>20*1024*1024:raise ValueError('Fail terlalu besar (maksimum 20 MB).')
    doc=Document(io.BytesIO(raw))
    normal=doc.styles['Normal'];normal.font.name='Times New Roman';normal.font.size=Pt(11)
    for sec in doc.sections:
        sec.page_width=Cm(21);sec.page_height=Cm(29.7)
        sec.top_margin=sec.bottom_margin=sec.left_margin=sec.right_margin=Cm(2.54)
        for p in sec.header.paragraphs:
            if 'ITARS 5.0' in p.text or p.text.strip()=='2026':
                updated=p.text.replace('ITARS 5.0',label)
                if p.text.strip()=='2026': updated=year
                if p.runs:
                    p.runs[0].text=updated
                    for r in p.runs[1:]:r.text=''
    paragraphs=doc.paragraphs
    kinds,abstract_index,kw_index=classify(paragraphs)
    for i,p in enumerate(paragraphs):
        kind=kinds.get(i,'body'); txt=normalize(p.text)
        if not txt:continue
        if kind in ('title','author','abstract_heading'):
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            set_font(p,14 if kind=='title' else 11,bold=True if kind in ('title','abstract_heading') else None)
        elif kind=='heading':
            p.alignment=WD_ALIGN_PARAGRAPH.LEFT;set_font(p,11,bold=True)
            p.paragraph_format.keep_with_next=True
        elif kind=='subheading':
            p.alignment=WD_ALIGN_PARAGRAPH.LEFT;set_font(p,11,bold=True)
            p.paragraph_format.keep_with_next=True
        elif kind=='caption':
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER;set_font(p,10)
        elif kind=='keywords':
            p.alignment=WD_ALIGN_PARAGRAPH.LEFT;set_font(p,11)
            # Italicize keyword values without rewriting runs or removing embedded elements.
            for r in p.runs:
                if 'Keywords:' not in r.text:r.italic=True
        else:
            p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            set_font(p,11)
        p.paragraph_format.space_after=Pt(0)
    # Original tables stay editable; format text only (do not rebuild cells).
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:set_font(p,10)
    buf=io.BytesIO();doc.save(buf)
    return buf.getvalue(),analyze(doc,kinds,abstract_index,kw_index)

def analyze(doc,kinds=None,abstract_index=None,kw_index=None):
    ps=doc.paragraphs
    if kinds is None:kinds,abstract_index,kw_index=classify(ps)
    text=[normalize(p.text) for p in ps]
    abstract=' '.join(text[i] for i in range((abstract_index or 0)+1,(kw_index if kw_index is not None else len(ps))) if kinds.get(i)=='abstract_body') if abstract_index is not None else ''
    words=len(re.findall(r'\b[\w]+(?:[-’\'][\w]+)*\b',abstract))
    kw=text[kw_index].split(':',1)[-1] if kw_index is not None else ''
    keywords=[x.strip() for x in kw.split(',') if x.strip()]
    heads=[t for t in text if is_section(t)]
    found=[x for x in SECTIONS if any(h.casefold() in z.casefold() for z in heads for h in [x])]
    figures=sum(bool(re.match(r'^(Figure|Fig\.)\s*\d+',t,re.I)) for t in text)
    captions=sum(bool(re.match(r'^Table\s*\d+',t,re.I)) for t in text)
    # Inline shapes count plus floating drawings; not treated as verified image quality.
    reports=[]
    def check(name,good,details):reports.append({'Semakan':name,'Status':'Lulus' if good else 'Perlu semakan','Butiran':details})
    check('Abstrak 250–300 perkataan',250<=words<=300,f'{words} perkataan' if abstract_index is not None else 'Tiada tajuk ABSTRACT dikesan')
    check('Sekurang-kurangnya 3 keywords',len(keywords)>=3,f'{len(keywords)} keywords')
    check('Seksyen teras',len(found)==len(SECTIONS),f"Dikesan: {', '.join(found) or 'tiada'}")
    check('Jadual boleh diedit', captions<=len(doc.tables),f'{captions} kapsyen jadual; {len(doc.tables)} jadual Word (semak secara visual)')
    check('Semakan APA 7',False,'Perlu semakan manusia; sistem tidak mengesahkan DOI, sumber atau setiap sitasi')
    check('Rajah dan persamaan',False,f'{figures} kapsyen rajah; semak kedudukan, imej dan persamaan dalam Word')
    check('Tajuk dan abstrak tanpa simbol khas',False,'Perlu penilaian penulis; tanda baca biasa tidak diubah')
    return {'checks':reports,'abstract_words':words,'keyword_count':len(keywords),'tables':len(doc.tables),'figure_captions':figures,'sections':heads}
