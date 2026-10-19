"""ITARS 6.0 V4 formatter. Mutates source DOCX in-place to preserve embedded objects.
The official Word reference template supplies layout and paragraph formatting exemplars.
This software does not claim to validate the academic accuracy of a paper.
"""
import io
import re
import zipfile
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn

TEMPLATE_PATH = Path(__file__).with_name('ITARS_6_Reference_Template.docx')
CORE = ('Introduction','Literature Review','Methodology','Discussion of Analysis and Findings','Conclusion and Future Research','References')
HEADING_RE = re.compile(r'^\s*(\d+(?:\.\d+)*)[.\s]+(.+?)\s*$')
CAPTION_RE = re.compile(r'^\s*(?:Figure|Fig\.|Table)\s*\d+[\w.]*\s*[:.\-]?\s*.+', re.I)

def clean(text):
    return re.sub(r'\s+', ' ', text or '').strip()

def section_name(text):
    t=clean(text).casefold()
    m=HEADING_RE.match(t)
    if m: t=m.group(2).strip().casefold()
    return next((v for v in CORE if v.casefold()==t), None)

def classify(ps):
    texts=[clean(p.text) for p in ps]
    a=next((i for i,t in enumerate(texts) if t.upper()=='ABSTRACT'),None)
    k=next((i for i,t in enumerate(texts) if a is not None and i>a and t.casefold().startswith('keywords:')),None)
    first=next((i for i,t in enumerate(texts) if section_name(t)),len(ps))
    classes={}
    metadata=[i for i,t in enumerate(texts) if t and i < (a if a is not None else first)]
    for n,i in enumerate(metadata):classes[i]='title' if n==0 else 'author'
    if a is not None:
        classes[a]='abstract_heading'
        for i in range(a+1, k if k is not None else first):
            if texts[i]:classes[i]='abstract'
    if k is not None:classes[k]='keywords'
    reference_start=next((i for i,t in enumerate(texts) if section_name(t)=='References'),len(ps))
    for i,t in enumerate(texts):
        if not t:continue
        if section_name(t):classes[i]='heading'
        elif CAPTION_RE.match(t):classes[i]='caption'
        elif i > first and i < reference_start:
            m=HEADING_RE.match(t)
            if m and len(t)<140 and not t.endswith(('.', ';')):
                classes[i]='subheading'
        if i>reference_start and i not in classes:classes[i]='reference'
    return classes,a,k

def extract_spec():
    t=Document(TEMPLATE_PATH)
    p=t.paragraphs
    def find(fn):
        return next(x for x in p if fn(clean(x.text)))
    ex={'title':find(lambda s:s=='TITLE'),
        'author':find(lambda s:s.startswith('Author1,')),
        'abstract_heading':find(lambda s:s=='ABSTRACT'),
        'abstract':find(lambda s:s.startswith('Abstract should between')),
        'keywords':find(lambda s:s.startswith('Keywords:')),
        'heading':find(lambda s:s=='Methodology'),
        'subheading':find(lambda s:s.startswith('1.1 Agent B')),
        'body':find(lambda s:s.startswith('Should discuss the problem statement')),
        'caption':find(lambda s:s.startswith('Figure 1.1:')),
        'reference':find(lambda s:s.startswith('Ruxton, C.'))}
    return t,ex

def point_size(p,default=12):
    for r in p.runs:
        if r.text.strip() and r.font.size: return r.font.size.pt
    if p.style.font.size:return p.style.font.size.pt
    return default

def apply_format(p,example,kind):
    # Use paragraph settings from Word template exemplar rather than hardcoded distances.
    source=example._p.pPr
    target=p._p
    existing=target.pPr
    if existing is not None:target.remove(existing)
    if source is not None:
        new_props=deepcopy(source)
        # Numbering IDs belong to the template numbering part and must not be copied
        # into another Word file. Keep manuscript-native numbering references instead.
        for child in list(new_props):
            if child.tag==qn('w:numPr'):
                new_props.remove(child)
        if existing is not None:
            for child in existing:
                if child.tag==qn('w:numPr'):
                    new_props.append(deepcopy(child))
        target.insert(0,new_props)
    if kind == 'body':
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    if kind=='reference':
        # APA hanging indent; this is a layout aid, not a claim of APA validity.
        p.paragraph_format.left_indent=Pt(18)
        p.paragraph_format.first_line_indent=Pt(-18)
    if kind in ('heading','subheading'):
        p.paragraph_format.keep_with_next=True
    size=point_size(example)
    if kind in ('body','heading','subheading','reference','caption','abstract_heading'):
        size=12 # Normal base font inherited by the reference exemplar
    for r in p.runs:
        r.font.name='Times New Roman'
        r.font.size=Pt(size)
        if kind in ('heading','subheading','title','abstract_heading'):
            r.bold=True
        if kind=='keywords':
            # Preserve existing run boundaries; best effort italicization of keyword values.
            r.italic=True
    if kind=='keywords':
        # Leave the label upright when it occupies its own run.
        for r in p.runs:
            if clean(r.text).casefold() in ('keywords:', 'keywords'):r.italic=False

def format_document(raw,label='ITARS 6.0',year='2027'):
    if not isinstance(raw,bytes) or not zipfile.is_zipfile(io.BytesIO(raw)):
        raise ValueError('The uploaded file is not a valid Microsoft Word .docx document.')
    if len(raw)>20*1024*1024:raise ValueError('File exceeds the 20 MB application limit.')
    template,examples=extract_spec()
    doc=Document(io.BytesIO(raw))
    template_section=template.sections[0]
    for sec in doc.sections:
        for attr in ('page_width','page_height','top_margin','bottom_margin','left_margin','right_margin','header_distance','footer_distance'):
            setattr(sec,attr,getattr(template_section,attr))
        # The official ITARS 6.0 header is authoritative. Replace source header
        # wholesale because Word can split the event name across many runs.
        hdr=sec.header
        for para in list(hdr.paragraphs):
            node=para._element
            node.getparent().remove(node)
        for para in template_section.header.paragraphs:
            hdr._element.append(deepcopy(para._element))
    normal=doc.styles['Normal']
    normal.font.name='Times New Roman'
    normal.font.size=Pt(12)
    classes,a,k=classify(doc.paragraphs)
    for i,p in enumerate(doc.paragraphs):
        if not clean(p.text):continue
        kind=classes.get(i,'body')
        apply_format(p,examples[kind],kind)
    # Word-native tables preserved; update run font only, never flatten to images.
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name='Times New Roman';r.font.size=Pt(12)
    out=io.BytesIO();doc.save(out)
    return out.getvalue(), analyze(doc,classes,a,k)

def analyze(doc,classes=None,a=None,k=None):
    ps=doc.paragraphs
    if classes is None:classes,a,k=classify(ps)
    texts=[clean(p.text) for p in ps]
    abstract=' '.join(texts[i] for i in range(a+1,k if k is not None else len(ps)) if classes.get(i)=='abstract') if a is not None else ''
    words=len(re.findall(r"\b[\w]+(?:[-’'][\w]+)*\b",abstract))
    kw=texts[k].split(':',1)[-1] if k is not None else ''
    keywords=[x.strip() for x in kw.split(',') if x.strip()]
    sections=[s for s in CORE if any(section_name(t)==s for t in texts)]
    figure_caps=sum(bool(re.match(r'^\s*(?:Figure|Fig\.)\s*\d+',t,re.I)) for t in texts)
    table_caps=sum(bool(re.match(r'^\s*Table\s*\d+',t,re.I)) for t in texts)
    checks=[]
    def add(name,status,details):checks.append({'check':name,'status':status,'details':details})
    add('Abstract length (250–300 words)','Pass' if 250<=words<=300 else 'Review required',f'{words} words' if a is not None else 'ABSTRACT heading not detected')
    add('At least three keywords','Pass' if len(keywords)>=3 else 'Review required',f'{len(keywords)} keyword(s) detected')
    add('Required core sections','Pass' if len(sections)==len(CORE) else 'Review required','Detected: '+(', '.join(sections) or 'none'))
    add('Editable Word tables','Review required',f'{len(doc.tables)} Word table(s), {table_caps} table caption(s). Check for image-based tables separately.')
    add('APA 7th edition references','Review required','Bibliographic accuracy, DOI validity, in-text matching and styling require author review.')
    add('Figures and equations','Review required',f'{figure_caps} figure caption(s). Verify picture positioning, equation integrity and numbering.')
    add('Special characters in title and abstract','Review required','Author must check symbols, formulae and special punctuation.')
    add('Template-derived paragraph formatting','Review required','Word paragraph examples applied. Verify layout in Microsoft Word, especially complex objects and list numbering.')
    return {'checks':checks,'abstract_words':words,'keyword_count':len(keywords),'tables':len(doc.tables),'figure_captions':figure_caps,'sections':sections}
