#!/usr/bin/env python3
"""Plain-text author/year citation cross-checker for .docx manuscripts.
Usage: python citation_checker.py manuscript.docx --out results_folder
Requires: Python 3.9+ standard library. Outputs portable CSV and JSON files.
Rule based: no claim of full citation identity validation; ambiguous cases flagged.
"""
import re, csv, json, argparse, unicodedata
from pathlib import Path
from collections import defaultdict,Counter
from difflib import SequenceMatcher

YEAR=r'(?:19|20)\d{2}[a-z]?'
YEAR_RE=re.compile(r'(?<!\d)((?:19|20)\d{2})([a-z]?)(?!\d)',re.I)

def norm(s):
    """Fold case, accents, whitespace and typographic punctuation."""
    s = unicodedata.normalize('NFKD', s).casefold()
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', s.replace('’', "'").replace('–', '-').replace('‐', '-')).strip()


def surname(s):
    s = re.sub(r'\([^)]*\)', '', norm(s))
    s = re.sub(r'\s+(?:jr|sr)\.?$', '', s)
    return re.sub(r"[^\w\s'-]", '', s).strip()


# Unicode letters, particles, compound names, generic parenthetical aliases.
CAP = r"[A-ZÀ-ÖØ-Þ][\w’'’-]*"
PART = r"(?:de|del|la|van|von|der|den|da|di|le|du|St\.)"
NAME = rf"(?:{PART}\s+)*{CAP}(?:\s+(?:{PART}\s+)*[A-ZÀ-ÖØ-Þ][\w’'-]*)*(?:\s*\([^()\d]+\))?"
AUTH = rf"{NAME}(?:\s+et\s+al\.?|(?:\s+(?:and|&)\s+|,\s*(?:and\s+)?){NAME})*"
NARR = re.compile(rf"(?<![\w])({AUTH})\s*\((({YEAR})(?:\s*,\s*{YEAR})*)\)")


def authors(text):
    text = re.sub(r'^(?:e\.g\.,?\s*|see also\s+|see\s+|cf\.\s+)', '', text.strip(), flags=re.I)
    text = re.sub(r'\bet\s+al\.?', '', text, flags=re.I)
    return [surname(x) for x in re.split(r'\s+(?:and|&)\s+|,\s*(?:and\s+)?', text) if surname(x)]


def get_ref(p, idx):
    text = p.strip()
    ym = YEAR_RE.search(text)
    prefix = text[:ym.start()].strip(' (,') if ym else text
    # Personal reference names normally have surname, initials. Both APA and
    # initials-first coauthors are supported. Unknown styles remain reviewable.
    # Match surnames immediately before initials, excluding preceding initials.
    personal = list(re.finditer(r"(?:^|[,;&]\s*|\band\s+)([^,;&.]+?),\s*(?=[A-Z]\.)", prefix))
    if personal:
        aa = [surname(m.group(1).strip().removeprefix('and ')) for m in personal]
        for m in re.finditer(r"(?:,|\band)\s*(?:[A-Z]\.\s*)+([\w’' -]+)", prefix):
            a = surname(m.group(1))
            if a and len(a)>1 and a not in aa: aa.append(a)
    else:
        org = re.split(r',\s*(?:January|February|March|April|May|June|July|August|September|October|November|December)(?=\b|\d)', prefix, flags=re.I)[0]
        aa = [surname(org.strip(' .,'))] if org else []
    year = ym.group(0).lower() if ym else ''
    title = text[ym.end():].lstrip(')., ') if ym else ''
    return dict(ref_id=idx, lead=aa[0] if aa else '', authors=aa, year=year,
                ref_text=text, ref_paragraph=idx, title=title,
                parse_review=not bool(ym and aa), normalized_key=f"{aa[0] if aa else '?'} {year}")


def extract_cites(p, idx):
    out = []
    def add(auth, ym, raw, kind, start, end):
        aa = authors(auth)
        if not aa or len(auth)>165 or norm(auth).strip(' ,') in {'in','since','during','between','from','until','by','year','years'}: return
        # Reject prose/numeric parentheticals; keep uncertain years in separate audit.
        clean = re.sub(r'^(?:see(?: also)?|e\.g\.,?|cf\.)\s*', '', auth.strip(), flags=re.I)
        if not re.match(r"(?:[A-ZÀ-ÖØ-Þ]|(?:de|van|von)\s)", clean): return
        out.append(dict(citation=raw, author=auth, authors=aa, lead=aa[0],
            year=ym.group(0).lower(), kind=kind, paragraph=idx, start=start, end=end,
            year_start=ym.start(), context=p, normalized_key=f"{aa[0]} {ym.group(0).lower()}"))
    stack=[]; groups=[]
    for j,ch in enumerate(p):
        if ch=='(': stack.append(j)
        elif ch==')' and stack:
            a=stack.pop()
            if not stack: groups.append((a,j))
    for a,b in groups:
        inner=p[a+1:b]
        if re.fullmatch(rf"\s*{YEAR}(?:\s*,\s*{YEAR})*\s*",inner): continue
        offset=a+1
        for part in inner.split(';'):
            ys=list(YEAR_RE.finditer(part))
            if ys:
                auth=part[:ys[0].start()].strip(' ,')
                for ym in ys:
                    # A second author-year pair without a semicolon is handled too.
                    prev=[y for y in ys if y.end()<=ym.start()]
                    if prev:
                        between=part[prev[-1].end():ym.start()].strip(' ,')
                        if between: auth=between
                    absolute=YEAR_RE.search(p, offset+ym.start())
                    add(auth,absolute,part.strip(),'parenthetical',offset,offset+len(part))
            offset+=len(part)+1
    for m in NARR.finditer(p):
        for ym in YEAR_RE.finditer(p,m.start(2),m.end(2)):
            add(m.group(1),ym,m.group(0),'narrative',m.start(),m.end())
    return sorted({(c['year_start'],c['lead']):c for c in out}.values(),key=lambda c:c['year_start'])


def read_paragraphs(path):
    """Main document XML order, including tables and text boxes; joins Word runs.
    Paragraph numbers are one-based XML w:p order (same as doc.paragraphs when
    no tables/text boxes exist). Notes and headers are explicitly out of scope.
    """
    from zipfile import ZipFile
    from xml.etree import ElementTree as ET
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with ZipFile(path) as z: root=ET.fromstring(z.read('word/document.xml'))
    return [''.join(n.text or '' if n.tag.endswith('}t') else '\n' for n in el.iter() if n.tag.endswith('}t') or n.tag.endswith('}br') or n.tag.endswith('}tab')) for el in root.findall('.//w:p',ns)]


def parse_references(ps, split):
    refs=[]
    for i,t in enumerate(ps[split+1:],split+2):
        if not t.strip(): continue
        if re.match(r'^(?:appendix|appendices|supplementary)\b',t.strip(),re.I): break
        # Join clear continuation paragraphs; retain uncertain fragments as review.
        if refs and (not YEAR_RE.search(t) or re.match(r'^\s*(?:https?://|doi:)',t)):
            prev=refs.pop(); refs.append(get_ref(prev['ref_text']+' '+t,prev['ref_id']))
        else: refs.append(get_ref(t,i))
    return refs


def match_citations(citations, refs):
    for c in citations:
        opts=[r for r in refs if r['lead']==c['lead'] and r['year']==c['year']]
        # Explicit coauthors must agree, not merely the first surname.
        compatible=[r for r in opts if len(c['authors'])==1 or all(a in r['authors'] for a in c['authors'])]
        status='Matched'; note='Author/year agrees; bibliographic truth not externally verified.'
        if opts and not compatible:
            status='NEEDS_REVIEW'; note='First author/year agrees but explicit coauthors differ or reference author parsing is incomplete.'
        elif len(compatible)>1:
            status='Ambiguous';note='Multiple compatible references; cannot select a work.'
        elif not opts:
            same=[r for r in refs if r['lead']==c['lead']]
            base=[r for r in same if r['year'][:4]==c['year'][:4]]
            near=[r for r in refs if r['year']==c['year'] and SequenceMatcher(None,r['lead'],c['lead']).ratio()>=.8]
            opts=base or same or near
            if opts:
                status='NEEDS_REVIEW';note='Year/suffix mismatch' if base else ('Possible publication-year mismatch' if same else 'Possible author spelling difference')
            else:
                status='Missing reference';note='No exact or plausible reference candidate in supplied list.'
        else: opts=compatible
        c.update(status=status,matched_ref=', '.join(str(r['ref_id']) for r in opts),note=note,
            confidence='HIGH' if status in ('Matched','Missing reference') else 'MEDIUM',
            suggested_action='Verify intended work against source; correct citation or supplied list.' if status!='Matched' else 'None')
    for r in refs:
        exact=[c for c in citations if c['status']=='Matched' and c['matched_ref']==str(r['ref_id'])]
        possible=[c for c in citations if c['status']!='Matched' and str(r['ref_id']) in c['matched_ref'].split(', ')]
        r.update(status='Cited' if exact else ('NEEDS_REVIEW' if possible or r['parse_review'] else 'Uncited reference'),
            cite_count=len(exact),citation_examples=' | '.join(c['citation'] for c in (exact+possible)[:3]),
            confidence='MEDIUM' if possible or r['parse_review'] else 'HIGH',
            note='Possible citation link or uncertain parsing' if possible or r['parse_review'] else 'No extracted citation' if not exact else 'Exact citation link',
            suggested_action='Check body and intended citation before removing.' if not exact else 'None')


def audit_docx(path):
    ps=read_paragraphs(path)
    split=next((i for i,t in enumerate(ps) if re.fullmatch(r'(?:[\d.]+\s*)?(?:references|bibliography|reference list)[:.]?',t.strip(),re.I)),None)
    if split is None: raise ValueError('Cannot find References/Bibliography heading; numbered citations are unsupported.')
    refs=parse_references(ps,split)
    body=[(i+1,p) for i,p in enumerate(ps[:split]) if p.strip()]
    cites=[c for i,p in body for c in extract_cites(p,i)]
    match_citations(cites,refs)
    dup=[]
    for i,r in enumerate(refs):
        for q in refs[:i]:
            doi=lambda x: re.findall(r'10\.\d{4,9}/[^\s]+',x.lower())
            same_doi=bool(doi(r['ref_text']) and set(doi(r['ref_text'])) & set(doi(q['ref_text'])))
            similar=r['lead']==q['lead'] and r['year']==q['year'] and SequenceMatcher(None,norm(r['title']),norm(q['title'])).ratio()>.9
            if same_doi or similar: dup.append(dict(ref_id=r['ref_id'],author_year=r['normalized_key'],reference=r['ref_text'],note=f"NEEDS_REVIEW: possible duplicate of paragraph {q['ref_id']}",confidence='MEDIUM'))
    # Independent coverage control: every plausible year token, including dates
    # and unrecognized formats, has an explicit disposition and full paragraph.
    coverage=[]
    for i,p in body:
        for ym in YEAR_RE.finditer(p):
            covered=[c for c in cites if c['paragraph']==i and c['year_start']==ym.start()]
            coverage.append(dict(paragraph=i,year=ym.group(0),start=ym.start(),context=p,
                status='COVERED' if covered else 'NEEDS_REVIEW',
                note='Extracted citation' if covered else 'Uncovered year: inspect for date, numeric data, or missed citation',confidence='MEDIUM'))
    stats=dict(body_paragraphs=len(body),reference_entries=len(refs),citation_instances=len(cites),
        distinct_author_years=len({(c['lead'],c['year']) for c in cites}),matched=sum(c['status']=='Matched' for c in cites),
        missing=sum(c['status']=='Missing reference' for c in cites),
        year_suffix_mismatch=sum('mismatch' in c['note'].lower() for c in cites),
        ambiguous=sum(c['status']=='Ambiguous' for c in cites),uncited=sum(r['status']=='Uncited reference' for r in refs),
        duplicate_candidates=len(dup),needs_review=sum(c['status']=='NEEDS_REVIEW' for c in cites)+sum(r['status']=='NEEDS_REVIEW' for r in refs)+sum(c['status']=='NEEDS_REVIEW' for c in coverage)+len(dup),
        year_tokens=len(coverage),uncovered_year_tokens=sum(c['status']=='NEEDS_REVIEW' for c in coverage))
    audit_docx.coverage=coverage
    return stats,cites,refs,dup

def save(path,outdir):
    stats,cites,refs,dup=audit_docx(path)
    outdir=Path(outdir);outdir.mkdir(parents=True,exist_ok=True)
    sets={
      'all_citations.csv':(cites,['status','citation','lead','year','kind','paragraph','context','matched_ref','note']),
      'all_references.csv':(refs,['status','lead','year','ref_paragraph','ref_text','cite_count','citation_examples','ref_id']),
      'citation_issues.csv':([x for x in cites if x['status']!='Matched'],['status','citation','lead','year','paragraph','context','matched_ref','note']),
      'uncited_references.csv':([x for x in refs if x['status']=='Uncited reference'],['lead','year','ref_paragraph','ref_text','ref_id']),
      'duplicate_candidates.csv':(dup,['author_year','ref_id','reference','note'])}
    for filename,(rows,headers) in sets.items():
        with (outdir/filename).open('w',encoding='utf-8-sig',newline='') as f:
            headers = headers + [k for k in ('normalized_key','confidence','suggested_action','note','start','end','year_start') if k not in headers and any(k in r for r in rows)]
            w=csv.DictWriter(f,fieldnames=headers,extrasaction='ignore');w.writeheader();w.writerows(rows)
    (outdir/'year_coverage.json').write_text(json.dumps(audit_docx.coverage,indent=2,ensure_ascii=False),encoding='utf-8')
    (outdir/'results.json').write_text(json.dumps(dict(summary=stats,citations=cites,references=refs,duplicate_candidates=dup),indent=2,ensure_ascii=False),encoding='utf-8')
    print(json.dumps(stats,indent=2,ensure_ascii=False))
    return stats,cites,refs,dup

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('docx');ap.add_argument('--out',default='citation_check_results');args=ap.parse_args()
    save(args.docx,args.out)
