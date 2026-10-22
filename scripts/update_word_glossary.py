#!/usr/bin/env python3
"""Writes the terms in glossary/additions.json into the Word glossary, so the .docx (and the PDF made from it) has the
same terms as glossary/glossary.json and www.mansiosec.com/glossary.

  python3 scripts/update_word_glossary.py      then rebuild the PDF:
  soffice --headless --convert-to pdf Common_Terms_Glossary_IT_OT_Security_v1.docx

The added terms go in their own sections, N) onwards, before Appendix A, grouped by category and sorted. Run it again
after changing additions.json: it replaces those sections rather than adding them twice. Where an addition corrects an
entry already in sections A) to M) (NIST CSF, the regulations), the old paragraph is removed so the term appears once.
make_glossary.py skips sections N) onwards, because additions.json is their source.
Needs python-docx.
"""
import copy, json, os, re
import docx
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
F = os.path.join(HERE, "Common_Terms_Glossary_IT_OT_Security_v1.docx")
ADD = json.load(open(os.path.join(HERE, "glossary", "additions.json"), encoding="utf-8"))
SECTIONS = [("IT security", "IT Security: Identity, Tools, Vulnerabilities and Response"), ("Network and server architecture", "Network and Infrastructure"),
            ("Regulations and frameworks", "Frameworks and Regulations"), ("Risk (SCF)", "Risk, Continuity and Recovery"),
            ("OT and ICS", "OT and Industrial Systems"), ("Core IT", "Documents and IT Operations"), ("ITIL", "Documents and IT Operations")]
# entries in sections A) to M) that an addition replaces, by the text before their first colon
REPLACED = ["NIST CSF", "Zero Trust Architecture (ZTA)", "Historian", "Cybersecurity Maturity Model Certification (CMMC) 2.0",
            "Health Insurance Portability and Accountability Act (HIPAA) Security Rule", "Payment Card Industry Data Security Standard (PCI DSS) v4.0.1",
            "Sarbanes-Oxley Act (SOX) §404", "System and Organization Controls (SOC) 2"]

d = docx.Document(F)
body = d.element.body
paras = lambda: [p for p in d.paragraphs]
heads = [p for p in paras() if p.style.name.startswith("Heading 1")]
appx = next(p for p in heads if p.text.startswith("Appendix A"))
# remove the sections written last time (N) and later, up to Appendix A)
gen = False
for p in paras():
    if p.style.name.startswith("Heading 1") and re.match(r"^[N-Z]\)\s", p.text): gen = True
    if p._p is appx._p: break
    if gen: body.remove(p._p)
# remove the entries the additions replace
gone = []
for p in paras():
    lab = p.text.split(":")[0].strip()
    if lab in REPLACED and ":" in p.text: gone.append(lab); p._p.getparent().remove(p._p)
# templates for a heading and an entry, copied from the existing document
heads = [p for p in paras() if p.style.name.startswith("Heading 1")]
H = next(p for p in heads if p.text.startswith("C)"))
E = next(p for p in paras() if p.text.startswith("SIEM:"))
def para(tpl, text):
    x = copy.deepcopy(tpl._p)
    for pp in x.findall(qn("w:pPr")):                     # no list bullet or numbering on the copies
        for n in pp.findall(qn("w:numPr")): pp.remove(n)
    for r in x.findall(qn("w:r")): x.remove(r)
    for h in x.findall(qn("w:hyperlink")): x.remove(h)
    rs = tpl._p.findall(qn("w:r"))
    r = copy.deepcopy(rs[0]) if rs else None
    if r is None: r = x.makeelement(qn("w:r"), {})
    for c in list(r):
        if c.tag != qn("w:rPr"): r.remove(c)
    for i, line in enumerate(text.split("\n")):
        if i: r.append(r.makeelement(qn("w:br"), {}))
        t = r.makeelement(qn("w:t"), {}); t.text = line; t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); r.append(t)
    x.append(r); return x
anchor = next(p for p in paras() if p.text.startswith("Appendix A"))._p
letter, done = ord("N"), set()
for cat, title in SECTIONS:
    if title in done: continue
    cats = [c for c, t in SECTIONS if t == title]; done.add(title)
    items = sorted([e for e in ADD if e["category"] in cats], key=lambda e: e["term"].lower())
    if not items: continue
    anchor.addprevious(para(H, f"{chr(letter)}) {title}")); letter += 1
    for e in items:
        name = e["term"] + (f" ({', '.join(e['aliases'])})" if e["aliases"] else "")
        anchor.addprevious(para(E, f"{name}: {e['def']}" + (f"\nExample: {e['example']}" if e.get("example") else "")))
d.save(F)
print(f"{len(ADD)} terms in {letter - ord('N')} sections; removed the replaced entries: {', '.join(gone) or 'none (already done)'}")
