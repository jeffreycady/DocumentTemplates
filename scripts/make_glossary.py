#!/usr/bin/env python3
"""Builds glossary/glossary.json from Common_Terms_Glossary_IT_OT_Security_v1.docx.

Run it after the Word glossary changes:  python3 scripts/make_glossary.py
glossary/glossary.json is what the free glossary on www.mansiosec.com/glossary shows. Terms suggested on GitHub
(see CONTRIBUTING.md) are added to glossary/additions.json, which this script merges in, so they survive a rebuild.
Needs python-docx (pip install python-docx).
"""
import json, os, re, sys
import docx

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(HERE, "Common_Terms_Glossary_IT_OT_Security_v1.docx")
OUT = os.path.join(HERE, "glossary", "glossary.json")
ADD = os.path.join(HERE, "glossary", "additions.json")
REM = os.path.join(HERE, "glossary", "removals.json")

CATEGORY = {"A": "OT and ICS", "C": "Core IT", "D": "IT security", "E": "Network and server architecture", "F": "ITIL",
            "G": "Regulations and frameworks", "H": "Risk (SCF)", "J": "SAP", "K": "PeopleSoft", "L": "AVEVA PI", "M": "Zenon SCADA"}
SKIP_SECTIONS = {"I"}                      # Regional Scope: one organization's footprint, not definitions
ATTR = {"definition", "architecture", "database support", "integration", "modules", "deployment", "key features", "data flow",
        "why it matters for security/ops", "tiers & components", "client connectivity modes", "example", "examples", "usage"}
ACR = re.compile(r"^[A-Z0-9][A-Z0-9&/\-|.]{1,12}$")

def clean(s):
    s = s.replace(" ", " ").replace("‑", "-")
    return re.sub(r"[ \t]+", " ", s).strip()

def split_source(d):
    m = re.search(r"\s*\[([a-z0-9.\-]+(?:\.\.\.[a-z0-9.\-]+)?)\]\s*$", d)
    return (d[:m.start()].strip(), m.group(1)) if m else (d, None)

def entry(term, d, cat, aliases=()):
    term = clean(term).rstrip(":").strip()
    d = clean(d)
    ex = None
    m = re.search(r"\s*(?:Example:\s*)(.+)$", d)
    if m: d, ex = d[:m.start()].strip(), m.group(1).strip()
    q = re.search(r"\s*[“\"]([^”\"]{8,})[”\"]\s*$", d)
    if q and not ex: d, ex = d[:q.start()].strip(), q.group(1).strip()
    d, src = split_source(d)
    al = []
    pm = re.match(r"^(.*?)\s*\(([^()]+)\)$", term)
    if pm:
        a, b = pm.group(1).strip(), pm.group(2).strip()
        if ACR.match(a) and not ACR.match(b): term, al = a, [b]          # LAN (Local Area Network)
        elif ACR.match(b.replace(" ", "")) or len(b) <= 12: term, al = a, [b]  # Zero Trust Architecture (ZTA)
    al = [x for x in al if not x.startswith("Level ")] + [x for x in aliases if x]
    if ex and re.search(r"\bCompany\b", ex): ex = None                  # examples written about one organization
    if not d or len(term) > 100: return None
    d = d[:1].upper() + d[1:]
    e = {"term": term, "aliases": sorted(set(al), key=al.index), "category": cat, "def": d}
    if ex: e["example"] = ex
    if src: e["source"] = src
    return e

def build():
    D = docx.Document(SRC)
    out, sec, block, reading = [], None, None, False
    for p in D.paragraphs:
        t = p.text
        if p.style.name.startswith("Heading 1"):
            m = re.match(r"^([A-M])\)\s", t)
            if m: sec, block, reading = m.group(1), None, False; continue
            if re.match(r"^[N-Z]\)\s", t): sec = None; continue   # written from additions.json by update_word_glossary.py
            if t.startswith("Appendix"): sec = None; continue
            if sec in ("L",) and ":" in t:       # AVEVA entries are styled as headings
                pass
            else: continue
        if reading and not t.strip(): reading = False             # a reading list ends at the next blank line
        if not sec or sec in SKIP_SECTIONS or not t.strip(): continue
        cat = CATEGORY[sec]
        lines = [clean(x) for x in t.split("\n") if clean(x)]
        head, rest = lines[0], lines[1:]
        if "\t" in t and t.strip().startswith("Level"):           # the Purdue table
            parts = [clean(x) for x in re.split(r"\t+", t) if clean(x)]
            if len(parts) >= 3 and parts[0] != "Level":
                e = entry(f"Purdue {parts[0]}: {parts[1]}", parts[2], cat, aliases=[parts[1]]); e and out.append(e)
            continue
        if reading: continue
        if ":" not in head:
            reading = head.lower().startswith("recommended reading")
            block = None if reading else (head if p.style.name.startswith("List") or len(head) < 60 else None)
            continue
        label, _, d = head.partition(":")
        d = " ".join([d.strip()] + rest)
        if label.strip().lower() in ATTR:
            if label.strip().lower() == "definition" and block:
                e = entry(block, d, cat); e and out.append(e)
            continue
        if sec == "F" and " " in label:                            # "ITIL Information Technology Infrastructure Library: ..."
            first, _, more = label.partition(" ")
            if ACR.match(first) and more[:1].isupper():
                e = entry(first, d, cat, aliases=[more]); e and out.append(e); continue
        sub = re.match(r"^\s*([A-Z][^:]{2,70}):\s*(.+)$", d)       # "IT: Information Technology: def"
        short = ACR.match(label.strip()) or (len(label.strip()) <= 20 and len(sub.group(1)) <= 50 if sub else False)
        if sub and short and "." not in sub.group(1) and not sub.group(1).startswith("Per "):
            e = entry(label.strip(), sub.group(2), cat, aliases=[sub.group(1)])
        else:
            e = entry(label, d, cat)
        if e: out.append(e)
    seen, res = {}, []
    for e in out:                                  # the same term in two sections: one entry, the fuller definition
        k = e["term"].lower()
        if k in seen:
            o = seen[k]
            o["aliases"] = sorted(set(o["aliases"]) | set(e["aliases"]), key=(o["aliases"] + e["aliases"]).index)
            if len(e["def"]) > len(o["def"]):
                for f in ("def", "category", "example", "source"):
                    if f in e: o[f] = e[f]
                    elif f in o and f != "category": del o[f]
            continue
        seen[k] = e; res.append(e)
    if os.path.exists(REM):                        # entries replaced in additions.json, duplicates, or notes rather than terms
        gone = {t.lower() for t in json.load(open(REM, encoding="utf-8"))["terms"]}
        missing = gone - {e["term"].lower() for e in res}
        assert not missing, "removals.json names terms the Word glossary doesn't have: " + ", ".join(sorted(missing))
        res = [e for e in res if e["term"].lower() not in gone]
    if os.path.exists(ADD):
        for e in json.load(open(ADD, encoding="utf-8")):
            k = e["term"].lower()
            res = [x for x in res if x["term"].lower() != k] + [e]
    names = {}
    for e in res:                                  # a name or alias used by two entries would make a document's glossary ambiguous
        for n in [e["term"]] + e["aliases"]:
            assert n.lower() not in names or names[n.lower()] == e["term"], f"'{n}' is used by both {names[n.lower()]} and {e['term']}"
            names[n.lower()] = e["term"]
    res.sort(key=lambda e: e["term"].lower())
    return res

if __name__ == "__main__":
    g = build()
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc = {"title": "Common Terms Glossary: IT, OT and Security", "license": "CC-BY-4.0", "categories": list(CATEGORY.values()),
           "source": "https://github.com/jeffreycady/DocumentTemplates", "terms": g}
    json.dump(doc, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(len(g), "terms ->", os.path.relpath(OUT, HERE))
    if "-v" in sys.argv:
        for e in g: print(f'{e["category"][:14]:14} | {e["term"][:40]:40} | {",".join(e["aliases"])[:30]:30} | {e["def"][:70]}')
