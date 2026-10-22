#!/usr/bin/env python3
"""Builds web/templates.json: which parts of each Word template the free form-filler on www.mansiosec.com/templates
fills in, and with what. Run it after changing a template:  python3 scripts/make_manifest.py

The form-filler opens the original .docx in the browser and changes only the places listed here. Each place records
the template text it expects to find ("expect"); if a template is edited so a place no longer matches, this script fails,
and the website skips that place rather than writing into the wrong paragraph.

Places ("ops"):
  after    keep the label at the start of a paragraph ("Version:") and write the answer after it
  para     replace a whole paragraph (and the extra guidance paragraphs in "also") with the answer, one paragraph per line
  replace  replace a piece of text inside a paragraph (or a table cell: t, r, c)
  steps    replace the example steps between two paragraphs with numbered steps
  rows     fill table rows from a list (approvers, owners, change log)
  glossary fill a Glossary section with the definitions of terms the finished document uses
  sdt      set a Word dropdown or date picker (the i-th content control in a paragraph)
  cut      remove everything after a table (notes and samples that aren't part of the template)
Every op is skipped when its answer is left blank, so the template's own guidance stays in place.
Needs python-docx.
"""
import json, os, re, sys
import docx
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(HERE, "web", "templates.json")

def load(f):
    d = docx.Document(os.path.join(HERE, f))
    paras, tables = [], []
    for el in d.element.body.iterchildren():
        if el.tag == qn("w:p"): paras.append(docx.text.paragraph.Paragraph(el, d))
        elif el.tag == qn("w:tbl"): tables.append(docx.table.Table(el, d))
    return d, paras, tables

def slug(s): return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")
def key(s):
    w = re.sub(r"[^A-Za-z0-9 ]+", " ", s).split()
    return (w[0].lower() + "".join(x.capitalize() for x in w[1:]))[:40] if w else "field"

class T:
    def __init__(self, tid, file, title, category, summary):
        self.d, self.P, self.Tb = load(file)
        self.m = {"id": tid, "file": file, "docx": tid + ".docx", "title": title, "category": category, "summary": summary,
                  "fields": [], "ops": []}
        self.keys = set()
    def field(self, k, label, hint="", type="textarea", group="", common=False):
        if k in self.keys: return k
        self.keys.add(k); f = {"key": k, "label": label, "type": type, "group": group}
        if hint: f["hint"] = hint
        if common: f["common"] = True
        self.m["fields"].append(f); return k
    def text(self, p): return self.P[p].text
    def after(self, p, label, k, flabel=None, type="text", group="", common=False, hint=None):
        t = self.text(p); assert t.startswith(label), (self.m["id"], p, t, label)
        rest = t[len(label):].strip()
        self.field(k, flabel or label.rstrip(": "), hint if hint is not None else rest, type, group, common)
        self.m["ops"].append({"op": "after", "p": p, "label": label, "key": k, "expect": t})
    def para(self, p, k, label, also=(), group="", type="textarea", hint=None):
        t = self.text(p)
        h = hint if hint is not None else " ".join([t] + [self.text(a) for a in also]).strip(" []{}<>")
        self.field(k, label, h, type, group)
        self.m["ops"].append({"op": "para", "p": p, "key": k, "also": list(also), "expect": t, "alsoExpect": [self.text(a) for a in also]})
    def replace(self, p, find, k, label, type="text", group="", common=False, cell=None, hint=None):
        if cell:
            t_, r_, c_ = cell; t = self.Tb[t_].rows[r_].cells[c_].text
            assert find in t, (self.m["id"], cell, t)
            self.field(k, label, hint if hint is not None else find.strip("[]{}<>"), type, group, common)
            self.m["ops"].append({"op": "replace", "t": t_, "r": r_, "c": c_, "find": find, "key": k, "expect": t}); return
        t = self.text(p); assert find in t, (self.m["id"], p, t, find)
        self.field(k, label, hint if hint is not None else find.strip("[]{}<>"), type, group, common)
        self.m["ops"].append({"op": "replace", "p": p, "find": find, "key": k, "expect": t})
    def rows(self, t, start, k, cols, label, group="Sign-off"):
        tb = self.Tb[t]; head = [c.text.strip() for c in tb.rows[0].cells]
        self.field(k, label, "", "rows", group); f = self.m["fields"][-1]; f["cols"] = cols; f["heads"] = head
        self.m["ops"].append({"op": "rows", "t": t, "start": start, "key": k, "cols": [c["key"] for c in cols], "expect": head, "max": len(tb.rows) - start})
    def glossary(self, p, also=(), label_keep=None):
        t = self.text(p)
        self.m["ops"].append({"op": "glossary", "p": p, "also": list(also), "label": label_keep, "expect": t, "alsoExpect": [self.text(a) for a in also]})
        self.m["glossary"] = True
    def steps(self, a, b, k, label="Steps", group="Procedure"):
        self.field(k, label, "One step per line. Each line becomes “Step 1”, “Step 2”, and so on.", "textarea", group)
        self.m["ops"].append({"op": "steps", "from": a, "to": b, "key": k, "expect": self.text(a), "toExpect": self.text(b)})
    def sdt(self, p, k, key_, label, group=""):
        """The k-th content control (a Word dropdown or date picker) in paragraph p."""
        sd = list(self.P[p]._p.iter(qn("w:sdt")))[k]; pr = sd.find(qn("w:sdtPr"))
        dd, dt = pr.find(qn("w:dropDownList")), pr.find(qn("w:date"))
        if dd is not None:
            opts = [li.get(qn("w:displayText")) or li.get(qn("w:value")) for li in dd.findall(qn("w:listItem"))]
            opts = [o for o in opts if o and o != "Choose an item."]
            self.field(key_, label, "", "select", group); self.m["fields"][-1]["options"] = opts
        else:
            fmt = dt.find(qn("w:dateFormat")).get(qn("w:val")) if dt is not None and dt.find(qn("w:dateFormat")) is not None else "M/d/yyyy"
            self.field(key_, label, "", "date", group)
            self.m["ops"].append({"op": "sdt", "p": p, "i": k, "key": key_, "date": fmt, "expect": self.text(p)}); return
        self.m["ops"].append({"op": "sdt", "p": p, "i": k, "key": key_, "expect": self.text(p)})
    def strip(self, p, find):
        assert find in self.text(p), (self.m["id"], p, find)
        self.m["ops"].append({"op": "strip", "p": p, "find": find, "expect": self.text(p)})
    def cut(self, t):
        self.m["ops"].append({"op": "cut", "t": t, "expect": [c.text.strip() for c in self.Tb[t].rows[0].cells]})
    def headings_before(self, p):
        for i in range(p - 1, -1, -1):
            s = self.P[i].style.name; t = self.text(i).strip()
            if t and (s.startswith("Heading") or s == "Title"): return t.rstrip(": ")
        return ""
    def enclosed(self, open_, close, label_from="heading", skip=(), group_from=None):
        """Every paragraph (or run of paragraphs) wrapped in open_...close becomes a 'para' field."""
        i, n = 0, len(self.P)
        while i < n:
            t = self.text(i).strip()
            if i in skip or not t.startswith(open_): i += 1; continue
            j = i
            while not self.text(j).strip().endswith(close) and j + 1 < n: j += 1
            label = self.headings_before(i)
            grp = group_from(i) if group_from else ""
            if label.lower() == "glossary": self.glossary(i, range(i + 1, j + 1))
            else: self.para(i, self.unique(key(label)), label, range(i + 1, j + 1), grp)
            i = j + 1
    def unique(self, k):
        b, n = k, 2
        while k in self.keys: k = f"{b}{n}"; n += 1
        return k
    def section(self, p):
        for i in range(p, -1, -1):
            if self.P[i].style.name == "Heading 1": return self.text(i).strip()
        return ""

PEOPLE = [{"key": "name", "label": "Name"}, {"key": "title", "label": "Title"}, {"key": "date", "label": "Date"}]
OWNERS = [{"key": "name", "label": "Owner"}, {"key": "title", "label": "Position"}, {"key": "date", "label": "Date"}]
CHANGES6 = [{"key": "desc", "label": "Description of change"}, {"key": "requestor", "label": "Requestor"}, {"key": "date", "label": "Date revised"},
            {"key": "decision", "label": "Decision"}, {"key": "ddate", "label": "Decision date"}, {"key": "comments", "label": "Comments"}]
REVS = [{"key": "version", "label": "Version"}, {"key": "date", "label": "Date"}, {"key": "desc", "label": "Description of changes"}, {"key": "name", "label": "Author"}]

out = []

# ---- High-Level and Low-Level Design
for tid, f, title, owner_p in [("hld", "High-Level Design (HLD) Document Template.docx", "High-Level Design (HLD)", None),
                               ("lld", "Low-Level Design (LLD) Document Template.docx", "Low-Level Design (LLD)", 4)]:
    t = T(tid, f, title, "Design", "The design of a system or solution: its purpose, architecture, components, requirements, risks and dependencies." if tid == "hld"
          else "The detailed design behind a high-level design: modules, data, interfaces, configuration and security.")
    G = "Document details"
    t.replace(0, "{Solution/Application}" if tid == "hld" else "{Solution}", "subject", "Solution or application", group=G, hint="The system or solution this design covers")
    t.after(2, "Version:", "version", "Version", group=G, common=True, hint="1.0")
    t.after(3, "Author(s):", "author", "Author(s)", group=G, common=True, hint="")
    if tid == "lld":
        t.after(4, "Solution Owner(s):", "owner", "Solution owner(s)", group=G, common=True, hint="")
        t.after(5, "Hight Level Design document:", "hldLink", "High-level design document", group=G, hint="Link to the HLD this LLD was made from")
        t.after(6, "Date:", "date", "Date", type="date", group=G, common=True, hint="")
    else:
        t.after(4, "Date:", "date", "Date", type="date", group=G, common=True, hint="")
    t.rows(0, 1, "owners", OWNERS, "Solution/application owners", group=G)
    t.rows(1, 1, "changelog", CHANGES6, "Change log", group=G)
    t.enclosed("{", "}", group_from=lambda i, t=t: t.section(i), skip={0, 2, 3, 4, 5, 6})
    out.append(t.m)

# ---- Method of Procedure and Standard Operating Procedure
for tid, f, title, word in [("mop", "Method of Procedure (MOP)_Template.docx", "Method of Procedure (MOP)", "MOP"),
                            ("sop", "Standard Operating Procedure (SOP) Template.docx", "Standard Operating Procedure (SOP)", "SOP")]:
    t = T(tid, f, title, "Operations", "Step-by-step instructions for a change or task, with roles, validation, rollback and approvals." if tid == "mop"
          else "A standard, repeatable procedure with roles, validation, rollback, approvals and document control.")
    G = "Document details"
    cell = t.Tb[0].rows[0].cells[0].text
    find = "[Template <ITSM Ticket>]" if tid == "mop" else "[Template]"
    t.replace(None, find, "reference", "Reference (ticket or document number)", cell=(0, 0, 0), group=G, hint="e.g. CHG0012345" if tid == "mop" else "e.g. SOP-014")
    t.replace(None, "<Application \u2013 Request>", "application", "Application or request", cell=(0, 0, 0), group=G, hint="The application or request this covers")
    t.para(1, "subject", "Procedure name", group=G, type="text", hint="The name of the procedure")
    steps_from = next(i for i, p in enumerate(t.P) if p.text.strip() == "Step 1")
    steps_to = next(i for i, p in enumerate(t.P) if p.text.strip().startswith("{Pic of the thing"))
    appr = next(i for i, p in enumerate(t.P) if p.style.name.startswith("Heading") and "Approvals" in p.text)
    t.enclosed("[", "]", group_from=lambda i, t=t, a=appr: "Procedure" if i < a else "Document administration", skip={1})
    t.steps(steps_from, steps_to, "steps")
    apr = 1
    t.rows(apr, 1, "approvers", PEOPLE, "Approvals")
    out.append(t.m)

# ---- Risk statement
t = T("risk-statement", "Risk_Statement_YYYY-XXXX_Template.docx", "Cybersecurity Risk Statement", "Risk",
      "One risk, written as a clear statement, with its assessment, controls, mitigation and review plan.")
G1, G2, G3, G4 = "Risk identification", "Risk statement", "Risk description and assessment", "Mitigation and review"
t.strip(0, " Template")
t.after(2, "Risk ID:", "riskId", "Risk ID", group=G1, hint="YYYY-XXXX, e.g. 2026-0007")
t.sdt(3, 0, "date", "Date identified", group=G1); t.m["fields"][-1]["common"] = True
t.after(4, "Risk Owner:", "owner", "Risk owner", group=G1, common=True, hint="Name / Department")
t.sdt(5, 0, "riskCategory", "Category", group=G1)
t.sdt(5, 1, "riskCategory2", "Second category (optional)", group=G1)
for find, k, lab in [("[threat actor or event]", "threat", "Threat actor or event"), ("[action or impact]", "action", "Action or impact"),
                     ("[vulnerability or condition]", "vulnerabilityCondition", "Vulnerability or condition"), ("[consequence or business impact]", "consequence", "Consequence or business impact")]:
    t.replace(7, find, k, lab, group=G2)
for p, lab, k in [(9, "Source:", "source"), (10, "Vulnerability:", "vulnerability"), (11, "Asset(‘s) at Risk:", "assets"), (12, "Potential Impact:", "impact")]:
    t.after(p, lab, k, lab.rstrip(":").replace("(‘s)", "s"), type="textarea", group=G3)
t.sdt(13, 0, "likelihood", "Likelihood", group=G3)
t.sdt(15, 0, "initialRisk", "Initial risk level", group=G3)
t.after(16, "Controls in Place:", "controls", "Controls in place", type="textarea", group=G3)
t.after(17, "Residual Risk:", "residualRisk", "Residual risk", type="textarea", group=G3)
t.sdt(18, 0, "residualLikelihood", "Likelihood (after controls)", group=G3)
t.para(20, "mitigation", "Mitigation strategies", group=G4)
t.sdt(22, 0, "reviewFrequency", "Review frequency", group=G4)
t.sdt(23, 0, "nextReview", "Next review date", group=G4)
t.after(24, "Responsible Party:", "responsible", "Responsible party", group=G4)
out.append(t.m)

# ---- Firewall audit: every "Label: guidance" line
t = T("firewall-audit", "Firewall_Audit_Template.docx", "Firewall Audit", "Audit",
      "A firewall audit report: scope, method, policy and configuration review, vulnerabilities, findings and recommendations.")
t.strip(0, " Template")
grp = ""
for i, p in enumerate(t.P):
    s = p.text.strip()
    if re.match(r"^\d+\.\s", s): grp = re.sub(r"^\d+\.\s*", "", s); continue
    if ":" in s and i > 0:
        lab = s.split(":")[0] + ":"
        if lab.startswith("Glossary"): t.glossary(i, label_keep=lab); continue
        t.after(i, lab, t.unique(key(lab)), lab.rstrip(":"), type="textarea", group=grp)
out.append(t.m)

# ---- Technology roadmap
t = T("technology-roadmap", "Technology_Roadmap_Template.docx", "Technology Roadmap", "Strategy",
      "A roadmap for a security technology: overview, governance, policies, objectives, phases, vision, initiatives and approvals.")
t.strip(0, " Template")
t.rows(0, 1, "approvers", PEOPLE, "Document approvals", group="Document details")
t.replace(3, "[Organization]", "company", "Organization", group="Document details", common=True)
nxt = lambda a: [j for j in range(a + 1, len(t.P)) if t.text(j).startswith("- ")][:0]
i = 0
while i < len(t.P):
    p = t.P[i]
    if p.style.name == "Normal" and i != 3 and not p.text.startswith("- ") and p.text.strip():
        also, j = [], i + 1
        while j < len(t.P) and t.text(j).startswith("- "): also.append(j); j += 1
        lab = t.headings_before(i); lab = re.sub(r"^\d+\.\s*", "", lab)
        t.para(i, t.unique(key(lab)), lab, also, group="Roadmap")
        i = j; continue
    i += 1
out.append(t.m)

# ---- Security requirements (the template part; the tips and the SASE sample after it are left out)
t = T("security-requirements", "Security Requirements Template.docx", "Security Requirements", "Requirements",
      "Security requirements for a project or system, by category, with roles, dependencies, acceptance criteria and approval.")
G = "Document details"
t.strip(0, " Template")
t.para(2, "subject", "Project or initiative name", group=G, type="text", hint="Name of the project, system or initiative")
t.para(4, "date", "Date", group=G, type="date"); t.m["fields"][-1]["common"] = True
t.para(6, "owner", "Owner", group=G, type="text", hint="The security manager or department responsible"); t.m["fields"][-1]["common"] = True
t.after(9, "Objective:", "objective", "Objective", type="textarea", group="Purpose")
t.after(10, "Scope:", "scope", "Scope", type="textarea", group="Purpose")
for p in range(14, 42):
    s = t.text(p).strip()
    m = re.match(r"^([a-g])\.\s+(.+)$", s)
    if m:
        lab = m.group(2); ex = [j for j in range(p + 2, p + 4) if t.text(j).startswith("Example")]
        t.para(p + 1, t.unique(key(lab)), lab, ex, group="Security requirements", hint=t.text(p + 1) + " " + " ".join(t.text(j) for j in ex))
for p, lab in [(44, "Security Team:"), (45, "Project/Team Leads:"), (46, "Other Stakeholders:")]:
    t.after(p, lab, t.unique(key(lab)), lab.rstrip(":"), type="textarea", group="Roles and responsibilities")
t.para(49, "dependencies", "Dependencies", group="Dependencies and acceptance")
t.para(52, "acceptance", "Acceptance criteria", [53, 54], group="Dependencies and acceptance")
t.para(57, "references", "References", group="Dependencies and acceptance")
for p, lab, k in [(60, "Prepared By:", "preparedBy"), (61, "Reviewed By:", "reviewedBy"), (62, "Approved By:", "approvedBy"), (63, "Date of Approval:", "approvalDate")]:
    t.after(p, lab, k, lab.rstrip(":"), type="date" if k == "approvalDate" else "text", group="Approval", hint="Name / Title" if k != "approvalDate" else "")
t.rows(0, 1, "revisions", REVS, "Revision history", group="Approval")
t.cut(0)
out.append(t.m)

# ---- check every op against its template, then write
for m in out:
    keys = {f["key"] for f in m["fields"]}
    for op in m["ops"]:
        assert op.get("key") in keys or op["op"] in ("glossary", "cut", "strip"), (m["id"], op)
os.makedirs(os.path.dirname(OUT), exist_ok=True)
json.dump({"version": 1, "license": "CC-BY-4.0", "credit": "Templates by Jeff Cady (JCady LLC), CC BY 4.0",
           "source": "https://github.com/jeffreycady/DocumentTemplates", "templates": out}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for m in out: print(f'{m["id"]:22} {len(m["fields"]):3} fields {len(m["ops"]):3} places' + ("  + glossary" if m.get("glossary") else ""))
