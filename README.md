# Document Templates

Free Word templates for IT and security documentation, plus a glossary of common IT, OT and security terms. They were written by Jeff Cady, CISSP ([JCady LLC](https://security.jcady.org)), and are free to use, change and share under [CC BY 4.0](LICENSE).

**Fill them in online:** [www.mansiosec.com/templates](https://www.mansiosec.com/templates) lets you add your company's details and logo once, fill in any template with a form, and download a finished Word file. Everything stays in your browser, and your company details save to a file you can reuse for every document.

**Search the glossary:** [www.mansiosec.com/glossary](https://www.mansiosec.com/glossary)

## What's here

| Template | What it's for |
|---|---|
| [High-Level Design (HLD)](High-Level%20Design%20(HLD)%20Document%20Template.docx) | The design of a system or solution: purpose, scope, architecture, components, non-functional requirements, risks and dependencies |
| [Low-Level Design (LLD)](Low-Level%20Design%20(LLD)%20Document%20Template.docx) | The detailed design behind an HLD: modules, data structures, interfaces, integration, configuration and security design |
| [Method of Procedure (MOP)](Method%20of%20Procedure%20(MOP)_Template.docx) | Step-by-step instructions for a change, with prerequisites, roles, validation, rollback, risks and approvals |
| [Standard Operating Procedure (SOP)](Standard%20Operating%20Procedure%20(SOP)%20Template.docx) | A standard, repeatable procedure with roles, validation, rollback, approvals and document control |
| [Cybersecurity Risk Statement](Risk_Statement_YYYY-XXXX_Template.docx) | One risk written as a clear statement ("There is a risk that..."), with its assessment, controls, mitigation and review plan |
| [Security Requirements](Security%20Requirements%20Template.docx) | Security requirements for a project or system by category, with roles, dependencies, acceptance criteria and approval. Includes tips and a worked example for a single-vendor SASE rollout |
| [Firewall Audit](Firewall_Audit_Template.docx) | A firewall audit report: scope, methodology, policy and configuration review, vulnerabilities, findings and recommendations |
| [Technology Roadmap](Technology_Roadmap_Template.docx) | A roadmap for a security technology: governance, policies, objectives, phases, vision and initiatives |

| Glossary | |
|---|---|
| [Common Terms Glossary: IT, OT and Security](Common_Terms_Glossary_IT_OT_Security_v1.docx) ([PDF](Common_Terms_Glossary_IT_OT_Security_v1.pdf)) | OT and ICS (including the Purdue model), core IT, IT security, network and server architecture, ITIL, regulations and frameworks, SCF risk terms, SAP, PeopleSoft, AVEVA PI and Zenon SCADA |
| [`glossary/glossary.json`](glossary/glossary.json) | The same glossary as data (term, other names, category, definition, example, source), made by `scripts/make_glossary.py` |

## Using the templates

- **In Word:** download a template and replace the text in `{braces}`, `[brackets]` and `<angle brackets>`, and the guidance after each label, with your own. Delete any guidance you don't need.
- **Online:** use [www.mansiosec.com/templates](https://www.mansiosec.com/templates). Anything you leave blank keeps the template's guidance, so you can finish it in Word.
- The Glossary section of the HLD, LLD, MOP, SOP and Firewall Audit can be filled in for you online with the definitions of the terms your document uses, from this glossary and your own company's terms.

## Suggesting a glossary term or a change

See [CONTRIBUTING.md](CONTRIBUTING.md). In short: search the [online glossary](https://www.mansiosec.com/glossary) and choose **Suggest it for everyone**, or [open a glossary issue](https://github.com/jeffreycady/DocumentTemplates/issues/new?template=glossary-term.yml). Each suggestion is checked before it's added.

## Repository layout

```
*.docx, *.pdf                the templates and the glossary (the originals)
glossary/glossary.json       the glossary as data (made from the Word glossary)
glossary/additions.json      terms added from suggestions (kept through rebuilds)
web/templates.json           the places in each template the online form fills in
scripts/make_manifest.py     rebuilds web/templates.json and checks it against the templates
scripts/make_glossary.py     rebuilds glossary/glossary.json
```

## License

[Creative Commons Attribution 4.0 International (CC BY 4.0)](LICENSE). You can use, adapt and share these templates, including commercially, as long as you give credit when you share them: "Templates by Jeff Cady (JCady LLC), CC BY 4.0." There's no share-alike condition: documents you fill in for your own organization don't have to be shared, under this or any other license.

These templates are a starting point, not legal or compliance advice. Check them against your organization's requirements, and have counsel review anything with legal effect.
