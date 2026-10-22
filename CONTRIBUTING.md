# Contributing

Thanks for helping improve these templates and the glossary. Everything here is shared under [CC BY 4.0](LICENSE), so anything you contribute is shared the same way.

## Suggest a glossary term

The easiest way: on [www.mansiosec.com/glossary](https://www.mansiosec.com/glossary), search for the term, then choose **Suggest it for everyone**. That opens a GitHub issue with the term already filled in.

Or [open a glossary issue here](https://github.com/jeffreycady/DocumentTemplates/issues/new?template=glossary-term.yml) yourself. Please:

- write the definition in your own words, in one or two plain sentences (don't copy it from a website or standard);
- add the expansion of an acronym, an example of use and a source if you have them;
- leave out anything confidential about your organization.

Each suggestion is checked before it goes in. Accepted terms (and corrections to existing ones, under the same term) are added to [`glossary/additions.json`](glossary/additions.json) (the same fields as `glossary/glossary.json`: `term`, `aliases`, `category`, `def`, and optional `example` and `source`), so they survive the next rebuild from the Word glossary. Entries from the Word glossary that are replaced, duplicated or aren't really terms are listed in [`glossary/removals.json`](glossary/removals.json). The build stops if two entries share a name or another name.

You can also send a pull request that edits `glossary/additions.json` directly.

## Suggest a change to a template

Open an issue describing the change, or send a pull request with the edited `.docx`. If you change a template, run the scripts below so the web form-filler still finds the places it fills in.

## Rebuilding the web files

The free form-filler and glossary on www.mansiosec.com read two files made from the Word documents here:

```
pip install python-docx
python3 scripts/make_manifest.py     # web/templates.json: the places in each template the form fills in
python3 scripts/make_glossary.py     # glossary/glossary.json: the glossary, from the Word glossary and additions.json
```

`make_manifest.py` checks every place against the template's current text and stops if one has moved, so an edited template can't be filled in the wrong spot.
