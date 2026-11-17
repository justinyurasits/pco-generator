# Change Order Generator

**[Live at draftconstructionchangeorder.com](https://draftconstructionchangeorder.com/)**

Turns a description of a change into a formatted construction change order —
scope language, pricing, and the structure a general contractor actually
issues — as a Word document and a PDF.

**Status:** active, live. Construction OS builds the same document as its first
governed slice, in [`construction-os-workbench`](https://github.com/justinyurasits/construction-os-workbench);
this repo is the standalone public tool.

## Why

A change order is the document that decides whether work performed gets paid
for. It's also written under time pressure by a project manager who is on a
job site, which is why change work routinely gets performed before anything
is authorized in writing, and why that margin is the most commonly eaten
margin in the business.

The slow part isn't the decision. It's producing the document.

## How it works

Pricing comes from a cost database rather than the model, so numbers are
looked up, not generated. Scope language and document structure follow an
established change order standard. The model handles the part it's good at —
turning an informal description of what changed into complete, specific scope
language — and nothing else.

The output is a draft. It is not an authorization, and the project manager
reviews and issues it.

## Precision note

This generates the **Potential Change Order** — the document issued to the
owner for pricing and approval, before any executed change order exists. The
sequence it belongs to is Change Event → PCO → written owner approval →
executed change order. The tool does not shortcut that, and shouldn't.

## Built with

Python, Flask and gunicorn, the Claude API, python-docx. The backend runs on
Render; the front end is served by GitHub Pages at the custom domain.

## Limits

**The pricing database is embedded at a fixed revision.** When rates change it
has to be updated and redeployed — there's no live sync. Numbers are looked
up rather than invented, but they're only as current as the last update.

**The rate sheet belongs to one company.** A different contractor, market, or
contract structure would get numbers grounded in a cost model that isn't
theirs.

**No document history.** Each generation is independent. Output is emailed and
that's the end of it — there's no way to retrieve a past draft or iterate on
one.

## Running and testing it

```bash
pip install -r requirements.txt
python app.py         # local server; needs ANTHROPIC_API_KEY and RESEND_API_KEY in the environment
python test_run.py    # runs the generator on realistic, messy field notes
```

`test_run.py` makes real API calls: it's an output-quality check, not a unit test suite.
