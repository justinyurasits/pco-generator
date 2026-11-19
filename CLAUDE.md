# pco-generator

<!-- standard: design/_template/repo/CLAUDE.md v1.2
     Sections 1–6 are the standard: copy them unchanged, filling in only the <angle brackets>.
     Section 7 is this repository's own. To change a standard rule here, say so under
     "Overrides" in section 7, with the reason. -->

## 1. What this is

The PCO Generator: turns a described change into a formatted potential change order
(scope language, line-item pricing, Word and PDF), live at draftconstructionchangeorder.com.
One of the Construction OS demos; standalone, sharing no code with the workbench.
Status: active, live.

## 2. Design

The design docs for this repository are in the `design` repository, cloned next to this
one, at `../design/construction-os-demos/pco-generator/`. Pull it before planning. In a git worktree,
`../design` means next to the main checkout, not next to the worktree.

Before proposing work, read in this order (skip any the project doesn't have yet):

1. `../design/construction-os-demos/brief.md`: why it exists and for whom
2. `../design/construction-os-demos/product-spec.md`: what the product must do, by capability ID
3. `../design/construction-os-demos/architecture.md` and `principles.md`: how the parts fit, and
   the rules every design decision follows
4. `../design/construction-os-demos/pco-generator/brief.md` and `spec.md`: who this component is for,
   and what it must do
5. `../design/construction-os-demos/pco-generator/design.md`: how it works, and the decision log
6. `../design/construction-os-demos/pco-generator/<slice>-design.md`: the build design for the slice
   the issue belongs to
7. `../design/construction-os-demos/pco-generator/plan.md`: the milestones still ahead

- **Design leads, code follows.** Never edit design docs from this repository. If the
  work needs a change to what the product does, to a requirement or to the architecture,
  stop and propose the change. Don't build ahead of the design.
- **When the design is silent, ask.** Don't improvise behavior, copy or scope.
- **When the design and this repository disagree, say so.** Don't guess which is right.
- A synced design file starts with a header naming its Claude Doc. If the two differ,
  the Claude Doc is current.

## 3. How work happens

- **The issue is the plan.** Work only from a GitHub issue, and read it first
  (`gh issue view <N>`). The issue authorizes the work; the spec and design constrain
  how it's done. If they conflict, stop and say so. With no issue, write the work up as
  a proposed issue and wait.
- One issue, one branch, one pull request. Branch name: `issue-<N>-<short-description>`.
- Anything outside the issue gets a new issue. Don't widen the scope.
- Never commit to `main`. Never merge: Justin reviews and merges.

## 4. Checks

- Run the checks listed in section 7 before opening a pull request, and paste the
  results into it.
- A failure is never ignored. Either the change is wrong or the check is; say which.
- Never change an expected result to make a check pass. Report the mismatch and stop.

## 5. Every pull request

Fill in `.github/pull_request_template.md`:

- `Closes #<N>`, and what changed and why, in plain English.
- How each acceptance criterion in the issue was checked.
- **Decisions (proposed):** choices made while coding: the decision, the options, why,
  and what would change it. Merging the pull request approves them.
- **Build notes:** what you did, what broke, what it taught.

## 6. Roll-up, ask before, never

**Roll-up.** Ask "Roll up to design?" when a pull request you opened is merged, when an
issue or milestone is done, and about every ten turns of work since you last asked.

- On yes: collect Decisions and Build notes from pull requests merged since the last
  roll-up (the latest "Roll-up:" line in the component's `build-log.md`). Propose
  decisions for the `design.md` decision log and findings for `build-log.md`, ending with
  the line `Roll-up: <repo> #<first>–#<last>`. For files synced from Claude Docs, add
  them as suggestions in the Claude Doc if you have the Docs connector; otherwise give
  Justin the text to paste. For repo-authored files, open a pull request in `design`.
- On no: don't ask again for another ten turns.
- Nothing enters the design record until Justin approves it.

**Ask before:**

- adding a dependency
- changing a prompt, tool description, rubric, or anything else that changes what a
  model decides
- anything not in the issue

**Never:**

- read, print or edit `.env`. Secrets live only there; `.env.example` holds key names only
- commit secrets, API keys or real customer data
- rewrite git history

## 7. This repository

- **Layout:** `app.py` is the Flask app (request handling, email delivery through
  Resend); `generator.py` holds the prompt, the embedded SYS-09 rates and the document
  build; `index.html` is the front end; `test_run.py` runs the generator on messy field
  notes; `Procfile` starts gunicorn; `CNAME` is the custom domain. README.md describes
  the same files for people; when one is added, moved or retired, update both.
- **Deploys:** merging to `main` deploys the live site: the backend on Render, the
  front end on GitHub Pages. Treat every merge as a release.
- **Run:** `pip install -r requirements.txt`, then `python app.py`. Needs
  `ANTHROPIC_API_KEY` and `RESEND_API_KEY` in the environment.
- **Checks:** `python -m py_compile app.py generator.py` on every pull request. Run
  `python test_run.py` when the prompt, the rates or the document layout change, and
  paste what changed in the output. It makes real API calls, so don't run it otherwise.
- **Conventions:**
  - Prices come from the SYS-09 rates or the contractor's own inputs, never from the
    model. A change that lets the model invent a number is a bug.
  - The output is a potential change order, a draft the project manager reviews and
    issues. It never authorizes work.
- **Overrides:** none.
