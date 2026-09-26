# AGENTS.md

Operating notes for AI agents in this repo. Product rationale lives in [PLAN.md](PLAN.md);
[CLAUDE.md](CLAUDE.md) covers the architecture at a high level. This file covers the traps.

## Hard constraints

- **Never run `clasp push` or `clasp deploy`.** Prohibited outright — uploading and deploying
  belong to the user. After changing code, stop and tell them what still needs pushing.
- **Never run `git commit`, `git pull`, or `git push`.** The user owns all commits and remote
  sync. Reading status and editing files is fine.
- **No build, test, lint, or typecheck exists.** No `package.json`, no local runner, no CI.
  Do not invent npm scripts, and do not describe a change as verified — Apps Script only
  executes inside Google's runtime, and the agent cannot reach that runtime. Verification is
  the user pushing and exercising the page in the Apps Script editor or the deployed web app.
- Read-only tooling is allowed: `git status` / `git log`, `clasp status`, and `clasp open`.

## Stack

Plain Google Apps Script (V8) plus self-contained HTML/CSS/JS. No bundler, no framework, no
dependencies. `clasp` is the only dev tool; `.clasp.json` pins this directory to a
**standalone** Apps Script project (see the next section).

## The Sheet is resolved by ID, not by binding — this is not optional

**Verified by failure:** `SpreadsheetApp.getActiveSpreadsheet()` returns `null` in this
project. The script is *standalone*, not container-bound, so every data function threw
`TypeError: Cannot read properties of null (reading 'getSheetByName')`.

`Code.gs` therefore defines `SPREADSHEET_ID` and a `getSheet()` helper
(`SpreadsheetApp.openById`), and all five data call sites go through `getSheet()`. Do not
"simplify" those calls back to `getActiveSpreadsheet()`.

A standalone script cannot be bound to a Sheet after the fact, so `openById` is the only fix.
Consequences worth remembering:

- `SPREADSHEET_ID` in `Code.gs` is set to the real Sheet ID. If the user ever rotates or
  recreates the Sheet, this constant is the single place to update.
- The Sheet must be shared with the Google account that executes the script, or `openById`
  throws instead. Opening it in that same account is enough for this single-owner setup.
- Because the script is not "on" the Sheet, it will not appear in the Sheet's Extensions
  menu, and `setupSheets()` must still be run once by hand.

## What `clasp push` actually uploads

*Reference only — the agent does not run this.* Only `.gs`, `.js`, `.html`, and
`appsscript.json` are uploaded: the extensions declared in `.clasp.json` (`.clasp.json` itself
is always local). So `clasp status` lists `CLAUDE.md`, `PLAN.md`, `.clasp.json`, and `.git\` as
"Untracked": that is clasp's wording for "not uploaded to Apps Script", **not** git status. Edits
to the Markdown docs never need a push.

Pushing is not deploying. `clasp push` updates the *dev* version only; the deployed `/exec` URL
keeps serving the last deployed version until the user runs `clasp deploy` (or Deploy > Manage
deployments > edit > New version). This is the usual cause of "I pushed and nothing changed", and
it means a code fix is not live until the user does both.

## Server function names are an unguarded contract

No bundler or compiler checks that `google.script.run.<fn>()` in the HTML matches
`function <fn>()` in `Code.gs`. Renaming or removing a server function breaks the page at
runtime with no static error. The current public surface, called from the HTML, is exactly:
`getCategorias`, `addExpense`, `getPresupuesto`, `getResumen`.

`include(filename)` in `Code.gs` is dead code — both HTML files are self-contained. There are
no partials; do not factor shared markup out without also wiring a real include mechanism.

## Code that does not do what it looks like

- `doGet` only tests `page === 'resumen'`; **every other value falls through to Formulario**.
  That is why `Resumen.html`'s `?page=form` link works at all. Adding a third page means
  changing that condition, not just adding an HTML file.
- `getResumen(anio, mes)` takes a month and year, but `Resumen.html` calls `getResumen()` with
  no arguments, so the summary is always the current month. Both parameters are effectively dead.
- Currency formatting is hardcoded in `Resumen.html` as `'$' + n.toLocaleString('es-CO', …)`.
  Changing currency or number separators means editing that one line; there is no locale config.

## `setupSheets()` is destructive — never run it casually

`setupSheets()` in `Code.gs` **deletes every sheet that is not** `Gastos`, `Categorías`, or
`Presupuesto`. If the user has added a tab for their own bookkeeping, this wipes it. It only
seeds headers when `getLastRow() === 0`, so re-running is safe for the three managed tabs and
nothing else.

## Editing gotchas

- `getResumen` buckets dates with `fecha.getFullYear()` / `getMonth()`, so month boundaries
  follow `appsscript.json`'s `timeZone` (`America/Bogota`), not the phone's locale.
- `setupSheets()` is the only place `CATEGORIAS_DEFAULT` is written; the `Categorías` and
  `Presupuesto` tabs are user-editable afterwards, so code must read them, not assume defaults.
- There is no `.gitignore`; `.clasp.json` (which contains the scriptId) is committed.
