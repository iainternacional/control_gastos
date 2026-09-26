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
`getCategorias`, `addCategoria`, `addExpense`, `getPresupuesto`, `getCategoriasIngresos`,
`addCategoriaIngreso`, `addIngreso`, `getResumen`.

`include(filename)` in `Code.gs` is dead code — all three HTML files are self-contained. There
are no partials; do not factor shared markup out without also wiring a real include mechanism.
`Formulario.html` and `Ingreso.html` are near-duplicates by design (same "+ agregar categoría"
UX, same layout) — a change to one's category-adding flow almost certainly needs the same
change in the other.

## Code that does not do what it looks like

- `doGet` now has three explicit branches (`resumen`, `ingreso`, else → `Formulario`) and injects
  `appUrl` into every template. Relative `<a href="?page=...">` links do **not** work reliably in
  Apps Script HtmlService — pages render inside a sandboxed `googleusercontent.com` iframe, so a
  relative href resolves against that iframe's URL, not the `/exec` URL. Every nav link must be
  built as `<?= appUrl ?>?page=...`, never a bare relative href.
- **`appUrl` is the hardcoded `APP_URL` constant, not `ScriptApp.getService().getUrl()`.**
  Verified by failure: this deployment is restricted to the `ipuc.org.co` Workspace domain, whose
  correct URL shape is `https://script.google.com/a/macros/ipuc.org.co/s/<id>/exec`.
  `ScriptApp.getService().getUrl()` returned a URL that did not match that shape, and navigating
  to it from the phone failed with Google's "No se puede abrir el archivo en estos momentos" —
  even though the original bookmarked `/a/macros/.../exec` URL loaded fine. `APP_URL` in
  `Code.gs` is now the single source of truth for the app's own base URL; if the user ever
  creates a new deployment (as opposed to a new version of the same one), `APP_URL` must be
  updated by hand to the new `/exec` URL, same as `SPREADSHEET_ID`.
- **Multi-account phones need `authuser` preserved, not just `appUrl`.** On a phone with several
  Google accounts signed in, the URL Google actually loads is `.../exec?authuser=N&page=...`.
  `ScriptApp.getService().getUrl()` never includes `authuser`, so a nav link built as plain
  `<?= appUrl ?>?page=X` silently drops back to account 0 on click — if that account lacks access
  to the Sheet, the target page fails to load with no visible error. `doGet` reads
  `e.parameter.authuser` and injects it as `template.authSuffix` (`'&authuser=N'` or `''`); every
  nav link must append `<?= authSuffix ?>` after its `?page=...`, in all three HTML files.
- `getResumen(anio, mes)` takes a month and year, but `Resumen.html` calls `getResumen()` with
  no arguments, so the summary is always the current month. Both parameters are effectively dead.
- `getResumen()`'s `porcentajeGastado` and each detalle item's `porcentajeIngreso` are `null`
  (not `0` or `NaN`) when `totalIngresos` is `0` for that month — the frontend must check for
  `null` before formatting, not just falsiness.
- Currency formatting is hardcoded in `Resumen.html` as `'$' + n.toLocaleString('es-CO', …)`.
  Changing currency or number separators means editing that one line; there is no locale config.
- `Ingresos`/`CategoriasIngresos` have no `Presupuesto`-style budget/limit concept — that only
  exists for expenses. Don't assume symmetry beyond the Fecha/Categoría/Descripción/Monto shape.

## `setupSheets()` is destructive — never run it casually

`setupSheets()` in `Code.gs` **deletes every sheet that is not** `Gastos`, `Categorías`,
`Presupuesto`, `Ingresos`, or `CategoriasIngresos`. If the user has added a tab for their own
bookkeeping, this wipes it. It only seeds headers when `getLastRow() === 0`, so re-running is
safe for the five managed tabs and nothing else.

## Editing gotchas

- `getResumen` buckets dates with `fecha.getFullYear()` / `getMonth()`, so month boundaries
  follow `appsscript.json`'s `timeZone` (`America/Bogota`), not the phone's locale.
- `setupSheets()` is the only place `CATEGORIAS_DEFAULT` is written; the `Categorías` and
  `Presupuesto` tabs are user-editable afterwards, so code must read them, not assume defaults.
- There is no `.gitignore`; `.clasp.json` (which contains the scriptId) is committed.
