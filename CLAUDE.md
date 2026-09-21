# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Control de Gastos: a couple's shared expense tracker. Single shared fund, no split-by-person accounting. Free, mobile-friendly, built entirely on Google Sheets + Google Apps Script (no external hosting or paid APIs). See [PLAN.md](PLAN.md) for the full rationale and decisions behind the architecture.

## Architecture

- **Data store**: the bound Google Sheet, with exactly three tabs — `Gastos` (Fecha, Categoría, Descripción, Monto), `Categorías` (editable list), `Presupuesto` (Categoría + límite mensual). `setupSheets()` in [Code.gs](Code.gs) creates/repairs these tabs and deletes any other sheet, so it's safe to re-run to reset structure.
- **Backend**: [Code.gs](Code.gs) is the only server-side file. `doGet(e)` routes between the two HTML views based on the `page` query param (`?page=resumen` → Resumen, anything else → Formulario). All spreadsheet reads/writes go through this file's functions (`getCategorias`, `addExpense`, `getPresupuesto`, `getResumen`); there is no other data-access layer.
- **Frontend**: [Formulario.html](Formulario.html) (expense entry) and [Resumen.html](Resumen.html) (monthly summary with a Google Charts pie chart) are self-contained HTML/CSS/JS files rendered via `HtmlService`. They call server functions through `google.script.run` — there is no REST API. Both files link to each other via the `?page=` query param, not client-side routing.
- **No build step**: this is plain HTML/CSS/JS and Apps Script (`.gs`), not a bundled frontend. Edit files directly.

## Development workflow

Code is written locally and pushed to the linked Apps Script project with `clasp` — it is not edited in the Apps Script web editor.

```
clasp push     # upload local files to the Apps Script project
clasp open     # open the project in the Apps Script editor (for manual testing/deploy)
```

There is no local runner, linter, or test suite — Apps Script code only runs inside Google's environment. To verify a change, push and exercise it from the Apps Script editor or the deployed web app.

Deployment is manual via `clasp deploy` or the Apps Script editor's "Deploy > Web app" flow (execute as owner, access shared with both partners or "anyone with the link"), per the plan's "Acceso desde el celular" section.

## Git

Never run `git commit`, `git pull`, or `git push`. Only the user runs these — staging/editing files is fine, but leave all commits and remote sync to them.
