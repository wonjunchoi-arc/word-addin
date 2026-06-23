# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Microsoft **Word** Office Add-in (task pane) for inserting clinical-study-report (CSR) adverse-event cases — section "12.2.2" — into a Word document. It is a thin front end: it fetches case data from a separate backend API and writes formatted HTML into the active document via Office.js. The repo scaffolding comes from the `OfficeDev/Office-Addin-TaskPane-JS` template, but `src/taskpane/` has been customized for the CSR workflow.

## Architecture

Two independent entry points, bundled separately by webpack (see `webpack.config.js` `entry`):

- **Task pane** (`src/taskpane/`) — the main UI. `taskpane.html` defines the markup (a status line, a "run pipeline" button, and a `<template id="case-item-tpl">` cloned per case). `taskpane.js` drives it. This is the only part with real application logic.
- **Commands** (`src/commands/`) — ribbon command handler. Still the unmodified template stub (`commands.js` shows a mailbox notification, which is dead code in a Word host). Touch only if adding ribbon buttons.

### Task pane data flow (`src/taskpane/taskpane.js`)

The add-in is a client of a backend API at the `API` constant (`https://localhost:8000`). All app behavior is HTTP calls to that backend plus Office.js document writes:

- `loadCases()` → `GET /cases` → renders the list (called on `Office.onReady` and after a pipeline run).
- `insertCase(subjectId)` → `GET /cases/{subjectId}` → returns `{ html }`, inserted at document end via `Word.run` + `context.document.body.insertHtml(...)` + `context.sync()`.
- `runPipeline()` → `POST /run` → then re-loads the case list.
- `setStatus(msg)` updates the `#status` element; all errors surface there (UI is Korean).

The backend is **not in this repo** — it must be running on `localhost:8000` for the add-in to do anything. If you change that origin, update the `API` constant **and** the matching `<AppDomain>` in `manifest.xml` (Office blocks cross-origin calls not declared there).

### manifest.xml

The contract Office loads. Host is `Document` (Word). Dev URLs point at `https://localhost:3000`; `webpack.config.js` rewrites `localhost:3000` → the `urlProd` value (`https://www.contoso.com/` — a placeholder, change before any real deployment) only in production builds. `AppDomains` whitelists `localhost:3000` (the add-in itself) and `localhost:8000` (the backend API). Permission level is `ReadWriteDocument`.

## Commands

- `npm run dev-server` — webpack dev server on https://localhost:3000 (HTTPS via `office-addin-dev-certs`).
- `npm start` — sideload + debug the add-in in Word desktop (`config.app_to_debug` in `package.json`); this also starts the dev server.
- `npm stop` — stop the debugging session / unload the add-in.
- `npm run build` — production bundle; `npm run build:dev` for development mode.
- `npm run validate` — validate `manifest.xml` (run after any manifest edit).
- `npm run lint` / `npm run lint:fix` — `office-addin-lint`.

There are no tests in this project.

## Notes

- Office.js is loaded from CDN in `taskpane.html`; do not bundle it. Any code using `Office`/`Word` must run inside or after `Office.onReady`.
- The dev server sets `Access-Control-Allow-Origin: *`, but the backend at `localhost:8000` must independently allow CORS from `localhost:3000` for fetches to succeed.
