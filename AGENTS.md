# Laburnum CC Availability Tracker — Agent Guide

Cross-tool project context for coding agents (OpenCode, Antigravity, and any
tool that reads `AGENTS.md`). Keep all tool-specific config (`.opencode/`,
`.antigravity/`, custom agents/skills) OUT of this file and its sibling docs.

## What this app is

A mobile-first web app for Laburnum Cricket Club players to submit availability
for each round, and for selectors to run the sheet (view/move players between
Available / Unavailable / Unknown lists per match date).

## Architecture (important)

- **Google Apps Script (`src/`) is the logic engine.** It owns the spreadsheet
  schema, availability retrieval/saving, household matching, list movement, and
  the selector features. Deployed with `clasp`.
- **Firebase** hosts only the static frontend (`public/`) and a thin Cloud
  Function proxy. The proxy forwards `/api` requests to the GAS web app URL
  server-side — no Sheets credentials ever appear in browser code.
- Both share one Google Sheet workbook.

## Security model (why no secrets live in the frontend)

- The GAS web app runs as the **deploying owner account**
  (`appsscript.json`: `executeAs=USER_DEPLOYING`, `access=ANYONE_ANONYMOUS`).
  The one-time owner authorization to the spreadsheet + Drive lives in the Apps
  Script project, never in the browser.
- Firebase hosts only static HTML/JS. The frontend calls the GAS web app URL
  (`https://script.google.com/macros/s/<id>/exec`) via `fetch` → `doPost`,
  exactly as it used to call `/api`. No service account, API key, or
  spreadsheet ID is embedded in frontend code.
- App-level auth still applies: players authenticate by registered phone;
  selector/admin access is gated by the `Admins` tab. Never treat the exec URL
  as a secret.
- CORS: browser calls must POST with `Content-Type: text/plain` (no preflight);
  `doPost` returns JSON via `ContentService`.

## Two-layer rule (enforced)

App scripts are split so the logic is testable in Node:

| Layer | Location | Rules |
| :--- | :--- | :--- |
| **Pure business logic** | `src/logic.js` | MUST NOT call `SpreadsheetApp`, `DriveApp`, `Ui`, `HtmlService`, etc. Plain JS in / plain JS out. |
| **GAS adapters** | `src/SheetLogic.gs`, `src/Setup.gs` | Thin wrappers that read/write the sheet and call the pure layer. |

- Add pure functions to `logic.js`, never call a GAS service inside it.
- `logic.js` has a guarded `module.exports` so Jest can `require()` it while
  Apps Script still treats the functions as globals.
- Pure logic is covered by Jest (`npm test`); GAS adapters are covered with
  lightweight global mocks (`tests/gasMock.js`, loaded via jest `setupFiles`).

## Current gaps (bring-in in progress)

- `public/index.html` (Firebase-hosted player UI) — payload shape and field names
  don't match GAS responses (e.g. expects `roundInfo`, receives `fixtureInfo`).
  GAS `doPost` already accepts old field aliases (`roundNum`, `playerId`), so
  just the rendering side needs updating.
- `public/admin.html` (selector dashboard) — calls retired `getAdminRounds`/
  `getAdminData` actions that GAS doesn't implement yet. Needs rewiring to the
  new selector features in SheetLogic.gs.
- Selector/admin gating (phone check against `Admins` tab) needs adding to
  GAS `doPost`.

## Commands

```bash
npm test          # Jest unit tests (pure logic, runs in <1s)
npm run test:watch
npm run push      # npm test && clasp push — refuses to deploy failing code
npm run deploy    # firebase deploy (static hosting only)
npx clasp open    # open the Apps Script editor
```

## Deploy notes

- `src/` is the clasp `rootDir` (see `.clasp.json`); `.claspignore` keeps
  package files and docs out of the Apps Script project.
- After any `clasp push`, bump the deployment version in the Apps Script editor
  (Deploy > Manage Deployments > New Version) or the live URL runs old code.
- Do not commit `.clasp.json`, `service-account.json`, or `.env` (secrets).

## Conventions

- Follow the heritage palette in `design.md` (`#4d0012` maroon / `#fac218` gold).
- Mobile-first, touch targets >= 48px, minimum font 16px, Hanken Grotesk.
- No emojis in code unless requested.
