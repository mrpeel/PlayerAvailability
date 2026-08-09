# Laburnum CC Availability Tracker - Deployment Guide

## Architecture

```
Browser (Firebase static hosting)
   │
   │  POST /api
   ▼
Cloud Function (thin proxy)
   │
   │  POST Content-Type: text/plain → GAS exec URL
   ▼
Google Apps Script web app (owner account)
   │
   ├── SheetLogic.gs  (doGet, doPost, getInitialData, saveAvailability, dialogs)
   ├── Setup.gs       (setupInitialSystem)
   └── logic.js       (pure helpers, tested with Jest)
   │
   ▼
Google Sheet + Drive
```

The GAS web app runs as the **deploying owner account** (`executeAs=USER_DEPLOYING`), so no service account, spreadsheet ID, or API key ever leaves the server. The Cloud Function is a ~20-line proxy — its only secret is the GAS web app URL.

## Prerequisites

1. **Google Account** — owns the sheet and script.
2. **Node.js >=22** and npm.
3. **Enable Apps Script API**: [script.google.com/home/usersettings](https://script.google.com/home/usersettings) — set to **On**.
4. **Firebase project** on the Blaze plan (required for Functions with secrets).

---

## 1. Google Apps Script — one-time setup

The script runs **inside a Google Sheet** (container-bound).  Every file in `src/`
is pushed to that script.  The sheet is your database; the script is your logic
engine.  A single `.clasp.json` at the repo root maps `src/` to the script.

### Step-by-step

**1.** In Google Drive, navigate to the folder where you want the sheet to live.

**2.  File > New > Google Sheets.**  Name it "Laburnum CC Availability".

**3.** Open the sheet, then **Extensions > Apps Script.**

**4.** In the Apps Script editor's URL bar, copy the **script ID** — the long string
   between `/d/` and `/edit`.

> Example: `https://script.google.com/d/1ABCxyz123456/edit` → script ID is `1ABCxyz123456`.

**5.** Back in the repo:

```bash
npx clasp login

# Write the config file (replace the ID with yours)
echo '{"scriptId":"<PASTE_YOUR_SCRIPT_ID>","rootDir":"src"}' > .clasp.json

# Push code to the script (tests run first)
npm run push
```

**6.** Refresh the Apps Script editor.  Select **`setupInitialSystem`** from the
   toolbar function dropdown and click **Run**.  Grant permissions when prompted.
   The `Players`, `Fixtures`, `Availability_Log`, and `Presentation_Staging` tabs
   appear in the sheet.

**.clasp.json** is **gitignored** (contains your private script ID).  Share it
with teammates via a password manager, never git.  On a new machine repeat steps
4–5 (copy the existing script ID, paste into a fresh `.clasp.json`).

## 2. Populate the sheet

1. Open the Google Sheet you created in step 1.
2. Populate the `Players` and `Fixtures` tabs with real data.
3. Add admin phone numbers to an `Admins` tab if using selector features.

## 3. Deploy the GAS web app

1. In the Apps Script editor: **Deploy > New Deployment**.
2. Type: **Web App**.
3. Settings:
   - **Description**: `Initial deployment`
   - **Execute as**: `Me (your-email)`
   - **Who has access**: `Anyone`
4. Click **Deploy**.
5. **Copy the Web App URL** — you'll need it for the next step.

## 4. Firebase — proxy function + hosting

```bash
npm install            # installs firebase-tools, jest, clasp etc.
cd functions && npm install && cd ..

# Check your Firebase project is selected
npx firebase use
# If wrong:  npx firebase use --add

# Set the GAS URL as a secret (paste the URL from step 3)
npx firebase functions:secrets:set GAS_EXEC_URL
# Paste the GAS web app URL when prompted and press Enter.

# Deploy everything (function + static hosting)
npm run deploy
```

The function proxy now lives at `https://<project>.web.app/api`. The frontends in `public/` call `/api` unchanged — they never see the GAS URL.

## 5. Updating the app

| What changed | Command | Extra step |
|:---|:---|:---|
| GAS backend (`src/`)| `npm run push` | Tests run first; `clasp push -f` pushes non-interactively. Then **bump the deployment version**: Apps Script editor > Deploy > Manage Deployments > pencil icon > New Version. The live URL runs old code until you do this. |
| Firebase (function + frontend) | `npm run deploy` | None — redeploys automatically. |

## 6. Shareable links

- Player availability: `https://<project>.web.app/?date=YYYY-MM-DD`
- Admin/selector dashboard: `https://<project>.web.app/admin.html?date=YYYY-MM-DD`
- The URL parameter `?date=2026-10-04` pre-loads that specific match date's tab.

The same URL works on the raw GAS link: `https://script.google.com/macros/s/<id>/exec?date=2026-10-04`.

## 7. Create a new match-date tab

Open the Google Sheet, then from the toolbar: **🏏 LCC Selection > Initialise new round tab**. Pick the first day of play. The `deployVerticalRoundSheet` function in `SheetLogic.gs` builds the full team + 3-list ledger.

## 8. Ongoing maintenance

- **Adding players**: Add rows to the `Players` tab with `ProfileID`, `First Name`, `Last Name`, `Full Name`, `GlobalStatus`, phone columns, `JuniorClass`. The `Full Name` column auto-fills from `First Name + Last Name` via `ARRAYFORMULA`.
- **Adding fixtures**: Add rows to the `Fixtures` tab with `RoundID`, `Game Date`, `Match Format`, and opponent/venue details.
- **Mid-week sync**: On an active date tab, run **LCC Selection > Pull new players from Master** to pick up new registrations.
- **Injuries**: **LCC Selection > Record player injury / absence** updates the master `Players` tab and moves the player to Unavailable on the current date tab.

## Troubleshooting

| Problem | Fix |
|:---|:---|
| `npm run push` fails with "Could not find script" | Check `.clasp.json` `scriptId` matches the script URL; verify `npx clasp login` |
| Live GAS app returns old code | You pushed code but didn't bump the deployment version (see section 5) |
| Proxy returns 500 | Check `GAS_EXEC_URL` is set: `npx firebase functions:secrets:access GAS_EXEC_URL` |
| Proxy returns 400 | GAS returned an error (missing tab, data issue) — check the response body |
| `functions:secrets:set` permission denied | You need a Firebase Blaze plan. Check permissions in IAM. |
