# Summary of the Laburnum Cricket Club (LCC) Selection System Documentation

The generated Markdown document (**`LCC_Cricket_Selection_System_Summary.md`**) provides a complete architectural blueprint, codebase specification, and unit testing guide for the Laburnum Cricket Club selection system. 

Here is an executive summary of what is contained inside the document:

---

### 1. Architectural Design & Philosophy
* **Date-Driven Ledger Pattern (`YYYY-MM-DD`):** Replaces sequential round numbers with absolute match dates (e.g., `2026-10-04`), creating frozen, standalone historical records for every weekend or midweek fixture.
* **Roster State vs. Selection Grid Separation:**
  * **Roster Availability State:** Managed across three closed-loop lists (*Available for Round*, *Unavailable for Round*, *Unknown Availability*). Every active club member exists in exactly one list.
  * **Selection Filter:** The *Available for Selection* panel dynamically filters *Available for Round* by subtracting players assigned to team slots. Selecting or unselecting a player in a team sheet never alters their underlying availability status.
* **PlayHQ `ProfileID` Key:** Uses PlayHQ GUIDs (e.g., `dbe01945-2937-4fa1-87fa-7e11e223e599`) as the primary relational key in hidden sheet columns and API payloads, while displaying clean player names to selectors.
* **4-Phone Household Matching:** Matches incoming logins and availability submissions against `Phone`, `Phone2`, `Phone3`, and `Phone4` to support junior players and family accounts.
* **Deterministic Headshots:** Stores photos in Google Drive as `<ProfileID>.jpg` for automatic loading across Sheets (`=IMAGE()`), Slides, and the Web App.
* **Presentation Staging Hub:** Google Slides links permanently to `Presentation_Staging`. Changing cell `B1` to point to a new date sheet updates all presentation slides instantly.

---

### 2. Sheet Setup & Database Schemas (`Setup.gs`)
* **Single Master Function:** `setupInitialSystem()` builds all initial database tables in sequence:
  * **`Players`:** Stores `ProfileID`, `First Name`, `Last Name`, `Full Name`, `GlobalStatus` (`Active`, `Injured`, `Long-Term Away`, `Inactive`), `ExpectedReturnDate`, `Phone`, `Phone2`, `Phone3`, `Phone4`, `Email`, `T20_Squad`, and `JuniorClass` (`Adult`, `U18` `*`, `U16_Y2` `^`, `U16_Y1` `!`, `U14` `~`).
  * **`Fixtures`:** Wide table capturing grade-by-grade opponents, venues, and formats (One Day / Two Day / T20).
  * **`Availability_Log`:** Audit log capturing timestamps, profile IDs, match dates, responses, and notes.
  * **`Presentation_Staging`:** Switchboard tab with `=INDIRECT()` formulas for live slide linking.

---

### 3. Operational Selection Tab Design (`YYYY-MM-DD`)
* **Horizontal Grid Layout (Left to Right):**
  * **Cols A–B:** Mobile-stacked vertical team selection blocks (1st XI down to 5th XI).
  * **Col D:** Dynamic virtual pool filtering out assigned team players.
  * **Cols F–I:** *Available for Round* (Col F = Hidden `ProfileID`, Col G = Player Name + Junior Tag, Col H = Notes, Col I = `🚫 Move to Unavailable`).
  * **Cols K–N:** *Unavailable for Round* (Col K = Hidden `ProfileID`, Col L = Player Name, Col M = Exemption Notes, Col N = `✅ Move to Available`).
  * **Cols P–S:** *Unknown Status* (Col P = Hidden `ProfileID`, Col Q = Player Name, Col R = Notes, Col S = `✅ Move to Available` / `🚫 Move to Unavailable`).

---

### 4. Sheet Logic & Runtime Workflows (`SheetLogic.gs`)
* **HTML5 Datepicker Dialog:** Modal window that visually prompts selectors for the first day of play.
* **Real-Time `onEdit` Trigger:** Intercepts emoji dropdown clicks in columns I, N, and S to shift 3-cell blocks (`ProfileID`, `Name`, `Notes`) between lists and compact source rows automatically.
* **Mid-Selection Two-Way Sync:**
  * `pullNewPlayersToActiveTab`: Pulls mid-week PlayHQ registrations into the active tab's *Unknown* list.
  * `showRecordInjuryDialog`: Updates `GlobalStatus = "Injured"` in `Players` tab and shifts the player to *Unavailable* on the active tab.
  * `showMarkInactiveDialog`: Updates `GlobalStatus = "Inactive"` in `Players` tab and shifts the player to *Unavailable*.
* **Real-Time Web App Integration (`doGet` & `doPost`):**
  * `getInitialData`: Fetches household players across all 4 phone fields, fixture info, and current selection status from the `YYYY-MM-DD` tab.
  * `saveAvailability` / `doPost`: Moves players between availability lists using hidden `ProfileID` columns in real time.

---

### 5. Unit Testing & QA Strategy
* **Core Logic vs. Adapter Pattern:** Decouples pure JavaScript business logic from Google Apps Script APIs (`SpreadsheetApp`, `DriveApp`) so unit tests run locally in Node.js using **Jest**.
* **Local Test Suites Included in Summary:**
  * **Phone Normalization:** Tests E164 formatting (`+614...`) and junior suffix tag appending.
  * **Household Matching:** Tests 4-phone lookup across `Phone1`–`Phone4`.
  * **Array List Movement:** Tests 3-cell row moving and list compaction logic.
* **Pre-Push Automation Pipeline:** Configures `"push": "npm test && clasp push"` in `package.json` to guarantee that code is only deployed to Google Apps Script after all unit tests pass.