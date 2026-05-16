# Google Sheets Setup Guide

This document outlines the required structure for the Google Sheet used by the Laburnum CC Availability Tracker.

## Sheet 1: `Players`
This sheet maintains the player database.

| Column Name | Description | Example |
| :--- | :--- | :--- |
| **ID** | Unique identifier for the player (can be numeric) | `101` |
| **FirstName** | Player's first name | `Jimi` |
| **FamilyName** | Player's last name | `Kloot` |
| **Phone** | Primary mobile number (normalized by app) | `0412345678` |
| **Phone2** | (Optional) Second family number | `0400000000` |
| **Phone3** | (Optional) Third family number | |
| **Phone4** | (Optional) Fourth family number | |
| **GlobalStatus** | Current state: `Active`, `Injured`, `Away` or `Finished` | `Active` |
| **ExpectedReturnDate** | Reference for players who are `Away` | `15/12/2026` |

> [!NOTE]
> The app matches users by checking all Phone columns. If a family shares a number, they will all appear in the dashboard.

---

## Sheet 2: `Rounds`
This sheet defines the matches for the season.

| Column Name | Description | Example |
| :--- | :--- | :--- |
| **RoundNum** | The round number (used in URL `?round=X`) | `3` |
| **Date1** | Date of Day 1 (format: DD/MM/YYYY) | `04/06/2026` |
| **Date2** | (Optional) Date of Day 2 for 2-Day games | `11/06/2026` |

> [!TIP]
> The app automatically identifies a round as a **2-Day** game if **Date2** has a value. If it is empty, the app treats it as a **1-Day** game.

---

## Sheet 3: `Availability`
This sheet stores the responses submitted by players. The app appends new rows here.

| Column Name | Description |
| :--- | :--- |
| **Timestamp** | ISO string of when the response was sent |
| **PlayerID** | Matches the **ID** from the `Players` sheet |
| **RoundNum** | The round being responded to |
| **Response** | The selection (e.g., `Available`, `Both Days`, `Unavailable`) |

---

## Technical Configuration
1. **Share the Sheet**: The Google Service Account (found in your `GOOGLE_SERVICE_ACCOUNT` secret) must be added as an **Editor** to this Google Sheet.
2. **Sheet Names**: Ensure the tabs are named exactly `Players`, `Rounds`, and `Availability` (case-sensitive).
3. **Headers**: The first row of each sheet must contain the exact column names listed above.

---

## Sheet 4: `Admins`
Used for restricted access to the Admin Dashboard.

| Column Name | Description |
| :--- | :--- |
| **Name** | Name of the admin |
| **Phone** | Mobile number (used for login) |

---

## Sheet 5: `Messages` (Or updated `Rounds`)
To support the Admin Dashboard, the app expects columns for match messages. You can add these to the **`Rounds`** sheet.

| Column Name | Description |
| :--- | :--- |
| **AvailabilityMessage** | The text for the availability call-out |
| **WallOfShameMessage** | The text for follow-ups (unresponsive players) |
