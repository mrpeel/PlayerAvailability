# Google Sheets Setup Guide

This document outlines the required structure for the Google Sheet used by the Laburnum CC Availability Tracker.

## Sheet 1: `Players`
Maintains the player database, synchronized from PlayHQ exports.

| Column Name | Description | Example |
| :--- | :--- | :--- |
| **ProfileID** | PlayHQ GUID or unique player ID | `dbe01945-2937-4fa1-87fa-7e11e223e599` |
| **FirstName** | Player's first name / preferred name | `Jimi` |
| **LastName** | Player's surname | `Kloot` |
| **FullName** | Complete display name | `Jimi Kloot` |
| **JuniorLevel** | Junior age group tag (`U18`, `U16`, `U14`, or blank) | `U18` |
| **T20Squad** | T20 squad designation (optional) | |
| **GlobalStatus** | Current state: `Active`, `Injured`, `Long-Term Away`, `Inactive` | `Active` |
| **ExpectedReturnDate** | Reference for injured or away players | `15/12/2026` |
| **Phone** | Primary mobile number (account holder) | `+61412345678` |
| **Phone2** | Second mobile number (Parent/Guardian 1) | `+61400000000` |
| **Phone3** | Third mobile number (Parent/Guardian 2) | |
| **Phone4** | Fourth mobile number | |
| **Email** | Contact email address | `player@example.com` |

> [!NOTE]
> The app matches players by checking all 4 phone columns. If a household shares a phone number, all associated players appear in the submission UI.

---

## Sheet 2: `Config`
Stores team mapping configuration and WhatsApp communication templates.

### Table 1: Team Configuration
Maps internal grade names to Play Cricket / PlayHQ team names for fixture matching:

| Internal Team Name | Competition | Play Cricket Team Name |
| :--- | :--- | :--- |
| `1st XI` | `BHRDCA Senior Competition` | `Laburnum - 1st XI` |
| `2nd XI` | `BHRDCA Senior Competition` | `Laburnum - 2nd XI` |
| `3rd XI` | `BHRDCA Senior Competition` | `Laburnum - 3rd XI` |
| `4th XI` | `BHRDCA Senior Competition` | `Laburnum - 4th XI` |
| `5th XI` | `BHRDCA Senior Competition` | `Laburnum - 5th XI` |
| `T20 1st XI` | `BHRDCA T20 Competition` | `Laburnum T20 1st XI` |
| `T20 2nd XI` | `BHRDCA T20 Competition` | `Laburnum T20 2nd XI` |

### Table 2: WhatsApp & System Templates

| Template Name | Template Text |
| :--- | :--- |
| `Availability Callout` | `🏏 *LCC ROUND AVAILABILITY* 🏏\nPlease submit your availability for this round: {url}` |
| `Wall of Shame` | `🚨 *WALL OF SHAME* 🚨\nThe following players have not yet entered their availability: {players}` |
| `Selection Announcement` | `🏏 *LABURNUM CC TEAMS - {date}* 🏏\n\n*1st XI vs {1st_opponent}* ({1st_venue}, {1st_format})\n{1st_team}\n...` |

---

## Sheet 3: `Fixtures`
Stores match fixtures by match date, imported automatically from PlayHQ (Home and Away CSV exports). `Game Date` is the primary key for the row, and each configured team has dedicated columns for its specific Round number, Match Format, Opponent, and Venue:

| Column Name | Description | Example |
| :--- | :--- | :--- |
| **Game Date** | Match date in `YYYY-MM-DD` format (primary key) | `2025-10-04` |
| **1st Round** | 1st XI round number | `Round 1` |
| **1st Format** | 1st XI match format (`One Day`, `Two Day`, `T20`) | `Two Day` |
| **1st Opponent** | 1st XI opponent | `Mitcham - 2nd XI` |
| **1st Venue** | 1st XI venue | `Kalang Park` |
| **2nd Round** | 2nd XI round number | `Round 1` |
| **2nd Format** | 2nd XI match format (`One Day`, `Two Day`, `T20`) | `Two Day` |
| **2nd Opponent** | 2nd XI opponent | `Bulleen Templestowe - 2nd XI` |
| **2nd Venue** | 2nd XI venue | `Ted Ajani Reserve` |
| **3rd Round** | 3rd XI round number | `Round 1` |
| **3rd Format** | 3rd XI match format (`One Day`, `Two Day`, `T20`) | `Two Day` |
| **3rd Opponent** | 3rd XI opponent | `Box Hill North Super Kings - 4th XI` |
| **3rd Venue** | 3rd XI venue | `Eley Park` |
| **4th Round** | 4th XI round number | `Round 1` |
| **4th Format** | 4th XI match format (`One Day`, `Two Day`, `T20`) | `Two Day` |
| **4th Opponent** | 4th XI opponent | `Nunawading - 5th XI` |
| **4th Venue** | 4th XI venue | `Mahoneys Reserve` |
| **5th Round** | 5th XI round number | `Round 1` |
| **5th Format** | 5th XI match format (`One Day`, `Two Day`, `T20`) | `Two Day` |
| **5th Opponent** | 5th XI opponent | `Blackburn North CC - 4th XI` |
| **5th Venue** | 5th XI venue | `Koonung Reserve` |

---

## Sheet 4: `Admins`
Used for restricted selector/admin access.

| Column Name | Description | Example |
| :--- | :--- | :--- |
| **Name** | Admin full name | `Neil Kloot` |
| **Phone** | Mobile number (normalized) | `+61417663518` |

---

## Sheet 5: `Availability_Log`
Audit log recording every availability submission from the web app.

| Column Name | Description |
| :--- | :--- |
| **Timestamp** | ISO timestamp of submission |
| **ProfileID** | Player PlayHQ GUID |
| **MatchDate** | Match date (`YYYY-MM-DD`) |
| **Response** | Selection (`Available`, `Unavailable`) |
| **Notes** | Optional player context notes |

