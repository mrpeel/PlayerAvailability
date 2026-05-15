# Laburnum CC Availability Tracker - Deployment Guide

This guide explains how to deploy the Availability Tracker as a Google Apps Script Web App.

## Prerequisites

1.  **Google Account**: You need access to a Google account to host the Spreadsheet and Script.
2.  **Node.js & NPM**: Installed on your local machine.
3.  **Enable Apps Script API**: Visit [script.google.com/home/usersettings](https://script.google.com/home/usersettings) and set the Google Apps Script API to **On**.

## Initial Setup

1.  **Login to Google**:
    ```bash
    npx clasp login
    ```

2.  **Create a New Script Project**:
    If you haven't already linked a project, run the following. If you get an error that a file already exists, you can add `--force`.
    ```bash
    npx clasp create --title "Laburnum CC Availability" --type webapp --rootDir src
    ```
    *Note: This will create a new Google Sheet in your Google Drive.*

3.  **Verify Configuration**:
    Check that `.clasp.json` in the root folder contains a valid `scriptId`.
    *Note: `.clasp.json` is included in `.gitignore` to prevent your private Script ID from being committed to a public repository. If you are setting this up on a new machine, running `clasp create` or `clasp link <ID>` will regenerate this file.*

4.  **Push the Code**:
    ```bash
    npm run push
    ```

## Configuring the Spreadsheet

1.  **Open the Script**:
    ```bash
    npx clasp open
    ```
2.  In the Apps Script editor, find the **`initializeSpreadsheet`** function in `Code.gs`.
3.  Select it from the toolbar and click **Run**.
4.  Grant the necessary permissions when prompted.
5.  Open the associated Google Sheet (it will have the same name as your script). You will now see the `Players`, `Rounds`, and `Availability` tabs with headers.

## Deployment as a Web App

1.  In the Apps Script editor, click **Deploy** > **New Deployment**.
2.  Select **Web App**.
3.  Set the following:
    *   **Description**: Initial Deployment
    *   **Execute as**: Me (your-email)
    *   **Who has access**: Anyone
4.  Click **Deploy**.
5.  Copy the **Web App URL**. This is the link you share with players.

### ⚠️ Updating the App (Crucial)
Google Apps Script does **not** automatically update your live URL when you push code. If you make changes and push them via `npm run push`:
1.  In the script editor, click **Deploy** > **Manage Deployments**.
2.  Click the **pencil icon** (Edit) on your active deployment.
3.  Under "Version", select **New Version**.
4.  Click **Deploy**.
*Without this, your live URL will continue to run the old code.*

## Shareable Links (Round Specific)

You can pre-load a specific round by appending `?r=ROUND_NUMBER` to your URL:
`https://script.google.com/.../exec?r=1`

## Ongoing Maintenance

*   **Adding Players**: Add them to the `Players` tab. Ensure the Phone column uses the format `04...` or `+614...`. The app will auto-normalize them.
*   **Adding Rounds**: Add new rounds to the `Rounds` tab. The app defaults to the latest round if none is specified in the URL.
*   **Wall of Shame**: You can call the `getWallOfShame(roundNum)` function from the script editor or a custom menu to see who hasn't responded yet.

## Troubleshooting "Could not find script"
If `npm run push` fails with "Could not find script":
1.  Check that the `scriptId` in `.clasp.json` matches the ID in your script URL (the string between `/d/` and `/edit`).
2.  Ensure you are logged in to the same Google account that owns the script.
