# Firebase Deployment Guide

This guide ensures you have everything needed to deploy the Laburnum CC Availability Tracker.

## Prerequisites
- **Node.js**: Version 20 (required by Firebase Functions).
- **Google Cloud Secrets**: Ensure `SPREADSHEET_ID` and `GOOGLE_SERVICE_ACCOUNT` are set in the Firebase Console or via CLI.

## Deployment Setup
We have configured local scripts to make deployment easier and avoid "command not found" errors.

1.  **Install Dependencies** (only needed once):
    ```bash
    npm install
    cd functions && npm install
    cd ..
    ```

2.  **Select Your Project**:
    You must tell Firebase which project to deploy to. Replace `your-project-id` with your actual Firebase Project ID (e.g., `lcc-availability`):
    ```bash
    npx firebase use --add
    ```

3.  **Deploy**:
    Run this from the **root** directory (where `firebase.json` is located):
    ```bash
    npm run deploy
    ```
    *Note: This is an alias for `npx firebase deploy`.*

## Troubleshooting
- **"command not found: firebase"**: Always use `npx firebase` or `npm run deploy` to use the local installation.
- **"No Hosting site detected"**: Ensure you have run `npx firebase use <project-id>` first so Firebase knows which project's hosting site to target.
- **404 Requested entity not found**: This usually means the Project ID in your command or config doesn't exist. Double-check your Project ID in the [Firebase Console](https://console.firebase.google.com/).

## Local Testing
1.  **Start the Emulator**:
    ```bash
    npx firebase emulators:start
    ```
2.  **Open the App**:
    Go to [http://localhost:5000](http://localhost:5000).
