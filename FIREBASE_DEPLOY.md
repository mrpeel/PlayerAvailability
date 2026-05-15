# Firebase Deployment for Laburnum CC Availability Tracker

This project has been migrated from Google Apps Script to Firebase to provide a premium, white-labeled experience without Google banners.

## Prerequisites
1.  **Firebase Project**: Create one at [console.firebase.google.com](https://console.firebase.google.com).
2.  **Service Account**:
    *   Go to Project Settings > Service Accounts.
    *   Click **Generate new private key**.
    *   Save this file as `functions/service-account.json`.
    *   **CRITICAL**: Copy the `client_email` from this JSON and share the Google Sheet with this email (Editor access).

## Initial Setup
Link your local directory to your Firebase project:
```bash
npx firebase-tools use --add
# Select your project from the list and give it an alias (e.g., 'default')
```

## Configuration
Use **Firebase Secrets Manager** to keep your credentials out of the codebase.

1.  **Set Spreadsheet ID**:
    ```bash
    # Run this from the ROOT directory, then paste your ID
    npx firebase-tools functions:secrets:set SPREADSHEET_ID
    ```

2.  **Set Google Credentials**:
    ```bash
    # Run this from the ROOT directory
    npx firebase-tools functions:secrets:set GOOGLE_SERVICE_ACCOUNT < functions/service-account.json
    ```

3.  **Local Development**:
    Place a `service-account.json` file inside the `functions/` directory. It is already added to `.gitignore` and will be used as a fallback for the local emulator.

## Deployment
1.  **Install dependencies**:
    ```bash
    cd functions && npm install
    ```
2.  **Login to Firebase**:
    ```bash
    npx firebase-tools login
    ```
3.  **Deploy**:
    ```bash
    npx firebase-tools deploy
    ```

## Local Testing
1.  **Configure local variables**:
    Create a file named `functions/.env` and add your spreadsheet ID:
    ```bash
    SPREADSHEET_ID=your_spreadsheet_id_here
    ```
    *(Note: The `service-account.json` you downloaded earlier is used automatically for auth during local testing)*.

2.  **Start the Emulator**:
    Run this from the root directory:
    ```bash
    npx firebase-tools emulators:start
    ```

3.  **Open the App**:
    Go to [http://localhost:5000](http://localhost:5000) in your browser.
    *The API will be running locally at http://localhost:5001 (mapped automatically).*
