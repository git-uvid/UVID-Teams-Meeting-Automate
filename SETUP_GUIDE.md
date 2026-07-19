# Setup, Testing, and Production Deployment Guide

This document outlines everything you need to replace in the code, how to test the add-in, and how to build it for production.

## 1. Required Configuration (Code Replacements)
Before testing or deploying, you must connect the add-in to your Microsoft 365 environment by updating the code in **`src/taskpane/taskpane.js`**.

### Entra ID (Azure AD) Setup
In your Entra ID portal, create a **Single Page Application (SPA)** App Registration.
- Add `https://localhost:3000/taskpane.html` as a Redirect URI for local testing.
- Add your production URL (e.g., GitHub Pages URL) as a Redirect URI for production.
- Grant API Permissions: `Microsoft Graph` > Delegated > `Sites.ReadWrite.All`.

**Update `taskpane.js` (around line 46):**
```javascript
const msalConfig = {
  auth: {
    clientId: "YOUR_ENTRA_CLIENT_ID", // Replace with your Application (client) ID
    authority: "https://login.microsoftonline.com/YOUR_TENANT_ID", // Replace with your Directory (tenant) ID
    redirectUri: "https://localhost:3000/taskpane.html" // Change this to your public URL when moving to production!
  }
};
```

### SharePoint List Setup
Ensure you have a SharePoint list with the following columns: `Title`, `MeetingType`, `Project`, and `Status`.

**Update `taskpane.js` (around line 58):**
```javascript
const siteId = "YOUR_SITE_ID"; // Replace with your target SharePoint Site ID
const listId = "YOUR_LIST_ID"; // Replace with your target SharePoint List ID
```

---

## 2. Local Testing
To test the add-in locally on your machine:

1. **Start the Development Server:**
   Run the following command in your terminal:
   ```bash
   npm run dev-server
   ```
   This hosts your application locally on `https://localhost:3000`.

2. **Sideload into Outlook:**
   - Go to [Outlook on the web](https://outlook.office.com/).
   - Click **New mail** to open the compose window.
   - Click the **Apps** icon (four squares) or the `...` menu, then select **Get Add-ins**.
   - Navigate to **My add-ins** > **Custom add-ins** > **Add a custom add-in** > **Add from file...**
   - Upload the `manifest.xml` file located in the root of this project.

3. **Verify:**
   - Click the "Schedule Meeting" button in the email compose toolbar.
   - Fill out the form and click "Schedule & Log".
   - You should see the MSAL login popup. Once authenticated, the meeting will inject into the email and log to SharePoint!

---

## 3. Production Deployment
When you are ready to move out of local testing and host this for real users:

1. **Update URLs:**
   - In `manifest.xml`, replace all instances of `https://localhost:3000` with your final hosting URL (e.g., `https://my-company.github.io/meeting-scheduler`).
   - In `src/taskpane/taskpane.js`, update the MSAL `redirectUri` to point to the production `taskpane.html`.

2. **Build the App:**
   Run the build script:
   ```bash
   npm run build
   ```
   This will bundle and optimize all files and output them into the **`dist/`** directory.

3. **Host the Files:**
   Upload the contents of the `dist/` directory to a static HTTPS server. Good free options include:
   - **GitHub Pages** (Push `dist` contents to a `gh-pages` branch).
   - **Azure Storage Static Website hosting**.

4. **Deploy the Manifest:**
   Distribute the updated `manifest.xml` file. Users can sideload it manually, or your Microsoft 365 IT Administrator can deploy it centrally to everyone in your organization via the **Microsoft 365 Admin Center** > **Settings** > **Integrated apps**.
