# Meeting Scheduler Outlook Web Add-in

This project is a modern Outlook Web Add-in designed to streamline scheduling and tracking meetings. It provides a simple task pane within the email compose window to capture meeting details and logs the output directly to a SharePoint list via the Microsoft Graph API.

## Features
- **Task Pane Form:** Inputs for Meeting Title, Date, and Meeting Type (Client vs Internal).
- **Email Injection:** Injects formatted meeting text into the body of the drafted email.
- **MSAL Authentication:** Uses MSAL.js for secure Azure AD/Entra ID authentication.
- **Graph API Integration:** Logs the meeting event into a target SharePoint list upon successful authentication.

## Quick Start
1. Run `npm install` to install dependencies.
2. Read the `SETUP_GUIDE.md` for instructions on configuring your Entra ID and SharePoint details.
3. Run `npm run dev-server` to start the local development server at `https://localhost:3000`.
4. Sideload the `manifest.xml` file into Outlook on the web to test the add-in.

## Build Requirements
- Node.js
- An Office 365 / Microsoft 365 account for testing.
- An Entra ID (Azure AD) App Registration with `Sites.ReadWrite.All` delegated permissions.
