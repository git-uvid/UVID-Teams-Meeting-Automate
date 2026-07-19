# Security Policy

## Authentication
This add-in uses **MSAL.js (Microsoft Authentication Library)** for handling user authentication via Azure AD / Entra ID. 
- It relies on the implicit or authorization code flow with PKCE inside the browser environment.
- Access tokens are securely managed in memory/session storage by MSAL.

## Secret Management
- **Never commit sensitive credentials** (such as Client Secrets, Certificates, or user tokens) into this repository.
- The `clientId` and `authority` used in `taskpane.js` are considered public identifiers for Single Page Applications (SPAs). However, always ensure your Entra ID App Registration is locked down to specific Redirect URIs.

## API Permissions
The add-in requests the **`Sites.ReadWrite.All`** delegated permission to write to the SharePoint list. 
- Ensure this permission is consented to in your Entra ID tenant.
- Do not add overly broad permissions (e.g., Application-level permissions) since this runs entirely on the client side.

## Reporting a Vulnerability
If you discover a security vulnerability within this add-in, please do not disclose it publicly. Review your internal company guidelines for reporting application security issues.
