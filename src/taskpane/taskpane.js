Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("injectAndLog").onclick = handleInjectAndLog;
  }
});

async function handleInjectAndLog() {
  const title = document.getElementById("meetingTitle").value;
  const date = document.getElementById("meetingDate").value;
  const type = document.getElementById("meetingType").value;
  const statusEl = document.getElementById("statusMessage");

  statusEl.innerText = "";
  statusEl.style.color = "#d13438";

  if (!title || !date) {
    statusEl.innerText = "Please fill in all fields.";
    return;
  }

  // Step B: Inject into email body
  const template = `Subject: Schedule Meeting\nTitle: ${title}\nDate: ${date}\nTime: \nTimezone: \nDuration: \nParticipants: \nType: ${type}`;

  Office.context.mailbox.item.body.setSelectedDataAsync(
    template,
    { coercionType: Office.CoercionType.Text },
    async (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        statusEl.innerText = "Failed to inject text: " + asyncResult.error.message;
        return;
      }

      try {
        statusEl.innerText = "Injected successfully! Authenticating...";
        statusEl.style.color = "#0078d4";

        await logToSharePoint(title, type);

        statusEl.innerText = "Successfully scheduled and logged to SharePoint!";
        statusEl.style.color = "#107c10"; // success green
      } catch (error) {
        statusEl.innerText = "Error logging to SharePoint: " + error.message;
        statusEl.style.color = "#d13438";
        console.error(error);
      }
    }
  );
}

async function logToSharePoint(title, meetingType) {
  // Step C: Initialize MSAL
  const msalConfig = {
    auth: {
      clientId: "YOUR_ENTRA_CLIENT_ID", // Replace with your Entra ID Client ID
      authority: "https://login.microsoftonline.com/YOUR_TENANT_ID", // Replace with your Tenant ID or 'common'
      redirectUri: "https://localhost:3000/taskpane.html",
    },
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);

  // Step D: Trigger loginPopup
  const loginRequest = {
    scopes: ["Sites.ReadWrite.All"],
  };

  const loginResponse = await msalInstance.loginPopup(loginRequest);
  const accessToken = loginResponse.accessToken;

  // Step E: Fetch POST to Graph API
  const siteId = "YOUR_SITE_ID"; // Replace with your SharePoint Site ID
  const listId = "YOUR_LIST_ID"; // Replace with your SharePoint List ID

  const graphEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;

  // Step F: Graph API payload structure
  const payload = {
    fields: {
      Title: title,
      MeetingType: meetingType,
      Project: "TBD",
      Status: "Requested",
    },
  };

  const response = await fetch(graphEndpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const errorData = await response.json();
    throw new Error(errorData.error?.message || "Graph API request failed");
  }
}
