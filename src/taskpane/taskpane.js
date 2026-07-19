/* global document, Office, msal, fetch, console */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("createListBtn").onclick = handleCreateList;
    document.getElementById("addEmailBtn").onclick = handleAddEmailOnly;
    document.getElementById("injectAndLog").onclick = handleInjectAndLog;
  }
});

const siteId = "YOUR_SITE_ID"; // Replace with your SharePoint Site ID
const listId = "YOUR_LIST_ID"; // Replace with your SharePoint List ID
const clientId = "YOUR_ENTRA_CLIENT_ID"; // Replace with your Entra ID Client ID
const tenantId = "YOUR_TENANT_ID"; // Replace with your Tenant ID or 'common'

async function getAccessToken(scopes) {
  const msalConfig = {
    auth: {
      clientId: clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri: "https://git-uvid.github.io/UVID-Teams-Meeting-Automate/taskpane.html",
    },
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);
  const loginRequest = { scopes: scopes };
  const loginResponse = await msalInstance.loginPopup(loginRequest);
  return loginResponse.accessToken;
}

function updateStatus(message, isError = false) {
  const statusEl = document.getElementById("statusMessage");
  statusEl.innerText = message;
  statusEl.style.color = isError ? "#d13438" : "#107c10";
}

// ==========================================
// BUTTON 1: Create List Only
// ==========================================
async function handleCreateList() {
  updateStatus("Authenticating for List Creation...", false);
  try {
    const accessToken = await getAccessToken(["Sites.Manage.All"]);
    updateStatus("Creating SharePoint List...", false);

    const graphEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`;
    
    const listPayload = {
      displayName: "Meetings Log",
      columns: [
        { name: "MeetingID", text: {} },
        { name: "SeriesID", text: {} },
        { name: "Project", text: {} },
        { name: "MeetingType", choice: { choices: ["Client", "Internal"] } },
        { name: "LeadEmail", text: {} },
        { name: "ScheduledDate", text: {} },
        { name: "Timezone", text: {} },
        { name: "Duration_x0028_minutes_x0029_", number: {} },
        { name: "Participants", text: {} },
        { name: "Recurring", boolean: {} },
        { name: "Status", choice: { choices: ["Scheduled", "Rescheduled", "Cancelled", "Completed"] } },
        { name: "TranscriptFiled", boolean: {} },
        { name: "DecisionRecordsCreated", boolean: {} },
        { name: "MeetingStatus", choice: { choices: ["Requested", "Created", "Rescheduled", "Cancelled", "Completed"] } }
      ],
      list: {
        template: "genericList"
      }
    };

    const response = await fetch(graphEndpoint, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(listPayload),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error?.message || "List creation failed");
    }

    const createdList = await response.json();
    console.log("Created List ID:", createdList.id);
    updateStatus(`List created successfully! (ID: ${createdList.id})`, false);
  } catch (error) {
    updateStatus("Error: " + error.message, true);
    console.error(error);
  }
}

// ==========================================
// BUTTON 2: Add Email Only (Test)
// ==========================================
async function handleAddEmailOnly() {
  const leadEmail = document.getElementById("leadEmail").value;
  const title = document.getElementById("meetingTitle").value || "Test Meeting";

  if (!leadEmail) {
    updateStatus("Please enter a Lead Email first.", true);
    return;
  }

  updateStatus("Authenticating for Email Insertion...", false);
  try {
    const accessToken = await getAccessToken(["Sites.ReadWrite.All"]);
    updateStatus("Inserting test item...", false);

    const graphEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
    const payload = {
      fields: {
        Title: title,
        LeadEmail: leadEmail
      }
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
      throw new Error(errorData.error?.message || "Item creation failed");
    }

    updateStatus("Test email inserted successfully!", false);
  } catch (error) {
    updateStatus("Error: " + error.message, true);
    console.error(error);
  }
}

// ==========================================
// BUTTON 3: Schedule & Log (Full)
// ==========================================
async function handleInjectAndLog() {
  // Capture all fields
  const fields = {
    Title: document.getElementById("meetingTitle").value,
    MeetingID: document.getElementById("meetingId").value,
    SeriesID: document.getElementById("seriesId").value,
    Project: document.getElementById("project").value,
    MeetingType: document.getElementById("meetingType").value,
    LeadEmail: document.getElementById("leadEmail").value,
    ScheduledDate: document.getElementById("scheduledDate").value,
    Timezone: document.getElementById("timezone").value,
    Duration_x0028_minutes_x0029_: parseInt(document.getElementById("duration").value, 10) || null,
    Participants: document.getElementById("participants").value,
    Recurring: document.getElementById("recurring").checked,
    Status: document.getElementById("status").value,
    TranscriptFiled: document.getElementById("transcriptFiled").checked,
    DecisionRecordsCreated: document.getElementById("decisionRecords").checked,
    MeetingStatus: document.getElementById("meetingStatus").value
  };

  if (!fields.Title) {
    updateStatus("Meeting Title is required.", true);
    return;
  }
  
  // Clean up any nulls to avoid Graph API errors
  if (fields.Duration_x0028_minutes_x0029_ === null) delete fields.Duration_x0028_minutes_x0029_;

  // Inject into email body
  const template = `Subject: Schedule Meeting\nTitle: ${fields.Title}\nMeeting ID: ${fields.MeetingID}\nDate: ${fields.ScheduledDate}\nTimezone: ${fields.Timezone}\nDuration: ${fields.Duration_x0028_minutes_x0029_ || ""} mins\nParticipants: ${fields.Participants}\nType: ${fields.MeetingType}\nStatus: ${fields.Status}`;

  updateStatus("Injecting into email...", false);
  Office.context.mailbox.item.body.setSelectedDataAsync(
    template,
    { coercionType: Office.CoercionType.Text },
    async (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        updateStatus("Failed to inject text: " + asyncResult.error.message, true);
        return;
      }

      try {
        updateStatus("Injected! Authenticating for full log...", false);
        const accessToken = await getAccessToken(["Sites.ReadWrite.All"]);
        updateStatus("Logging to SharePoint...", false);

        const graphEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
        const payload = { fields: fields };

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

        updateStatus("Successfully scheduled and logged!", false);
      } catch (error) {
        updateStatus("Error logging to SharePoint: " + error.message, true);
        console.error(error);
      }
    }
  );
}
