/*
 * UVID Teams Meeting Automate - Taskpane Logic
 */

let isMeetingFound = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Mail) {
    document.getElementById("flowAction").addEventListener("change", handleActionChange);
    document.getElementById("injectAndLog").addEventListener("click", handleInjectAndLog);
    
    // Initial setup
    handleActionChange();
  }
});

function handleActionChange() {
  const action = document.getElementById("flowAction").value;
  
  // Hide all sections first
  document.getElementById("idSection").style.display = "none";
  document.getElementById("createFields").style.display = "none";
  document.getElementById("recurringFields").style.display = "none";
  document.getElementById("rescheduleFields").style.display = "none";
  document.getElementById("reasonField").style.display = "none";
  document.getElementById("addParticipantFields").style.display = "none";
  document.getElementById("updateTitleFields").style.display = "none";
  document.getElementById("timezoneFieldShared").style.display = "none";
  document.getElementById("newDateSection").style.display = "none";
  document.getElementById("newDaySection").style.display = "none";

  if (action === "Schedule Meeting") {
    document.getElementById("createFields").style.display = "block";
  } 
  else if (action === "Schedule Recurring Meeting") {
    document.getElementById("createFields").style.display = "block";
    document.getElementById("recurringFields").style.display = "block";
  }
  else if (action === "Reschedule Meeting") {
    document.getElementById("idSection").style.display = "block";
    document.getElementById("rescheduleFields").style.display = "block";
    document.getElementById("newDateSection").style.display = "block";
    document.getElementById("reasonField").style.display = "block";
    document.getElementById("timezoneFieldShared").style.display = "block";
  }
  else if (action === "Reschedule Series") {
    document.getElementById("idSection").style.display = "block";
    document.getElementById("rescheduleFields").style.display = "block";
    document.getElementById("newDaySection").style.display = "block";
    document.getElementById("reasonField").style.display = "block";
    document.getElementById("timezoneFieldShared").style.display = "block";
  }
  else if (action === "Cancel Meeting" || action === "Cancel Series") {
    document.getElementById("idSection").style.display = "block";
    document.getElementById("reasonField").style.display = "block";
  }
  else if (action === "Add Participant") {
    document.getElementById("idSection").style.display = "block";
    document.getElementById("addParticipantFields").style.display = "block";
  }
  else if (action === "Update Meeting") {
    document.getElementById("idSection").style.display = "block";
    document.getElementById("updateTitleFields").style.display = "block";
  }
}

// Function to format date to "15 July 2026"
function formatDateText(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  const options = { day: 'numeric', month: 'long', year: 'numeric' };
  return d.toLocaleDateString('en-GB', options);
}

// Function to format time to "2:00 PM"
function formatTimeText(timeStr) {
  if (!timeStr) return "";
  const parts = timeStr.split(":");
  if (parts.length < 2) return timeStr;
  let h = parseInt(parts[0], 10);
  let m = parts[1];
  let ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12;
  h = h ? h : 12; 
  return `${h}:${m} ${ampm}`;
}

async function handleInjectAndLog() {
  const action = document.getElementById("flowAction").value;
  const statusEl = document.getElementById("statusMessage");
  statusEl.innerText = "Constructing Email...";
  statusEl.style.color = "blue";
  
  // Format the email body based on action
  let emailBody = "";
  
  if (action === "Schedule Meeting") {
    const title = document.getElementById("meetingTitle").value;
    const date = formatDateText(document.getElementById("startDate").value);
    const time = formatTimeText(document.getElementById("startTime").value);
    const duration = document.getElementById("duration").value;
    const participants = document.getElementById("participants").value.replace(/[\s,]+/g, ';');
    const type = document.getElementById("meetingType").value;
    const timezone = document.getElementById("timezone").value;
    
    emailBody = `Title: ${title}\nDate: ${date}\nTime: ${time}\nDuration: ${duration}\nParticipants: ${participants}\nType: ${type}\nTimezone: ${timezone}`;
  }
  else if (action === "Schedule Recurring Meeting") {
    const title = document.getElementById("meetingTitle").value;
    const date = formatDateText(document.getElementById("startDate").value);
    const time = formatTimeText(document.getElementById("startTime").value);
    const duration = document.getElementById("duration").value;
    const participants = document.getElementById("participants").value.replace(/[\s,]+/g, ';');
    const type = document.getElementById("meetingType").value;
    const timezone = document.getElementById("timezone").value;
    const freq = document.getElementById("recurrenceFrequency").value;
    const day = document.getElementById("recurrenceDay").value;
    const endDate = formatDateText(document.getElementById("seriesEndDate").value);
    
    emailBody = `Title: ${title}\nDate: ${date}\nTime: ${time}\nTimezone: ${timezone}\nDuration: ${duration}\nParticipants: ${participants}\nType: ${type}\nFrequency: ${freq}\nDay: ${day}\nEnd date: ${endDate}`;
  }
  else if (action === "Reschedule Meeting") {
    const id = document.getElementById("meetingId").value;
    const date = formatDateText(document.getElementById("newDate").value);
    const time = formatTimeText(document.getElementById("newTime").value);
    const timezone = document.getElementById("timezoneEdit").value;
    const reason = document.getElementById("reason").value;
    
    emailBody = `Meeting ID: ${id}\nNew Date: ${date}\nNew Time: ${time}\nTimezone: ${timezone}\nReason: ${reason}`;
  }
  else if (action === "Reschedule Series") {
    const id = document.getElementById("meetingId").value;
    const day = document.getElementById("newDay").value;
    const time = formatTimeText(document.getElementById("newTime").value);
    const timezone = document.getElementById("timezoneEdit").value;
    const reason = document.getElementById("reason").value;
    
    emailBody = `Series ID: ${id}\nNew Day: ${day}\nNew Time: ${time}\nTimezone: ${timezone}\nReason: ${reason}`;
  }
  else if (action === "Cancel Meeting") {
    const id = document.getElementById("meetingId").value;
    const reason = document.getElementById("reason").value;
    
    emailBody = `Meeting ID: ${id}\nReason: ${reason}`;
  }
  else if (action === "Cancel Series") {
    const id = document.getElementById("meetingId").value;
    const reason = document.getElementById("reason").value;
    
    emailBody = `Series ID: ${id}\nReason: ${reason}`;
  }
  else if (action === "Add Participant") {
    const id = document.getElementById("meetingId").value;
    const newParticipant = document.getElementById("newParticipantEmail").value;
    const name = document.getElementById("newParticipantName").value;
    
    emailBody = `Meeting ID: ${id}\nNew Participant: ${newParticipant}`;
    if (name) emailBody += `\nName: ${name}`;
  }
  else if (action === "Update Meeting") {
    const id = document.getElementById("meetingId").value;
    const newTitle = document.getElementById("newTitle").value;
    
    emailBody = `Meeting ID: ${id}\nNew Title: ${newTitle}`;
  }
  
  // Set Email To, Subject, and Body using Office.js
  Office.context.mailbox.item.to.setAsync([{ emailAddress: "connect@uvidconsulting.com" }], (res) => {
    if(res.status === Office.AsyncResultStatus.Failed) {
      console.error(res.error);
    }
  });
  
  Office.context.mailbox.item.subject.setAsync(action, (res) => {
    if(res.status === Office.AsyncResultStatus.Failed) {
      console.error(res.error);
    }
  });
  
  Office.context.mailbox.item.body.setSelectedDataAsync(emailBody, { coercionType: Office.CoercionType.Text }, async (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      statusEl.innerText = "Error injecting content: " + asyncResult.error.message;
      statusEl.style.color = "red";
      return;
    }
    
    // Log to SharePoint
    statusEl.innerText = "Logging to SharePoint...";
    try {
      await logToSharePoint();
      statusEl.innerText = "Success! Email populated and logged to SharePoint.";
      statusEl.style.color = "green";
    } catch (err) {
      console.error(err);
      statusEl.innerText = "Email populated, but failed to log to SharePoint.";
      statusEl.style.color = "orange";
    }
  });
}

async function logToSharePoint() {
  const token = await getAccessToken();
  const siteId = "30e62035-7798-466d-89df-155e97b1a206"; 
  const listId = "f5fbbf94-1a93-41bb-92cc-7d722b51cc13";
  
  const action = document.getElementById("flowAction").value;
  
  const payload = {
    fields: {
      Title: document.getElementById("meetingTitle")?.value || document.getElementById("newTitle")?.value || action,
      MeetingID: document.getElementById("meetingId")?.value || "NEW_ID", 
      Project: document.getElementById("project")?.value || "Unknown",
      MeetingType: document.getElementById("meetingType")?.value || "Client",
      LeadEmail: document.getElementById("leadEmail")?.value || "Unknown",
      Timezone: document.getElementById("timezone")?.value || document.getElementById("timezoneEdit")?.value || "UTC",
      Duration_x0028_minutes_x0029_: parseInt(document.getElementById("duration")?.value) || 0,
      Participants: document.getElementById("participants")?.value.replace(/[\s,]+/g, ';') || document.getElementById("newParticipantEmail")?.value || "",
      Recurring: action.includes("Recurring") || action.includes("Series"),
      Status: "Requested"
    }
  };
  
  const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to log to SP: ${errorText}`);
  }
}

async function getAccessToken() {
  const msalConfig = {
    auth: {
      clientId: "8b7e28a4-0ef6-43b4-a4f6-8bb0b1d8f16b", 
      authority: "https://login.microsoftonline.com/common",
      redirectUri: "http://localhost:3000"
    }
  };
  const msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    const request = { scopes: ["Sites.ReadWrite.All"], account: accounts[0] };
    try {
      const response = await msalInstance.acquireTokenSilent(request);
      return response.accessToken;
    } catch (e) {
      const response = await msalInstance.acquireTokenPopup(request);
      return response.accessToken;
    }
  } else {
    const request = { scopes: ["Sites.ReadWrite.All"] };
    const response = await msalInstance.acquireTokenPopup(request);
    return response.accessToken;
  }
}

// Utility: email validation
document.querySelectorAll('.email-input, .multi-email-input').forEach(el => {
  el.addEventListener('input', function() {
    const val = this.value.trim();
    const isMulti = this.classList.contains('multi-email-input');
    const errText = this.parentElement.querySelector('.error-text');
    let hasError = false;
    if (val) {
      if (isMulti) {
        const emails = val.split(/[\s,;]+/);
        emails.forEach(e => {
          if (e && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) hasError = true;
        });
      } else {
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(val)) hasError = true;
      }
    }
    if (hasError) {
      this.classList.add('error-border');
      if (errText) errText.style.display = 'block';
    } else {
      this.classList.remove('error-border');
      if (errText) errText.style.display = 'none';
    }
  });
});
