/*
 * UVID Teams Meeting Automate - Taskpane Logic
 */

let isMeetingFound = false;
let userEmail = "TestUser@uvidconsulting.com";

// Setup event listeners immediately so it works in local browser testing too
document.addEventListener("DOMContentLoaded", () => {
  // Listeners for Tab Navigation
  const tabs = document.querySelectorAll('input[name="tabNav"]');
  tabs.forEach(tab => {
    tab.addEventListener("change", handleTabChange);
  });

  document.getElementById("flowActionNew")?.addEventListener("change", handleActionChangeNew);
  document.getElementById("flowAction")?.addEventListener("change", handleActionChangeEdit);
  
  document.getElementById("startTime")?.addEventListener("change", calculateEndTime);
  document.getElementById("duration")?.addEventListener("input", calculateEndTime);

  document.getElementById("injectAndLog")?.addEventListener("click", handleInjectAndLog);
  
  // Initial setup
  handleTabChange();
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Mail) {
    if (Office.context.mailbox && Office.context.mailbox.userProfile) {
      userEmail = Office.context.mailbox.userProfile.emailAddress;
    }
  }
  const genBy = document.getElementById("generatedBy");
  if (genBy) {
    genBy.value = userEmail;
  }
});

function calculateEndTime() {
  const startTime = document.getElementById("startTime").value;
  const duration = parseInt(document.getElementById("duration").value);
  const endTimeEl = document.getElementById("endTime");

  if (startTime && !isNaN(duration)) {
    const parts = startTime.split(":");
    let date = new Date();
    date.setHours(parseInt(parts[0], 10));
    date.setMinutes(parseInt(parts[1], 10));
    
    // Add duration in minutes
    date.setMinutes(date.getMinutes() + duration);
    
    let endHours = String(date.getHours()).padStart(2, '0');
    let endMins = String(date.getMinutes()).padStart(2, '0');
    endTimeEl.value = `${endHours}:${endMins}`;
  } else {
    if (endTimeEl) endTimeEl.value = "";
  }
}

function handleTabChange() {
  const activeTabEl = document.querySelector('input[name="tabNav"]:checked');
  if (!activeTabEl) return;
  const activeTab = activeTabEl.value;

  if (activeTab === "newMeeting") {
    document.getElementById("newMeetingContainer").style.display = "block";
    document.getElementById("editMeetingContainer").style.display = "none";
    handleActionChangeNew();
  } else {
    document.getElementById("newMeetingContainer").style.display = "none";
    document.getElementById("editMeetingContainer").style.display = "block";
    handleActionChangeEdit();
  }
}

function handleActionChangeNew() {
  const action = document.getElementById("flowActionNew").value;
  if (action === "Schedule Recurring Meeting") {
    document.getElementById("recurringFields").style.display = "block";
  } else {
    document.getElementById("recurringFields").style.display = "none";
  }
}

function handleActionChangeEdit() {
  const action = document.getElementById("flowAction").value;
  
  // Hide edit sections first
  document.getElementById("rescheduleFields").style.display = "none";
  document.getElementById("reasonField").style.display = "none";
  document.getElementById("addParticipantFields").style.display = "none";
  document.getElementById("updateTitleFields").style.display = "none";
  document.getElementById("timezoneFieldShared").style.display = "none";
  document.getElementById("newDateSection").style.display = "none";
  document.getElementById("newDaySection").style.display = "none";

  if (action === "Reschedule Meeting") {
    document.getElementById("rescheduleFields").style.display = "block";
    document.getElementById("newDateSection").style.display = "block";
    document.getElementById("reasonField").style.display = "block";
    document.getElementById("timezoneFieldShared").style.display = "block";
  }
  else if (action === "Reschedule Series") {
    document.getElementById("rescheduleFields").style.display = "block";
    document.getElementById("newDaySection").style.display = "block";
    document.getElementById("reasonField").style.display = "block";
    document.getElementById("timezoneFieldShared").style.display = "block";
  }
  else if (action === "Cancel Meeting" || action === "Cancel Series") {
    document.getElementById("reasonField").style.display = "block";
  }
  else if (action === "Add Participant") {
    document.getElementById("addParticipantFields").style.display = "block";
  }
  else if (action === "Update Meeting") {
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
  const activeTab = document.querySelector('input[name="tabNav"]:checked').value;
  let action = "";
  
  if (activeTab === "newMeeting") {
    action = document.getElementById("flowActionNew").value;
  } else {
    action = document.getElementById("flowAction").value;
  }

  const statusEl = document.getElementById("statusMessage");
  statusEl.innerText = "Constructing Email...";
  statusEl.style.color = "blue";
  
  // Format the email body based on action
  let emailBody = "";
  
  if (action === "Schedule Meeting") {
    const title = document.getElementById("meetingTitle").value;
    const subject = document.getElementById("meetingSubject").value;
    const bodyMsg = document.getElementById("meetingEventMessage").value;
    const date = formatDateText(document.getElementById("startDate").value);
    const time = formatTimeText(document.getElementById("startTime").value);
    const endTime = formatTimeText(document.getElementById("endTime").value);
    const duration = document.getElementById("duration").value;
    const reqAttendees = document.getElementById("requiredAttendees").value.replace(/[\s,]+/g, ';');
    const optAttendees = document.getElementById("optionalAttendees").value.replace(/[\s,]+/g, ';');
    const type = document.getElementById("meetingType").value;
    const timezone = document.getElementById("timezone").value;
    
    emailBody = `Title: ${title}\nSubject: ${subject}\nMessage: ${bodyMsg}\nDate: ${date}\nStart Time: ${time}\nEnd Time: ${endTime}\nDuration: ${duration}\nRequired Attendees: ${reqAttendees}\nOptional Attendees: ${optAttendees}\nType: ${type}\nTimezone: ${timezone}`;
  }
  else if (action === "Schedule Recurring Meeting") {
    const title = document.getElementById("meetingTitle").value;
    const subject = document.getElementById("meetingSubject").value;
    const bodyMsg = document.getElementById("meetingEventMessage").value;
    const date = formatDateText(document.getElementById("startDate").value);
    const time = formatTimeText(document.getElementById("startTime").value);
    const endTime = formatTimeText(document.getElementById("endTime").value);
    const duration = document.getElementById("duration").value;
    const reqAttendees = document.getElementById("requiredAttendees").value.replace(/[\s,]+/g, ';');
    const optAttendees = document.getElementById("optionalAttendees").value.replace(/[\s,]+/g, ';');
    const type = document.getElementById("meetingType").value;
    const timezone = document.getElementById("timezone").value;
    const freq = document.getElementById("recurrenceFrequency").value;
    const interval = document.getElementById("recurrenceInterval").value;
    const endDate = formatDateText(document.getElementById("seriesEndDate").value);
    
    emailBody = `Title: ${title}\nSubject: ${subject}\nMessage: ${bodyMsg}\nDate: ${date}\nStart Time: ${time}\nEnd Time: ${endTime}\nTimezone: ${timezone}\nDuration: ${duration}\nRequired Attendees: ${reqAttendees}\nOptional Attendees: ${optAttendees}\nType: ${type}\nFrequency: ${freq}\nInterval: ${interval}\nEnd date: ${endDate}`;
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
  if (window.Office && window.Office.context && window.Office.context.mailbox) {
    Office.context.mailbox.item.to.setAsync([{ emailAddress: "connect@uvidconsulting.com" }], (res) => {
      if(res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
    });
    
    Office.context.mailbox.item.subject.setAsync(action, (res) => {
      if(res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
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
        await logToSharePoint(action, activeTab);
        statusEl.innerText = "Success! Email populated and logged to SharePoint.";
        statusEl.style.color = "green";
      } catch (err) {
        console.error(err);
        statusEl.innerText = "Email populated, but failed to log to SharePoint.";
        statusEl.style.color = "orange";
      }
    });
  } else {
    // If testing in browser without Office.js
    console.log("Subject:", action);
    console.log("Body:", emailBody);
    statusEl.innerText = "Testing outside Outlook. Check console for output.";
    statusEl.style.color = "orange";
  }
}

async function logToSharePoint(action, activeTab) {
  const token = await getAccessToken();
  const siteId = "30e62035-7798-466d-89df-155e97b1a206"; 
  const listId = "f5fbbf94-1a93-41bb-92cc-7d722b51cc13";
  
  let payload = {};

  if (activeTab === "newMeeting") {
    const isRecurring = action === "Schedule Recurring Meeting";
    // Format start time correctly as ISO string if possible, or just the current string
    let startDateVal = document.getElementById("startDate").value; // YYYY-MM-DD
    let startTimeVal = document.getElementById("startTime").value; // HH:MM
    let startIso = "";
    if (startDateVal && startTimeVal) {
      startIso = startDateVal + "T" + startTimeVal + ":00Z"; // Rough approximation, real SP expects ISO
    }

    let endDateVal = document.getElementById("startDate").value; // Assuming same day
    let endTimeVal = document.getElementById("endTime").value;
    let endIso = "";
    if (endDateVal && endTimeVal) {
      endIso = endDateVal + "T" + endTimeVal + ":00Z";
    }

    payload = {
      fields: {
        Title: document.getElementById("meetingTitle")?.value || "Unknown",
        MeetingSubject: document.getElementById("meetingSubject")?.value || "",
        MeetingEventmessagecontent: document.getElementById("meetingEventMessage")?.value || "",
        MeetingID: "NEW_ID",
        SeriesID: "NEW_ID",
        Project: document.getElementById("project")?.value || "Unknown",
        MeetingType: document.getElementById("meetingType")?.value || "Client",
        LeadEmail: document.getElementById("leadEmail")?.value || "Unknown",
        Timezone: document.getElementById("timezone")?.value || "UTC",
        Duration_x0028_minutes_x0029_: parseInt(document.getElementById("duration")?.value) || 0,
        Requiredattendees: document.getElementById("requiredAttendees")?.value.replace(/[\s,]+/g, ';') || "",
        Optionalattendees: document.getElementById("optionalAttendees")?.value.replace(/[\s,]+/g, ';') || "",
        Recurring: isRecurring,
        Recurrencepattern: isRecurring ? document.getElementById("recurrenceFrequency")?.value : "",
        Recurrenceinterval: isRecurring ? (parseInt(document.getElementById("recurrenceInterval")?.value) || 1) : 0,
        Starttime: startIso || null,
        Endtime: endIso || null,
        GeneratedBy: document.getElementById("generatedBy")?.value || "Unknown",
        Status: "Requested" // Initial SP status is requested, flow changes it
      }
    };
  } else {
    // Edit mode: we just log a basic record or we would ideally update the SP item, but for now we push a log row
    payload = {
      fields: {
        Title: document.getElementById("newTitle")?.value || action,
        MeetingID: document.getElementById("meetingId")?.value || "",
        SeriesID: document.getElementById("meetingId")?.value || "",
        Status: "Requested"
      }
    };
  }
  
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
document.addEventListener("DOMContentLoaded", () => {
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
});
