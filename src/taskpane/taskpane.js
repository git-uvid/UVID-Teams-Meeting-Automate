/*
 * UVID Teams Meeting Automate - Taskpane Logic
 */

let isMeetingFound = false;
let userEmail = "";

Office.onReady((info) => {
  if (window.Office && window.Office.context && window.Office.context.mailbox && window.Office.context.mailbox.userProfile) {
    userEmail = window.Office.context.mailbox.userProfile.emailAddress;
  } else {
    userEmail = "TestUser@uvidconsulting.com";
  }
  
  const genBy = document.getElementById("generatedBy");
  if (genBy) {
    genBy.value = userEmail;
  }

  // Listeners for Tab Navigation
  const tabs = document.querySelectorAll('input[name="tabNav"]');
  tabs.forEach(tab => {
    tab.addEventListener("change", handleTabChange);
  });

  document.getElementById("flowActionNew")?.addEventListener("change", handleActionChangeNew);
  document.getElementById("flowAction")?.addEventListener("change", handleActionChangeEdit);
  
  document.getElementById("startTime")?.addEventListener("change", calculateEndTime);
  document.getElementById("duration")?.addEventListener("input", calculateEndTime);
  
  document.getElementById("startDate")?.addEventListener("change", (e) => {
    const rStart = document.getElementById("recurrenceStartDate");
    if (rStart) {
      rStart.value = e.target.value;
    }
  });

  document.getElementById("injectEmailOnly")?.addEventListener("click", () => handleSubmission(true, false));
  document.getElementById("logSpOnly")?.addEventListener("click", () => handleSubmission(false, true));
  document.getElementById("injectAndLog")?.addEventListener("click", () => handleSubmission(true, true));
  
  // Initial setup
  handleTabChange();

  // Fetch projects on load
  fetchProjects();
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

function escapeHtml(unsafe) {
  if (!unsafe) return "";
  return unsafe.toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

async function handleSubmission(doEmail, doSP) {
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
    const title = escapeHtml(document.getElementById("meetingTitle").value);
    const subject = escapeHtml(document.getElementById("meetingSubject").value);
    const bodyMsg = escapeHtml(document.getElementById("meetingEventMessage").value).replace(/\n/g, "<br>");
    const date = escapeHtml(formatDateText(document.getElementById("startDate").value));
    const time = escapeHtml(formatTimeText(document.getElementById("startTime").value));
    const endTime = escapeHtml(formatTimeText(document.getElementById("endTime").value));
    const duration = escapeHtml(document.getElementById("duration").value);
    const reqAttendees = escapeHtml(document.getElementById("requiredAttendees").value.replace(/[\s,]+/g, ';'));
    const optAttendees = escapeHtml(document.getElementById("optionalAttendees").value.replace(/[\s,]+/g, ';'));
    const type = escapeHtml(document.getElementById("meetingType").value);
    const timezone = escapeHtml(document.getElementById("timezone").value);
    
    emailBody = `Title: ${title}<br>Subject: ${subject}<br>Message: ${bodyMsg}<br>Date: ${date}<br>Start Time: ${time}<br>End Time: ${endTime}<br>Duration: ${duration}<br>Required Attendees: ${reqAttendees}<br>Optional Attendees: ${optAttendees}<br>Type: ${type}<br>Timezone: ${timezone}`;
  }
  else if (action === "Schedule Recurring Meeting") {
    const title = escapeHtml(document.getElementById("meetingTitle").value);
    const subject = escapeHtml(document.getElementById("meetingSubject").value);
    const bodyMsg = escapeHtml(document.getElementById("meetingEventMessage").value).replace(/\n/g, "<br>");
    const date = escapeHtml(formatDateText(document.getElementById("startDate").value));
    const time = escapeHtml(formatTimeText(document.getElementById("startTime").value));
    const endTime = escapeHtml(formatTimeText(document.getElementById("endTime").value));
    const duration = escapeHtml(document.getElementById("duration").value);
    const reqAttendees = escapeHtml(document.getElementById("requiredAttendees").value.replace(/[\s,]+/g, ';'));
    const optAttendees = escapeHtml(document.getElementById("optionalAttendees").value.replace(/[\s,]+/g, ';'));
    const type = escapeHtml(document.getElementById("meetingType").value);
    const timezone = escapeHtml(document.getElementById("timezone").value);
    const freq = escapeHtml(document.getElementById("recurrenceFrequency").value);
    const interval = escapeHtml(document.getElementById("recurrenceInterval").value);
    const recurStartDate = escapeHtml(formatDateText(document.getElementById("recurrenceStartDate").value));
    const recurEndDate = escapeHtml(formatDateText(document.getElementById("recurrenceEndDate").value));
    
    emailBody = `Title: ${title}<br>Subject: ${subject}<br>Message: ${bodyMsg}<br>Date: ${date}<br>Start Time: ${time}<br>End Time: ${endTime}<br>Timezone: ${timezone}<br>Duration: ${duration}<br>Required Attendees: ${reqAttendees}<br>Optional Attendees: ${optAttendees}<br>Type: ${type}<br>Frequency: ${freq}<br>Interval: ${interval}<br>Recurrence Start Date: ${recurStartDate}<br>Recurrence End Date: ${recurEndDate}`;
  }
  else if (action === "Reschedule Meeting") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const date = escapeHtml(formatDateText(document.getElementById("newDate").value));
    const time = escapeHtml(formatTimeText(document.getElementById("newTime").value));
    const timezone = escapeHtml(document.getElementById("timezoneEdit").value);
    const reason = escapeHtml(document.getElementById("reason").value);
    
    emailBody = `Meeting ID: ${id}<br>New Date: ${date}<br>New Time: ${time}<br>Timezone: ${timezone}<br>Reason: ${reason}`;
  }
  else if (action === "Reschedule Series") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const day = escapeHtml(document.getElementById("newDay").value);
    const time = escapeHtml(formatTimeText(document.getElementById("newTime").value));
    const timezone = escapeHtml(document.getElementById("timezoneEdit").value);
    const reason = escapeHtml(document.getElementById("reason").value);
    
    emailBody = `Series ID: ${id}<br>New Day: ${day}<br>New Time: ${time}<br>Timezone: ${timezone}<br>Reason: ${reason}`;
  }
  else if (action === "Cancel Meeting") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const reason = escapeHtml(document.getElementById("reason").value);
    
    emailBody = `Meeting ID: ${id}<br>Reason: ${reason}`;
  }
  else if (action === "Cancel Series") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const reason = escapeHtml(document.getElementById("reason").value);
    
    emailBody = `Series ID: ${id}<br>Reason: ${reason}`;
  }
  else if (action === "Add Participant") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const newParticipant = escapeHtml(document.getElementById("newParticipantEmail").value);
    const name = escapeHtml(document.getElementById("newParticipantName").value);
    
    emailBody = `Meeting ID: ${id}<br>New Participant: ${newParticipant}`;
    if (name) emailBody += `<br>Name: ${name}`;
  }
  else if (action === "Update Meeting") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const newTitle = escapeHtml(document.getElementById("newTitle").value);
    
    emailBody = `Meeting ID: ${id}<br>New Title: ${newTitle}`;
  }
  
  
  if (window.Office && window.Office.context && window.Office.context.mailbox) {
    
    if (doEmail) {
      statusEl.innerText = "Constructing Email...";
      statusEl.style.color = "blue";
      
      // Clear To, CC, BCC
      Office.context.mailbox.item.to.setAsync([{ emailAddress: "connect@uvidconsulting.com" }], (res) => {
        if(res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
      });
      Office.context.mailbox.item.cc.setAsync([], (res) => {
        if(res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
      });
      Office.context.mailbox.item.bcc.setAsync([], (res) => {
        if(res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
      });
      
      Office.context.mailbox.item.subject.setAsync(action, (res) => {
        if(res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
      });
      
      // Append Regards and Disclaimer
      const name = userEmail ? userEmail.split('@')[0].replace('.', ' ') : 'User';
      const capitalizedName = name.replace(/\b\w/g, l => l.toUpperCase());
      const finalEmailBody = emailBody + `<br><br>Best regards,<br>${capitalizedName}<br><br><i>[Auto-generated email by UVID Teams Meeting Automate]</i>`;
      
      Office.context.mailbox.item.body.setAsync(finalEmailBody, { coercionType: Office.CoercionType.Html }, async (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          statusEl.innerText = "Error injecting content: " + asyncResult.error.message;
          statusEl.style.color = "red";
          return;
        }
        
        if (doSP) {
          await executeSPLog(action, activeTab, statusEl, "Success! Email populated and logged to SharePoint.", "Email populated, but failed to log to SharePoint.");
        } else {
          statusEl.innerText = "Success! Email constructed.";
          statusEl.style.color = "green";
        }
      });
    } else if (doSP) {
      await executeSPLog(action, activeTab, statusEl, "Success! Logged to SharePoint.", "Failed to log to SharePoint.");
    }
  } else {
    // If testing in browser without Office.js
    console.log("Subject:", action);
    console.log("Body:", emailBody);
    
    if (doSP) {
      await executeSPLog(action, activeTab, statusEl, "Testing outside Outlook. Logged to SharePoint.", "Testing outside Outlook. SP Log failed.");
    } else {
      statusEl.innerText = "Testing outside Outlook. Email generated in console.";
      statusEl.style.color = "orange";
    }
  }
}

async function executeSPLog(action, activeTab, statusEl, successMsg, failMsg) {
  statusEl.innerText = "Logging to SharePoint...";
  statusEl.style.color = "blue";
  try {
    await logToSharePoint(action, activeTab);
    statusEl.innerText = successMsg;
    statusEl.style.color = "green";
  } catch (err) {
    console.error(err);
    statusEl.innerText = failMsg;
    statusEl.style.color = "red";
  }
}

async function logToSharePoint(action, activeTab) {
  const token = await getAccessToken();
  const siteId = "key65akcdgsfg2zhwxauifkam1a.sharepoint.com"; 
  const listId = "093329a2-701b-4e93-9789-f561ef47ddce";
  
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
        Recurrencestartdate: (isRecurring && document.getElementById("recurrenceStartDate")?.value) ? document.getElementById("recurrenceStartDate").value + "T00:00:00Z" : null,
        Recurrenceenddate: (isRecurring && document.getElementById("recurrenceEndDate")?.value) ? document.getElementById("recurrenceEndDate").value + "T00:00:00Z" : null,
        Starttime: startIso || null,
        Endtime: endIso || null,
        GeneratedBy: document.getElementById("generatedBy")?.value || "Unknown"
      }
    };
  } else {
    // Edit mode: we just log a basic record or we would ideally update the SP item, but for now we push a log row
    payload = {
      fields: {
        Title: document.getElementById("newTitle")?.value || action,
        MeetingID: document.getElementById("meetingId")?.value || "",
        SeriesID: document.getElementById("meetingId")?.value || ""
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
      clientId: "aae90a8e-999d-4b18-ba63-f9abfb54ee68", 
      authority: "https://login.microsoftonline.com/bcbce4e4-0e70-42c4-bf70-cf41ea55f075",
      redirectUri: "https://git-uvid.github.io/UVID-Teams-Meeting-Automate/taskpane.html"
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

async function fetchProjects() {
  const projectSelect = document.getElementById("project");
  if (!projectSelect) return;
  
  try {
    const token = await getAccessToken();
    const siteId = "key65akcdgsfg2zhwxauifkam1a.sharepoint.com";
    const listId = "42699509-b16d-4e35-a7d7-0ff95b9b4f0c";
    
    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json"
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch projects: ${await response.text()}`);
    }

    const data = await response.json();
    
    // Clear existing options
    projectSelect.innerHTML = '<option value="">Select a Project</option>';
    
    if (data.value && data.value.length > 0) {
      data.value.forEach(item => {
        const title = item.fields.Title || item.fields.ProjectName;
        if (title) {
          const option = document.createElement("option");
          option.value = title;
          option.text = title;
          projectSelect.appendChild(option);
        }
      });
    } else {
      projectSelect.innerHTML = '<option value="">No projects found</option>';
    }
  } catch (error) {
    console.error("Error fetching projects:", error);
    projectSelect.innerHTML = '<option value="">Error loading projects</option>';
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
