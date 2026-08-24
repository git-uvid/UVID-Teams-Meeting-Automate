/* eslint-env browser, es2022 */
/* global Office, console, msal, document, window, fetch */

/*
 * UVID Teams Meeting Automate - Taskpane Logic
 */

let userEmail = "";

Office.onReady(() => {
  if (
    window.Office &&
    window.Office.context &&
    window.Office.context.mailbox &&
    window.Office.context.mailbox.userProfile
  ) {
    userEmail = window.Office.context.mailbox.userProfile.emailAddress;
  } else {
    userEmail = "TestUser@uvidconsulting.com";
  }

  const genBy = document.getElementById("generatedBy");
  if (genBy) {
    genBy.value = userEmail;
  }

  document.getElementById("signInBtn")?.addEventListener("click", authenticateUser);

  document.getElementById("btnCancelEdit")?.addEventListener("click", () => {
    document.getElementById("editMeetingForm").style.display = "none";
    document.getElementById("idSection").style.display = "block";
    document.getElementById("btnCancelEdit").style.display = "none";
    document.getElementById("fetchMeetingStatus").innerText = "";
  });

  document.querySelectorAll(".btn-unlock").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      const targetId = e.target.getAttribute("data-target");
      if (targetId) {
        const inputEl = document.getElementById(targetId);
        if (inputEl) {
          inputEl.disabled = false;
          inputEl.focus();
        }
      }
    });
  });

  // Listeners for Tab Navigation
  const tabs = document.querySelectorAll('input[name="tabNav"]');
  tabs.forEach((tab) => {
    tab.addEventListener("change", handleTabChange);
  });

  document.getElementById("flowActionNew")?.addEventListener("change", handleActionChangeNew);

  document.getElementById("startTime")?.addEventListener("change", calculateEndTime);
  document.getElementById("duration")?.addEventListener("input", calculateEndTime);

  document.getElementById("startDate")?.addEventListener("change", (e) => {
    const rStart = document.getElementById("recurrenceStartDate");
    if (rStart) {
      rStart.value = e.target.value;
    }
  });

  document
    .getElementById("injectEmailOnly")
    ?.addEventListener("click", () => handleSubmission(true, false));
  document
    .getElementById("logSpOnly")
    ?.addEventListener("click", () => handleSubmission(false, true));
  document
    .getElementById("injectAndLog")
    ?.addEventListener("click", () => handleSubmission(true, true));

  document.getElementById("btnFetchMeeting")?.addEventListener("click", fetchMeetingDetails);

  // Initial setup
  handleTabChange();

  // Fetch projects on load
  fetchProjects();
  fetchTimezones();
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

    let endHours = String(date.getHours()).padStart(2, "0");
    let endMins = String(date.getMinutes()).padStart(2, "0");
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

document.getElementById("recurrenceFrequency")?.addEventListener("change", (e) => {
  const daysSection = document.getElementById("selectedDaysOfWeekSection");
  if (daysSection) {
    daysSection.style.display = e.target.value === "Weekly" ? "block" : "none";
  }
});

// Function to format date to "15 July 2026"
function formatDateText(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  const options = { day: "numeric", month: "long", year: "numeric" };
  return d.toLocaleDateString("en-GB", options);
}

// Function to format time to "2:00 PM"
function formatTimeText(timeStr) {
  if (!timeStr) return "";
  const parts = timeStr.split(":");
  if (parts.length < 2) return timeStr;
  let h = parseInt(parts[0], 10);
  let m = parts[1];
  let ampm = h >= 12 ? "PM" : "AM";
  h = h % 12;
  h = h ? h : 12;
  return `${h}:${m} ${ampm}`;
}

function escapeHtml(unsafe) {
  if (!unsafe) return "";
  return unsafe
    .toString()
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
    const bodyMsg = escapeHtml(document.getElementById("meetingEventMessage").value).replace(
      /\n/g,
      "<br>"
    );
    const date = escapeHtml(formatDateText(document.getElementById("startDate").value));
    const time = escapeHtml(formatTimeText(document.getElementById("startTime").value));
    const endTime = escapeHtml(formatTimeText(document.getElementById("endTime").value));
    const duration = escapeHtml(document.getElementById("duration").value);
    const reqAttendees = escapeHtml(
      document.getElementById("requiredAttendees").value.replace(/[\s,]+/g, ";")
    );
    const optAttendees = escapeHtml(
      document.getElementById("optionalAttendees").value.replace(/[\s,]+/g, ";")
    );
    const type = escapeHtml(document.getElementById("meetingType").value);
    const timezone = escapeHtml(document.getElementById("timezone").value);

    emailBody = `Title: ${title}<br>Subject: ${subject}<br>Message: ${bodyMsg}<br>Date: ${date}<br>Start Time: ${time}<br>End Time: ${endTime}<br>Duration: ${duration}<br>Required Attendees: ${reqAttendees}<br>Optional Attendees: ${optAttendees}<br>Type: ${type}<br>Timezone: ${timezone}`;
  } else if (action === "Schedule Recurring Meeting") {
    const title = escapeHtml(document.getElementById("meetingTitle").value);
    const subject = escapeHtml(document.getElementById("meetingSubject").value);
    const bodyMsg = escapeHtml(document.getElementById("meetingEventMessage").value).replace(
      /\n/g,
      "<br>"
    );
    const date = escapeHtml(formatDateText(document.getElementById("startDate").value));
    const time = escapeHtml(formatTimeText(document.getElementById("startTime").value));
    const endTime = escapeHtml(formatTimeText(document.getElementById("endTime").value));
    const duration = escapeHtml(document.getElementById("duration").value);
    const reqAttendees = escapeHtml(
      document.getElementById("requiredAttendees").value.replace(/[\s,]+/g, ";")
    );
    const optAttendees = escapeHtml(
      document.getElementById("optionalAttendees").value.replace(/[\s,]+/g, ";")
    );
    const type = escapeHtml(document.getElementById("meetingType").value);
    const timezone = escapeHtml(document.getElementById("timezone").value);
    const freq = escapeHtml(document.getElementById("recurrenceFrequency").value);
    const interval = escapeHtml(document.getElementById("recurrenceInterval").value);
    const recurStartDate = escapeHtml(
      formatDateText(document.getElementById("recurrenceStartDate").value)
    );
    const recurEndDate = escapeHtml(
      formatDateText(document.getElementById("recurrenceEndDate").value)
    );

    emailBody = `Title: ${title}<br>Subject: ${subject}<br>Message: ${bodyMsg}<br>Date: ${date}<br>Start Time: ${time}<br>End Time: ${endTime}<br>Timezone: ${timezone}<br>Duration: ${duration}<br>Required Attendees: ${reqAttendees}<br>Optional Attendees: ${optAttendees}<br>Type: ${type}<br>Frequency: ${freq}<br>Interval: ${interval}<br>Recurrence Start Date: ${recurStartDate}<br>Recurrence End Date: ${recurEndDate}`;
  } else if (action === "Reschedule Meeting") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const date = escapeHtml(formatDateText(document.getElementById("newDate").value));
    const time = escapeHtml(formatTimeText(document.getElementById("newTime").value));
    const timezone = escapeHtml(document.getElementById("timezoneEdit").value);
    const reason = escapeHtml(document.getElementById("reason").value);

    emailBody = `Meeting ID: ${id}<br>New Date: ${date}<br>New Time: ${time}<br>Timezone: ${timezone}<br>Reason: ${reason}`;
  } else if (action === "Reschedule Series") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const day = escapeHtml(document.getElementById("newDay").value);
    const time = escapeHtml(formatTimeText(document.getElementById("newTime").value));
    const timezone = escapeHtml(document.getElementById("timezoneEdit").value);
    const reason = escapeHtml(document.getElementById("reason").value);

    emailBody = `Series ID: ${id}<br>New Day: ${day}<br>New Time: ${time}<br>Timezone: ${timezone}<br>Reason: ${reason}`;
  } else if (action === "Cancel Meeting") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const reason = escapeHtml(document.getElementById("reason").value);

    emailBody = `Meeting ID: ${id}<br>Reason: ${reason}`;
  } else if (action === "Cancel Series") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const reason = escapeHtml(document.getElementById("reason").value);

    emailBody = `Series ID: ${id}<br>Reason: ${reason}`;
  } else if (action === "Add Participant") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const newParticipant = escapeHtml(document.getElementById("newParticipantEmail").value);
    const name = escapeHtml(document.getElementById("newParticipantName").value);

    emailBody = `Meeting ID: ${id}<br>New Participant: ${newParticipant}`;
    if (name) emailBody += `<br>Name: ${name}`;
  } else if (action === "Update Meeting") {
    const id = escapeHtml(document.getElementById("meetingId").value);
    const newTitle = escapeHtml(document.getElementById("newTitle").value);

    emailBody = `Meeting ID: ${id}<br>New Title: ${newTitle}`;
  }

  if (window.Office && window.Office.context && window.Office.context.mailbox) {
    if (doEmail) {
      statusEl.innerText = "Constructing Email...";
      statusEl.style.color = "blue";

      // Clear To, CC, BCC
      Office.context.mailbox.item.to.setAsync(
        [{ emailAddress: "connect@uvidconsulting.com" }],
        (res) => {
          if (res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
        }
      );
      Office.context.mailbox.item.cc.setAsync([], (res) => {
        if (res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
      });
      Office.context.mailbox.item.bcc.setAsync([], (res) => {
        if (res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
      });

      Office.context.mailbox.item.subject.setAsync(action, (res) => {
        if (res.status === Office.AsyncResultStatus.Failed) console.error(res.error);
      });

      // Append Regards and Disclaimer
      const name = userEmail ? userEmail.split("@")[0].replace(".", " ") : "User";
      const capitalizedName = name.replace(/\b\w/g, (l) => l.toUpperCase());
      const finalEmailBody =
        emailBody +
        `<br><br>Best regards,<br>${capitalizedName}<br><br><i>[Auto-generated email by UVID Teams Meeting Automate]</i>`;

      Office.context.mailbox.item.body.setAsync(
        finalEmailBody,
        { coercionType: Office.CoercionType.Html },
        async (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            statusEl.innerText = "Error injecting content: " + asyncResult.error.message;
            statusEl.style.color = "red";
            return;
          }

          if (doSP) {
            await executeSPLog(
              action,
              activeTab,
              statusEl,
              "Success! Email populated and logged to SharePoint.",
              "Email populated, but failed to log to SharePoint."
            );
          } else {
            statusEl.innerText = "Success! Email constructed.";
            statusEl.style.color = "green";
          }
        }
      );
    } else if (doSP) {
      await executeSPLog(
        action,
        activeTab,
        statusEl,
        "Success! Logged to SharePoint.",
        "Failed to log to SharePoint."
      );
    }
  } else {
    // If testing in browser without Office.js
    console.log("Subject:", action);
    console.log("Body:", emailBody);

    if (doSP) {
      await executeSPLog(
        action,
        activeTab,
        statusEl,
        "Testing outside Outlook. Logged to SharePoint.",
        "Testing outside Outlook. SP Log failed."
      );
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
        Requiredattendees:
          document.getElementById("requiredAttendees")?.value.replace(/[\s,]+/g, ";") || "",
        Optionalattendees:
          document.getElementById("optionalAttendees")?.value.replace(/[\s,]+/g, ";") || "",
        Recurring: isRecurring,
        Recurrencepattern: isRecurring ? document.getElementById("recurrenceFrequency")?.value : "",
        SelectedDaysOfWeek:
          isRecurring && document.getElementById("recurrenceFrequency")?.value === "Weekly"
            ? JSON.stringify(
                Array.from(document.querySelectorAll('input[name="recurrenceDays"]:checked')).map(
                  (cb) => cb.value
                )
              )
            : "[]",
        Recurrenceinterval: isRecurring
          ? parseInt(document.getElementById("recurrenceInterval")?.value) || 1
          : 0,
        Recurrencestartdate:
          isRecurring && document.getElementById("recurrenceStartDate")?.value
            ? document.getElementById("recurrenceStartDate").value + "T00:00:00Z"
            : null,
        Recurrenceenddate:
          isRecurring && document.getElementById("recurrenceEndDate")?.value
            ? document.getElementById("recurrenceEndDate").value + "T00:00:00Z"
            : null,
        Starttime: startIso || null,
        Endtime: endIso || null,
        GeneratedBy: document.getElementById("generatedBy")?.value || "Unknown",
      },
    };
  } else {
    // Edit mode: construct payload ONLY with modified (unlocked) fields
    payload = { fields: {} };

    payload.fields.MeetingID = document.getElementById("meetingId")?.value || "";
    payload.fields.SeriesID =
      document.getElementById("masterEventId")?.value ||
      document.getElementById("meetingId")?.value ||
      "";
    payload.fields.MeetingSubject = document.getElementById("flowActionEdit")?.value || action;
    payload.fields.GeneratedBy = document.getElementById("triggeredBy")?.value || "Unknown";

    const getIfUnlocked = (id) => {
      const el = document.getElementById(id);
      return el && !el.disabled ? el.value : undefined;
    };

    const title = getIfUnlocked("editTitle");
    if (title !== undefined) payload.fields.Title = title;

    const tz = getIfUnlocked("editTimezone");
    if (tz !== undefined) payload.fields.Timezone = tz;

    const duration = getIfUnlocked("editDuration");
    if (duration !== undefined)
      payload.fields.Duration_x0028_minutes_x0029_ = parseInt(duration) || 0;

    const attendees = getIfUnlocked("editRequiredAttendees");
    if (attendees !== undefined)
      payload.fields.Requiredattendees = attendees.replace(/[\s,]+/g, ";");

    const reason = getIfUnlocked("editReason");
    if (reason !== undefined) payload.fields.MeetingEventmessagecontent = "Reason: " + reason;

    const dateStr = getIfUnlocked("editDate");
    const timeStr = getIfUnlocked("editTime");
    if (dateStr !== undefined || timeStr !== undefined) {
      const d = document.getElementById("editDate")?.value;
      const t = document.getElementById("editTime")?.value;
      if (d && t) payload.fields.Starttime = d + "T" + t + ":00Z";
    }
  }

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to log to SP: ${errorText}`);
  }
}

let msalInstance;

async function initMsal() {
  if (!msalInstance) {
    const msalConfig = {
      auth: {
        clientId: "aae90a8e-999d-4b18-ba63-f9abfb54ee68",
        authority: "https://login.microsoftonline.com/bcbce4e4-0e70-42c4-bf70-cf41ea55f075",
        redirectUri: window.location.href.split("?")[0].split("#")[0],
      },
    };
    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();
  }
}

async function getAccessToken(interactive = false) {
  await initMsal();
  const accounts = msalInstance.getAllAccounts();
  const request = { scopes: ["Sites.ReadWrite.All"] };

  if (accounts.length > 0) {
    request.account = accounts[0];
    if (!interactive) {
      try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
      } catch (e) {
        console.warn("Silent token acquisition failed.", e);
        throw e;
      }
    }
  }

  // If interactive is true OR no accounts exist yet, prompt popup
  const response = await msalInstance.acquireTokenPopup(request);
  return response.accessToken;
}

async function fetchProjects() {
  const projectSelect = document.getElementById("project");
  if (!projectSelect) return;

  try {
    const token = await getAccessToken();
    const siteId = "key65akcdgsfg2zhwxauifkam1a.sharepoint.com";
    const listId = "42699509-b16d-4e35-a7d7-0ff95b9b4f0c";

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
        },
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch projects: ${await response.text()}`);
    }

    const data = await response.json();

    // Clear existing options
    projectSelect.innerHTML = '<option value="">Select a Project</option>';

    if (data.value && data.value.length > 0) {
      data.value.forEach((item) => {
        // Skip completed projects
        if (
          item.fields.IsCompleted === true ||
          String(item.fields.IsCompleted).toLowerCase() === "true" ||
          item.fields.IsCompleted === 1
        ) {
          return;
        }

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

async function fetchMeetingDetails() {
  const meetingIdInput = document.getElementById("meetingId");
  const meetingId = meetingIdInput ? meetingIdInput.value.trim() : "";
  const statusEl = document.getElementById("fetchMeetingStatus");

  if (!meetingId) {
    if (statusEl) {
      statusEl.innerText = "Please enter a Meeting ID.";
      statusEl.style.color = "red";
    }
    return;
  }

  if (statusEl) {
    statusEl.innerText = "Fetching details...";
    statusEl.style.color = "blue";
  }

  try {
    const token = await getAccessToken();
    const siteId = "key65akcdgsfg2zhwxauifkam1a.sharepoint.com";
    const listId = "093329a2-701b-4e93-9789-f561ef47ddce";

    // Use $filter on fields/Meeting_ID. Use Prefer header to allow non-indexed queries just in case.
    // Also orderBy Created desc to get the latest if multiple exist.
    const queryUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields&$filter=fields/Meeting_ID eq '${meetingId}'&$orderby=fields/Createdat desc&$top=1`;

    const response = await fetch(queryUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      },
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch meeting: ${await response.text()}`);
    }

    const data = await response.json();

    if (data.value && data.value.length > 0) {
      const itemFields = data.value[0].fields;

      if (statusEl) {
        statusEl.innerText = `Found: ${itemFields.Title || "Untitled"}`;
        statusEl.style.color = "green";
      }

      const idSection = document.getElementById("idSection");
      const editMeetingForm = document.getElementById("editMeetingForm");
      const btnCancelEdit = document.getElementById("btnCancelEdit");
      if (idSection) idSection.style.display = "none";
      if (editMeetingForm) editMeetingForm.style.display = "block";
      if (btnCancelEdit) btnCancelEdit.style.display = "block";

      document
        .querySelectorAll(
          "#editMeetingForm input, #editMeetingForm select, #editMeetingForm textarea"
        )
        .forEach((el) => {
          if (
            el.id !== "flowActionEdit" &&
            el.id !== "triggeredBy" &&
            el.id !== "masterEventId" &&
            el.id !== "recurrenceInstanceId"
          ) {
            el.disabled = true;
          }
        });

      const trig = document.getElementById("triggeredBy");
      if (trig) trig.value = userEmail;

      const setVal = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.value = val || "";
      };

      setVal("editTitle", itemFields.Title);
      setVal("editDuration", itemFields.Duration_x0028_minutes_x0029_ || itemFields.Duration || "");
      setVal("editTimezone", itemFields.Timezone);
      setVal(
        "editRequiredAttendees",
        itemFields.Requiredattendees ? itemFields.Requiredattendees.replace(/;/g, "\n") : ""
      );

      if (itemFields.Starttime) {
        const startDateStr = itemFields.Starttime.split("T")[0];
        let startTimeStr = itemFields.Starttime.split("T")[1];
        if (startTimeStr) startTimeStr = startTimeStr.substring(0, 5);
        setVal("editDate", startDateStr);
        setVal("editTime", startTimeStr);
      } else {
        setVal("editDate", "");
        setVal("editTime", "");
      }

      if (
        itemFields.Recurring === true ||
        String(itemFields.Recurring).toLowerCase() === "true" ||
        itemFields.Recurring === 1
      ) {
        const recSection = document.getElementById("recurrenceEditSection");
        if (recSection) recSection.style.display = "block";
        setVal("masterEventId", itemFields.MeetingID || "");
      } else {
        const recSection = document.getElementById("recurrenceEditSection");
        if (recSection) recSection.style.display = "none";
        setVal("masterEventId", "");
        setVal("recurrenceInstanceId", "");
      }
    } else {
      if (statusEl) {
        statusEl.innerText = "Meeting ID not found in SharePoint.";
        statusEl.style.color = "red";
      }
    }
  } catch (error) {
    console.error("Error fetching meeting details:", error);
    if (statusEl) {
      statusEl.innerText = "Error fetching details. Check console.";
      statusEl.style.color = "red";
    }
  }
}

// Utility: email validation
document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll(".email-input, .multi-email-input").forEach((el) => {
    el.addEventListener("input", function () {
      const val = this.value.trim();
      const isMulti = this.classList.contains("multi-email-input");
      const errText = this.parentElement.querySelector(".error-text");
      let hasError = false;
      if (val) {
        if (isMulti) {
          const emails = val.split(/[\s,;]+/);
          emails.forEach((e) => {
            if (e && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) hasError = true;
          });
        } else {
          if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(val)) hasError = true;
        }
      }
      if (hasError) {
        this.classList.add("error-border");
        if (errText) errText.style.display = "block";
      } else {
        this.classList.remove("error-border");
        if (errText) errText.style.display = "none";
      }
    });
  });
});

async function fetchTimezones() {
  const newTzSelect = document.getElementById("timezone");
  const editTzSelect = document.getElementById("editTimezone");
  if (!newTzSelect || !editTzSelect) return;

  try {
    const token = await getAccessToken();
    const siteId = "key65akcdgsfg2zhwxauifkam1a.sharepoint.com";
    const listId = "a7c59750-ef7f-4395-bdba-e65dd47f2a90";

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
        },
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch time zones: ${await response.text()}`);
    }

    const data = await response.json();

    newTzSelect.innerHTML = "";
    editTzSelect.innerHTML = "";

    if (data.value && data.value.length > 0) {
      data.value.forEach((item) => {
        if (
          item.fields.Active === true ||
          String(item.fields.Active).toLowerCase() === "true" ||
          item.fields.Active === 1
        ) {
          const tzText = item.fields.TimeZone || "Unknown";
          const newTzVal = item.fields.TeamsTimeZone || tzText;
          const editTzVal = item.fields.CalenderTimeZone || tzText;

          const newOption = document.createElement("option");
          newOption.value = newTzVal;
          newOption.text = tzText;
          newTzSelect.appendChild(newOption);

          const editOption = document.createElement("option");
          editOption.value = editTzVal;
          editOption.text = tzText;
          editTzSelect.appendChild(editOption);
        }
      });

      // Fallback if none active
      if (newTzSelect.options.length === 0) {
        newTzSelect.innerHTML = '<option value="">No active time zones</option>';
        editTzSelect.innerHTML = '<option value="">No active time zones</option>';
      }
    } else {
      newTzSelect.innerHTML = '<option value="">No time zones found</option>';
      editTzSelect.innerHTML = '<option value="">No time zones found</option>';
    }
  } catch (error) {
    console.error("Error fetching time zones:", error);
    if (
      error.message === "Interaction required" ||
      error.name === "BrowserAuthError" ||
      String(error).includes("popup_window_error")
    ) {
      newTzSelect.innerHTML = '<option value="">Sign in required</option>';
      editTzSelect.innerHTML = '<option value="">Sign in required</option>';
      const authContainer = document.getElementById("authContainer");
      if (authContainer) authContainer.style.display = "block";
    } else {
      newTzSelect.innerHTML = '<option value="">Error loading time zones</option>';
      editTzSelect.innerHTML = '<option value="">Error loading time zones</option>';
    }
  }
}

async function authenticateUser() {
  try {
    await getAccessToken(true);
    const authContainer = document.getElementById("authContainer");
    if (authContainer) authContainer.style.display = "none";

    // Refresh dropdowns
    fetchProjects();
    fetchTimezones();
  } catch (error) {
    console.error("Authentication failed:", error);
  }
}
