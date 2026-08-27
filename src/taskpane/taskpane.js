/* eslint-env browser, es2022 */
/* global Office, console, msal, document, window, fetch, Option, Quill */

/*
 * UVID Teams Meeting Automate - Taskpane Logic
 */

let userEmail = "";
let editEditor;

Office.onReady(() => {
  // Initialize Quill Editor
  const quill = new window.Quill("#editor-container", {
    theme: "snow",
    modules: {
      toolbar: [
        ["bold", "italic", "underline"],
        [{ list: "ordered" }, { list: "bullet" }],
        ["link"],
      ],
    },
  });

  quill.on("text-change", function () {
    const messageInput = document.getElementById("meetingEventMessage");
    if (messageInput) {
      messageInput.value = quill.root.innerHTML;
    }
  });

  editEditor = new Quill("#edit-editor-container", {
    theme: "snow",
    placeholder: "Compose meeting invite body...",
    modules: {
      toolbar: [
        [{ header: [1, 2, false] }],
        ["bold", "italic", "underline"],
        [{ list: "ordered" }, { list: "bullet" }],
        ["link", "clean"],
      ],
    },
  });

  editEditor.on("text-change", function () {
    const editMessageInput = document.getElementById("editMeetingEventMessage");
    if (editMessageInput) {
      editMessageInput.value = editEditor.root.innerHTML;
    }
  });

  document.getElementById("edit-editor-container").__quill = editEditor;

  document.querySelectorAll(".btn-unlock").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      const targetId = e.currentTarget.getAttribute("data-target");
      if (targetId) {
        const inputEl = document.getElementById(targetId);
        if (inputEl) {
          inputEl.disabled = false;
          inputEl.focus();
        }

        // Special case for Sender Email
        if (targetId === "editSenderEmail") {
          const extractBtn = document.getElementById("btnExtractEmailEdit");
          if (extractBtn) extractBtn.disabled = false;
        }
      }
    });
  });

  document.querySelectorAll(".btn-unlock-editor").forEach((btn) => {
    btn.addEventListener("click", () => {
      const overlay = document.getElementById("edit-editor-overlay");
      if (overlay) overlay.style.display = "none";
      if (editEditor) editEditor.focus();
    });
  });

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
  const senderEmailEl = document.getElementById("senderEmail");
  if (genBy) {
    genBy.value = userEmail;
  }
  if (senderEmailEl) {
    senderEmailEl.value = userEmail;
  }

  document.getElementById("signInBtn")?.addEventListener("click", authenticateUser);

  document.getElementById("btnCancelEdit")?.addEventListener("click", () => {
    handleSubmission(false, true, "Cancel Meeting");
  });

  // Listeners for Tab Navigation
  const tabs = document.querySelectorAll('input[name="tabNav"]');
  tabs.forEach((tab) => {
    tab.addEventListener("change", handleTabChange);
  });

  document.getElementById("flowActionNew")?.addEventListener("change", handleActionChangeNew);

  document.getElementById("startTime")?.addEventListener("change", calculateEndTime);
  const durationEl = document.getElementById("duration");
  durationEl?.addEventListener("input", calculateEndTime);

  document.getElementById("startDate")?.addEventListener("change", (e) => {
    const rStart = document.getElementById("recurrenceStartDate");
    if (rStart) {
      rStart.value = e.target.value;
    }
  });

  document.getElementById("editTime")?.addEventListener("change", calculateEditEndTime);
  document.getElementById("editDuration")?.addEventListener("change", calculateEditEndTime);
  document.getElementById("editDuration")?.addEventListener("input", function () {
    if (this.value.length > 2) {
      this.value = this.value.slice(0, 3);
    }
    if (parseInt(this.value) > 999) {
      this.value = 999;
    }
    calculateEditEndTime();
  });

  document
    .getElementById("newInjectEmailOnly")
    ?.addEventListener("click", () => handleSubmission(true, false));
  document
    .getElementById("newLogSpOnly")
    ?.addEventListener("click", () => handleSubmission(false, true));
  document
    .getElementById("newInjectAndLog")
    ?.addEventListener("click", () => handleSubmission(true, true));

  document
    .getElementById("editInjectEmailOnly")
    ?.addEventListener("click", () => handleSubmission(true, false));
  document
    .getElementById("editLogSpOnly")
    ?.addEventListener("click", () => handleSubmission(false, true));
  document
    .getElementById("editInjectAndLog")
    ?.addEventListener("click", () => handleSubmission(true, true));

  document.getElementById("btnFetchMeeting")?.addEventListener("click", fetchMeetingDetails);
  document
    .getElementById("btnConfirmRecurrenceSelection")
    ?.addEventListener("click", handleRecurrenceSelection);
  document
    .getElementById("btnExtractEmail")
    ?.addEventListener("click", extractEmailAndAddToAttendees);
  document.getElementById("senderEmail")?.addEventListener("change", extractEmailAndAddToAttendees);

  // Duration validation
  const validateDuration = (e) => {
    let val = parseInt(e.target.value);
    if (isNaN(val)) return;
    if (val < 15) val = 15;
    const remainder = val % 5;
    if (remainder !== 0) {
      val = val + (5 - remainder);
    }
    e.target.value = val;
    // Trigger any dependent calculations
    e.target.dispatchEvent(new window.Event("input"));
  };
  document.getElementById("duration")?.addEventListener("change", validateDuration);
  document.getElementById("editDuration")?.addEventListener("change", validateDuration);

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

function calculateEditEndTime() {
  const editTime = document.getElementById("editTime").value;
  const editDuration = parseInt(document.getElementById("editDuration").value);
  const editEndTimeEl = document.getElementById("editEndTime");

  if (editTime && !isNaN(editDuration)) {
    const parts = editTime.split(":");
    let date = new Date();
    date.setHours(parseInt(parts[0], 10));
    date.setMinutes(parseInt(parts[1], 10));

    // Add duration in minutes
    date.setMinutes(date.getMinutes() + editDuration);

    let endHours = String(date.getHours()).padStart(2, "0");
    let endMins = String(date.getMinutes()).padStart(2, "0");
    editEndTimeEl.value = `${endHours}:${endMins}`;

    // Fire input event to trigger checkIfEdited
    editEndTimeEl.dispatchEvent(new window.Event("input"));
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

function extractEmailAndAddToAttendees() {
  const senderEmailEl = document.getElementById("senderEmail");
  const reqAttendeesEl = document.getElementById("requiredAttendees");
  const optAttendeesEl = document.getElementById("optionalAttendees");

  const emailRegex = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi;

  const cleanField = (el, isMulti) => {
    if (!el || !el.value) return [];
    const matches = el.value.match(emailRegex);
    let extracted = [];
    if (matches) {
      if (isMulti) {
        extracted = [...new Set(matches)];
        el.value = extracted.join(";\n");
      } else {
        extracted = [matches[0]];
        el.value = matches[0];
      }
      el.dispatchEvent(new window.Event("input"));
    }
    return extracted;
  };

  cleanField(senderEmailEl, false);
  cleanField(reqAttendeesEl, true);
  cleanField(optAttendeesEl, true);
}

function extractEmailAndAddToAttendeesEdit() {
  const senderEmailEl = document.getElementById("editSenderEmail");
  const reqAttendeesEl = document.getElementById("editRequiredAttendees");
  const optAttendeesEl = document.getElementById("editOptionalAttendees");

  const emailRegex = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi;

  const cleanField = (el, isMulti) => {
    if (!el || !el.value) return [];
    const matches = el.value.match(emailRegex);
    let extracted = [];
    if (matches) {
      if (isMulti) {
        extracted = [...new Set(matches)];
        el.value = extracted.join(";\n");
      } else {
        extracted = [matches[0]];
        el.value = matches[0];
      }
      el.dispatchEvent(new window.Event("input"));
    }
    return extracted;
  };

  cleanField(senderEmailEl, false);
  cleanField(reqAttendeesEl, true);
  cleanField(optAttendeesEl, true);
}

document
  .getElementById("btnExtractEmailEdit")
  ?.addEventListener("click", extractEmailAndAddToAttendeesEdit);
document
  .getElementById("editSenderEmail")
  ?.addEventListener("change", extractEmailAndAddToAttendeesEdit);
document
  .getElementById("editRequiredAttendees")
  ?.addEventListener("change", extractEmailAndAddToAttendeesEdit);
document
  .getElementById("editOptionalAttendees")
  ?.addEventListener("change", extractEmailAndAddToAttendeesEdit);

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

async function handleSubmission(doEmail, doSP, overrideAction = null) {
  const activeTab = document.querySelector('input[name="tabNav"]:checked').value;
  let action = "";

  if (activeTab === "newMeeting") {
    action = document.getElementById("flowActionNew").value;
  } else {
    action = overrideAction || "Change Details";
  }

  const statusEl = document.getElementById("statusMessage");
  statusEl.innerText = "Constructing Email...";
  statusEl.style.color = "blue";

  // Format the email body based on action
  let emailBody = "";

  if (activeTab === "newMeeting") {
    if (action === "Schedule Meeting") {
      const subject = escapeHtml(document.getElementById("meetingSubject").value);
      const bodyMsg = document.getElementById("meetingEventMessage").value;
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

      emailBody = `Subject: ${subject}<br>Message: ${bodyMsg}<br>Date: ${date}<br>Start Time: ${time}<br>End Time: ${endTime}<br>Duration: ${duration}<br>Required Attendees: ${reqAttendees}<br>Optional Attendees: ${optAttendees}<br>Type: ${type}<br>Timezone: ${timezone}`;
    } else if (action === "Schedule Recurring Meeting") {
      const subject = escapeHtml(document.getElementById("meetingSubject").value);
      const bodyMsg = document.getElementById("meetingEventMessage").value;
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
      const recurStartDate = escapeHtml(
        formatDateText(document.getElementById("recurrenceStartDate").value)
      );
      const recurEndDate = escapeHtml(
        formatDateText(document.getElementById("recurrenceEndDate").value)
      );

      emailBody = `Subject: ${subject}<br>Message: ${bodyMsg}<br>Date: ${date}<br>Start Time: ${time}<br>End Time: ${endTime}<br>Timezone: ${timezone}<br>Duration: ${duration}<br>Required Attendees: ${reqAttendees}<br>Optional Attendees: ${optAttendees}<br>Type: ${type}<br>Frequency: ${freq}<br>Recurrence Start Date: ${recurStartDate}<br>Recurrence End Date: ${recurEndDate}`;
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
  } else {
    // Edit mode
    const subject = escapeHtml(document.getElementById("editTitle").value);
    const bodyMsg = document.getElementById("editMeetingEventMessage").value;
    const reason = escapeHtml(document.getElementById("editReason").value);
    emailBody = `Action: ${action}<br>Subject: ${subject}<br>Message: ${bodyMsg}<br>Reason: ${reason}`;
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
  let listId = "093329a2-701b-4e93-9789-f561ef47ddce"; // Default Meetings Log list

  let payload = {};

  if (activeTab === "newMeeting") {
    const isRecurring = action === "Schedule Recurring Meeting";
    // Format start time correctly as ISO string if possible, or just the current string
    let startDateVal = document.getElementById("startDate").value; // YYYY-MM-DD
    let startTimeVal = document.getElementById("startTime").value; // HH:MM
    let startIso = "";
    if (startDateVal && startTimeVal) {
      startIso = startDateVal + "T" + startTimeVal + ":00"; // Real SP expects ISO, omitting Z uses site local time
    }

    let endDateVal = document.getElementById("startDate").value; // Assuming same day
    let endTimeVal = document.getElementById("endTime").value;
    let endIso = "";
    if (endDateVal && endTimeVal) {
      endIso = endDateVal + "T" + endTimeVal + ":00";
    }

    let reqEmails =
      document.getElementById("requiredAttendees")?.value.replace(/[\s,]+/g, ";") || "";
    const senderEmail = document.getElementById("senderEmail")?.value;
    if (senderEmail && !reqEmails.includes(senderEmail)) {
      reqEmails = reqEmails ? reqEmails + ";" + senderEmail : senderEmail;
    }

    payload = {
      fields: {
        MeetingSubject: document.getElementById("meetingSubject")?.value || "",
        MeetingEventmessagecontent: document.getElementById("meetingEventMessage")?.value || "",
        MeetingID: "NEW_ID",
        Project: document.getElementById("project")?.value || "Unknown",
        MeetingType: document.getElementById("meetingType")?.value || "Client",
        LeadEmail: senderEmail || "Unknown",
        Timezone: document.getElementById("timezone")?.value || "UTC",
        Duration_x0028_minutes_x0029_: parseInt(document.getElementById("duration")?.value) || 0,
        Requiredattendees: reqEmails,
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
        Recurrencestartdate:
          isRecurring && document.getElementById("recurrenceStartDate")?.value
            ? document.getElementById("recurrenceStartDate").value + "T00:00:00"
            : null,
        Recurrenceenddate:
          isRecurring && document.getElementById("recurrenceEndDate")?.value
            ? document.getElementById("recurrenceEndDate").value + "T00:00:00"
            : null,
        Starttime: startIso || null,
        Endtime: endIso || null,
        GeneratedBy: document.getElementById("generatedBy")?.value || "Unknown",
      },
    };
  } else {
    // Edit mode: Send ALL data to the Edit Meetings list
    listId = "82fed015-2b74-4b6d-96e7-f95002bcb0ab"; // Edit Meetings List

    payload = { fields: {} };

    // Get all values directly from the fields regardless of disabled state
    const getValue = (id) => document.getElementById(id)?.value || "";

    let reqEmailsEdit = getValue("editRequiredAttendees").replace(/[\s,]+/g, ";");
    const senderEmailEdit = getValue("editSenderEmail");
    if (senderEmailEdit && !reqEmailsEdit.includes(senderEmailEdit)) {
      reqEmailsEdit = reqEmailsEdit ? reqEmailsEdit + ";" + senderEmailEdit : senderEmailEdit;
    }

    payload.fields.ActionCategory = action;
    payload.fields.MeetingSubject = getValue("editTitle");
    payload.fields.Timezone = getValue("editTimezone");
    payload.fields.RequiredAttendees = reqEmailsEdit;
    payload.fields.OptionalAttendees = getValue("editOptionalAttendees").replace(/[\s,]+/g, ";");

    // Formatting StartTime / EndTime to SharePoint expected format (ISO-like string)
    // The inputs are date type and time type. If both exist, combine them.
    const editDate = getValue("editDate");
    const editTime = getValue("editTime");
    const editEndTime = getValue("editEndTime");

    let bodyContent = getValue("editMeetingEventMessage");
    if (typeof editEditor !== "undefined" && editEditor) {
      bodyContent = editEditor.root.innerHTML;
    }
    payload.fields.Body = bodyContent;

    if (editDate && editTime) {
      payload.fields.StartTime = editDate + "T" + editTime + ":00";
    } else {
      payload.fields.StartTime = "";
    }

    if (editDate && editEndTime) {
      // Assuming same day end time for simplicity in UI, if it crosses midnight it would be an issue but UI lacks end date
      payload.fields.EndTime = editDate + "T" + editEndTime + ":00";
    } else {
      payload.fields.EndTime = "";
    }

    payload.fields.Notes = getValue("editReason");

    // Grab meeting ID and event ID
    payload.fields.MeetingID = getValue("masterEventId");
    payload.fields.EventID = getValue("recurrenceInstanceId");

    // Determine TypeOfEvent based on selection
    let typeOfEvent = "";
    if (document.getElementById("recurrenceSelectionContainer")?.style.display === "block") {
      const selectedRadio = document.querySelector('input[name="eventSelection"]:checked');
      if (selectedRadio) {
        if (selectedRadio.value === "MASTER") {
          typeOfEvent = "Master Series ID";
        } else {
          typeOfEvent = "Single Event ID";
        }
      }
    } else {
      // If it's a one-off meeting (no recurrence selection container), it's a single event
      typeOfEvent = "Single Event ID";
    }
    payload.fields.TypeOfEvent = typeOfEvent;
    
    // Silently add the user who triggered the action
    payload.fields.TriggeredBy = userEmail;
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
  } else {
    if (!interactive) {
      const err = new Error("Interaction required");
      err.name = "InteractionRequiredAuthError";
      throw err;
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

    // Clear existing options safely
    projectSelect.options.length = 0;
    projectSelect.options.add(new Option("Select a Project", ""));

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
          projectSelect.options.add(new Option(title, title));
        }
      });
    } else {
      projectSelect.options.length = 0;
      projectSelect.options.add(new Option("No projects found", ""));
    }
  } catch (error) {
    console.error("Error fetching projects:", error);
    projectSelect.options.length = 0;
    if (
      error.message === "Interaction required" ||
      error.name === "BrowserAuthError" ||
      error.name === "InteractionRequiredAuthError" ||
      String(error).includes("popup_window_error") ||
      String(error).includes("interaction_in_progress")
    ) {
      projectSelect.options.add(new Option("Sign in required", ""));
      const authContainer = document.getElementById("authContainer");
      if (authContainer) authContainer.style.display = "block";
    } else {
      projectSelect.options.add(new Option("Error loading projects", ""));
    }
  }
}

let currentFetchedEvents = { master: null, singles: [] };

async function fetchMeetingDetails() {
  const meetingIdInput = document.getElementById("meetingId");
  const meetingId = meetingIdInput ? meetingIdInput.value.trim() : "";
  const statusEl = document.getElementById("fetchMeetingStatus");

  if (!meetingId) {
    if (statusEl) {
      statusEl.innerText = "Please enter a 12-digit Meeting ID.";
      statusEl.style.color = "red";
    }
    return;
  }

  if (statusEl) {
    statusEl.innerText = "Fetching meeting log details...";
    statusEl.style.color = "blue";
  }

  try {
    const token = await getAccessToken();
    const siteId = "key65akcdgsfg2zhwxauifkam1a.sharepoint.com";

    // Step 1: Query Meetings Log (list 1) to get the Teams Meeting ID
    const logListId = "093329a2-701b-4e93-9789-f561ef47ddce";
    const logQueryUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${logListId}/items?expand=fields&$filter=fields/Meeting_ID eq '${meetingId}'&$top=1`;

    const logResponse = await fetch(logQueryUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      },
    });

    if (!logResponse.ok)
      throw new Error(`Failed to fetch meeting log: ${await logResponse.text()}`);

    const logData = await logResponse.json();
    if (!logData.value || logData.value.length === 0) {
      if (statusEl) {
        statusEl.innerText = "Meeting ID not found in Meetings Log.";
        statusEl.style.color = "red";
      }
      return;
    }

    const logFields = logData.value[0].fields;
    const teamsMeetingId = logFields.MeetingID; // Teams ID
    let isRecurring = logFields.Recurring;
    if (typeof isRecurring === "string") isRecurring = isRecurring.toLowerCase() === "true";
    if (isRecurring === 1) isRecurring = true;

    // Fallback: If the Recurring boolean was corrupted/overwritten by a background flow,
    // we can reliably determine it's recurring if it has a Recurrencepattern.
    if (!isRecurring && logFields.Recurrencepattern) {
      isRecurring = true;
    }

    if (statusEl) statusEl.innerText = "Fetching detailed events...";

    // Step 2: Query DetailedEventID (list 2) using TeamsMeetingID
    const detailedListId = "f7439cde-302a-4c2f-b517-74ab117167a7";
    const detailedQueryUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${detailedListId}/items?expand=fields&$filter=fields/TeamsMeetingID eq '${teamsMeetingId}'`;

    const detailedResponse = await fetch(detailedQueryUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      },
    });

    if (!detailedResponse.ok)
      throw new Error(`Failed to fetch detailed events: ${await detailedResponse.text()}`);

    const detailedData = await detailedResponse.json();
    if (!detailedData.value || detailedData.value.length === 0) {
      if (statusEl) {
        statusEl.innerText = "No detailed events found for this meeting.";
        statusEl.style.color = "red";
      }
      return;
    }

    // Enrich detailed events with parent log fields (like SenderEmail and Duration) that only exist in List 1
    const allEvents = detailedData.value.map((item) => {
      const enrichedFields = { ...item.fields };
      // Copy over fields from List 1 that aren't in List 2 but are needed in UI
      enrichedFields.LeadEmail = logFields.LeadEmail;
      enrichedFields.Duration = logFields.Duration_x0028_minutes_x0029_;
      enrichedFields.Recurrencepattern = logFields.Recurrencepattern;
      return enrichedFields;
    });

    // Step 3: Handle One-off vs Recurring
    if (!isRecurring) {
      // One-off
      if (statusEl) {
        statusEl.innerText = `Selected details: Single Event ID\nFound: ${allEvents[0].Subject || "Untitled"}`;
        statusEl.style.color = "green";
      }
      populateEditForm(allEvents[0]);
    } else {
      // Recurring - Show selection UI
      if (statusEl) {
        statusEl.innerText = `Found Recurring Series: ${allEvents[0].Subject || "Untitled"}`;
        statusEl.style.color = "green";
      }

      currentFetchedEvents.master =
        allEvents.find((e) => e.TypeOfEvent === "Master Series ID") || null;
      currentFetchedEvents.singles = allEvents.filter((e) => e.TypeOfEvent === "Single Event ID");

      showRecurrenceSelection();
    }
  } catch (error) {
    console.error("Error fetching meeting details:", error);
    if (statusEl) {
      statusEl.innerText = "Error fetching details. Check console.";
      statusEl.style.color = "red";
    }
  }
}

function showRecurrenceSelection() {
  const recurrenceSelectionContainer = document.getElementById("recurrenceSelectionContainer");
  const editMeetingForm = document.getElementById("editMeetingForm");
  const btnCancelEdit = document.getElementById("btnCancelEdit");

  if (editMeetingForm) editMeetingForm.style.display = "block";
  if (btnCancelEdit) btnCancelEdit.style.display = "block";

  // Hide the actual edit fields until they make a selection
  document
    .querySelectorAll("#editMeetingForm > .form-group:not(#recurrenceSelectionContainer)")
    .forEach((el) => {
      el.style.display = "none";
    });

  if (recurrenceSelectionContainer) recurrenceSelectionContainer.style.display = "block";

  const allEventsList = document.getElementById("allEventsList");
  if (allEventsList) {
    allEventsList.innerHTML = ""; // Clear existing

    // Add Master Series option if present
    if (currentFetchedEvents.master) {
      const label = document.createElement("label");
      label.style.display = "block";
      label.style.marginBottom = "6px";
      label.style.fontWeight = "bold";
      label.style.cursor = "pointer";

      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = "eventSelection";
      radio.value = "MASTER";
      radio.checked = true; // Default select master
      radio.style.marginRight = "6px";

      const mEvent = currentFetchedEvents.master;
      let mStart = "Unknown",
        mEnd = "Unknown",
        mStartTime = "Unknown",
        mEndTime = "Unknown";

      if (mEvent.StartTime) {
        const p = mEvent.StartTime.split("T");
        mStart = formatDateText(p[0]);
        mStartTime = p.length > 1 ? formatTimeText(p[1]) : "";
      }

      const endDateVal = mEvent.RecurrenceEndDate || mEvent.EndTime;
      if (endDateVal) {
        const p = endDateVal.split("T");
        mEnd = formatDateText(p[0]);
      }
      if (mEvent.EndTime) {
        const p = mEvent.EndTime.split("T");
        mEndTime = p.length > 1 ? formatTimeText(p[1]) : "";
      }

      const pattern = mEvent.Recurrencepattern ? ` (${mEvent.Recurrencepattern})` : "";
      const text = `${mStart} - ${mEnd} ${mStartTime} - ${mEndTime} ${mEvent.Subject || "Untitled"}${pattern}`;

      label.appendChild(radio);
      label.appendChild(document.createTextNode(text));
      allEventsList.appendChild(label);
    }

    // Add individual events
    if (currentFetchedEvents.singles && currentFetchedEvents.singles.length > 0) {
      currentFetchedEvents.singles.forEach((event) => {
        const label = document.createElement("label");
        label.style.display = "block";
        label.style.marginBottom = "6px";
        label.style.marginLeft = "20px";
        label.style.cursor = "pointer";

        const radio = document.createElement("input");
        radio.type = "radio";
        radio.name = "eventSelection";
        radio.value = event.EventID;
        radio.style.marginRight = "6px";

        let iDate = "Unknown",
          iStartTime = "Unknown",
          iEndTime = "Unknown";
        if (event.StartTime) {
          const p = event.StartTime.split("T");
          iDate = formatDateText(p[0]);
          iStartTime = p.length > 1 ? formatTimeText(p[1]) : "";
        }
        if (event.EndTime) {
          const p = event.EndTime.split("T");
          iEndTime = p.length > 1 ? formatTimeText(p[1]) : "";
        }

        const text = `${iDate} ${iStartTime} - ${iEndTime} ${event.Subject || "Untitled"}`;

        label.appendChild(radio);
        label.appendChild(document.createTextNode(text));
        allEventsList.appendChild(label);
      });
    }
  }
}

function handleRecurrenceSelection() {
  const selectedRadio = document.querySelector('input[name="eventSelection"]:checked');

  const statusEl = document.getElementById("fetchMeetingStatus");
  if (!selectedRadio) {
    if (statusEl) {
      statusEl.innerText = "Please select an event to edit.";
      statusEl.style.color = "red";
    }
    return;
  }

  let selectedEvent = null;
  if (selectedRadio.value === "MASTER") {
    selectedEvent = currentFetchedEvents.master;
  } else {
    selectedEvent = currentFetchedEvents.singles.find((e) => e.EventID === selectedRadio.value);
  }

  if (selectedEvent) {
    populateEditForm(selectedEvent);
    if (statusEl) {
      statusEl.innerText = `Selected details: ${selectedRadio.value === "MASTER" ? "Master Series ID" : "Single Event ID"}`;
      statusEl.style.color = "green";
    }
  } else {
    if (statusEl) {
      statusEl.innerText = "Could not load the selected event data.";
      statusEl.style.color = "red";
    }
  }
}

let initialEditData = {};

function checkIfEdited() {
  const getVal = (id) => document.getElementById(id)?.value || "";

  const currentData = {
    title: getVal("editTitle"),
    body: getVal("editMeetingEventMessage"),
    date: getVal("editDate"),
    time: getVal("editTime"),
    endTime: getVal("editEndTime"),
    duration: getVal("editDuration"),
    timezone: getVal("editTimezone"),
    senderEmail: getVal("editSenderEmail"),
    reqAttendees: getVal("editRequiredAttendees"),
    optAttendees: getVal("editOptionalAttendees"),
    reason: getVal("editReason"),
  };

  let hasChanges = false;
  for (const key in currentData) {
    if (currentData[key] !== initialEditData[key]) {
      hasChanges = true;
      break;
    }
  }

  const btnEmailOnly = document.getElementById("editInjectEmailOnly");
  const btnLogSpOnly = document.getElementById("editLogSpOnly");
  const btnInjectLog = document.getElementById("editInjectAndLog");

  const isNotesEmpty = !currentData.reason.trim();

  if (btnEmailOnly) btnEmailOnly.disabled = !hasChanges || isNotesEmpty;
  if (btnLogSpOnly) btnLogSpOnly.disabled = !hasChanges || isNotesEmpty;
  if (btnInjectLog) btnInjectLog.disabled = !hasChanges || isNotesEmpty;
}

function populateEditForm(event) {
  // Show edit fields
  document
    .querySelectorAll("#editMeetingForm > .form-group:not(#recurrenceSelectionContainer)")
    .forEach((el) => {
      el.style.display = "block";
    });
  const recurrenceSelectionContainer = document.getElementById("recurrenceSelectionContainer");
  if (recurrenceSelectionContainer) recurrenceSelectionContainer.style.display = "none";

  const editMeetingForm = document.getElementById("editMeetingForm");
  const btnCancelEdit = document.getElementById("btnCancelEdit");

  if (editMeetingForm) editMeetingForm.style.display = "block";
  if (btnCancelEdit) btnCancelEdit.style.display = "block";

  const trig = document.getElementById("triggeredBy");
  if (trig) trig.value = userEmail;

  if (!event) return;

  const setValue = (id, val) => {
    const el = document.getElementById(id);
    if (el) {
      el.value = val || "";
    }
  };

  setValue("editTitle", event.Subject || event.Title);

  const editEditorEl = document.getElementById("edit-editor-container");
  if (editEditorEl && window.Quill) {
    editEditorEl.__quill = Quill.find(editEditorEl) || null;
    if (!editEditorEl.__quill && typeof editEditor !== "undefined") {
      editEditorEl.__quill = editEditor;
    }
    if (editEditorEl.__quill) {
      editEditorEl.__quill.root.innerHTML = event.Body || "";
    }
  }
  setValue("editMeetingEventMessage", event.Body || "");

  // Attach listeners to trigger checkIfEdited
  const inputsToCheck = [
    "editTitle",
    "editDate",
    "editTime",
    "editEndTime",
    "editDuration",
    "editTimezone",
    "editSenderEmail",
    "editRequiredAttendees",
    "editOptionalAttendees",
    "editReason",
  ];
  inputsToCheck.forEach((id) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener("input", checkIfEdited);
      el.addEventListener("change", checkIfEdited);
      el.disabled = true; // Start locked
    }
  });

  if (typeof editEditor !== "undefined" && editEditor) {
    editEditor.on("text-change", checkIfEdited);
  }

  const startDate = event.StartTime;
  if (startDate) {
    const startDateParts = startDate.split(" ");
    if (startDateParts.length > 1) {
      // Format: "08/29/2026 00:26:00" -> MM/DD/YYYY HH:MM:SS
      const dParts = startDateParts[0].split("/"); // [08, 29, 2026]
      if (dParts.length === 3) {
        setValue(
          "editDate",
          `${dParts[2]}-${dParts[0].padStart(2, "0")}-${dParts[1].padStart(2, "0")}`
        );
      }
      setValue("editTime", startDateParts[1].substring(0, 5));
    } else if (startDate.includes("T")) {
      const parts = startDate.split("T");
      setValue("editDate", parts[0]);
      setValue("editTime", parts[1].substring(0, 5));
    }
  }

  const endDate = event.EndTime;
  if (endDate) {
    const endDateParts = endDate.split(" ");
    if (endDateParts.length > 1) {
      setValue("editEndTime", endDateParts[1].substring(0, 5));
    } else if (endDate.includes("T")) {
      setValue("editEndTime", endDate.split("T")[1].substring(0, 5));
    }
  }

  setValue("editTimezone", event.Timezone || event.TimeZone);
  setValue(
    "editRequiredAttendees",
    event.RequiredAttendees ? event.RequiredAttendees.replace(/;/g, "\n") : ""
  );
  setValue(
    "editOptionalAttendees",
    event.OptionalAttendees ? event.OptionalAttendees.replace(/;/g, "\n") : ""
  );
  setValue("editReason", event.CancelReason || "");
  setValue("masterEventId", event.TeamsMeetingID || "");
  setValue("recurrenceInstanceId", event.EventID || "");
  setValue("editSenderEmail", event.LeadEmail || "");
  setValue("editDuration", event.Duration || "");

  // Lock all inputs initially
  const inputsToLock = [
    "editTitle",
    "editDate",
    "editTime",
    "editEndTime",
    "editDuration",
    "editTimezone",
    "editSenderEmail",
    "editRequiredAttendees",
    "editOptionalAttendees",
    "editReason",
  ];
  inputsToLock.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.disabled = true;
  });

  const extractBtn = document.getElementById("btnExtractEmailEdit");
  if (extractBtn) extractBtn.disabled = true;

  const overlay = document.getElementById("edit-editor-overlay");
  if (overlay) overlay.style.display = "block";

  // Store initial state
  const getVal = (id) => document.getElementById(id)?.value || "";
  initialEditData = {
    title: getVal("editTitle"),
    body: getVal("editMeetingEventMessage"),
    date: getVal("editDate"),
    time: getVal("editTime"),
    endTime: getVal("editEndTime"),
    duration: getVal("editDuration"),
    timezone: getVal("editTimezone"),
    senderEmail: getVal("editSenderEmail"),
    reqAttendees: getVal("editRequiredAttendees"),
    optAttendees: getVal("editOptionalAttendees"),
    reason: getVal("editReason"),
  };

  checkIfEdited();
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

    newTzSelect.options.length = 0;
    editTzSelect.options.length = 0;

    if (data.value && data.value.length > 0) {
      data.value.forEach((item) => {
        if (
          item.fields.Active === true ||
          String(item.fields.Active).toLowerCase() === "true" ||
          item.fields.Active === 1
        ) {
          const tzText = item.fields.TimeZone || "Unknown";
          const calTzVal = item.fields.CalenderTimeZone || tzText;

          newTzSelect.options.add(new Option(tzText, calTzVal));
          editTzSelect.options.add(new Option(tzText, calTzVal));
        }
      });

      // Fallback if none active
      if (newTzSelect.options.length === 0) {
        newTzSelect.options.add(new Option("No active time zones", ""));
        editTzSelect.options.add(new Option("No active time zones", ""));
      }
    } else {
      newTzSelect.options.add(new Option("No time zones found", ""));
      editTzSelect.options.add(new Option("No time zones found", ""));
    }
  } catch (error) {
    console.error("Error fetching time zones:", error);
    newTzSelect.options.length = 0;
    editTzSelect.options.length = 0;
    if (
      error.message === "Interaction required" ||
      error.name === "BrowserAuthError" ||
      error.name === "InteractionRequiredAuthError" ||
      String(error).includes("popup_window_error") ||
      String(error).includes("interaction_in_progress")
    ) {
      newTzSelect.options.add(new Option("Sign in required", ""));
      editTzSelect.options.add(new Option("Sign in required", ""));
      const authContainer = document.getElementById("authContainer");
      if (authContainer) authContainer.style.display = "block";
    } else {
      newTzSelect.options.add(new Option("Error loading time zones", ""));
      editTzSelect.options.add(new Option("Error loading time zones", ""));
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
