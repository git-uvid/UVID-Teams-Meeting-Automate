with open('tz_chunk.txt', 'r', encoding='utf-8') as f:
    tz_chunk = f.read()

html_start = """<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>UVID Teams Meeting Automate</title>
    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="https://alcdn.msauth.net/browser/2.30.0/js/msal-browser.min.js"></script>
    <style>
      body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; padding: 10px; margin: 0; background-color: #ffffff; color: #000000; }
      h2 { font-size: 16px; font-weight: 600; color: #000; margin-bottom: 15px; text-align: center; }
      .form-group { margin-bottom: 12px; }
      label { display: block; font-weight: 600; font-size: 13px; margin-bottom: 4px; color: #000; }
      input[type="text"], input[type="date"], input[type="time"], input[type="number"], select, textarea {
        width: 100%; box-sizing: border-box; padding: 6px 8px; border: 1px solid #000;
        border-radius: 4px; font-size: 13px; font-family: inherit; background-color: #fff; color: #000;
      }
      textarea { resize: vertical; min-height: 50px; }
      .checkbox-group { display: flex; align-items: center; }
      .checkbox-group input { width: auto; margin-right: 8px; }
      .checkbox-group label { margin-bottom: 0; font-weight: normal; }
      .button-group { margin-top: 20px; display: flex; flex-direction: column; gap: 8px; }
      .action-btn {
        width: 100%; padding: 8px; font-size: 14px; font-weight: 600; border: 1px solid #000; border-radius: 4px; cursor: pointer;
      }
      .primary { background-color: #0f6cbd; color: white; border-color: #0f6cbd; }
      .primary:hover { background-color: #0c5696; }
      .secondary { background-color: #f3f2f1; color: #000; }
      .secondary:hover { background-color: #edebe9; }
      .status { margin-top: 15px; font-size: 13px; font-weight: 600; text-align: center; color: #000; }
      .small-btn { margin-left: auto; cursor: pointer; background: #fff; border: 1px solid #000; border-radius: 4px; font-size: 12px; color: #000; }
      .help-text { display: block; font-size: 11px; color: #333; margin-top: 2px; }
      .error-border { border-color: #d13438 !important; outline: 1px solid #d13438; }
      .error-text { color: #d13438; font-size: 11px; margin-top: 2px; display: none; }
      .section-box { border: 1px solid #ccc; padding: 10px; border-radius: 4px; margin-bottom: 15px; background-color: #fafafa; }
    </style>
  </head>
  <body>
    <main id="app-body">
      <h2>Meeting Setup & Testing</h2>

      <div class="form-group section-box">
        <label for="flowAction">Flow Action (Select what you want to do)</label>
        <select id="flowAction">
          <option value="Schedule Meeting">Schedule Meeting (One-off)</option>
          <option value="Schedule Recurring Meeting">Schedule Recurring Meeting</option>
          <option value="Reschedule Meeting">Reschedule Meeting</option>
          <option value="Reschedule Series">Reschedule Series</option>
          <option value="Cancel Meeting">Cancel Meeting</option>
          <option value="Cancel Series">Cancel Series</option>
          <option value="Add Participant">Add Participant</option>
          <option value="Update Meeting">Update Meeting (Change Title)</option>
        </select>
        <span class="help-text">This automatically sets the email subject.</span>
      </div>

      <!-- ID Lookup Field (For Edit modes) -->
      <div class="form-group section-box" id="idSection" style="display: none;">
        <label for="meetingId">Meeting ID / Series ID</label>
        <input type="text" id="meetingId" placeholder="Paste ID from SharePoint or confirmation email" />
        <span class="help-text">Required for updates, rescheduling, or cancelling.</span>
      </div>

      <div class="ms-TextField">

        <!-- CREATE FIELDS -->
        <div id="createFields">
          <div class="form-group">
            <label for="meetingTitle">Meeting Title</label>
            <input type="text" id="meetingTitle" />
            <span class="help-text">The title of the meeting.</span>
          </div>
          <div class="form-group">
            <label for="project">Project</label>
            <input type="text" id="project" />
          </div>
          <div class="form-group">
            <label for="meetingType">Meeting Type</label>
            <select id="meetingType">
              <option value="Client">Client</option>
              <option value="Internal">Internal</option>
            </select>
          </div>
          <div class="form-group">
            <label for="leadEmail">Lead Email</label>
            <input type="text" id="leadEmail" class="email-input" />
            <div class="error-text">Invalid email format.</div>
          </div>
          <div class="form-group">
            <label for="participants">Participants</label>
            <textarea id="participants" class="multi-email-input" placeholder="Paste emails here"></textarea>
            <span class="help-text">Emails separated by commas, spaces, or newlines. Output uses semicolons.</span>
            <div class="error-text">Some emails are invalid.</div>
          </div>
          
          <div class="form-group">
            <label for="startDate">Date</label>
            <input type="date" id="startDate" />
          </div>
          <div class="form-group">
            <label for="startTime">Time</label>
            <input type="time" id="startTime" />
          </div>
          <div class="form-group">
            <label for="duration">Duration (minutes)</label>
            <input type="number" id="duration" />
          </div>
          <div class="form-group">
            <label for="timezone">Timezone <button id="refreshTimezoneBtn" class="small-btn">?</button></label>
"""

html_end = """          </div>
          
          <!-- RECURRING FIELDS -->
          <div id="recurringFields" style="display: none; padding: 10px; border-left: 3px solid #0f6cbd; margin-left: 5px;">
            <div class="form-group">
              <label for="recurrenceFrequency">Frequency</label>
              <select id="recurrenceFrequency">
                <option value="Weekly">Weekly</option>
                <option value="Fortnightly">Fortnightly</option>
                <option value="Monthly">Monthly</option>
              </select>
            </div>
            <div class="form-group">
              <label for="recurrenceDay">Day</label>
              <select id="recurrenceDay">
                <option value="Monday">Monday</option>
                <option value="Tuesday">Tuesday</option>
                <option value="Wednesday">Wednesday</option>
                <option value="Thursday">Thursday</option>
                <option value="Friday">Friday</option>
              </select>
            </div>
            <div class="form-group">
              <label for="seriesEndDate">End Date</label>
              <input type="date" id="seriesEndDate" />
            </div>
          </div>

          <div class="form-group">
            <label for="status">Status</label>
            <select id="status">
              <option value="Scheduled">Scheduled</option>
              <option value="Rescheduled">Rescheduled</option>
              <option value="Cancelled">Cancelled</option>
              <option value="Completed">Completed</option>
            </select>
          </div>
          
          <div class="form-group">
            <label for="meetingStatus">Meeting Status</label>
            <select id="meetingStatus">
              <option value="Requested" selected>Requested</option>
              <option value="Created">Created</option>
              <option value="Rescheduled">Rescheduled</option>
              <option value="Cancelled">Cancelled</option>
              <option value="Completed">Completed</option>
            </select>
          </div>

          <!-- Add-in internal fields -->
          <div style="border: 1px dashed #ccc; padding: 5px; margin-top: 15px;">
            <div class="form-group">
              <label for="meetingSubject">Invite Subject</label>
              <input type="text" id="meetingSubject" />
            </div>
            <div class="form-group">
              <label for="meetingEventmessagecontent">Invite Message</label>
              <textarea id="meetingEventmessagecontent"></textarea>
            </div>
            <div class="form-group">
              <label for="requiredAttendees">Required Attendees</label>
              <textarea id="requiredAttendees" class="multi-email-input"></textarea>
              <div class="error-text">Some emails are invalid.</div>
            </div>
            <div class="form-group">
              <label for="optionalAttendees">Optional Attendees</label>
              <textarea id="optionalAttendees" class="multi-email-input"></textarea>
              <div class="error-text">Some emails are invalid.</div>
            </div>
          </div>
        </div>

        <!-- RESCHEDULE FIELDS -->
        <div id="rescheduleFields" style="display: none;">
          <div class="form-group" id="newDateSection">
            <label for="newDate">New Date</label>
            <input type="date" id="newDate" />
          </div>
          <div class="form-group" id="newDaySection" style="display: none;">
            <label for="newDay">New Day</label>
            <select id="newDay">
              <option value="Monday">Monday</option>
              <option value="Tuesday">Tuesday</option>
              <option value="Wednesday">Wednesday</option>
              <option value="Thursday">Thursday</option>
              <option value="Friday">Friday</option>
            </select>
          </div>
          <div class="form-group">
            <label for="newTime">New Time</label>
            <input type="time" id="newTime" />
          </div>
          <div class="form-group">
            <label for="reason">Reason (Optional)</label>
            <input type="text" id="reason" placeholder="e.g. Client requested change" />
          </div>
        </div>

        <!-- ADD PARTICIPANT FIELDS -->
        <div id="addParticipantFields" style="display: none;">
          <div class="form-group">
            <label for="newParticipantEmail">New Participant Email</label>
            <input type="text" id="newParticipantEmail" class="email-input" />
            <div class="error-text">Invalid email.</div>
          </div>
          <div class="form-group">
            <label for="newParticipantName">Name (Optional)</label>
            <input type="text" id="newParticipantName" />
          </div>
        </div>

        <!-- UPDATE TITLE FIELDS -->
        <div id="updateTitleFields" style="display: none;">
          <div class="form-group">
            <label for="newTitle">New Title</label>
            <input type="text" id="newTitle" />
          </div>
        </div>

      </div>

      <div class="button-group">
        <button id="injectAndLog" class="action-btn primary">Construct Email</button>
      </div>

      <div id="statusMessage" class="status"></div>
    </main>
    <script type="text/javascript" src="taskpane.js"></script>
  </body>
</html>
"""

html = html_start + tz_chunk + html_end

with open('src/taskpane/taskpane.html', 'w', encoding='utf-8') as f:
    f.write(html)
