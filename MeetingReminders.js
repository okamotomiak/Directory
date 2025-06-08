// Functions for managing meeting reminders

/**
 * Create or get the Meeting Reminders sheet with headers
 */
function getReminderSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Meeting Reminders');
  if (!sheet) {
    sheet = ss.insertSheet('Meeting Reminders');
    const headers = ['Meeting Name', 'Next Reminder', 'Recurrence', 'Recipient Tags', 'Message'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#4a90e2')
         .setFontColor('#ffffff')
         .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 150);
  }
  return sheet;
}

/**
 * Show a popup dialog for creating a reminder
 */
function showReminderDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ReminderSidebar')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Meeting Reminders');
}

/**
 * Add a new meeting reminder from the sidebar
 */
function addMeetingReminder(form) {
  const sheet = getReminderSheet();
  sheet.appendRow([
    form.meetingName,
    new Date(form.nextReminder),
    form.recurrence,
    form.recipients,
    form.message
  ]);
}

/**
 * Send reminder emails if any are due and update next reminder date
 */
function sendMeetingReminders() {
  const sheet = getReminderSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    let [name, nextReminder, recurrence, recipientTags, message] = data[i];
    if (nextReminder && now >= new Date(nextReminder)) {
      const tags = recipientTags ? recipientTags.split(',').map(t => t.trim()).filter(String) : [];
      const emailSet = new Set();
      tags.forEach(tag => {
        const emails = getEmailsByTag(tag);
        emails.forEach(e => emailSet.add(e));
      });
      const allEmails = Array.from(emailSet).join(',');
      if (allEmails) {
        GmailApp.sendEmail(allEmails, `Reminder: ${name}`, message);
      }

      // Calculate next reminder date
      const next = new Date(nextReminder);
      switch (recurrence) {
        case 'Daily':
          next.setDate(next.getDate() + 1);
          break;
        case 'Weekly':
          next.setDate(next.getDate() + 7);
          break;
        case 'Monthly':
          next.setMonth(next.getMonth() + 1);
          break;
        default:
          // no recurrence, clear next reminder to stop sending
          nextReminder = '';
          break;
      }

      // Write next reminder date back
      sheet.getRange(i + 1, 2).setValue(nextReminder ? next : '');
    }
  }
}

/**
 * Create a daily trigger for sending reminders
 */
function createReminderTrigger() {
  ScriptApp.newTrigger('sendMeetingReminders')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
}
