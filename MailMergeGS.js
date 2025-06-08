// Functions for sending mail merges

// Display HTML interface for mail merge
function showMailMergeDialog() {
  const html = HtmlService.createHtmlOutputFromFile('MailMergeModal')

    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Mail Merge');
}

// Send personalized emails based on one or more tags
function sendMailMerge(tagOrTags, subject, body) {
  try {
    const tags = Array.isArray(tagOrTags) ? tagOrTags : [tagOrTags];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leadership Directory');
    const data = sheet.getDataRange().getValues();
    let count = 0;
    const sent = new Set();
    for (let i = 1; i < data.length; i++) {
      const emailTags = data[i][8];
      const email = data[i][3];
      const status = data[i][10];
      const fullName = data[i][1];
      if (status === 'Active' && emailTags && tags.some(t => emailTags.includes(t)) && !sent.has(email)) {
        const personalizedBody = body
          .replace(/{{\s*Full Name\s*}}/g, fullName)
          .replace(/{{\s*Email\s*}}/g, email);
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: personalizedBody
        });
        sent.add(email);
        count++;
      }
    }
    return `Sent ${count} emails.`;
  } catch (error) {
    console.error('Error sending mail merge:', error);
    throw error;
  }
}

// Retrieve list of tags from the Email Tag Reference sheet
function getAllTags() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Tag Reference');
  if (!sheet) {
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  return values.map(r => r[0]).filter(String);
}

// Backwards compatible name used by older HTML files
function getAvailableTags() {
  return getAllTags();
}
