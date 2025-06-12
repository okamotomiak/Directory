// Functions for sending mail merges

/**
 * Open the Mail Merge modal dialog in the spreadsheet UI.
 */
function showMailMergeDialog() {
  const html = HtmlService.createHtmlOutputFromFile('MailMergeModal')

    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Mail Merge');
}

/**
 * Send personalized emails to leaders matching the provided tag or tags.
 *
 * @param {string|string[]} tagOrTags Single tag or array of tags.
 * @param {string} subject Email subject line.
 * @param {string} body HTML body with {{Full Name}} and {{Email}} variables.
 * @return {string} Status message summarizing how many emails were sent.
 */
function sendMailMerge(tagOrTags, subject, body) {
  try {
    const tags = Array.isArray(tagOrTags) ? tagOrTags : [tagOrTags];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leadership Directory');
    const data = sheet.getDataRange().getValues();
    let count = 0;
    const sent = new Set();
    data.slice(1).forEach(row => {
      const emailTags = row[8];
      const email = row[3];
      const status = row[10];
      const fullName = row[1];
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
    });
    return `Sent ${count} emails.`;
  } catch (error) {
    console.error('Error sending mail merge:', error);
    throw error;
  }
}

/**
 * Retrieve list of tags from the "Email Tag Reference" sheet.
 *
 * @return {string[]} Array of tag strings.
 */
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

/**
 * Backwards compatible name used by older HTML files.
 *
 * @return {string[]} Array of tag strings.
 */
function getAvailableTags() {
  return getAllTags();
}

/**
 * Get or create the sheet used to store mail merge templates.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Sheet storing templates.
 */
function getTemplateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Mail Templates');
  if (!sheet) {
    sheet = ss.insertSheet('Mail Templates');
    const headers = ['Template Name', 'Subject', 'Body'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4a90e2')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 200);
  }
  return sheet;
}

/**
 * Save or update a template in the Mail Templates sheet.
 *
 * @param {string} name Template name.
 * @param {string} subject Email subject.
 * @param {string} body Email body HTML.
 */
function saveMailTemplate(name, subject, body) {
  const sheet = getTemplateSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[subject, body]]);
      return;
    }
  }
  sheet.appendRow([name, subject, body]);
}

/**
 * Retrieve all saved mail merge templates.
 *
 * @return {{name:string,subject:string,body:string}[]} Templates.
 */
function listMailTemplates() {
  const sheet = getTemplateSheet();
  const data = sheet.getDataRange().getValues();
  const templates = [];
  data.slice(1).forEach(row => {
    const [name, subject, body] = row;
    if (name) templates.push({ name, subject, body });
  });
  return templates;
}

/**
 * Retrieve a single template by name.
 *
 * @param {string} name Template name.
 * @return {{name:string,subject:string,body:string}|null} Template data or null.
 */
function getMailTemplate(name) {
  const sheet = getTemplateSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      return { name: data[i][0], subject: data[i][1], body: data[i][2] };
    }
  }
  return null;
}
