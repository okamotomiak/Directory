/**
 * FWWPU Northeast Leadership Directory Creator
 * Creates formatted sheets in the active Google Spreadsheet for leadership contact
 * info and email automation
*/
// Color constants
const COLOR_HEADER_BG = "#4a90e2";
const COLOR_HEADER_TEXT = "#ffffff";
const COLOR_STATUS_ACTIVE = "#d9ead3";
const COLOR_STATUS_INACTIVE = "#f4cccc";
const COLOR_STATUS_TRANSITION = "#fce5cd";
const COLOR_LINK = "#1155cc";

// Column index constants
const COL_ID = 1;
const COL_FULL_NAME = 2;
const COL_TITLE = 3;
const COL_EMAIL = 4;
const COL_PHONE = 5;
const COL_LOCATION = 6;
const COL_BIO = 7;
const COL_PHOTO_URL = 8;
const COL_EMAIL_TAGS = 9;
const COL_WEBSITE_DISPLAY = 10;
const COL_STATUS = 11;
const COL_LAST_UPDATED = 12;

/**
 * Add the custom "Leadership Tools" menu to the spreadsheet UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Leadership Tools')
    .addItem('Create Directory', 'createLeadershipDirectory')
    .addSeparator()
    .addItem('Create Leader Signup Form', 'createLeaderForm')
    .addItem('Import New Leader Responses', 'importLeaderFormResponses')
    .addSeparator()
    .addItem('Send Mail Merge…', 'showMailMergeDialog')
    .addSeparator()
    .addItem('Meeting Reminder Setup…', 'showReminderDialog')
    .addItem('Send Meeting Reminders Now', 'sendMeetingReminders')
    .addToUi();
}
/**
 * Automatically assign an ID when a new row is added to the
 * Leadership Directory sheet.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Leadership Directory') return;
  const row = e.range.getRow();
  if (row <= 1) return;
  const idCell = sheet.getRange(row, 1);
  if (idCell.getValue()) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  let max = 0;
  data.forEach(r => {
    const v = r[0];
    if (typeof v === 'number' && v > max) max = v;
  });
  idCell.setValue(max + 1);
}


/**
 * Generate the Leadership Directory and related reference sheets in the
 * active spreadsheet.
 *
 * @return {string} URL of the spreadsheet containing the directory.
 */
function createLeadershipDirectory() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.insertSheet();
    sheet.setName('Leadership Directory');

    setupDirectoryHeaders(sheet);
    insertSampleLeaders(sheet);
    applyDirectoryFormatting(sheet);
    createTagReferenceSheet(spreadsheet);
    createInstructionsSheet(spreadsheet);

    console.log('Leadership Directory created successfully in the active spreadsheet');
    return spreadsheet.getUrl();
  } catch (error) {
    console.error('Error creating leadership directory:', error);
    throw error;
  }
}

/**
 * Create and format the header row in the directory sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to format.
 */
function setupDirectoryHeaders(sheet) {
  const headers = [
    'ID', 'Full Name', 'Title', 'Email', 'Phone', 'Location',
    'Bio', 'Photo URL', 'Email Tags', 'Website Display', 'Status', 'Last Updated'
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground(COLOR_HEADER_BG)
    .setFontColor(COLOR_HEADER_TEXT)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');

  const columnWidths = [50, 150, 180, 200, 120, 120, 300, 150, 250, 100, 100, 120];
  columnWidths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  sheet.setFrozenRows(1);
}

/**
 * Populate the directory with sample leader entries.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to receive data.
 */
function insertSampleLeaders(sheet) {
  const sampleData = [
    [
      1,
      'Rev. Michael Johnson',
      'Regional Director',
      'm.johnson@familyfed.org',
      '(617) 555-0123',
      'Boston, MA',
      `Rev. Johnson has served the Northeast region for 15 years, focusing on family ministry and community building. He leads our regional initiatives and coordinates with state leaders.`,
      'https://drive.google.com/file/d/sample1',
      'regional-director,pastor,massachusetts,board-member,emergency-contact',
      true,
      'Active',
      new Date()
    ],
    [
      2,
      'Mrs. Sarah Chen',
      'Education Director',
      's.chen@familyfed.org',
      '(212) 555-0456',
      'New York, NY',
      `Mrs. Chen oversees educational programs across the region with 12 years of experience. She specializes in youth development and character education initiatives.`,
      'https://drive.google.com/file/d/sample2',
      'regional-director,education-director,youth-leader,new-york,monthly-reports',
      true,
      'Active',
      new Date()
    ],
    [
      3,
      'Rev. David Rodriguez',
      'Massachusetts State Leader',
      'd.rodriguez@familyfed.org',
      '(413) 555-0789',
      'Springfield, MA',
      `Rev. Rodriguez leads Massachusetts communities with focus on family unity and social service. He coordinates local centers and statewide programs.`,
      'https://drive.google.com/file/d/sample3',
      'state-leader,pastor,massachusetts,community-outreach,monthly-reports',
      true,
      'Active',
      new Date()
    ],
    [
      4,
      'Mrs. Jennifer Kim',
      'Connecticut State Leader',
      'j.kim@familyfed.org',
      '(860) 555-0321',
      'Hartford, CT',
      `Mrs. Kim focuses on strengthening family bonds and organizing interfaith dialogue events throughout Connecticut.`,
      'https://drive.google.com/file/d/sample4',
      'state-leader,connecticut,interfaith-relations,family-ministry',
      true,
      'Active',
      new Date()
    ],
    [
      5,
      'Mrs. Maria Santos',
      'Youth Ministry Leader',
      'm.santos@familyfed.org',
      '(203) 555-0987',
      'New Haven, CT',
      `Mrs. Santos coordinates youth programs and summer camps across the region. She's passionate about empowering young people to become future leaders.`,
      'https://drive.google.com/file/d/sample5',
      'youth-leader,ministry-leader,connecticut,training-updates,event-coordination',
      true,
      'Active',
      new Date()
    ]
  ];
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData)
           .setVerticalAlignment('top')
           .setWrap(true);
  sheet.getRange(2, COL_BIO, sampleData.length, 1).setWrap(true);
  sheet.getRange(2, COL_EMAIL_TAGS, sampleData.length, 1).setWrap(true);
  sheet.setRowHeights(2, sampleData.length, 80);
}

/**
 * Apply data validation, conditional formatting and layout tweaks to
 * the directory sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to format.
 */
function applyDirectoryFormatting(sheet) {
  const statusRange = sheet.getRange(2, COL_STATUS, 1000, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'Transitioning'])
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(statusRule);

  const displayRange = sheet.getRange(2, COL_WEBSITE_DISPLAY, 1000, 1);
  const displayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'])
    .setAllowInvalid(false)
    .build();
  displayRange.setDataValidation(displayRule);

  const allDataRange = sheet.getRange(2, COL_ID, sheet.getLastRow() - 1, COL_LAST_UPDATED);
  allDataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.getRange(1, 1, sheet.getLastRow(), COL_LAST_UPDATED).setBorder(true, true, true, true, true, true);

  sheet.getRange(2, COL_ID, 1000, 1).setHorizontalAlignment('center');
  sheet.getRange(2, COL_WEBSITE_DISPLAY, 1000, 1).setHorizontalAlignment('center');
  sheet.getRange(2, COL_STATUS, 1000, 1).setHorizontalAlignment('center');
  sheet.getRange(2, COL_LAST_UPDATED, 1000, 1).setHorizontalAlignment('center');
  sheet.getRange(2, COL_EMAIL, 1000, 1).setFontColor(COLOR_LINK);

  const activeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Active')
    .setBackground(COLOR_STATUS_ACTIVE)
    .setRanges([sheet.getRange(2, COL_STATUS, 1000, 1)])
    .build();
  const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Inactive')
    .setBackground(COLOR_STATUS_INACTIVE)
    .setRanges([sheet.getRange(2, COL_STATUS, 1000, 1)])
    .build();
  const transitioningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Transitioning')
    .setBackground(COLOR_STATUS_TRANSITION)
    .setRanges([sheet.getRange(2, COL_STATUS, 1000, 1)])
    .build();
  sheet.setConditionalFormatRules([activeRule, inactiveRule, transitioningRule]);
}

/**
 * Build the "Email Tag Reference" sheet with example tags and descriptions.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 */
function createTagReferenceSheet(ss) {
  const tagSheet = ss.insertSheet('Email Tag Reference');
  const tagHeaders = ['Tag Category', 'Tag', 'Description', 'Example Usage'];
  tagSheet.getRange(1, 1, 1, tagHeaders.length).setValues([tagHeaders])
    .setBackground(COLOR_HEADER_BG)
    .setFontColor(COLOR_HEADER_TEXT)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  const tagData = [
    ['Leadership Level', 'regional-director', 'Regional leadership positions', 'Monthly board meetings'],
    ['Leadership Level', 'state-leader', 'State-level leadership', 'State coordination calls'],
    ['Leadership Level', 'local-leader', 'Local center leadership', 'Local event notifications'],
    ['Leadership Level', 'assistant-leader', 'Assistant/deputy positions', 'Training updates'],
    ['Ministry Type', 'pastor', 'All pastoral staff', 'Pastoral care updates'],
    ['Ministry Type', 'youth-leader', 'Youth ministry leaders', 'Youth program updates'],
    ['Ministry Type', 'womens-ministry', "Women's ministry leaders", "Women's event coordination"],
    ['Ministry Type', 'education-director', 'Education leadership', 'Educational program updates'],
    ['Ministry Type', 'family-ministry', 'Family ministry focus', 'Family program notifications'],
    ['Ministry Type', 'community-outreach', 'Community outreach leaders', 'Outreach opportunities'],
    ['Geographic', 'massachusetts', 'Massachusetts-based leaders', 'MA-specific updates'],
    ['Geographic', 'connecticut', 'Connecticut-based leaders', 'CT-specific updates'],
    ['Geographic', 'new-york', 'New York-based leaders', 'NY-specific updates'],
    ['Geographic', 'new-hampshire', 'New Hampshire-based leaders', 'NH-specific updates'],
    ['Geographic', 'vermont', 'Vermont-based leaders', 'VT-specific updates'],
    ['Geographic', 'rhode-island', 'Rhode Island-based leaders', 'RI-specific updates'],
    ['Geographic', 'maine', 'Maine-based leaders', 'ME-specific updates'],
    ['Communication', 'board-member', 'Board members', 'Board meeting notices'],
    ['Communication', 'monthly-reports', 'Receives monthly reports', 'Monthly statistical updates'],
    ['Communication', 'emergency-contact', 'Emergency notifications', 'Urgent communications'],
    ['Communication', 'training-updates', 'Training notifications', 'Leadership development'],
    ['Communication', 'event-coordination', 'Event coordinators', 'Event planning updates']
  ];
  tagSheet.getRange(2, 1, tagData.length, tagHeaders.length).setValues(tagData)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  tagSheet.setColumnWidths(1, tagHeaders.length, 200);
  tagSheet.setFrozenRows(1);
  tagSheet.getRange(1, 1, tagData.length + 1, tagHeaders.length)
    .setBorder(true, true, true, true, true, true);
}

/**
 * Generate a sheet with usage instructions for the directory and tags.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 */
function createInstructionsSheet(ss) {
  const instructionsSheet = ss.insertSheet('Instructions');
  const instructions = [
    ['FWWPU Northeast Leadership Directory - Instructions'],
    [''],
    ['PURPOSE:'],
    ['This sheet manages leadership contact information for both website display and email automation.'],
    [''],
    ['MAIN COLUMNS:'],
    ['• ID: Simple sequential number'],
    ['• Full Name: Name as it appears on website'],
    ['• Title: Leadership position'],
    ['• Email & Phone: Primary contact information'],
    ['• Location: City, State for website display'],
    ['• Bio: 2-3 sentence description for website'],
    ['• Photo URL: Link to headshot photo (store in Google Drive)'],
    ['• Email Tags: Comma-separated tags for targeted emails'],
    ['• Website Display: TRUE/FALSE - show on public website?'],
    ['• Status: Active/Inactive/Transitioning'],
    [''],
    ['EMAIL TAG EXAMPLES:'],
    ['• pastor,state-leader,massachusetts,monthly-reports'],
    ['• youth-leader,connecticut,training-updates,event-coordination'],
    ['• regional-director,board-member,emergency-contact'],
    [''],
    ['MAIL MERGE USAGE:'],
    ['1. Filter by tag (e.g., contains "pastor")'],
    ['2. Export filtered email addresses'],
    ['3. Use with Gmail, Mailchimp, or other mail merge tools'],
    [''],
    ['WEBSITE USAGE:'],
    ['1. Filter where Website Display = TRUE'],
    ['2. Export Name, Title, Location, Bio, Photo URL'],
    ['3. Use data to populate website leadership page'],
    [''],
    ['MAINTENANCE:'],
    ['• Update Last Updated column when making changes'],
    ['• Review contact info monthly'],
    ['• Check photo links periodically'],
    ['• Keep email tags consistent (see Email Tag Reference sheet)']
  ];
  instructionsSheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
  instructionsSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  instructionsSheet.getRange(3, 1).setFontWeight('bold');
  instructionsSheet.getRange(6, 1).setFontWeight('bold');
  instructionsSheet.getRange(18, 1).setFontWeight('bold');
  instructionsSheet.getRange(24, 1).setFontWeight('bold');
  instructionsSheet.getRange(29, 1).setFontWeight('bold');
  instructionsSheet.setColumnWidth(1, 600);
}
/**
 * Append a leader entry to the directory sheet.
 *
 * @param {string} name Leader's full name.
 * @param {string} title Leadership title.
 * @param {string} email Primary email address.
 * @param {string} phone Phone number.
 * @param {string} location City and state.
 * @param {string} bio Short biography text.
 * @param {string} photoUrl Link to a photo in Google Drive.
 * @param {string} emailTags Comma-separated email tags.
 * @param {boolean} [websiteDisplay=true] Whether to show on the website.
 * @param {string} [status='Active'] Status value.
 */
function addNewLeader(name, title, email, phone, location, bio, photoUrl, emailTags, websiteDisplay = true, status = 'Active') {
  try {
    // Open the existing spreadsheet (replace with your actual spreadsheet ID if needed)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leadership Directory');
    
    // Find the next available row
    const lastRow = sheet.getLastRow();
    const nextId = lastRow; // Since row 1 is headers, this gives us the right ID
    
    // Add the new leader data
    const newRow = [
      nextId,
      name,
      title,
      email,
      phone,
      location,
      bio,
      photoUrl,
      emailTags,
      websiteDisplay,
      status,
      new Date()
    ];
    
    sheet.getRange(lastRow + 1, COL_ID, 1, newRow.length).setValues([newRow]);
    
    console.log('New leader added: ' + name);
    
  } catch (error) {
    console.error('Error adding new leader:', error);
    throw error;
  }
}

/**
 * Retrieve email addresses for active leaders containing the given tag.
 *
 * @param {string} tag Tag to search for.
 * @return {string[]} Matching email addresses.
 */
function getEmailsByTag(tag) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leadership Directory');
    const data = sheet.getDataRange().getValues();
    const emails = [];
    data.slice(1).forEach(row => {
      const emailTags = row[COL_EMAIL_TAGS];
      const email = row[COL_EMAIL];
      const status = row[COL_STATUS];
      if (status === 'Active' && emailTags && emailTags.includes(tag)) {
        emails.push(email);
      }
    });
    
    console.log(`Found ${emails.length} emails for tag: ${tag}`);
    return emails;
    
  } catch (error) {
    console.error('Error getting emails by tag:', error);
    throw error;
  }
}

