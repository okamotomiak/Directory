/**
 * FWWPU Northeast Leadership Directory Creator
 * Creates formatted sheets in the active Google Spreadsheet for leadership contact
 * info and email automation
*/

// Add a custom menu to the spreadsheet when it is opened
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Leadership Tools')
    .addItem('Create Directory', 'createLeadershipDirectory')
    .addSeparator()
    .addItem('Send Mail Merge…', 'showMailMergeDialog')
    .addSeparator()
    .addItem('Meeting Reminder Setup…', 'showReminderSidebar')
    .addItem('Send Meeting Reminders Now', 'sendMeetingReminders')
    .addToUi();
}

function createLeadershipDirectory() {
  try {
    // Use the active spreadsheet instead of creating a new one
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Insert a new sheet for the directory
    const sheet = spreadsheet.insertSheet();
    sheet.setName('Leadership Directory');
    
    // Set up headers
    const headers = [
      'ID', 'Full Name', 'Title', 'Email', 'Phone', 'Location', 
      'Bio', 'Photo URL', 'Email Tags', 'Website Display', 'Status', 'Last Updated'
    ];
    
    // Add headers to row 1
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // Format headers
    headerRange.setBackground('#4a90e2')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setFontSize(12)
             .setHorizontalAlignment('center');
    
    // Set column widths
    const columnWidths = [50, 150, 180, 200, 120, 120, 300, 150, 250, 100, 100, 120];
    columnWidths.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width);
    });
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Add sample data
    const sampleData = [
      [
        1, 
        'Rev. Michael Johnson', 
        'Regional Director', 
        'm.johnson@familyfed.org', 
        '(617) 555-0123', 
        'Boston, MA',
        'Rev. Johnson has served the Northeast region for 15 years, focusing on family ministry and community building. He leads our regional initiatives and coordinates with state leaders.',
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
        'Mrs. Chen oversees educational programs across the region with 12 years of experience. She specializes in youth development and character education initiatives.',
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
        'Rev. Rodriguez leads Massachusetts communities with focus on family unity and social service. He coordinates local centers and statewide programs.',
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
        'Mrs. Kim focuses on strengthening family bonds and organizing interfaith dialogue events throughout Connecticut.',
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
        'Mrs. Santos coordinates youth programs and summer camps across the region. She\'s passionate about empowering young people to become future leaders.',
        'https://drive.google.com/file/d/sample5',
        'youth-leader,ministry-leader,connecticut,training-updates,event-coordination',
        true,
        'Active',
        new Date()
      ]
    ];
    
    // Add sample data starting from row 2
    const dataRange = sheet.getRange(2, 1, sampleData.length, headers.length);
    dataRange.setValues(sampleData);
    
    // Format data rows
    dataRange.setVerticalAlignment('top')
             .setWrap(true);
    
    // Format Bio column (G) for text wrapping
    sheet.getRange(2, 7, sampleData.length, 1).setWrap(true);
    
    // Format Email Tags column (I) for text wrapping
    sheet.getRange(2, 9, sampleData.length, 1).setWrap(true);
    
    // Set row heights for better readability
    sheet.setRowHeights(2, sampleData.length, 80);
    
    // Add data validation for Status column (K)
    const statusRange = sheet.getRange(2, 11, 1000, 1); // Apply to many rows for future entries
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Inactive', 'Transitioning'])
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusRule);
    
    // Add data validation for Website Display column (J) - boolean
    const displayRange = sheet.getRange(2, 10, 1000, 1);
    const displayRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TRUE', 'FALSE'])
      .setAllowInvalid(false)
      .build();
    displayRange.setDataValidation(displayRule);
    
    // Apply alternating row colors
    const allDataRange = sheet.getRange(2, 1, sampleData.length, headers.length);
    allDataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    
    // Add borders around all data
    const fullRange = sheet.getRange(1, 1, sampleData.length + 1, headers.length);
    fullRange.setBorder(true, true, true, true, true, true);
    
    // Center align specific columns
    sheet.getRange(2, 1, 1000, 1).setHorizontalAlignment('center'); // ID column
    sheet.getRange(2, 10, 1000, 1).setHorizontalAlignment('center'); // Website Display
    sheet.getRange(2, 11, 1000, 1).setHorizontalAlignment('center'); // Status
    sheet.getRange(2, 12, 1000, 1).setHorizontalAlignment('center'); // Last Updated
    
    // Format email column
    sheet.getRange(2, 4, 1000, 1).setFontColor('#1155cc'); // Make emails blue
    
    // Add conditional formatting for Status column
    const activeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Active')
      .setBackground('#d9ead3')
      .setRanges([sheet.getRange(2, 11, 1000, 1)])
      .build();
    
    const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Inactive')
      .setBackground('#f4cccc')
      .setRanges([sheet.getRange(2, 11, 1000, 1)])
      .build();
    
    const transitioningRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Transitioning')
      .setBackground('#fce5cd')
      .setRanges([sheet.getRange(2, 11, 1000, 1)])
      .build();
    
    sheet.setConditionalFormatRules([activeRule, inactiveRule, transitioningRule]);
    
    // Create a second sheet with Email Tag Reference
    const tagSheet = spreadsheet.insertSheet('Email Tag Reference');
    
    // Add tag reference data
    const tagHeaders = ['Tag Category', 'Tag', 'Description', 'Example Usage'];
    tagSheet.getRange(1, 1, 1, tagHeaders.length).setValues([tagHeaders]);
    
    const tagData = [
      ['Leadership Level', 'regional-director', 'Regional leadership positions', 'Monthly board meetings'],
      ['Leadership Level', 'state-leader', 'State-level leadership', 'State coordination calls'],
      ['Leadership Level', 'local-leader', 'Local center leadership', 'Local event notifications'],
      ['Leadership Level', 'assistant-leader', 'Assistant/deputy positions', 'Training updates'],
      ['Ministry Type', 'pastor', 'All pastoral staff', 'Pastoral care updates'],
      ['Ministry Type', 'youth-leader', 'Youth ministry leaders', 'Youth program updates'],
      ['Ministry Type', 'womens-ministry', 'Women\'s ministry leaders', 'Women\'s event coordination'],
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
    
    tagSheet.getRange(2, 1, tagData.length, tagHeaders.length).setValues(tagData);
    
    // Format tag reference sheet
    tagSheet.getRange(1, 1, 1, tagHeaders.length)
            .setBackground('#4a90e2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontSize(12)
            .setHorizontalAlignment('center');
    
    tagSheet.setColumnWidths(1, tagHeaders.length, 200);
    tagSheet.setFrozenRows(1);
    
    // Apply banding to tag sheet
    tagSheet.getRange(2, 1, tagData.length, tagHeaders.length)
            .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    
    // Add borders to tag sheet
    tagSheet.getRange(1, 1, tagData.length + 1, tagHeaders.length)
            .setBorder(true, true, true, true, true, true);
    
    // Create instructions sheet
    const instructionsSheet = spreadsheet.insertSheet('Instructions');
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
    
    // Log success
    console.log('Leadership Directory created successfully in the active spreadsheet');

    // Return the current spreadsheet URL for reference
    return spreadsheet.getUrl();
    
  } catch (error) {
    console.error('Error creating leadership directory:', error);
    throw error;
  }
}

/**
 * Helper function to add a new leader
 * Call this function with leader details to add them to the sheet
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
    
    sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    
    console.log('New leader added: ' + name);
    
  } catch (error) {
    console.error('Error adding new leader:', error);
    throw error;
  }
}

/**
 * Helper function to get email list by tag
 * Returns array of email addresses for leaders with specified tag
 */
function getEmailsByTag(tag) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leadership Directory');
    const data = sheet.getDataRange().getValues();
    const emails = [];
    
    // Skip header row, start from row 1 (index 1)
    for (let i = 1; i < data.length; i++) {
      const emailTags = data[i][8]; // Email Tags column (I)
      const email = data[i][3]; // Email column (D)
      const status = data[i][10]; // Status column (K)
      
      // Check if leader is active and has the specified tag
      if (status === 'Active' && emailTags && emailTags.includes(tag)) {
        emails.push(email);
      }
    }
    
    console.log(`Found ${emails.length} emails for tag: ${tag}`);
    return emails;
    
  } catch (error) {
    console.error('Error getting emails by tag:', error);
    throw error;
  }
}
// Display HTML interface for mail merge
function showMailMergeDialog() {
  const html = HtmlService.createHtmlOutputFromFile('MailMerge')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Mail Merge');
}

// Send personalized emails based on tag
function sendMailMerge(tag, subject, body) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leadership Directory');
    const data = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const emailTags = data[i][8];
      const email = data[i][3];
      const status = data[i][10];
      const fullName = data[i][1];
      if (status === 'Active' && emailTags && emailTags.includes(tag)) {
        const personalizedBody = body.replace(/{{\s*Full Name\s*}}/g, fullName);
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: personalizedBody
        });
        count++;
      }
    }
    return `Sent ${count} emails.`;
  } catch (error) {
    console.error('Error sending mail merge:', error);
    throw error;
  }
}

// Retrieve available tags for the mail merge dropdown
function getAvailableTags() {
  try {
    const tagSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Tag Reference');
    if (!tagSheet) {
      return [];
    }
    const lastRow = tagSheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    const tags = tagSheet.getRange(2, 2, lastRow - 1, 1).getValues()
      .flat()
      .filter(String);
    // Remove duplicates
    return Array.from(new Set(tags));
  } catch (error) {
    console.error('Error getting available tags:', error);
    throw error;
  }
}

