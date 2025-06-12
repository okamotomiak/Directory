// Functions for creating a leader signup form and importing responses

/**
 * Create a Google Form for new leaders to submit their information.
 * The form responses will be stored in the active spreadsheet under
 * a sheet named 'New Leader Responses'.
 *
 * @return {string} The public URL of the created form.
 */
function createLeaderForm() {
  // Remove any existing submit triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'handleLeaderFormSubmit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  const form = FormApp.create('Leadership Directory Entry Form');

  form.addTextItem().setTitle('Full Name').setRequired(true);
  form.addTextItem().setTitle('Title').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addTextItem().setTitle('Phone');
  form.addTextItem().setTitle('Location');
  form.addParagraphTextItem().setTitle('Bio');
  form.addTextItem().setTitle('Photo URL');
  form.addTextItem().setTitle('Email Tags');
  form.addMultipleChoiceItem()
      .setTitle('Website Display')
      .setChoices([
        form.createChoice('TRUE', true),
        form.createChoice('FALSE', false)
      ])
      .setRequired(true);
  form.addMultipleChoiceItem()
      .setTitle('Status')
      .setChoices([
        form.createChoice('Active', true),
        form.createChoice('Inactive', false),
        form.createChoice('Transitioning', false)
      ])
      .setRequired(true);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Automatically process new submissions
  ScriptApp.newTrigger('handleLeaderFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();

  // Rename the default response sheet and add a Processed column
  const responseSheet = ss.getSheetByName('Form Responses 1');
  if (responseSheet) {
    responseSheet.setName('New Leader Responses');
    const lastCol = responseSheet.getLastColumn();
    responseSheet.insertColumnAfter(lastCol);
    responseSheet.getRange(1, lastCol + 1).setValue('Processed');
  }

  console.log('Leader form created: ' + form.getPublishedUrl());
  return form.getPublishedUrl();
}

/**
 * Import unprocessed form responses into the Leadership Directory sheet and
 * mark them as processed.
 */
function importLeaderFormResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName('New Leader Responses');
  if (!responseSheet) {
    console.log('No response sheet found.');
    return;
  }

  const data = responseSheet.getDataRange().getValues();
  const processedCol = responseSheet.getLastColumn();
  data.slice(1).forEach((row, idx) => {
    if (row[processedCol - 1]) return; // already processed

    const [timestamp, name, title, email, phone, location, bio,
           photoUrl, emailTags, websiteDisplay, status] = row;

    addNewLeader(name, title, email, phone, location, bio, photoUrl,
                 emailTags, websiteDisplay === 'TRUE', status);

    responseSheet.getRange(idx + 2, processedCol).setValue('DONE');
  });
}

/**
 * Trigger handler to process form submissions as they occur.
 *
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e Form submit event.
 */
function handleLeaderFormSubmit(e) {
  importLeaderFormResponses();
}
