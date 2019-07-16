// Creates a menu in Google Sheets to send a test email and/or the final email
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Send Email')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Project Email')
         .addItem('Send Test Email', 'sendProjectToSelf')
         .addItem('Send Email To List Provided', 'sendProjectToList'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Program Email')
         .addItem('Send Test Email', 'sendProgramToSelf')
         .addItem('Send Email To List Provided', 'sendProgramToList'))
      .addToUi();
}

// Updates Program sheet when the number of projects is edited
function onEdit(e){
  if (e.range.getA1Notation() === projectCountCell) {
    var projectCountValue = sh.getRange(projectCountCell).getValue();
    sh.getRange('A' + projectStartCell + ':B200').clear().clearNote().setDataValidation(null);

    // Copies the project template range as many times as specified in projectCountValue
    for (i = 0; i < projectCountValue; i++) {
      projectTemplateRange.copyTo(sh.getRange('A' + projectStartCell), SpreadsheetApp.CopyPasteType.PASTE_NORMAL);
      projectStartCell = projectStartCell + (projectTemplateRows + 1);
  }
 }
}
