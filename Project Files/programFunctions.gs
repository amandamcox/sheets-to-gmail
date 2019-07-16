function sendProgramToSelf() {
  setupProgramEmail('self');
}

function sendProgramToList() {
  setupProgramEmail('list');
}

function setupProgramEmail(recipient) {
  programTemplate.logoUrl = sh.getRange('B3').getValues().toString();
  programTemplate.programName = sh.getRange('B4').getValues().toString();
  programTemplate.programLink = sh.getRange('B5').getValues().toString();
  programTemplate.execSummary = sh.getRange('B6').getValues().toString();
  programTemplate.programStage = sh.getRange('B7').getValues().toString();
  programTemplate.programStatus = sh.getRange('B8').getValues().toString();
  programTemplate.programJust = sh.getRange('B9').getValues().toString();
  programTemplate.pathToGreen = sh.getRange('B10:B14').getValues();
  programTemplate.targetDate = sh.getRange('B15').getValues().toString();
  programTemplate.imageTitle = sh.getRange('B16').getValues().toString();
  var imageUrl = sh.getRange('B17').getValues().toString();
  programTemplate.footerText = sh.getRange('B18').getValues().toString();
  programTemplate.footerLink = sh.getRange('B19').getValues().toString();
  var attachmentUrl = sh.getRange('B20').getValues().toString();
  var projectNumber = sh.getRange('E3').getValues();
  var subject = programTemplate.programName + " | Status Report | " + getDate();

  // Gets project data, adds key, and inserts into array for each
  programTemplate.projectValues = [];
  var startCell = programLastRow + 2;
  for (i=0; i < projectNumber; i++) {
    var projectObject = {};
    var projectArray = sh.getRange('B' + startCell + ':B'+ (startCell + projectTemplateRows - 1)).getValues();
    projectObject.projectName = projectArray[0].toString();
    projectObject.projectLink = projectArray[1].toString();
    projectObject.projectStatus = projectArray[2].toString();
    projectObject.statusJust = projectArray[3].toString();
    projectObject.pathToGreen = [projectArray[4].toString(), projectArray[5].toString(), projectArray[6].toString(), projectArray[7].toString(), projectArray[8].toString()];
    projectObject.projectActs = [projectArray[9].toString(), projectArray[10].toString(), projectArray[11].toString(), projectArray[12].toString(), projectArray[13].toString()];
    projectObject.projectMiles = [projectArray[14].toString(), projectArray[15].toString(), projectArray[16].toString(), projectArray[17].toString(), projectArray[18].toString()];
    programTemplate.projectValues.push(projectObject);
    startCell += projectTemplateRows + 1;
  }

  // Determines whether to send a test email to current user or to the email list provided
  var emailList = (recipient === 'self') ? Session.getActiveUser().getEmail() : sh.getRange('B21').getValues();

  sendEmail(programTemplate, imageUrl, attachmentUrl, subject, emailList);
}