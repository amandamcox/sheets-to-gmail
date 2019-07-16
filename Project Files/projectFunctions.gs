/**
 * @OnlyCurrentDoc
*/

function sendProjectToSelf() {
  setupProjectEmail('self');
}

function sendProjectToList() {
  setupProjectEmail('list');
}

function setupProjectEmail(recipient) {
  projectTemplate.logoUrl = sh.getRange('B3').getValues().toString();
  projectTemplate.projectName = sh.getRange('B4').getValues().toString();
  projectTemplate.projectLink = sh.getRange('B5').getValues().toString();
  projectTemplate.projectStatus = sh.getRange('B6').getValues().toString();
  projectTemplate.projectStage = sh.getRange('B7').getValues().toString();
  projectTemplate.targetDate = sh.getRange('B8').getValues().toString();
  projectTemplate.execSummary = sh.getRange('B9').getValues().toString();
  projectTemplate.pathToGreen = sh.getRange('B10:B14').getValues();
  projectTemplate.recentActivity = sh.getRange('B15:B19').getValues();
  projectTemplate.upcomingMiles = sh.getRange('B20:B24').getValues();
  projectTemplate.imageTitle = sh.getRange('B25').getValues().toString();
  projectTemplate.footerText = sh.getRange('B27').getValues().toString();
  projectTemplate.footerLink = sh.getRange('B28').getValues().toString();
  var imageUrl = sh.getRange('B26').getValues().toString();
  var attachmentUrl = sh.getRange('B29').getValues().toString();
  var subject = projectTemplate.projectName + " | Status Report | " + getDate();

  // Determines whether to send a test email to current user or to the email list provided
  var emailList = (recipient === 'self') ? Session.getActiveUser().getEmail() : sh.getRange('B30').getValues();

  sendEmail(projectTemplate, imageUrl, attachmentUrl, subject, emailList);
}