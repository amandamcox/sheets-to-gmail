// Global Variables
var sh = SpreadsheetApp.getActiveSpreadsheet();
var projectTemplate = HtmlService.createTemplateFromFile('ProjectTemplate');
var programTemplate = HtmlService.createTemplateFromFile('ProgramTemplate');
var projectCountCell = 'E3';
var programLastRow = 21;
var projectStartCell = programLastRow + 2;
var projectTemplateRange = sh.getRange("G1:H19");
var projectTemplateRows = projectTemplateRange.getLastRow();

function sendEmail(templateType, imageUrl, attachmentUrl, subject, emailList) {
  var template = templateType.evaluate().getContent();

  // Gets Blobs for image and attachment if they exist
  if (imageUrl) {
   var imageBlob = getDriveId(imageUrl);
  }
  if (attachmentUrl) {
   var attachBlob = getDriveId(attachmentUrl);
  }

  // Sends email to current user with image/attachments, if applicable
  if (imageBlob && attachBlob) {
    MailApp.sendEmail(emailList, subject, '', {
      htmlBody: template,
      attachments: attachBlob,
      inlineImages: {userimage: imageBlob}
    });
  }
  else if (!imageBlob && attachBlob) {
    MailApp.sendEmail(emailList, subject, '', {
      htmlBody: template,
      attachments: attachBlob
    });
  }
  else if (imageBlob && !attachBlob) {
    MailApp.sendEmail(emailList, subject, '', {
      htmlBody: template,
      inlineImages: {userimage: imageBlob}
    });
  }
  else {
    MailApp.sendEmail(emailList, subject, '', {
      htmlBody: template
    });
  }
}

function getDate() {
  var currentDate = new Date();
  var day = currentDate.getDate();
  var month = currentDate.getMonth();
  var year = currentDate.getFullYear();
  var fullmonths = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  var dateString = fullmonths[month] + " " + day + ", " + year;
  return dateString;
}

function getDriveId(url) {
  var index = url.indexOf('=') + 1;
  var fileId = url.substring(index);
  var blob = DriveApp.getFileById(fileId).getBlob();
  return blob;
}