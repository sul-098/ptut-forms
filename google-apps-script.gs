// Averon Global — Student Registration
// Paste this into Extensions > Apps Script on your Google Sheet, then deploy as a Web App.
// See the setup guide for step-by-step instructions.

var PAYMENT_FOLDER_NAME = 'Averon Global - Payment Screenshots';

function getPaymentFolder_() {
  var folders = DriveApp.getFoldersByName(PAYMENT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(PAYMENT_FOLDER_NAME);
}

function saveScreenshot_(data) {
  if (!data.screenshotBase64) return '';
  var blob = Utilities.newBlob(
    Utilities.base64Decode(data.screenshotBase64),
    data.screenshotMimeType || 'application/octet-stream',
    data.screenshotFileName || 'payment-screenshot'
  );
  var file = getPaymentFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Registrations') || ss.getActiveSheet();
  var p = JSON.parse(e.postData.contents);
  var screenshotUrl = saveScreenshot_(p);

  sheet.appendRow([
    new Date(),          // Timestamp (server-side)
    p.fullName,
    p.email,
    p.mobile,
    p.university,
    p.degree,
    p.course,
    p.schedule,
    p.fee,
    p.transactionId,
    screenshotUrl,
    '',                   // Payment Verified — fill in manually after checking
    p.motivation,
    p.background,
    p.courseRequests,
    p.otherCourseRequest,
    p.referral,
    p.comments
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}
