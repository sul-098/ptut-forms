// Averon Global — Student Registration
// Paste this into Extensions > Apps Script on your Google Sheet, then deploy as a Web App.
// See the setup guide for step-by-step instructions.

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Registrations') || ss.getActiveSheet();
  var p = e.parameter;

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
