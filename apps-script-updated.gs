/**
 * YANSII — Google Apps Script backend
 * Handles two form types from the website:
 *   1. Waitlist signups  → "Waitlist" sheet tab  (email, timestamp, source)
 *   2. Feedback submissions → "Feedback" sheet tab (type, message, name, email, timestamp, source)
 *
 * SETUP:
 * 1. In your Google Sheet, make sure you have two tabs named exactly:
 *      "Waitlist"  — headers in row 1: Email | Timestamp | Source
 *      "Feedback"  — headers in row 1: Type | Message | Name | Email | Timestamp | Source
 *    (Create the "Feedback" tab if it doesn't exist yet — right-click any
 *    tab at the bottom → Duplicate, or use the + button, then rename it.)
 *
 * 2. Extensions → Apps Script → replace the existing doPost with this one.
 *
 * 3. Deploy → Manage deployments → edit (pencil icon) → New version → Deploy.
 *    (Redeploying keeps the same URL, so you do NOT need to update the
 *    fetch() URL in your HTML files.)
 */

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (data.formType === 'feedback') {
    var sheet = ss.getSheetByName('Feedback') || ss.insertSheet('Feedback');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Type', 'Message', 'Name', 'Email', 'Timestamp', 'Source']);
    }
    sheet.appendRow([
      data.feedbackType || 'general',
      data.message || '',
      data.name || '',
      data.email || '',
      new Date().toISOString(),
      data.source || 'yansii.in/feedback'
    ]);
  } else {
    // Default / existing behavior: waitlist signup
    var sheet = ss.getSheetByName('Waitlist') || ss.getActiveSheet();
    sheet.appendRow([
      data.email || '',
      new Date().toISOString(),
      data.source || 'website'
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
