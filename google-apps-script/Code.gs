/**
 * TourTripsDubai — Contact Form backend.
 * Deploy this bound to a Google Sheet (Extensions > Apps Script), as a Web App.
 * See README.md in this folder for step-by-step setup.
 */

const SHEET_NAME = 'Enquiries';
const ADMIN_EMAIL = 'info@tourtripsdubai.com';
const SITE_NAME = 'TourTripsDubai';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    appendToSheet(data);
    notifyAdmin(data);
    notifyCustomer(data);
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Page', 'Full Name', 'Email Address', 'WhatsApp Number',
      'Nationality', 'Travel Date', 'Package / Service', 'Number of Adults',
      'Number of Children', 'Special Requests / Message']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findValue(data, keys) {
  for (const k of keys) {
    if (data[k]) return data[k];
  }
  return '';
}

function appendToSheet(data) {
  const sheet = getSheet();
  sheet.appendRow([
    new Date(),
    data['Page'] || '',
    findValue(data, ['Full Name', 'Full Name (As Per Passport)']),
    findValue(data, ['Email Address', 'Email']),
    findValue(data, ['WhatsApp Number']),
    findValue(data, ['Nationality']),
    findValue(data, ['Travel Date']),
    findValue(data, ['Package / Service']),
    findValue(data, ['Number of Adults']),
    findValue(data, ['Number of Children']),
    findValue(data, ['Special Requests', 'Special Requests / Message', 'Message'])
  ]);
}

function formatSummary(data) {
  const skip = ['Page', 'Submitted At'];
  return Object.keys(data)
    .filter(k => !skip.includes(k) && data[k])
    .map(k => k + ': ' + data[k])
    .join('\n');
}

function notifyAdmin(data) {
  const name = findValue(data, ['Full Name', 'Full Name (As Per Passport)']) || 'Website Visitor';
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: 'New Website Enquiry — ' + name,
    body: 'You have a new enquiry from ' + SITE_NAME + ':\n\n' + formatSummary(data)
  });
}

function notifyCustomer(data) {
  const email = findValue(data, ['Email Address', 'Email']);
  if (!email) return;
  const name = findValue(data, ['Full Name', 'Full Name (As Per Passport)']);
  MailApp.sendEmail({
    to: email,
    subject: 'We received your enquiry — ' + SITE_NAME,
    body: 'Hi ' + (name || 'there') + ',\n\n' +
      'Thank you for reaching out to ' + SITE_NAME + '. We have received your enquiry ' +
      'and one of our travel experts will get back to you shortly.\n\n' +
      'Here is a summary of what you submitted:\n\n' + formatSummary(data) + '\n\n' +
      'Warm regards,\n' + SITE_NAME + ' Team'
  });
}
