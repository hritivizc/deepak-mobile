const SHEET_NAME = 'Deepak Mobile Leads';

function doGet(e) {
  const sheet = getOrCreateSheet();
  const params = e && e.parameter ? e.parameter : {};
  const phone = params.phone || params.mobile || params.mobileNumber || params.clientPhone || '';

  sheet.appendRow([
    new Date(),
    params.brand || 'Deepak Mobile',
    params.name || '',
    phone,
    params.source || 'QR',
    e && e.queryString ? e.queryString : ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Brand', 'Name', 'Phone', 'Source', 'Raw Query']);
    sheet.getRange('A1:F1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}
