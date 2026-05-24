const SHEET_NAME = 'Deepak Mobile Leads';

function doGet(e) {
  const sheet = getOrCreateSheet();
  const params = e && e.parameter ? e.parameter : {};

  sheet.appendRow([
    new Date(),
    params.brand || 'Deepak Mobile',
    params.name || '',
    params.phone || '',
    params.source || 'QR'
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
    sheet.appendRow(['Timestamp', 'Brand', 'Name', 'Phone', 'Source']);
    sheet.getRange('A1:E1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}
