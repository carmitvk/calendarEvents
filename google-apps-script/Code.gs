const SPREADSHEET_ID = '1V3wMw6tGR1XgiHqOov-nPGG-6CXL-e8o3qvA4uxTX5M';
const WRITE_TOKEN = 'REPLACE_WITH_A_RANDOM_SECRET';

function doPost(request) {
  try {
    const payload = JSON.parse(request.postData.contents || '{}');
    if (payload.token !== WRITE_TOKEN) return json({ ok: false, error: 'Unauthorized' });
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
    const values = sheet.getDataRange().getValues();
    const headerRow = values.findIndex(row => row.some(cell => String(cell).includes('תאריך לועזי')));
    if (headerRow < 0) return json({ ok: false, error: 'Header row not found' });
    const header = values[headerRow].map(String);
    const dateColumn = header.findIndex(value => value.includes('תאריך לועזי'));
    const titleColumn = header.findIndex(value => value.includes('שם החוגגת'));
    const classColumn = header.findIndex(value => value.includes('כיתה'));
    if (dateColumn < 0 || titleColumn < 0 || classColumn < 0) return json({ ok: false, error: 'Required columns not found' });

    if (payload.action === 'upsert') {
      const event = payload.event || {};
      const row = values.slice(headerRow + 1).findIndex(item => toDateKey(item[dateColumn]) === event.date);
      const rowNumber = row >= 0 ? headerRow + 2 + row : sheet.getLastRow() + 1;
      sheet.getRange(rowNumber, dateColumn + 1).setValue(new Date(event.date + 'T12:00:00'));
      sheet.getRange(rowNumber, titleColumn + 1).setValue(event.title || '');
      sheet.getRange(rowNumber, classColumn + 1).setValue(event.className || '');
      return json({ ok: true });
    }
    if (payload.action === 'delete') {
      const row = values.slice(headerRow + 1).findIndex(item => toDateKey(item[dateColumn]) === payload.date);
      if (row >= 0) {
        sheet.getRange(headerRow + 2 + row, titleColumn + 1).clearContent();
        sheet.getRange(headerRow + 2 + row, classColumn + 1).clearContent();
      }
      return json({ ok: true });
    }
    return json({ ok: false, error: 'Unknown action' });
  } catch (error) {
    return json({ ok: false, error: String(error) });
  }
}

function toDateKey(value) {
  if (!(value instanceof Date)) return String(value || '').slice(0, 10);
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
