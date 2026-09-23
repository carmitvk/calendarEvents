const SPREADSHEET_ID = '1V3wMw6tGR1XgiHqOov-nPGG-6CXL-e8o3qvA4uxTX5M';
const WRITE_TOKEN = 'REPLACE_WITH_A_RANDOM_SECRET';

function fillMissingGregorianDates() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headerRow = values.findIndex(row => row.some(cell => String(cell).includes('תאריך לועזי')));
  if (headerRow < 0) throw new Error('Header row not found');

  const header = values[headerRow].map(String);
  const gregorianColumn = header.findIndex(value => value.includes('תאריך לועזי'));
  const hebrewColumn = header.findIndex(value => value.includes('תאריך עברי'));
  if (gregorianColumn < 0 || hebrewColumn < 0) throw new Error('Date columns not found');

  let filled = 0;
  for (let row = headerRow + 1; row < values.length; row += 1) {
    const hebrewValue = values[row][hebrewColumn];
    if (!(hebrewValue instanceof Date)) continue;
    sheet.getRange(row + 1, gregorianColumn + 1)
      .setNumberFormat('@')
      .setValue(formatDateText(hebrewValue));
    filled += 1;
  }
  if (values.length > headerRow + 1) {
    sheet.getRange(headerRow + 2, gregorianColumn + 1, values.length - headerRow - 1, 1)
      .setNumberFormat('@');
  }
  Logger.log('Filled Gregorian dates: ' + filled);
}

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
    const hebrewColumn = header.findIndex(value => value.includes('תאריך עברי'));
    const titleColumn = header.findIndex(value => value.includes('שם החוגגת'));
    const classColumn = header.findIndex(value => value.includes('כיתה'));
    let ownerColumn = header.findIndex(value => value.includes('תז') || value.includes('ת.ז') || value.includes('ת״ז'));
    if (dateColumn < 0 || titleColumn < 0 || classColumn < 0) return json({ ok: false, error: 'Required columns not found' });
    if (ownerColumn < 0) {
      ownerColumn = header.length;
      sheet.getRange(headerRow + 1, ownerColumn + 1).setValue('ת״ז');
    }

    if (payload.action === 'upsert') {
      const event = payload.event || {};
      const role = payload.role || 'user';
      const ownerId = String(payload.userId || event.ownerId || '').trim();
      const sharedClasses = ['שכבתי', 'בית ספרי', 'בנות השכבה'];
      if (!ownerId) return json({ ok: false, error: 'Owner ID is required' });
      if (role === 'user' && sharedClasses.includes(String(event.className || ''))) {
        return json({ ok: false, error: 'Only managers can create shared events' });
      }
      if (role === 'user') {
        const alreadyCreated = values.slice(headerRow + 1).some(item => String(item[ownerColumn] || '').trim() === ownerId && String(item[titleColumn] || '').trim());
        if (alreadyCreated) return json({ ok: false, error: 'A regular user may create only one event' });
      }
      const row = values.slice(headerRow + 1).findIndex(item =>
        toDateKey(item[dateColumn]) === event.date ||
        (hebrewColumn >= 0 && toDateKey(item[hebrewColumn]) === event.date)
      );
      const rowNumber = row >= 0 ? headerRow + 2 + row : sheet.getLastRow() + 1;
      sheet.getRange(rowNumber, dateColumn + 1)
        .setNumberFormat('@')
        .setValue(formatDateText(event.date))
        .setNumberFormat('@');
      if (row < 0) {
        if (hebrewColumn >= 0) {
          sheet.getRange(rowNumber - 1, hebrewColumn + 1)
            .copyTo(sheet.getRange(rowNumber, hebrewColumn + 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          sheet.getRange(rowNumber, hebrewColumn + 1).setValue(new Date(event.date + 'T12:00:00'));
        }
      }
      sheet.getRange(rowNumber, titleColumn + 1).setValue(event.title || '');
      sheet.getRange(rowNumber, classColumn + 1).setValue(event.className || '');
      sheet.getRange(rowNumber, ownerColumn + 1).setNumberFormat('@').setValue(ownerId);
      return json({ ok: true });
    }
    if (payload.action === 'delete') {
      const role = payload.role || 'user';
      const ownerId = String(payload.userId || '').trim();
      const row = values.slice(headerRow + 1).findIndex(item => toDateKey(item[dateColumn]) === payload.date);
      if (row >= 0) {
        const existingOwnerId = String(values[headerRow + 1 + row][ownerColumn] || '').trim();
        if (role !== 'super_user' && existingOwnerId !== ownerId) {
          return json({ ok: false, error: 'You can delete only your own event' });
        }
        sheet.getRange(headerRow + 2 + row, titleColumn + 1).clearContent();
        sheet.getRange(headerRow + 2 + row, classColumn + 1).clearContent();
        sheet.getRange(headerRow + 2 + row, ownerColumn + 1).clearContent();
      }
      return json({ ok: true });
    }
    return json({ ok: false, error: 'Unknown action' });
  } catch (error) {
    return json({ ok: false, error: String(error) });
  }
}

function toDateKey(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const text = String(value || '').trim();
  const match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  return match
    ? `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`
    : text.slice(0, 10);
}

function formatDateText(value) {
  if (typeof value === 'string') {
    const iso = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
  }
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) throw new Error('Invalid Gregorian date');
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${day}/${month}/${year}`;
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
