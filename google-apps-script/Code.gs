const SHEET_NAME = 'Records';
const HEADERS = [
  'id', 'createdAt', 'date', 'week', 'ward', 'patientTotal',
  'ventilator', 'hfov', 'hhhnc', 'o2', 'roomAir', 'homeVent',
  'patientGray', 'patientDarkRed', 'patientLightRed', 'patientYellow',
  'patientGreen', 'staffHN', 'staffRN', 'staffAW', 'cat', 'topic',
  'supervisee', 'supervisor', 'goal', 'reality', 'option', 'whatNext',
  'study', 'status'
];

function doGet() {
  try {
    const sheet = getOrCreateSheet_();
    const records = readRecords_(sheet);
    return jsonResponse_({ success: true, records: records });
  } catch (error) {
    return jsonResponse_({ success: false, error: error.message });
  }
}

function doPost(event) {
  const lock = LockService.getScriptLock();

  try {
    if (!event || !event.postData || !event.postData.contents) {
      throw new Error('Request body is required');
    }

    const record = validateRecord_(JSON.parse(event.postData.contents));
    lock.waitLock(10000);

    const sheet = getOrCreateSheet_();
    if (recordExists_(sheet, record.id)) {
      return jsonResponse_({ success: true, duplicate: true, id: record.id });
    }

    const row = HEADERS.map(header => safeCellValue_(record[header]));
    sheet.appendRow(row);
    SpreadsheetApp.flush();

    return jsonResponse_({ success: true, id: record.id });
  } catch (error) {
    return jsonResponse_({ success: false, error: error.message });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Apps Script must be attached to a Google Sheet');
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#0f766e')
      .setFontColor('#ffffff');
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
    if (currentHeaders.join('|') !== HEADERS.join('|')) {
      throw new Error('Records sheet headers do not match the current application schema');
    }
  }

  return sheet;
}

function readRecords_(sheet) {
  if (sheet.getLastRow() < 2) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  const timezone = Session.getScriptTimeZone();

  return rows
    .filter(row => row[0] !== '')
    .map(row => HEADERS.reduce((record, header, index) => {
      let value = row[index];
      if (value instanceof Date) {
        value = header === 'date'
          ? Utilities.formatDate(value, timezone, 'yyyy-MM-dd')
          : value.toISOString();
      }
      record[header] = value;
      return record;
    }, {}));
}

function recordExists_(sheet, id) {
  if (sheet.getLastRow() < 2) return false;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
  return ids.includes(String(id));
}

function validateRecord_(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) {
    throw new Error('Invalid record');
  }

  const record = {};
  HEADERS.forEach(header => {
    record[header] = input[header] ?? '';
  });

  ['id', 'date', 'ward', 'cat', 'topic', 'supervisor', 'status'].forEach(field => {
    if (String(record[field]).trim() === '') throw new Error(`${field} is required`);
  });

  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(record.date))) throw new Error('Invalid date');
  if (!['SCH-ICU2', 'SCH-ICU3', 'PCICU'].includes(record.ward)) throw new Error('Invalid ward');
  if (!['Environment', 'Flow', 'Infection', 'Risk', 'Equipment'].includes(record.cat)) throw new Error('Invalid category');
  if (!['done', 'pending'].includes(record.status)) throw new Error('Invalid status');

  return record;
}

function safeCellValue_(value) {
  if (typeof value === 'string' && /^[=+\-@]/.test(value)) return `'${value}`;
  return value;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
