const SPREADSHEET_ID = '14RmxRKMLayZsJP0OE-PDamipeqhUIgvVXgVXrUrfv3U';
const SHEET_NAME = 'Contest Entries';
const FOLDER_NAME = 'Kelly Drawing Contest Entries';

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      message: 'Drawing contest endpoint is active.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = e.parameter || {};

    if (!data.name || !data.email || !data.drawingData || !data.challengeType) {
      throw new Error('Name, email, drawing, and challenge type are required.');
    }

    const sheet = getSheet();
    const folder = getFolder();
    const entryId = Utilities.getUuid();
    const drawingUrl = saveDrawing(folder, data.drawingData, entryId, data.name);

    sheet.appendRow([
      new Date(),
      entryId,
      clean(data.contestName),
      clean(data.challengeType),
      clean(data.name),
      clean(data.email),
      clean(data.phone),
      data.guardianConfirmation === 'true',
      data.marketingConsent === 'true',
      drawingUrl,
      'Pending'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, entryId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: String(error.message || error)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  const headers = [
    'Timestamp',
    'Entry ID',
    'Contest',
    'Challenge Type',
    'Name',
    'Email',
    'Phone',
    'Age/Guardian Confirmed',
    'Marketing Consent',
    'Drawing Link',
    'Judging Status'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else if (sheet.getRange(1, 4).getValue() !== 'Challenge Type') {
    sheet.insertColumnAfter(3);
    sheet.getRange(1, 4).setValue('Challenge Type').setFontWeight('bold');
  }

  return sheet;
}

function getFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
}

function saveDrawing(folder, dataUrl, entryId, entrantName) {
  if (!String(dataUrl).startsWith('data:image/png;base64,')) {
    throw new Error('Invalid drawing image.');
  }

  const base64Data = String(dataUrl).split(',')[1];
  const imageBytes = Utilities.base64Decode(base64Data);
  const safeName = clean(entrantName)
    .replace(/[^a-z0-9_-]/gi, '_')
    .slice(0, 40) || 'entrant';

  const blob = Utilities.newBlob(
    imageBytes,
    'image/png',
    entryId + '_' + safeName + '.png'
  );

  return folder.createFile(blob).getUrl();
}

function clean(value) {
  return String(value || '').trim().replace(/^[=+\-@]/, "'");
}
