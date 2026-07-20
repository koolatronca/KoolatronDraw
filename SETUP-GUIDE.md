const SPREADSHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = 'Contest Entries';
const FOLDER_NAME = 'Kelly Drawing Contest Entries';

function doGet() {
  return ContentService.createTextOutput('Drawing contest endpoint is active.');
}

function doPost(e) {
  try {
    const data = e.parameter || {};
    if (!data.name || !data.email || !data.drawingData) throw new Error('Missing required fields.');
    if (data.guardianConfirmation !== 'true') throw new Error('Age/guardian confirmation required.');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) throw new Error('Invalid email.');

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp','Entry ID','Name','Email','Phone','Guardian confirmed','Marketing consent','Drawing link','Status']);
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,9).setFontWeight('bold');
    }

    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
    const entryId = Utilities.getUuid();
    const base64 = data.drawingData.split(',')[1];
    const bytes = Utilities.base64Decode(base64);
    const safeName = String(data.name).replace(/[^a-z0-9_-]/gi,'_').slice(0,40);
    const file = folder.createFile(Utilities.newBlob(bytes,'image/png',entryId+'_'+safeName+'.png'));

    const clean = v => String(v || '').trim().replace(/^[=+\-@]/, "'");
    sheet.appendRow([
      new Date(), entryId, clean(data.name), clean(data.email), clean(data.phone),
      data.guardianConfirmation === 'true', data.marketingConsent === 'true', file.getUrl(), 'Pending'
    ]);

    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({ok:true,entryId})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
