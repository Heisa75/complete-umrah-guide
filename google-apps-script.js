/**
 * ============================================================
 * UMRAH ACTIVITY LOG — Google Apps Script
 * ============================================================
 */
const SHEET_NAME = 'Umrah Activity Log';

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'Umrah Log API is running. Use POST to submit logs.' })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName('Logs') || ss.insertSheet('Logs');
    
    if (sheet.getLastRow() === 0) {
      const headers = ['Session ID', 'Umrah Started', 'Log #', 'Timestamp', 'Type', 'Activity', 'Elapsed Duration', 'Total Umrah Duration', 'Device Info'];
      sheet.appendRow(headers);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold').setBackground('#1B4332').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
    
    const sessionId = data.sessionId || 'N/A';
    const umrahStarted = data.umrahStarted || 'N/A';
    const totalDuration = data.totalDuration || 'N/A';
    const deviceInfo = data.deviceInfo || 'N/A';
    
    const logs = data.logs || [];
    logs.forEach((log, index) => {
      sheet.appendRow([sessionId, umrahStarted, index + 1, log.time || '', log.type || '', log.activity || '', log.duration || '', totalDuration, deviceInfo]);
    });
    
    writeSummary(ss, data);
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: `Logged ${logs.length} entries`, sheetUrl: ss.getUrl() })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function writeSummary(ss, data) {
  let summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('Summary');
    const headers = ['Date', 'Session ID', 'Umrah Started', 'Total Duration', 'Total Log Entries', 'Tawaf Rounds Logged', 'Saei Rounds Logged', 'Steps Completed', 'Device'];
    summarySheet.appendRow(headers);
    summarySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#C9A84C').setFontColor('#1B4332');
    summarySheet.setFrozenRows(1);
  }
  
  const logs = data.logs || [];
  summarySheet.appendRow([
    new Date().toISOString(), data.sessionId || 'N/A', data.umrahStarted || 'N/A', data.totalDuration || 'N/A',
    logs.length, logs.filter(l => l.activity && l.activity.includes('Tawaf')).length, logs.filter(l => l.activity && l.activity.includes("Sa'i")).length,
    logs.filter(l => l.type === 'complete').length, data.deviceInfo || 'N/A'
  ]);
}

function getOrCreateSpreadsheet() {
  const files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create(SHEET_NAME);
}