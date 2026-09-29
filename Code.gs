/* Google Apps Script web app for the shared SI Prep history.
 * Deploy this script from the Google Sheet listed in README.md.
 */
const SPREADSHEET_ID = '10umc1V8iAGp4wrD5IUpVuUeFgO_G8z9YFAjvj0CcRxA';
const SHEET_NAME = 'Attempts';
const ACCESS_CODE = '2911';

function doGet(e) {
  if (!authorised_(e.parameter)) return json_({ error: 'Unauthorised' });
  return json_({ attempts: getAttempts_() });
}

function doPost(e) {
  let payload;
  try { payload = JSON.parse(e.postData.contents); }
  catch (_) { return json_({ error: 'Invalid request' }); }
  if (!authorised_(payload)) return json_({ error: 'Unauthorised' });
  if (payload.action !== 'save' || !validAttempt_(payload.attempt)) {
    return json_({ error: 'Invalid attempt' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const attempts = getAttempts_();
    if (!attempts.some(a => String(a.id) === String(payload.attempt.id))) {
      const a = payload.attempt;
      sheet_().appendRow([
        String(a.id), a.exam, a.ts, a.date, a.total, a.correct,
        a.total ? a.correct / a.total : 0, JSON.stringify(a)
      ]);
      attempts.push(a);
    }
    return json_({ attempts: attempts });
  } finally {
    lock.releaseLock();
  }
}

function authorised_(data) {
  return data && String(data.code) === ACCESS_CODE;
}

function validAttempt_(a) {
  return a && a.id != null && Number.isInteger(a.exam) && Array.isArray(a.sections) &&
    Number.isFinite(a.total) && Number.isFinite(a.correct) && typeof a.date === 'string';
}

function sheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Missing ${SHEET_NAME} sheet`);
  return sheet;
}

function getAttempts_() {
  const sheet = sheet_();
  const last = sheet.getLastRow();
  if (last < 5) return [];
  return sheet.getRange(5, 8, last - 4, 1).getValues()
    .flat()
    .filter(Boolean)
    .map(value => JSON.parse(value));
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
