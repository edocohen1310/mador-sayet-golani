// ================================================================
// Google Apps Script — הערכת תרגילים מסלול תדהר
//
// הוראות התקנה:
//   1. צור Google Spreadsheet חדש
//   2. Extensions → Apps Script
//   3. מחק את הקוד הקיים, הדבק קובץ זה
//   4. שמור (Ctrl+S)
//   5. Deploy → New deployment → Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   6. לחץ Deploy, אשר הרשאות
//   7. העתק את ה-Web app URL לתוך index.html (משתנה SCRIPT_URL)
// ================================================================

const SHEET_NAME = 'הערכות';

const HEADERS = [
  'id', 'timestamp', 'coachName', 'exerciseName', 'team',
  'terrainType', 'date', 'scores', 'preserve1', 'preserve2',
  'improve1', 'improve2', 'avgScore',
];

// ---- GET: מחזיר את כל ההערכות ----
function doGet(e) {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();

  if (data.length <= 1) return respond([]);

  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    // Parse scores JSON string back to object
    if (typeof obj.scores === 'string' && obj.scores) {
      try { obj.scores = JSON.parse(obj.scores); } catch (_) { obj.scores = {}; }
    }
    return obj;
  });

  return respond(rows);
}

// ---- POST: שומר הערכה חדשה ----
function doPost(e) {
  const sheet = getSheet();
  const data  = JSON.parse(e.postData.contents);

  sheet.appendRow([
    data.id          || '',
    data.timestamp   || new Date().toISOString(),
    data.coachName   || '',
    data.exerciseName|| '',
    data.team        || '',
    data.terrainType || '',
    data.date        || '',
    JSON.stringify(data.scores || {}),
    data.preserve1   || '',
    data.preserve2   || '',
    data.improve1    || '',
    data.improve2    || '',
    data.avgScore    || '0',
  ]);

  return respond({ success: true });
}

// ---- Helpers ----
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#E8F0FE');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
