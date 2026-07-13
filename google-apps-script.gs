/**
 * דמות הבוגר – כלי רפלקטיבי למיפוי מצב קיים
 * קוד Google Apps Script לקליטת תשובות השאלון לגיליון Google Sheets.
 *
 * התקנה: ראו הוראות מפורטות בקובץ README.md
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var headers = data.headers || [];
    var values = data.values || [];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('תשובות') || ss.insertSheet('תשובות');

    // בפעם הראשונה – כתיבת שורת כותרות
    if (sheet.getLastRow() === 0 && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow(values);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// בדיקה ידנית מתוך עורך ה-Apps Script (לא חובה)
function testAppend() {
  var e = { postData: { contents: JSON.stringify({
    headers: ['תאריך', 'שם בית הספר', 'שאלה לדוגמה'],
    values: [new Date().toLocaleString('he-IL'), 'בית ספר לדוגמה', '3 - במידה רבה']
  }) } };
  Logger.log(doPost(e).getContent());
}
