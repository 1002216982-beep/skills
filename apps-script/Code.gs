/**
 * דמות הבוגר – כלי רפלקטיבי למיפוי מצב קיים
 * פרויקט Google Apps Script: מגיש את השאלון כ-Web App ושומר את התשובות לגיליון.
 *
 * התקנה:
 * 1. script.new ← פרויקט חדש
 * 2. הדביקו קובץ זה ב-Code.gs
 * 3. הוסיפו קובץ HTML בשם Index והדביקו בו את תוכן Index.html
 * 4. Deploy ← New deployment ← Web app ← Execute as: Me ← Who has access: Anyone
 * 5. פתחו את כתובת ה-Web app – זהו השאלון.
 *
 * הגיליון נוצר אוטומטית בפעם הראשונה שנשלחת תשובה, ונשמר ב-Drive שלכם
 * בשם "מיפוי מצב קיים – תשובות". (אפשר גם להדביק ID של גיליון קיים למטה.)
 */

var SHEET_ID = '1AEKJowdnP_aLlQE8oAV7K3YiK9-h9lxxEJvzQjWzF9g'; // הגיליון שאליו נאספות התשובות. ריק = ייווצר גיליון חדש אוטומטית.
var SHEET_NAME = 'תשובות';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('דמות הבוגר – כלי רפלקטיבי למיפוי מצב קיים')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet_() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* נמחק – ניצור חדש */ }
  }
  var ss = SpreadsheetApp.create('מיפוי מצב קיים – תשובות');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

/** נקרא מהדפדפן דרך google.script.run */
function saveResponse(headers, values) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    // מחיקת לשוניות ברירת-מחדל ריקות (למשל "גיליון1") כדי שהתשובות יופיעו מיד בפתיחת הקובץ
    ss.getSheets().forEach(function (s) {
      if (s.getSheetId() !== sheet.getSheetId() && s.getLastRow() === 0 && s.getLastColumn() === 0) {
        try { ss.deleteSheet(s); } catch (e) { /* לשונית אחרונה – מתעלמים */ }
      }
    });
    if (sheet.getLastRow() === 0 && headers && headers.length) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow(values);
    return { status: 'ok', url: ss.getUrl() };
  } finally {
    lock.releaseLock();
  }
}
