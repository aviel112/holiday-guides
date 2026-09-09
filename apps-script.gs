/**
 * מדריכי תשרי — Google Apps Script
 * שורה אחת לכל ליד, שמתעדכנת לאורך כל המסע: הרשמה → קריאת מדריכים → ספר מתכונים.
 * בכל רגע אפשר לראות בעמודת "שלב אחרון" איפה כל אחד נעצר.
 *
 * ── התקנה ─────────────────────────────────────────────────────────
 * אפשרות א׳ — גיליון חדש:
 *   פותחים גיליון (sheets.new) → Extensions → Apps Script → מדביקים.
 * אפשרות ב׳ — גיליון קיים שכבר יש בו סקריפט אחר:
 *   script.google.com → New project → מדביקים כאן →
 *   ממלאים את SPREADSHEET_ID למטה עם המזהה של הגיליון הקיים.
 *   נוצר טאב חדש בשם SHEET_NAME, בלי לגעת בשום דבר קיים.
 *
 * 1. מדביקים את הקוד (לפי אחת האפשרויות למעלה).
 * 2. שומרים (Cmd+S).
 * 3. מריצים פעם אחת את setupSheet (Run) — נוצרות הכותרות והצבעים.
 * 4. Deploy → New deployment → Web app:
 *      Description: holiday-guides
 *      Execute as: Me (aviamira5@gmail.com)
 *      Who has access: Anyone
 *    Deploy → מאשרים הרשאות (Advanced → Go to project → Allow).
 * 5. מעתיקים את ה-Web app URL (מסתיים ב-/exec)
 *    ומדביקים אותו ב-SHEET_URL בתוך index.html.
 *
 * ── עדכון קוד בעתיד ───────────────────────────────────────────────
 * Deploy → Manage deployments → עיפרון → Version: New version → Deploy
 * (ה-URL נשאר אותו דבר).
 *
 * ── שלוש הפעולות שהדף שולח ────────────────────────────────────────
 * action:'lead'         — נרשם בשער → יוצר שורה (או מעדכן קיימת לפי טלפון)
 * action:'progress'     — פתח מדריך → מעדכן אילו מדריכים נפתחו ואת השלב
 * action:'recipe-book'  — לחץ על ספר המתכונים → מסמן ✓ + חותמת זמן
 *
 * כל פעולה מזוהה לפי הטלפון בלבד — אז הכל נוחת על אותה שורה.
 *
 * ── מוכן ל-ManyChat ───────────────────────────────────────────────
 * קישור מה-DM:  .../holiday-guides/?utm_source=manychat&utm_campaign=tishrei-dm
 * או POST ישיר מ-ManyChat ל-URL הזה עם {"action":"lead","firstName":...,"phone":...}
 * — אותה שורה, אותו גיליון, בלי מאגר כפול.
 */

/* ===== לאן כותבים =====
   SPREADSHEET_ID — ריק = הגיליון שאליו הסקריפט מחובר (Extensions → Apps Script מתוך הגיליון).
                    אם ממלאים מזהה — הסקריפט יכול לרוץ עצמאי ולכתוב לגיליון קיים כלשהו.
   המזהה הוא החלק באמצע הכתובת:
   docs.google.com/spreadsheets/d/<<< המזהה כאן >>>/edit
   SHEET_NAME — שם הטאב. אם הוא לא קיים — נוצר אוטומטית, בלי לגעת בטאבים אחרים. */
var SPREADSHEET_ID = '1E2TYEEND0mof5yj-mfKV1lD7eVDbDxYXxEnLMBBX0MI';  // "לידים - מחשבון מאקרו חגים"
var SHEET_NAME     = 'מדריכי תשרי';

var HEADERS = [
  'תאריך הרשמה',        // A
  'שם פרטי',            // B
  'שם משפחה',           // C
  'גיל',                // D
  'טלפון',              // E
  'אינסטגרם',           // F
  'מקור (funnel)',      // G
  'שלב אחרון',          // H  ← כאן רואים איפה נעצר
  'מדריכים שנפתחו',     // I
  'כמה מדריכים',        // J
  'נכנס לספר מתכונים',  // K
  'תאריך כניסה לספר',   // L
  'פעילות אחרונה',      // M
  'מקור תנועה',         // N  utm_source
  'קמפיין'              // O  utm_campaign
];

var COL = {
  DATE:1, FIRST:2, LAST:3, AGE:4, PHONE:5, IG:6, FUNNEL:7,
  STAGE:8, GUIDES:9, GUIDES_N:10, BOOK:11, BOOK_DATE:12,
  SEEN:13, UTM_SRC:14, UTM_CMP:15
};

/* שמות ידידותיים למדריכים — מגיעים מהדף כמזהים (g1..g4) */
var GUIDE_NAMES = {
  g1: 'שולחן החג',
  g2: 'ראש השנה',
  g3: 'יום כיפור',
  g4: 'סוכות'
};

/* סולם השלבים — מספר גבוה יותר לא יורד אחורה */
var STAGES = [
  '1 · נרשם בלבד',
  '2 · פתח מדריך אחד',
  '3 · קרא כמה מדריכים',
  '4 · עבר לספר המתכונים'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // מונע דריסת שורות כששניים פעילים באותו רגע
  } catch (err) {
    return _json({ ok: false, error: 'busy' });
  }

  try {
    var data = {};
    if (e && e.postData && e.postData.contents) data = JSON.parse(e.postData.contents);
    else if (e && e.parameter) data = e.parameter;

    var action = data.action || 'lead';
    var phone  = _normPhone(data.phone);
    if (!phone) return _json({ ok: false, error: 'no-phone' });

    var sheet = _getSheet();
    var row   = _findRowByPhone(sheet, phone);

    /* ---- 1. הרשמה ---- */
    if (action === 'lead') {
      if (row) {
        // חזר לדף ונרשם שוב — מעדכנים פרטים בלי לאבד את ההתקדמות
        sheet.getRange(row, COL.FIRST).setValue(data.firstName || '');
        sheet.getRange(row, COL.LAST).setValue(data.lastName || '');
        sheet.getRange(row, COL.AGE).setValue(data.age || '');
        sheet.getRange(row, COL.IG).setValue(data.instagram || data.ig || '');
      } else {
        sheet.appendRow([
          new Date(),
          data.firstName || '',
          data.lastName  || '',
          data.age       || '',
          "'" + phone,                    // גרש מוביל = נשמר כטקסט, האפס המוביל לא נעלם
          data.instagram || data.ig || '',
          data.funnel    || 'holiday-guides',
          STAGES[0], '', 0, '', '', new Date(),
          data.utm_source   || '',
          data.utm_campaign || ''
        ]);
        row = sheet.getLastRow();
      }
      _touch(sheet, row);
      _paintStage(sheet, row);
      return _json({ ok: true, row: row, stage: sheet.getRange(row, COL.STAGE).getValue() });
    }

    if (!row) return _json({ ok: false, error: 'lead-not-found' });

    /* ---- 2. התקדמות בקריאה ---- */
    if (action === 'progress') {
      var ids = String(data.guides || '').split(',')
        .map(function (x) { return x.trim(); })
        .filter(function (x) { return GUIDE_NAMES[x]; });

      // מאחדים עם מה שכבר רשום, כדי שגלישה בכמה ביקורים תצטבר
      var existing = String(sheet.getRange(row, COL.GUIDES).getValue() || '')
        .split(',').map(function (x) { return x.trim(); }).filter(String);

      var names = ids.map(function (id) { return GUIDE_NAMES[id]; });
      var merged = [];
      existing.concat(names).forEach(function (n) {
        if (merged.indexOf(n) === -1) merged.push(n);
      });

      sheet.getRange(row, COL.GUIDES).setValue(merged.join(', '));
      sheet.getRange(row, COL.GUIDES_N).setValue(merged.length);
      _raiseStage(sheet, row, merged.length >= 2 ? 2 : (merged.length === 1 ? 1 : 0));
      _touch(sheet, row);
      return _json({ ok: true, row: row, guides: merged.length });
    }

    /* ---- 3. מעבר לספר המתכונים ---- */
    if (action === 'recipe-book') {
      if (!sheet.getRange(row, COL.BOOK).getValue()) {   // רק בפעם הראשונה
        sheet.getRange(row, COL.BOOK).setValue('✓');
        sheet.getRange(row, COL.BOOK_DATE).setValue(new Date());
      }
      _raiseStage(sheet, row, 3);
      _touch(sheet, row);
      return _json({ ok: true, row: row, marked: true });
    }

    return _json({ ok: false, error: 'unknown-action' });

  } catch (err) {
    return _json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// בדיקה שהדיפלוי חי — פותחים את ה-URL בדפדפן
function doGet() {
  return ContentService.createTextOutput('OK — holiday-guides endpoint is live');
}

/* ---------- עזרים ---------- */

// שלב רק עולה, אף פעם לא יורד — כדי שרענון דף לא "יוריד" מישהו בחזרה
function _raiseStage(sheet, row, level) {
  var cell = sheet.getRange(row, COL.STAGE);
  var cur  = STAGES.indexOf(String(cell.getValue()));
  if (level > cur) cell.setValue(STAGES[level]);
  _paintStage(sheet, row);
}

// צביעה לפי שלב — כדי לראות במבט אחד איפה כולם נעצרו
function _paintStage(sheet, row) {
  var colors = ['#fde8e4', '#fdf1dc', '#e9f2fb', '#e7f4f0'];  // אדמדם → ירקרק
  var idx = STAGES.indexOf(String(sheet.getRange(row, COL.STAGE).getValue()));
  if (idx >= 0) sheet.getRange(row, COL.STAGE).setBackground(colors[idx]);
}

function _touch(sheet, row) {
  sheet.getRange(row, COL.SEEN).setValue(new Date());
}

// מנרמל כל פורמט ל-05XXXXXXXX, כדי שההצלבה בין האירועים תמיד תתפוס
function _normPhone(raw) {
  var d = String(raw || '').replace(/\D/g, '');
  if (d.indexOf('00972') === 0) d = '0' + d.slice(5);
  else if (d.indexOf('972') === 0) d = '0' + d.slice(3);
  return d;
}

function _findRowByPhone(sheet, phone) {
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var col = sheet.getRange(2, COL.PHONE, last - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (_normPhone(col[i][0]) === phone) return i + 2;
  }
  return null;
}

function _getSheet() {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  var current = sheet.getLastRow() === 0
    ? []
    : sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];

  if (current.join('|') !== HEADERS.join('|')) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
         .setFontWeight('bold').setBackground('#0f6f5c').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.getRange(1, COL.BOOK, sheet.getMaxRows(), 1).setHorizontalAlignment('center');
    sheet.getRange(1, COL.GUIDES_N, sheet.getMaxRows(), 1).setHorizontalAlignment('center');
    sheet.setColumnWidth(COL.DATE, 150);
    sheet.setColumnWidth(COL.STAGE, 165);
    sheet.setColumnWidth(COL.GUIDES, 220);
    sheet.setColumnWidth(COL.BOOK_DATE, 150);
    sheet.setColumnWidth(COL.SEEN, 150);
  }
  return sheet;
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* מציג חלונית כשרצים מתוך גיליון, ומדפיס ללוג כשהסקריפט עצמאי */
function _out(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch (e) { Logger.log(msg); }
}

/* ---------- כלים להפעלה ידנית מתוך העורך ---------- */

// מריצים פעם אחת כדי לייצר את הגיליון והכותרות בלי לחכות לליד ראשון
function setupSheet() {
  _getSheet();
}

// דוח נשירה — כמה נעצרו בכל שלב
function showFunnel() {
  var sheet = _getSheet();
  var last = sheet.getLastRow();
  if (last < 2) { _out('אין עדיין לידים.'); return; }

  var stages = sheet.getRange(2, COL.STAGE, last - 1, 1).getValues();
  var total = stages.length;
  var counts = [0, 0, 0, 0];
  stages.forEach(function (r) {
    var i = STAGES.indexOf(String(r[0]));
    if (i >= 0) counts[i]++;
  });

  var lines = ['סה״כ נרשמו: ' + total, ''];
  for (var i = 0; i < STAGES.length; i++) {
    var pct = Math.round(counts[i] / total * 100);
    lines.push('נעצרו ב' + STAGES[i] + ' — ' + counts[i] + ' (' + pct + '%)');
  }
  // כמה הגיעו לפחות עד כל שלב (מצטבר)
  var reachedBook = counts[3];
  var reachedRead = counts[1] + counts[2] + counts[3];
  lines.push('');
  lines.push('פתחו לפחות מדריך אחד: ' + reachedRead + ' (' + Math.round(reachedRead / total * 100) + '%)');
  lines.push('הגיעו לספר המתכונים: ' + reachedBook + ' (' + Math.round(reachedBook / total * 100) + '%)');

  _out(lines.join('\n'));
}

/*
  ===== סטטוס פריסה =====
  גיליון: "לידים - מחשבון מאקרו חגים" (הקיים) — טאב חדש בשם "מדריכי תשרי"
  spreadsheet id: 1E2TYEEND0mof5yj-mfKV1lD7eVDbDxYXxEnLMBBX0MI
  הסקריפט עצמאי (script.google.com), לא מחובר לגיליון — כדי לא להתנגש
  בסקריפט הקיים של מחשבון המאקרו שכבר יושב על אותו גיליון.
  Web App URL: (להדביק כאן אחרי הפריסה, ולהזין ב-SHEET_URL בתוך index.html)
  הגיליון הוא היעד היחיד — אין CRM ואין מאגר מקביל.
  הדף שולח כל אירוע ב-Content-Type: text/plain — בקשה פשוטה בלי preflight.
*/
