/**
 * Книжный вернисаж — Google Sheets backend.
 * Владелец таблицы: https://docs.google.com/spreadsheets/d/1_r42GFPgdUFHpSTU3oiJ_p8dXwA0kN86HnpRtFYvjso/edit
 * Разместите этот код в «Расширения → Apps Script» созданной таблицы
 * и опубликуйте как веб-приложение: Выполнять от имени "Я", Доступ "Все".
 * Не публикуйте здесь секретные токены/пароли.
 */
const SPREADSHEET_ID = '1_r42GFPgdUFHpSTU3oiJ_p8dXwA0kN86HnpRtFYvjso';
const SHEET_NAME = 'Ответы';
const HEADERS = ['ID','Дата и время','Фамилия','Имя','Эмоция','Выбранные темы','Рекомендация'];

function sheet_() {
  const file = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = file.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = file.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}
function safe_(value, max) {
  let v = String(value == null ? '' : value).trim().slice(0, max);
  // Не позволяем вводимому пользователем тексту исполняться как формула.
  if (/^[=+\-@]/.test(v)) v = "'" + v;
  return v;
}
function json_(value, callback) {
  const json = JSON.stringify(value);
  if (callback && /^[a-zA-Z_$][\w$]{0,63}$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) {
  try {
    const p = e && e.parameter || {};
    if (p.action === 'status') return json_({ok: true, service: 'book-vernissage'}, p.callback);
    if (p.action !== 'read') return json_({ok: false, error: 'unknown_action'}, p.callback);
    const sheet = sheet_();
    const n = sheet.getLastRow();
    if (n < 2) return json_([], p.callback);
    const rows = sheet.getRange(2, 1, n - 1, 7).getDisplayValues();
    const participants = rows.filter(r => r[2] || r[3]).map(r => ({
      id: r[0], timestamp: r[1], surname: r[2], firstname: r[3],
      emotion: r[4], topics: r[5], recommendation: r[6]
    }));
    return json_(participants, p.callback);
  } catch (error) {
    return json_({ok: false, error: String(error)}, e && e.parameter && e.parameter.callback);
  }
}
function doPost(e) {
  const p = e && e.parameter || {};
  if (p.action !== 'write') return json_({ok: false, error: 'unknown_action'});
  const surname = safe_(p.surname, 100);
  const firstname = safe_(p.firstname, 100);
  const emotion = safe_(p.emotion, 150);
  const topics = safe_(p.topics, 1500);
  const recommendation = safe_(p.recommendation, 300);
  if (!surname || !firstname || !emotion || !topics || !recommendation) {
    return json_({ok: false, error: 'missing_required_fields'});
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = sheet_();
    const rowCount = sheet.getLastRow();
    if (rowCount > 1) {
      const names = sheet.getRange(2, 3, rowCount - 1, 2).getDisplayValues();
      const isDuplicate = names.some(r =>
        String(r[0]).trim().toLocaleLowerCase() === surname.toLocaleLowerCase() &&
        String(r[1]).trim().toLocaleLowerCase() === firstname.toLocaleLowerCase()
      );
      if (isDuplicate) return json_({ok: false, error: 'duplicate_participant'});
    }
    const id = Utilities.getUuid();
    const stamp = Utilities.formatDate(new Date(), 'Asia/Almaty', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([id, stamp, surname, firstname, emotion, topics, recommendation]);
    SpreadsheetApp.flush();
    return json_({ok: true, id: id});
  } catch (error) {
    return json_({ok: false, error: String(error)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
