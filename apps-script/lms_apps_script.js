/****************************************************
 * بک‌اند سامانه آموزشی درس‌نامه نقد احمدالحسن بصری
 * Backend: Google Apps Script + Google Sheets (کاملاً رایگان)
 *
 * راه‌اندازی:
 * ۱) یک گوگل‌شیت تازه بسازید
 * ۲) از منوی Extensions گزینه Apps Script را باز کنید
 * ۳) کل این کد را جایگزین کد پیش‌فرض کنید و ذخیره (Ctrl+S)
 * ۴) یک بار تابع setup() را اجرا کنید (اجازه دسترسی بدهید)
 * ۵) در شیت Config گذرواژه استاد و کد کلاس را بنویسید
 * ۶) Deploy > New deployment > نوع Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * ۷) آدرس وب‌اپ (ختم‌شونده به /exec) را در بخش تنظیمات سامانه وارد کنید
 ****************************************************/

const SHEETS = {
  config: 'Config',
  students: 'Students',
  progress: 'Progress',
  quiz: 'Quiz',
  questions: 'Questions',
  announcements: 'Announcements',
};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = {
    Config: ['key', 'value'],
    Students: ['studentId', 'name', 'code', 'registeredAt', 'lastActive'],
    Progress: ['ts', 'studentId', 'name', 'chapterId', 'lessonIdx', 'done'],
    Quiz: ['ts', 'studentId', 'name', 'chapterId', 'qIndex', 'question', 'rating', 'answer'],
    Questions: ['ts', 'studentId', 'name', 'text', 'answered'],
    Announcements: ['ts', 'text'],
  };
  Object.keys(SHEETS).forEach(k => {
    const name = SHEETS[k];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
      sh.getRange(1, 1, 1, headers[name].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  const cfg = ss.getSheetByName(SHEETS.config);
  const keys = cfg.getRange('A2:A').getValues().flat().filter(String);
  if (!keys.includes('teacherPassword')) cfg.appendRow(['teacherPassword', '1234']);
  if (!keys.includes('classCode')) cfg.appendRow(['classCode', 'QAEM1404']);
  Logger.log('راه‌اندازی کامل شد. گذرواژه و کد کلاس را در شیت Config تغییر دهید.');
}

function getConfig(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.config);
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) return String(vals[i][1]).trim();
  }
  return '';
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function newId() {
  return 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function touchStudent(ss, studentId, name) {
  const sh = ss.getSheetByName(SHEETS.students);
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === studentId) {
      sh.getRange(i + 1, 5).setValue(new Date());
      return;
    }
  }
}

/* ---------------- POST ---------------- */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();

    if (data.action === 'register') {
      const code = String(data.code || '').trim();
      if (code !== getConfig('classCode')) return json({ ok: false, error: 'wrong_code' });
      const id = newId();
      ss.getSheetByName(SHEETS.students)
        .appendRow([id, String(data.name || '').trim(), code, now, now]);
      return json({ ok: true, studentId: id });
    }

    if (data.action === 'progress') {
      ss.getSheetByName(SHEETS.progress).appendRow(
        [now, data.studentId, data.name, data.chapterId, data.lessonIdx, data.done ? 'TRUE' : 'FALSE']);
      touchStudent(ss, data.studentId, data.name);
      return json({ ok: true });
    }

    if (data.action === 'quiz') {
      const sh = ss.getSheetByName(SHEETS.quiz);
      (data.answers || []).forEach((a, idx) => {
        sh.appendRow([now, data.studentId, data.name, data.chapterId, idx, a.q, a.rating, a.text || '']);
      });
      touchStudent(ss, data.studentId, data.name);
      return json({ ok: true });
    }

    if (data.action === 'ask') {
      ss.getSheetByName(SHEETS.questions)
        .appendRow([now, data.studentId, data.name, String(data.text || '').trim(), 'FALSE']);
      touchStudent(ss, data.studentId, data.name);
      return json({ ok: true });
    }

    if (data.action === 'markAnswered') {
      const sh = ss.getSheetByName(SHEETS.questions);
      const vals = sh.getDataRange().getValues();
      const target = Number(data.ts);
      for (let i = 1; i < vals.length; i++) {
        const rowTs = vals[i][0] instanceof Date ? vals[i][0].getTime() : Number(vals[i][0]);
        if (String(vals[i][1]) === String(data.studentId) && Math.abs(rowTs - target) < 2000) {
          sh.getRange(i + 1, 5).setValue('TRUE');
          return json({ ok: true });
        }
      }
      return json({ ok: false, error: 'not_found' });
    }

    return json({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ---------------- GET ---------------- */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const a = (e.parameter.action || '').trim();

    if (a === 'announcements') {
      const vals = ss.getSheetByName(SHEETS.announcements).getDataRange().getValues();
      const list = [];
      for (let i = vals.length - 1; i >= 1 && list.length < 20; i--) {
        list.push({ ts: vals[i][0], text: String(vals[i][1] || '') });
      }
      return json({ ok: true, announcements: list });
    }

    if (a === 'dashboard') {
      if (String(e.parameter.password || '') !== getConfig('teacherPassword')) {
        return json({ ok: false, error: 'auth' });
      }
      return json({ ok: true, data: buildDashboard(ss) });
    }

    return json({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function sheetRows(name) {
  const vals = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getDataRange().getValues();
  return vals.slice(1).filter(r => r.some(c => String(c).trim() !== ''));
}

function buildDashboard(ss) {
  const students = sheetRows(SHEETS.students).map(r => ({
    studentId: String(r[0]), name: String(r[1]),
    registeredAt: r[3], lastActive: r[4],
  }));
  const progress = sheetRows(SHEETS.progress).map(r => ({
    ts: r[0], studentId: String(r[1]), name: String(r[2]),
    chapterId: String(r[3]), lessonIdx: r[4], done: String(r[5]) === 'TRUE',
  }));
  const quiz = sheetRows(SHEETS.quiz).map(r => ({
    ts: r[0], studentId: String(r[1]), name: String(r[2]),
    chapterId: String(r[3]), qIndex: r[4], question: String(r[5]),
    rating: String(r[6]), answer: String(r[7] || ''),
  }));
  const questions = sheetRows(SHEETS.questions).map(r => ({
    ts: r[0], studentId: String(r[1]), name: String(r[2]),
    text: String(r[3]), answered: String(r[4]) === 'TRUE',
  }));
  return { students, progress, quiz, questions };
}
