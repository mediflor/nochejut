/**
 * פורטל הצוות – פ.ד מדיפלור בע"מ
 * מודול "מעקב פניות מטופלים": עמית מתקשר → רחלי קובעת תור → ד"ר פלור מחליטה → לאו מבקר.
 * כלל ברזל: אין במערכת שום מידע רפואי. רק שם, טלפון, מזהה פנימי, תאריכים וסטטוסים.
 * הקובץ עצמאי: הוא משתמש רק ב-empByToken_ מהקוד הקיים, ולא משנה דבר באפליקציית הנוכחות.
 */

const PT_SHEET = 'פניות מטופלים';
const PT_LOG = 'יומן פניות';
const PT_ROLES = 'תפקידי פורטל';

const PT_HEAD = ['מזהה פנימי', 'נוצר', 'שם מטופל', 'טלפון', 'סוג פנייה', 'סטטוס', 'זמין מתאריך',
  'ניסיונות שיחה', 'תוצאת שיחה', 'הסבר (ללא מידע רפואי)', 'שיחה אחרונה', 'ניסיונות קביעה',
  'תאריך תור', 'שעת תור', 'הערת מזכירה', 'החלטת רופאה', 'הנחיה', 'הגעה', 'ביקר ברבעון',
  'עודכן', 'עודכן ע"י', 'רשימה'];
const PC = { ID: 0, CREATED: 1, NAME: 2, PHONE: 3, TYPE: 4, STATUS: 5, AVAIL: 6, CALLS: 7, CALLRES: 8,
  NOTE: 9, LASTCALL: 10, BOOKTRIES: 11, APPTDATE: 12, APPTTIME: 13, BOOKNOTE: 14, DOCDEC: 15,
  INSTR: 16, ARRIVAL: 17, VISITED: 18, UPDATED: 19, UPDBY: 20, LIST: 21 };

const ST = { CALL: 'לשיחה', BOOK: 'לקביעת תור', DOC: 'להחלטת רופאה', BOOKED: 'תור נקבע',
  DONE: 'הושלם', NOSHOW: 'לא הגיע', CLOSED: 'סגור' };
const PT_TYPES = ['מדדים', 'חיסון', 'ביקור', 'חידוש מרשם'];
const PT_INTERNAL = 'חידוש מרשם';
const PT_MAX_CALLS = 3;
const PT_MAX_BOOK = 3;

const ROLE = { CALLER: 'מתקשר', BOOKER: 'קובעת תורים', DOCTOR: 'רופאה', MANAGER: 'מנהל' };
const PT_DEFAULT_ROLES = [['E001', ROLE.MANAGER], ['E007', ROLE.DOCTOR], ['E005', ROLE.CALLER], ['E004', ROLE.BOOKER]];

/* ====================== דף ====================== */

function portalDoGet_(e) {
  const token = (e && e.parameter && e.parameter.k) || '';
  const emp = empByToken_(token);
  const t = HtmlService.createTemplateFromFile('PortalPage');
  t.token = emp ? token : '';
  t.invalid = !emp;
  return t.evaluate()
    .setTitle('פורטל הצוות – מדיפלור')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** נקודת כניסה יחידה של הפורטל מהדפדפן */
function portalApi(token, op, payload) {
  const emp = empByToken_(token);
  if (!emp) throw new Error('הקישור אינו תקף או שהעובד/ת אינו/ה פעיל/ה. פנו למנהל.');
  payload = payload || {};
  const roles = ptRoles_(emp);
  const need = function (r) {
    if (roles.indexOf(r) < 0 && roles.indexOf(ROLE.MANAGER) < 0) throw new Error('אין הרשאה לפעולה זו.');
  };
  switch (op) {
    case 'me': return { name: emp.name, roles: roles, types: PT_TYPES, counts: ptCounts_() };
    case 'list': return ptList_(payload.view, roles);
    case 'callResult': need(ROLE.CALLER); return ptWrite_(emp, function () { return ptCallResult_(emp, payload); });
    case 'bookResult': need(ROLE.BOOKER); return ptWrite_(emp, function () { return ptBookResult_(emp, payload); });
    case 'arrival': need(ROLE.BOOKER); return ptWrite_(emp, function () { return ptArrival_(emp, payload); });
    case 'visited': return ptWrite_(emp, function () { return ptVisited_(emp, payload); });
    case 'doctorDecision': need(ROLE.DOCTOR); return ptWrite_(emp, function () { return ptDoctor_(emp, payload); });
    case 'importList': need(ROLE.MANAGER); return ptWrite_(emp, function () { return ptImport_(emp, payload); });
    case 'stats': need(ROLE.MANAGER); return ptStats_();
  }
  throw new Error('פעולה לא מוכרת');
}

/* ====================== תפקידים ====================== */

function ptRoles_(emp) {
  const sh = ptSheet_(PT_ROLES, ['מזהה עובד', 'שם', 'תפקיד בפורטל'], function (s) {
    const emps = readEmps_();
    PT_DEFAULT_ROLES.forEach(function (d) {
      const em = emps.filter(function (x) { return x.id === d[0]; })[0];
      s.appendRow([d[0], em ? em.name : '', d[1]]);
    });
  });
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === emp.id) {
      String(v[i][2] || '').split(/[,،]/).forEach(function (r) { r = r.trim(); if (r) out.push(r); });
    }
  }
  if (emp.admin && out.indexOf(ROLE.MANAGER) < 0) out.push(ROLE.MANAGER);
  return out;
}

/* ====================== קריאה ====================== */

function ptRows_() {
  const sh = ptSheet_(PT_SHEET, PT_HEAD);
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (!String(v[i][PC.ID]).trim()) continue;
    out.push(ptObj_(v[i], i + 1));
  }
  return out;
}

function ptObj_(r, row) {
  return {
    row: row, id: String(r[PC.ID]), created: ptStr_(r[PC.CREATED]), name: String(r[PC.NAME] || ''),
    phone: ptPhone_(r[PC.PHONE]), type: String(r[PC.TYPE] || ''), status: String(r[PC.STATUS] || ''),
    avail: ptDate_(r[PC.AVAIL]), calls: Number(r[PC.CALLS]) || 0, callRes: String(r[PC.CALLRES] || ''),
    note: String(r[PC.NOTE] || ''), lastCall: ptStr_(r[PC.LASTCALL]), bookTries: Number(r[PC.BOOKTRIES]) || 0,
    apptDate: ptDate_(r[PC.APPTDATE]), apptTime: ptTime_(r[PC.APPTTIME]), bookNote: String(r[PC.BOOKNOTE] || ''),
    docDec: String(r[PC.DOCDEC] || ''), instr: String(r[PC.INSTR] || ''), arrival: String(r[PC.ARRIVAL] || ''),
    visited: String(r[PC.VISITED] || '') === 'כן', updated: ptStr_(r[PC.UPDATED]), updBy: String(r[PC.UPDBY] || ''),
    list: String(r[PC.LIST] || '')
  };
}

function ptList_(view, roles) {
  const today = ptToday_();
  const all = ptRows_();
  const isMgr = roles.indexOf(ROLE.MANAGER) >= 0;
  const can = function (r) { return isMgr || roles.indexOf(r) >= 0; };
  let rows;
  if (view === 'calls' && can(ROLE.CALLER)) {
    rows = all.filter(function (p) { return p.status === ST.CALL && (!p.avail || p.avail <= today); });
    rows.sort(function (a, b) { return (a.visited - b.visited) || (a.calls - b.calls) || (a.id < b.id ? -1 : 1); });
  } else if (view === 'book' && can(ROLE.BOOKER)) {
    rows = all.filter(function (p) { return p.status === ST.BOOK && (!p.avail || p.avail <= today); });
    rows.sort(function (a, b) { return (a.visited - b.visited) || (a.bookTries - b.bookTries) || (a.id < b.id ? -1 : 1); });
  } else if (view === 'followup' && can(ROLE.BOOKER)) {
    rows = all.filter(function (p) { return p.status === ST.BOOKED && p.apptDate && p.apptDate <= today; });
  } else if (view === 'upcoming' && can(ROLE.BOOKER)) {
    rows = all.filter(function (p) { return p.status === ST.BOOKED && p.apptDate > today; });
    rows.sort(function (a, b) { return a.apptDate < b.apptDate ? -1 : 1; });
  } else if (view === 'doctor' && can(ROLE.DOCTOR)) {
    rows = all.filter(function (p) { return p.status === ST.DOC; });
  } else if (view === 'all' && isMgr) {
    rows = all.slice().reverse();
  } else {
    throw new Error('אין הרשאה לרשימה זו.');
  }
  return { rows: rows.map(ptPub_), counts: ptCounts_(all), today: today };
}

function ptCounts_(all) {
  all = all || ptRows_();
  const today = ptToday_();
  const c = { calls: 0, book: 0, doctor: 0, followup: 0, upcoming: 0 };
  all.forEach(function (p) {
    if (p.status === ST.CALL && (!p.avail || p.avail <= today)) c.calls++;
    if (p.status === ST.BOOK && (!p.avail || p.avail <= today)) c.book++;
    if (p.status === ST.DOC) c.doctor++;
    if (p.status === ST.BOOKED && p.apptDate && p.apptDate <= today) c.followup++;
    if (p.status === ST.BOOKED && p.apptDate > today) c.upcoming++;
  });
  return c;
}

function ptStats_() {
  const all = ptRows_();
  const by = {};
  all.forEach(function (p) { by[p.status] = (by[p.status] || 0) + 1; });
  const called = all.filter(function (p) { return p.callRes && p.callRes !== 'לא ענה'; }).length;
  const interested = all.filter(function (p) { return p.callRes === 'מעוניין'; }).length;
  const booked = all.filter(function (p) { return p.apptDate; }).length;
  const arrived = all.filter(function (p) { return p.arrival === 'הגיע'; }).length;
  const noshow = all.filter(function (p) { return p.arrival === 'לא הגיע'; }).length;
  const newQ = all.filter(function (p) { return p.apptDate && !p.visited; }).length;
  return {
    total: all.length, byStatus: by, called: called, interested: interested, booked: booked,
    arrived: arrived, noshow: noshow, newInQuarter: newQ,
    responseRate: called ? Math.round(100 * interested / called) : 0,
    arrivalRate: (arrived + noshow) ? Math.round(100 * arrived / (arrived + noshow)) : 0,
    counts: ptCounts_(all)
  };
}

/* ====================== כתיבה ====================== */

function ptWrite_(emp, fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function ptFind_(id) {
  const p = ptRows_().filter(function (x) { return x.id === String(id); })[0];
  if (!p) throw new Error('הפנייה לא נמצאה. רעננו את הרשימה.');
  return p;
}

function ptSet_(p, emp, fields) {
  const sh = ptSheet_(PT_SHEET, PT_HEAD);
  const range = sh.getRange(p.row, 1, 1, PT_HEAD.length);
  const r = range.getValues()[0];
  Object.keys(fields).forEach(function (k) { r[PC[k]] = fields[k]; });
  // שמירה כטקסט: טלפון (אפס מוביל) ושעת תור
  r[PC.PHONE] = "'" + ptPhone_(r[PC.PHONE]);
  const tm = ptTime_(r[PC.APPTTIME]);
  r[PC.APPTTIME] = tm ? "'" + tm : '';
  r[PC.UPDATED] = ptNow_();
  r[PC.UPDBY] = emp.name;
  range.setValues([r]);
}

function ptLog_(emp, id, action, details) {
  const sh = ptSheet_(PT_LOG, ['זמן', 'מזהה פנייה', 'עובד/ת', 'פעולה', 'פרטים']);
  sh.appendRow([ptNow_(), id, emp.name, action, details || '']);
}

function ptCallResult_(emp, p) {
  const x = ptFind_(p.id);
  if (x.status !== ST.CALL) throw new Error('הפנייה כבר טופלה. רעננו את הרשימה.');
  const note = ptClean_(p.note);
  const now = ptNow_();
  if (p.result === 'interested') {
    ptSet_(x, emp, { STATUS: ST.BOOK, CALLS: x.calls + 1, CALLRES: 'מעוניין', LASTCALL: now, NOTE: note, AVAIL: '' });
    ptLog_(emp, x.id, 'מעוניין – הועבר לקביעת תור', note);
  } else if (p.result === 'not_interested') {
    if (!note) throw new Error('יש לכתוב הסבר קצר (ללא מידע רפואי).');
    ptSet_(x, emp, { STATUS: ST.DOC, CALLS: x.calls + 1, CALLRES: 'לא מעוניין', LASTCALL: now, NOTE: note });
    ptLog_(emp, x.id, 'לא מעוניין – הועבר לרופאה', note);
  } else if (p.result === 'no_answer') {
    const n = x.calls + 1;
    if (n >= PT_MAX_CALLS) {
      ptSet_(x, emp, { STATUS: ST.DOC, CALLS: n, CALLRES: 'לא ענה', LASTCALL: now, NOTE: note || ('לא ענה ' + n + ' פעמים') });
      ptLog_(emp, x.id, 'לא ענה ' + n + ' פעמים – הועבר לרופאה', note);
    } else {
      ptSet_(x, emp, { CALLS: n, CALLRES: 'לא ענה', LASTCALL: now, AVAIL: ptNextWorkday_(), NOTE: note || x.note });
      ptLog_(emp, x.id, 'לא ענה – ניסיון ' + n, note);
    }
  } else if (p.result === 'done_internal') {
    if (x.type !== PT_INTERNAL) throw new Error('סימון "בוצע" מיועד לחידושי מרשמים בלבד.');
    ptSet_(x, emp, { STATUS: ST.DONE, CALLRES: 'בוצע (פנימי)', LASTCALL: now, NOTE: note });
    ptLog_(emp, x.id, 'חידוש מרשם בוצע', note);
  } else {
    throw new Error('תוצאה לא מוכרת');
  }
  return { ok: true };
}

function ptBookResult_(emp, p) {
  const x = ptFind_(p.id);
  if (x.status !== ST.BOOK) throw new Error('הפנייה כבר טופלה. רעננו את הרשימה.');
  const note = ptClean_(p.note);
  if (p.result === 'booked') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(p.date || ''))) throw new Error('יש לבחור תאריך תור.');
    if (!/^\d{2}:\d{2}$/.test(String(p.time || ''))) throw new Error('יש לבחור שעת תור.');
    ptSet_(x, emp, { STATUS: ST.BOOKED, BOOKTRIES: x.bookTries + 1, APPTDATE: p.date, APPTTIME: "'" + p.time, BOOKNOTE: note, AVAIL: '' });
    ptLog_(emp, x.id, 'תור נקבע ' + p.date + ' ' + p.time, note);
  } else if (p.result === 'not_reached') {
    const n = x.bookTries + 1;
    if (n >= PT_MAX_BOOK) {
      ptSet_(x, emp, { STATUS: ST.DOC, BOOKTRIES: n, BOOKNOTE: note || ('לא הושג ' + n + ' פעמים לקביעת תור') });
      ptLog_(emp, x.id, 'לא הושג ' + n + ' פעמים – הועבר לרופאה', note);
    } else {
      ptSet_(x, emp, { BOOKTRIES: n, AVAIL: ptNextWorkday_(), BOOKNOTE: note || x.bookNote });
      ptLog_(emp, x.id, 'לא הושג – ניסיון ' + n + ', חוזר ביום העבודה הבא', note);
    }
  } else if (p.result === 'declined') {
    if (!note) throw new Error('יש לכתוב הסבר קצר (ללא מידע רפואי).');
    ptSet_(x, emp, { STATUS: ST.DOC, BOOKTRIES: x.bookTries + 1, BOOKNOTE: note });
    ptLog_(emp, x.id, 'ויתר על התור – הועבר לרופאה', note);
  } else {
    throw new Error('תוצאה לא מוכרת');
  }
  return { ok: true };
}

function ptArrival_(emp, p) {
  const x = ptFind_(p.id);
  if (x.status !== ST.BOOKED) throw new Error('לפנייה זו אין תור פתוח.');
  if (p.arrived) {
    ptSet_(x, emp, { STATUS: ST.DONE, ARRIVAL: 'הגיע', VISITED: 'כן' });
    ptLog_(emp, x.id, 'המטופל הגיע לתור', '');
  } else {
    ptSet_(x, emp, { STATUS: ST.DOC, ARRIVAL: 'לא הגיע' });
    ptLog_(emp, x.id, 'המטופל לא הגיע – הועבר לרופאה', '');
  }
  return { ok: true };
}

function ptVisited_(emp, p) {
  const x = ptFind_(p.id);
  ptSet_(x, emp, { VISITED: p.value ? 'כן' : '' });
  ptLog_(emp, x.id, p.value ? 'סומן: כבר ביקר ברבעון' : 'בוטל סימון ביקור ברבעון', '');
  return { ok: true };
}

function ptDoctor_(emp, p) {
  const x = ptFind_(p.id);
  if (x.status !== ST.DOC) throw new Error('הפנייה כבר טופלה. רעננו את הרשימה.');
  const note = ptClean_(p.note);
  if (p.decision === 'retry') {
    const days = Math.max(1, Math.min(365, Number(p.days) || 30));
    const d = new Date(); d.setDate(d.getDate() + days);
    ptSet_(x, emp, { STATUS: ST.CALL, AVAIL: ptFmt_(d), CALLS: 0, BOOKTRIES: 0, DOCDEC: 'לנסות שוב בעוד ' + days + ' ימים', INSTR: note, ARRIVAL: '' });
    ptLog_(emp, x.id, 'רופאה: לנסות שוב בעוד ' + days + ' ימים', note);
  } else if (p.decision === 'instruct') {
    if (!note) throw new Error('יש לכתוב הנחיה לעמית.');
    ptSet_(x, emp, { STATUS: ST.CALL, AVAIL: '', CALLS: 0, BOOKTRIES: 0, DOCDEC: 'הנחיה לעמית', INSTR: note, ARRIVAL: '' });
    ptLog_(emp, x.id, 'רופאה: הנחיה לעמית', note);
  } else if (p.decision === 'book') {
    ptSet_(x, emp, { STATUS: ST.BOOK, AVAIL: '', BOOKTRIES: 0, DOCDEC: 'להעביר לקביעת תור', INSTR: note, ARRIVAL: '' });
    ptLog_(emp, x.id, 'רופאה: להעביר לקביעת תור', note);
  } else if (p.decision === 'close') {
    ptSet_(x, emp, { STATUS: ST.CLOSED, DOCDEC: 'סגור', INSTR: note });
    ptLog_(emp, x.id, 'רופאה: סגירת הפנייה', note);
  } else {
    throw new Error('החלטה לא מוכרת');
  }
  return { ok: true };
}

/** ייבוא רשימה: כל שורה "שם, טלפון" או "שם<TAB>טלפון" (אפשר גם סוג פנייה בעמודה שלישית) */
function ptImport_(emp, p) {
  const type = PT_TYPES.indexOf(p.type) >= 0 ? p.type : PT_TYPES[2];
  const listName = ptClean_(p.listName) || ('רשימה ' + ptToday_());
  const lines = String(p.text || '').split(/\r?\n/);
  const sh = ptSheet_(PT_SHEET, PT_HEAD);
  const existing = ptRows_();
  const phones = {};
  existing.forEach(function (x) { if (x.status !== ST.CLOSED && x.status !== ST.DONE) phones[x.phone] = true; });
  let seq = existing.reduce(function (m, x) { const n = Number(String(x.id).replace(/\D/g, '')); return n > m ? n : m; }, 0);
  const now = ptNow_();
  const rows = [];
  const skipped = [];
  lines.forEach(function (line, i) {
    line = line.trim();
    if (!line) return;
    const parts = line.split(/\t|,|;/).map(function (s) { return s.trim(); }).filter(Boolean);
    if (parts.length < 2) { skipped.push((i + 1) + ': חסר טלפון'); return; }
    let name = parts[0], phone = parts[1];
    if (/\d{7,}/.test(name.replace(/\D/g, '')) && !/\d/.test(phone)) { const t = name; name = phone; phone = t; }
    phone = phone.replace(/[^\d+]/g, '');
    if (phone.length < 9) { skipped.push((i + 1) + ': טלפון לא תקין'); return; }
    if (phones[phone]) { skipped.push((i + 1) + ': כבר קיים ברשימה פעילה'); return; }
    phones[phone] = true;
    const rowType = PT_TYPES.indexOf(parts[2]) >= 0 ? parts[2] : type;
    seq++;
    const r = new Array(PT_HEAD.length).fill('');
    r[PC.ID] = 'P-' + ('00000' + seq).slice(-5);
    r[PC.CREATED] = now;
    r[PC.NAME] = ptClean_(name).slice(0, 80);
    r[PC.PHONE] = "'" + phone;
    r[PC.TYPE] = rowType;
    r[PC.STATUS] = ST.CALL;
    r[PC.CALLS] = 0;
    r[PC.BOOKTRIES] = 0;
    r[PC.UPDATED] = now;
    r[PC.UPDBY] = emp.name;
    r[PC.LIST] = listName;
    rows.push(r);
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, PT_HEAD.length).setValues(rows);
  ptLog_(emp, '', 'ייבוא רשימה', listName + ': ' + rows.length + ' נוספו, ' + skipped.length + ' דולגו');
  return { added: rows.length, skipped: skipped };
}

/* ====================== עזר ====================== */

function ptSheet_(name, headers, init) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#fbe3e5');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    if (init) init(sh);
  }
  return sh;
}

function ptNow_() { return Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm'); }
function ptToday_() { return Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd'); }
function ptFmt_(d) { return Utilities.formatDate(d, 'Asia/Jerusalem', 'yyyy-MM-dd'); }
function ptDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return ptFmt_(v);
  const s = String(v).trim();
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : s;
}
function ptTime_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Jerusalem', 'HH:mm');
  return String(v).replace(/^'/, '').slice(0, 5);
}
function ptStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm');
  return String(v);
}
/** יום העבודה הבא (ראשון–חמישי) */
function ptNextWorkday_() {
  const d = new Date();
  do { d.setDate(d.getDate() + 1); } while (Number(Utilities.formatDate(d, 'Asia/Jerusalem', 'u')) >= 5 && Number(Utilities.formatDate(d, 'Asia/Jerusalem', 'u')) <= 6);
  return ptFmt_(d);
}
function ptPhone_(v) {
  let s = String(v == null ? '' : v).replace(/^'/, '').replace(/[^\d+]/g, '');
  if (/^[1-9]\d{7,8}$/.test(s)) s = '0' + s;
  return s;
}
function ptClean_(s) { return String(s || '').replace(/[<>]/g, '').trim().slice(0, 300); }
function ptPub_(p) { const o = {}; Object.keys(p).forEach(function (k) { if (k !== 'row') o[k] = p[k]; }); return o; }
