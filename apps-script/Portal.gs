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
  'עודכן', 'עודכן ע"י', 'רשימה', 'ת.ז. מטופל'];
const PC = { ID: 0, CREATED: 1, NAME: 2, PHONE: 3, TYPE: 4, STATUS: 5, AVAIL: 6, CALLS: 7, CALLRES: 8,
  NOTE: 9, LASTCALL: 10, BOOKTRIES: 11, APPTDATE: 12, APPTTIME: 13, BOOKNOTE: 14, DOCDEC: 15,
  INSTR: 16, ARRIVAL: 17, VISITED: 18, UPDATED: 19, UPDBY: 20, LIST: 21, TZ: 22 };

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
  const area = ptAreaApi_(emp, roles, op, payload);
  if (area) return area;
  switch (op) {
    case 'me': return { name: emp.name, roles: roles, types: PT_TYPES, counts: ptCounts_(), unread: ptMyArea_(emp).unread };
    case 'list': return ptList_(payload.view, roles);
    case 'callResult': need(ROLE.CALLER); return ptWrite_(emp, function () { return ptCallResult_(emp, payload); });
    case 'bookResult': need(ROLE.BOOKER); return ptWrite_(emp, function () { return ptBookResult_(emp, payload); });
    case 'arrival': need(ROLE.BOOKER); return ptWrite_(emp, function () { return ptArrival_(emp, payload); });
    case 'visited': return ptWrite_(emp, function () { return ptVisited_(emp, payload); });
    case 'doctorDecision': need(ROLE.DOCTOR); return ptWrite_(emp, function () { return ptDoctor_(emp, payload); });
    case 'addPatient': need(ROLE.CALLER); return ptWrite_(emp, function () { return ptAdd_(emp, payload); });
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
    list: String(r[PC.LIST] || ''), tz: String(r[PC.TZ] || '').replace(/\D/g, '')
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
  // כותרת העמודה החדשה (ת.ז.) בגיליון קיים
  if (sh.getLastColumn() < PT_HEAD.length) sh.getRange(1, PT_HEAD.length).setValue(PT_HEAD[PT_HEAD.length - 1]).setFontWeight('bold').setBackground('#fbe3e5');
  const phones = {};
  existing.forEach(function (x) {
    if (x.status !== ST.CLOSED && x.status !== ST.DONE) { if (x.phone) phones[x.phone] = true; if (x.tz) phones['tz' + x.tz] = true; }
  });
  let seq = existing.reduce(function (m, x) { const n = Number(String(x.id).replace(/\D/g, '')); return n > m ? n : m; }, 0);
  const now = ptNow_();
  const rows = [];
  const skipped = [];
  lines.forEach(function (line, i) {
    line = line.trim();
    if (!line) return;
    // שורה: שם + טלפון ו/או ת.ז. (9 ספרות, לא מתחיל ב-0) + סוג פנייה – בכל סדר, מופרד בפסיק/טאב/מקף
    line = line.replace(/^[-–•\s]+/, '');
    const parts = line.split(/\t|,|;|\s+[-–]\s+/).map(function (s) { return s.trim(); }).filter(Boolean);
    let name = '', phone = '', tz = '', rowType = type;
    parts.forEach(function (s) {
      const d = s.replace(/[^\d]/g, '');
      if (PT_TYPES.indexOf(s) >= 0) rowType = s;
      else if (/^[\d\s\-+()]+$/.test(s) && d.length === 9 && d[0] !== '0' && !tz) tz = d;
      else if (/^[\d\s\-+()]+$/.test(s) && d.length >= 9) phone = s.replace(/[^\d+]/g, '');
      else if (!name) {
        // "שם 123456789" באותו חלק
        const m = s.match(/^(.*?)\s+(\d{9})$/);
        if (m && m[2][0] !== '0') { name = m[1]; tz = m[2]; } else name = s;
      }
    });
    if (!name) { skipped.push((i + 1) + ': חסר שם'); return; }
    if (!phone && !tz) { skipped.push((i + 1) + ': חסר טלפון או ת.ז.'); return; }
    if ((phone && phones[phone]) || (tz && phones['tz' + tz])) { skipped.push((i + 1) + ': כבר קיים ברשימה פעילה'); return; }
    if (phone) phones[phone] = true;
    if (tz) phones['tz' + tz] = true;
    seq++;
    const r = new Array(PT_HEAD.length).fill('');
    r[PC.ID] = 'P-' + ('00000' + seq).slice(-5);
    r[PC.CREATED] = now;
    r[PC.NAME] = ptClean_(name).slice(0, 80);
    r[PC.PHONE] = phone ? "'" + phone : '';
    r[PC.TZ] = tz ? "'" + tz : '';
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

/** בדיקת ספרת ביקורת של ת.ז. ישראלית */
function ptTzValid_(tz) {
  tz = String(tz || '').replace(/\D/g, '');
  if (tz.length < 5 || tz.length > 9) return false;
  tz = ('000000000' + tz).slice(-9);
  let s = 0;
  for (let i = 0; i < 9; i++) { let d = Number(tz[i]) * ((i % 2) + 1); s += d > 9 ? d - 9 : d; }
  return s % 10 === 0;
}

/** עמית מוסיף מטופל שכבר דיבר איתו (או שצריך להתקשר אליו) – עם ת.ז., בלי חובת טלפון */
function ptAdd_(emp, p) {
  const name = ptClean_(p.name).slice(0, 80);
  const tzRaw = String(p.tz || '').replace(/\D/g, '');
  const phone = String(p.phone || '').replace(/[^\d+]/g, '');
  const type = PT_TYPES.indexOf(p.type) >= 0 ? p.type : PT_TYPES[0];
  const mode = ['interested', 'tocall', 'hasappt'].indexOf(p.mode) >= 0 ? p.mode : 'interested';
  if (!name) throw new Error('חסר שם מטופל.');
  if (!ptTzValid_(tzRaw)) throw new Error('מספר ת.ז. לא תקין. בדקו שוב את הספרות.');
  if (phone && phone.length < 9) throw new Error('מספר הטלפון לא תקין.');
  const tz = ('000000000' + tzRaw).slice(-9);
  const sh = ptSheet_(PT_SHEET, PT_HEAD);
  if (sh.getLastColumn() < PT_HEAD.length) sh.getRange(1, PT_HEAD.length).setValue(PT_HEAD[PT_HEAD.length - 1]).setFontWeight('bold').setBackground('#fbe3e5');
  const existing = ptRows_();
  const dup = existing.filter(function (x) { return x.tz === tz && x.status !== ST.CLOSED && x.status !== ST.DONE; })[0];
  if (dup) throw new Error('המטופל כבר קיים ברשימה פעילה (' + dup.id + ', ' + dup.status + ').');
  const seq = existing.reduce(function (m, x) { const n = Number(String(x.id).replace(/\D/g, '')); return n > m ? n : m; }, 0) + 1;
  const now = ptNow_();
  const note = ptClean_(p.note).slice(0, 300);
  const r = new Array(PT_HEAD.length).fill('');
  r[PC.ID] = 'P-' + ('00000' + seq).slice(-5);
  r[PC.CREATED] = now;
  r[PC.NAME] = name;
  r[PC.PHONE] = phone ? "'" + phone : '';
  r[PC.TZ] = "'" + tz;
  r[PC.TYPE] = type;
  r[PC.BOOKTRIES] = 0;
  if (mode === 'tocall') {
    r[PC.STATUS] = ST.CALL; r[PC.CALLS] = 0; r[PC.NOTE] = note;
  } else {
    r[PC.STATUS] = ST.BOOK; r[PC.CALLS] = 1; r[PC.CALLRES] = 'מעוניין'; r[PC.LASTCALL] = now;
    r[PC.NOTE] = mode === 'hasappt' ? ('כבר יש לו תור – לבצע בהגעה. רחלי: להזין את מועד התור ולסמן "תור נקבע".' + (note ? ' ' + note : '')) : note;
  }
  r[PC.UPDATED] = now;
  r[PC.UPDBY] = emp.name;
  r[PC.LIST] = 'נוסף ע"י ' + emp.name;
  sh.getRange(sh.getLastRow() + 1, 1, 1, PT_HEAD.length).setValues([r]);
  ptLog_(emp, r[PC.ID], 'הוספת מטופל', mode);
  return { id: r[PC.ID] };
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

/* ====================== אזור אישי: הודעות וצפי שעות שבועי ====================== */

const PL_SHEET = 'צפי שעות';
const PL_HEAD = ['שבוע (ראשון)', 'מזהה עובד', 'שם', 'יום', 'תאריך', 'משעה', 'עד שעה', 'הערה', 'הוגש'];
const MSG_SHEET = 'הודעות לעובדים';
const MSG_HEAD = ['מזהה', 'נשלח', 'מאת', 'אל (מזהה עובד / כולם)', 'נושא', 'תוכן', 'נקרא ע"י'];
const PL_DAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי'];

function ptAreaApi_(emp, roles, op, p) {
  const isMgr = roles.indexOf(ROLE.MANAGER) >= 0;
  switch (op) {
    case 'myArea': return ptMyArea_(emp);
    case 'savePlan': return ptWrite_(emp, function () { return ptSavePlan_(emp, p); });
    case 'markRead': return ptWrite_(emp, function () { return ptMarkRead_(emp, p.id); });
    case 'sendMessage':
      if (!isMgr) throw new Error('אין הרשאה לפעולה זו.');
      return ptWrite_(emp, function () { return ptSendMessage_(emp, p); });
    case 'teamArea':
      if (!isMgr) throw new Error('אין הרשאה לפעולה זו.');
      return ptTeamArea_(p.week);
  }
  return null;
}

/** ראשון של השבוע שמתחיל אחרי היום (במוצ"ש = מחר) */
function ptNextSunday_() {
  const d = new Date();
  do { d.setDate(d.getDate() + 1); } while (Utilities.formatDate(d, 'Asia/Jerusalem', 'u') !== '7');
  return ptFmt_(d);
}
function ptThisSunday_() {
  const d = new Date();
  while (Utilities.formatDate(d, 'Asia/Jerusalem', 'u') !== '7') d.setDate(d.getDate() - 1);
  return ptFmt_(d);
}
function ptAddDays_(ymd, n) {
  const p = ymd.split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2], 12);
  d.setDate(d.getDate() + n);
  return ptFmt_(d);
}

function ptPlanRows_() {
  const v = ptSheet_(PL_SHEET, PL_HEAD).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][1]) continue;
    out.push({ row: i + 1, week: ptDate_(v[i][0]), emp: String(v[i][1]), name: String(v[i][2]), day: String(v[i][3]),
      date: ptDate_(v[i][4]), from: ptTime_(v[i][5]), to: ptTime_(v[i][6]), note: String(v[i][7] || ''), submitted: ptStr_(v[i][8]) });
  }
  return out;
}

function ptPlanFor_(empId, week) {
  return ptPlanRows_().filter(function (r) { return r.emp === empId && r.week === week; })
    .map(function (r) { const o = ptPub_(r); return o; });
}

function ptMyArea_(emp) {
  const next = ptNextSunday_(), cur = ptThisSunday_();
  const msgs = ptMessages_().filter(function (m) { return m.to === 'כולם' || m.to === emp.id; })
    .map(function (m) { return { id: m.id, sent: m.sent, from: m.from, subject: m.subject, body: m.body, read: m.readBy.indexOf(emp.id) >= 0 }; })
    .reverse().slice(0, 30);
  return {
    name: emp.name, nextWeek: next, curWeek: cur,
    planNext: ptPlanFor_(emp.id, next), planCur: ptPlanFor_(emp.id, cur),
    needsPlan: !emp.admin && !emp.contractor,
    messages: msgs, unread: msgs.filter(function (m) { return !m.read; }).length
  };
}

function ptSavePlan_(emp, p) {
  const week = String(p.week || '');
  if ([ptNextSunday_(), ptThisSunday_()].indexOf(week) < 0) throw new Error('אפשר להגיש צפי רק לשבוע הנוכחי או לשבוע הבא.');
  const days = (p.days || []).filter(function (d) { return d && d.on; });
  const rows = [];
  days.forEach(function (d) {
    const idx = Number(d.idx);
    if (!(idx >= 0 && idx < PL_DAYS.length)) return;
    if (!/^\d{2}:\d{2}$/.test(d.from || '') || !/^\d{2}:\d{2}$/.test(d.to || '')) throw new Error('יש למלא שעת התחלה וסיום ליום ' + PL_DAYS[idx] + '.');
    if (d.to <= d.from) throw new Error('שעת הסיום ביום ' + PL_DAYS[idx] + ' צריכה להיות אחרי שעת ההתחלה.');
    rows.push([week, emp.id, emp.name, PL_DAYS[idx], ptAddDays_(week, idx), "'" + d.from, "'" + d.to, ptClean_(d.note), ptNow_()]);
  });
  const sh = ptSheet_(PL_SHEET, PL_HEAD);
  const old = ptPlanRows_().filter(function (r) { return r.emp === emp.id && r.week === week; })
    .map(function (r) { return r.row; }).sort(function (a, b) { return b - a; });
  old.forEach(function (r) { sh.deleteRow(r); });
  if (!rows.length) rows.push([week, emp.id, emp.name, '—', '', '', '', 'לא מתוכננת עבודה השבוע', ptNow_()]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, PL_HEAD.length).setValues(rows);
  ptLog_(emp, '', 'הוגש צפי שעות', 'שבוע ' + week + ': ' + days.length + ' ימים');
  return { ok: true };
}

function ptMessages_() {
  const v = ptSheet_(MSG_SHEET, MSG_HEAD).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    out.push({ row: i + 1, id: String(v[i][0]), sent: ptStr_(v[i][1]), from: String(v[i][2]), to: String(v[i][3]),
      subject: String(v[i][4] || ''), body: String(v[i][5] || ''), readBy: String(v[i][6] || '').split(',').filter(Boolean) });
  }
  return out;
}

function ptMarkRead_(emp, id) {
  const m = ptMessages_().filter(function (x) { return x.id === String(id); })[0];
  if (!m || (m.to !== 'כולם' && m.to !== emp.id)) throw new Error('ההודעה לא נמצאה.');
  if (m.readBy.indexOf(emp.id) < 0) {
    m.readBy.push(emp.id);
    ptSheet_(MSG_SHEET, MSG_HEAD).getRange(m.row, 7).setValue(m.readBy.join(','));
  }
  return { ok: true };
}

function ptSendMessage_(emp, p) {
  const subject = ptClean_(p.subject), body = String(p.body || '').replace(/[<>]/g, '').trim().slice(0, 2000);
  if (!subject || !body) throw new Error('יש למלא נושא ותוכן.');
  const to = p.to === 'כולם' ? 'כולם' : String(p.to || '');
  if (to !== 'כולם' && !readEmps_().some(function (x) { return x.id === to && x.active; })) throw new Error('עובד לא נמצא.');
  const sh = ptSheet_(MSG_SHEET, MSG_HEAD);
  const id = 'M-' + Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyMMddHHmmss');
  sh.appendRow([id, ptNow_(), emp.name, to, subject, body, '']);
  return { ok: true, id: id };
}

function ptTeamArea_(week) {
  const next = ptNextSunday_(), cur = ptThisSunday_();
  week = [next, cur].indexOf(week) >= 0 ? week : next;
  const emps = readEmps_().filter(function (e) { return e.active && !/בדיקה/.test(e.name); });
  const plans = ptPlanRows_().filter(function (r) { return r.week === week; });
  const msgs = ptMessages_();
  return {
    week: week, nextWeek: next, curWeek: cur,
    people: emps.map(function (e) { return { id: e.id, name: e.name, role: e.role, needsPlan: !e.admin && !e.contractor }; }),
    plans: plans.map(ptPub_),
    messages: msgs.slice(-20).reverse().map(function (m) {
      return { id: m.id, sent: m.sent, to: m.to, subject: m.subject, readBy: m.readBy.length };
    })
  };
}
