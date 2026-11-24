/*************************************************************
 *  NAFES SCIENCE QUEST — Google Sheets backend
 *  by Miss Basma Ahmed · Andalus International Schools
 *
 *  1) غيّري كلمة سر المعلمة هنا قبل النشر
 *  2) Deploy > New deployment > Web app
 *     Execute as: Me   |   Who has access: Anyone
 *************************************************************/
const TEACHER_PASSWORD = 'Basma2026';          // كلمة سر واجهة المعلمة
const SECRET = 'nafes-' + 'change-this-to-any-random-words-123';

const WEEK_LABELS = ['Week 1','Week 2','Week 3','Week 4','Week 5','Week 6','Week 7',
  'Week 8','Week 9','Week 10','Week 11','Grand Review','Mock Exam 1','Mock Exam 2'];
const PASS = 0.6;

const SH = {
  students: 'Students',
  attempts: 'Attempts',
  best: 'Best Scores',
  first: 'First Attempts',
  questions: 'Questions'
};
const HEAD = {
  students: ['Name','Password Hash','Salt','Joined','Last Active','Key','Coins','Stars','Profile'],
  attempts: ['Timestamp','Student','Week','Mission','Score','Total','Percent','Attempt #',
             'First Attempt?','Duration (sec)','Max Streak','Hints Used','Wrong Question IDs','Corrected IDs','Week Name'],
  summary: ['Student'].concat(WEEK_LABELS).concat(['Weeks Passed','Total Attempts','Last Active']),
  questions: ['ID','Week','Domain','Lesson','Question','A','B','C','D','Correct (A-D)']
};

/* ---------------- entry points ---------------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.q) {
    let out;
    try { out = route_(JSON.parse(e.parameter.q)); }
    catch (err) { out = { ok: false, error: String(err && err.message ? err.message : err) }; }
    return json_(out);
  }
  return json_({ ok: true, msg: 'Nafes Science Quest API is running ✔' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    out = route_(req);
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function route_(r) {
  setup();
  switch (r.action) {
    case 'ping':          return { ok: true };
    case 'register':      return register_(r);
    case 'login':         return login_(r);
    case 'submit':        return submit_(r);
    case 'sync':          { const st = auth_(r); return { ok: true, attempts: studentAttempts_(r.name), profile: readProfile_(st.row) }; }
    case 'saveProfile':   return saveProfile_(r);
    case 'getQuestions':  return getQuestions_();
    case 'teacherData':   return teacherData_(r);
    case 'resetPassword': return resetPassword_(r);
    case 'saveQuestions': return saveQuestions_(r);
    default:              return { ok: false, error: 'Unknown action' };
  }
}

/* ---------------- setup (runs automatically) ---------------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  makeSheet_(ss, SH.students, HEAD.students, '#5B3FD9');
  makeSheet_(ss, SH.attempts, HEAD.attempts, '#14B8A6');
  const best = makeSheet_(ss, SH.best, HEAD.summary, '#FF6B8B');
  const first = makeSheet_(ss, SH.first, HEAD.summary, '#4CC9F0');
  makeSheet_(ss, SH.questions, HEAD.questions, '#FFC93C');
  const stu = ss.getSheetByName(SH.students);
  if (stu.getRange(1, 7).getValue() !== 'Coins') {
    stu.getRange(1, 7, 1, 3).setValues([['Coins', 'Stars', 'Profile']]).setFontWeight('bold')
      .setBackground('#5B3FD9').setFontColor('#ffffff').setHorizontalAlignment('center');
    stu.hideColumns(9);
  }
  [best, first].forEach(function (s) {
    if (s.getConditionalFormatRules().length) return;
    const rng = s.getRange(2, 2, 500, WEEK_LABELS.length);
    rng.setNumberFormat('0%').setHorizontalAlignment('center');
    s.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(B2),B2<0.6)')
        .setBackground('#FFD6DE').setFontColor('#B4002C').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(B2),B2>=0.6,B2<0.8)')
        .setBackground('#FFF1C2').setFontColor('#8A5A00').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(B2),B2>=0.8)')
        .setBackground('#D4F7E9').setFontColor('#05704C').setRanges([rng]).build()
    ]);
  });
}

function makeSheet_(ss, name, head, color) {
  let s = ss.getSheetByName(name);
  if (s) return s;
  s = ss.insertSheet(name);
  s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold')
    .setBackground(color).setFontColor('#ffffff').setHorizontalAlignment('center');
  s.setFrozenRows(1);
  s.setFrozenColumns(name === SH.questions ? 0 : 1);
  s.setColumnWidth(1, 160);
  if (name === SH.students) { s.hideColumns(2, 2); s.hideColumns(6); }
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  return s;
}

/* ---------------- helpers ---------------- */
function key_(name) { return String(name || '').trim().replace(/\s+/g, ' ').toLowerCase(); }
function hash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8));
}
function token_(name, passHash) { return hash_(key_(name) + '|' + passHash + '|' + SECRET).slice(0, 32); }

function findStudent_(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(SH.students);
  const n = s.getLastRow();
  if (n < 2) return null;
  const vals = s.getRange(2, 1, n - 1, 6).getValues();
  const k = key_(name);
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][5] === k || key_(vals[i][0]) === k) {
      return { row: i + 2, name: vals[i][0], hash: vals[i][1], salt: vals[i][2] };
    }
  }
  return null;
}

function studentAttempts_(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(SH.attempts);
  const n = s.getLastRow();
  if (n < 2) return [];
  const k = key_(name);
  return s.getRange(2, 1, n - 1, HEAD.attempts.length).getValues()
    .filter(function (r) { return key_(r[1]) === k; })
    .map(rowToAttempt_);
}

function rowToAttempt_(r) {
  return {
    ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    student: r[1], week: Number(r[2]), mission: Number(r[3]), score: Number(r[4]), total: Number(r[5]),
    attemptNo: Number(r[7]), dur: Number(r[9]) || 0, streak: Number(r[10]) || 0, hints: Number(r[11]) || 0,
    wrong: r[12] ? String(r[12]).split(',').filter(String) : [],
    fixed: r[13] ? String(r[13]).split(',').filter(String) : []
  };
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- students ---------------- */
function register_(r) {
  const name = String(r.name || '').trim().replace(/\s+/g, ' ');
  const pass = String(r.pass || '');
  if (name.length < 2) return { ok: false, error: 'Please write your name.' };
  if (pass.length < 4) return { ok: false, error: 'Password must be at least 4 characters.' };
  return withLock_(function () {
    if (findStudent_(name)) return { ok: false, error: 'This name is already registered. Try logging in, or add your last name.' };
    const salt = Utilities.getUuid().slice(0, 8);
    const h = hash_(salt + pass);
    const now = new Date();
    SpreadsheetApp.getActive().getSheetByName(SH.students).appendRow([name, h, salt, now, now, key_(name)]);
    return { ok: true, name: name, token: token_(name, h), attempts: [] };
  });
}

function login_(r) {
  const st = findStudent_(r.name);
  if (!st) return { ok: false, error: 'Name not found. Are you a new student? Tap "New student".' };
  if (hash_(st.salt + String(r.pass || '')) !== st.hash) return { ok: false, error: 'Wrong password. Ask Miss Basma if you forgot it.' };
  SpreadsheetApp.getActive().getSheetByName(SH.students).getRange(st.row, 5).setValue(new Date());
  return { ok: true, name: st.name, token: token_(st.name, st.hash), attempts: studentAttempts_(st.name), profile: readProfile_(st.row) };
}

/* ---------------- coins, stars & profile (synced across devices) ---------------- */
function readProfile_(row) {
  const v = SpreadsheetApp.getActive().getSheetByName(SH.students).getRange(row, 9).getValue();
  if (!v) return null;
  try { return JSON.parse(v); } catch (e) { return null; }
}
function saveProfile_(r) {
  const st = auth_(r);
  const prof = JSON.stringify(r.profile || {});
  if (prof.length > 45000) return { ok: false, error: 'Profile too large' };
  return withLock_(function () {
    SpreadsheetApp.getActive().getSheetByName(SH.students).getRange(st.row, 7, 1, 3)
      .setValues([[Number(r.coins) || 0, Number(r.stars) || 0, prof]]);
    return { ok: true };
  });
}

function auth_(r) {
  const st = findStudent_(r.name);
  if (!st || token_(st.name, st.hash) !== r.token) throw new Error('Session expired. Please log in again.');
  return st;
}

/* ---------------- results ---------------- */
function submit_(r) {
  const st = auth_(r);
  const a = r.attempt || {};
  return withLock_(function () {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SH.attempts);
    const mine = studentAttempts_(st.name);
    const prev = mine.filter(function (x) { return x.week === Number(a.week) && x.mission === Number(a.mission); }).length;
    const total = Number(a.total) || 0, score = Number(a.score) || 0;
    sh.appendRow([new Date(), st.name, Number(a.week), Number(a.mission), score, total,
      total ? score / total : 0, prev + 1, prev === 0 ? 'Yes' : 'No', Number(a.dur) || 0,
      Number(a.streak) || 0, Number(a.hints) || 0, (a.wrong || []).join(','), (a.fixed || []).join(','), a.weekName || '']);
    sh.getRange(sh.getLastRow(), 7).setNumberFormat('0%');
    if (r.missionCounts) PropertiesService.getScriptProperties().setProperty('missionCounts', JSON.stringify(r.missionCounts));
    updateSummary_(st.name);
    ss.getSheetByName(SH.students).getRange(st.row, 5).setValue(new Date());
    return { ok: true, attemptNo: prev + 1 };
  });
}

function updateSummary_(name) {
  const ss = SpreadsheetApp.getActive();
  const atts = studentAttempts_(name);
  let mc = {};
  try { mc = JSON.parse(PropertiesService.getScriptProperties().getProperty('missionCounts') || '{}'); } catch (e) {}
  const bestRow = [name], firstRow = [name], bestNotes = [''], firstNotes = [''];
  let passed = 0;
  for (let w = 1; w <= WEEK_LABELS.length; w++) {
    const per = {};
    atts.filter(function (a) { return a.week === w; }).forEach(function (a) {
      const m = per[a.mission] || (per[a.mission] = { first: null, best: null });
      if (!m.first) m.first = a;
      if (!m.best || a.score / a.total > m.best.score / m.best.total) m.best = a;
    });
    const ms = Object.keys(per);
    if (!ms.length) { bestRow.push(''); firstRow.push(''); bestNotes.push(''); firstNotes.push(''); continue; }
    let bs = 0, bt = 0, fs = 0, ft = 0, ok = 0;
    ms.forEach(function (k) {
      bs += per[k].best.score; bt += per[k].best.total; fs += per[k].first.score; ft += per[k].first.total;
      if (per[k].best.score / per[k].best.total >= PASS) ok++;
    });
    const need = Number(mc[w]) || ms.length;
    if (ok >= need) passed++;
    bestRow.push(bt ? bs / bt : ''); firstRow.push(ft ? fs / ft : '');
    const note = ms.length + ' of ' + need + ' missions done';
    bestNotes.push(note); firstNotes.push(note);
  }
  const last = new Date();
  bestRow.push(passed, atts.length, last); firstRow.push(passed, atts.length, last);
  writeSummaryRow_(ss.getSheetByName(SH.best), name, bestRow, bestNotes);
  writeSummaryRow_(ss.getSheetByName(SH.first), name, firstRow, firstNotes);
}

function writeSummaryRow_(s, name, row, notes) {
  const n = s.getLastRow();
  let at = n + 1;
  if (n >= 2) {
    const names = s.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) if (key_(names[i][0]) === key_(name)) { at = i + 2; break; }
  }
  s.getRange(at, 1, 1, row.length).setValues([row]);
  s.getRange(at, 1, 1, notes.length).setNotes([notes]);
  s.getRange(at, 2, 1, WEEK_LABELS.length).setNumberFormat('0%');
  s.getRange(at, row.length).setNumberFormat('dd/mm/yyyy hh:mm');
}

/* ---------------- questions ---------------- */
function getQuestions_() {
  const s = SpreadsheetApp.getActive().getSheetByName(SH.questions);
  const n = s.getLastRow();
  if (n < 2) return { ok: true, questions: [] };
  const L = { A: 0, B: 1, C: 2, D: 3 };
  const qs = s.getRange(2, 1, n - 1, 10).getValues()
    .filter(function (r) { return r[0] && r[4] && r[5] && r[6]; })
    .map(function (r) {
      const o = [r[5], r[6], r[7], r[8]].map(String).filter(function (x) { return x.trim() !== ''; });
      return { id: String(r[0]), w: Number(r[1]) || 1, d: String(r[2]).toLowerCase(), l: String(r[3]), q: String(r[4]),
               o: o, a: L[String(r[9]).trim().toUpperCase()] || 0 };
    });
  return { ok: true, questions: qs };
}

function saveQuestions_(r) {
  if (r.teacherPass !== TEACHER_PASSWORD) return { ok: false, error: 'Wrong teacher password.' };
  const s = SpreadsheetApp.getActive().getSheetByName(SH.questions);
  const rows = (r.questions || []).map(function (q) {
    const o = q.o.concat(['', '', '', '']).slice(0, 4);
    return [q.id, q.w, q.d, q.l, q.q, o[0], o[1], o[2], o[3], 'ABCD'[q.a]];
  });
  return withLock_(function () {
    if (!r.append && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 10).clearContent();
    const start = r.append ? Math.max(2, s.getLastRow() + 1) : 2;
    if (rows.length) s.getRange(start, 1, rows.length, 10).setValues(rows);
    s.setColumnWidth(5, 420); s.setColumnWidths(6, 4, 180);
    return { ok: true, count: rows.length, total: s.getLastRow() - 1 };
  });
}

/* ---------------- teacher ---------------- */
function teacherData_(r) {
  if (r.teacherPass !== TEACHER_PASSWORD) return { ok: false, error: 'Wrong teacher password.' };
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName(SH.students);
  const students = st.getLastRow() < 2 ? [] : st.getRange(2, 1, st.getLastRow() - 1, 9).getValues().map(function (v) {
    let av = 0; try { av = v[8] ? (JSON.parse(v[8]).avatar || 0) : 0; } catch (e) {}
    return { name: v[0], joined: v[3] instanceof Date ? v[3].toISOString() : '', last: v[4] instanceof Date ? v[4].toISOString() : '',
             coins: Number(v[6]) || 0, stars: Number(v[7]) || 0, avatar: av };
  });
  const at = ss.getSheetByName(SH.attempts);
  const attempts = at.getLastRow() < 2 ? [] : at.getRange(2, 1, at.getLastRow() - 1, HEAD.attempts.length).getValues().map(rowToAttempt_);
  return { ok: true, students: students, attempts: attempts, sheetUrl: ss.getUrl() };
}

function resetPassword_(r) {
  if (r.teacherPass !== TEACHER_PASSWORD) return { ok: false, error: 'Wrong teacher password.' };
  const st = findStudent_(r.name);
  if (!st) return { ok: false, error: 'Student not found.' };
  if (String(r.newPass || '').length < 4) return { ok: false, error: 'New password must be at least 4 characters.' };
  const salt = Utilities.getUuid().slice(0, 8);
  SpreadsheetApp.getActive().getSheetByName(SH.students).getRange(st.row, 2, 1, 2).setValues([[hash_(salt + r.newPass), salt]]);
  return { ok: true };
}
