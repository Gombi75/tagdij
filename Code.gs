/**
 * VESC tagdíjak – Google Sheets + Apps Script backend
 *
 * A Google Sheethez kötött szkript (Bővítmények → Apps Script).
 * Lapok: Csomagok, Tagok, Befizetések (a setup() hozza létre őket).
 *
 * 1) Írd át a KEZDO_JELSZO értékét, mentsd, majd futtasd egyszer a setup() függvényt.
 * 2) Telepítés → Új telepítés → Webalkalmazás
 *    – Végrehajtás mint: Én (vesctriatlon@gmail.com)
 *    – Hozzáférés: Bárki
 * 3) A kapott /exec URL-t másold az index.html API_URL sorába.
 */

const KEZDO_JELSZO = 'valtoztasd-meg';   // a setup() ezt menti el jelszónak
const SH = { P: 'Csomagok', M: 'Tagok', B: 'Befizetések', L: 'Napló' };
const L_HEAD = ['Időpont', 'Ki', 'Művelet', 'Tag', 'Részletek', 'Előtte', 'Utána'];
const HO = ['jan','febr','márc','ápr','máj','jún','júl','aug','szep','okt','nov','dec'];
// Tag-adatlap mezői: kulcs → oszlopnév a Tagok lapon (hiányzó oszlopot a szkript magától létrehoz)
const FIELDS = { email: 'E-mail', phone: 'Telefon', birth: 'Születési dátum', joined: 'Belépés', left: 'Kilépés',
                 license: 'Licencszám', shirt: 'Pólóméret', note: 'Megjegyzés' };
const DATE_FIELDS = ['birth', 'joined', 'left'];
const B_HEAD = ['ID', 'TagID', 'Tag neve', 'Összeg', 'Dátum', 'Év', 'Tól (hó)', 'Ig (hó)', 'Mód', 'Megjegyzés', 'Rögzítve'];

/* ---------- web API ---------- */
function doGet(e) {
  return out_(() => {
    const p = e.parameter;
    check_(p.key, p.by);
    if (p.action === 'log') return readLog_(Number(p.n) || 1000);
    return all_();
  });
}
function doPost(e) {
  return out_(() => {
    const b = JSON.parse(e.postData.contents || '{}');
    check_(b.key, b.by);
    const by = String(b.by || '').trim() || 'ismeretlen';
    if (b.action === 'login') { log_(by, 'Belépés', '', '', '', ''); return true; }
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const before = all_();
      const r = act_(b);
      SpreadsheetApp.flush();
      logDiff_(by, before, all_());               // minden változás a Napló lapra
      return r;
    } finally { lock.releaseLock(); }
  });
}
function out_(fn) {
  let res;
  try { res = { ok: true, data: fn() }; }
  catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}
function check_(k, by) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!key) throw new Error('Nincs beállítva jelszó – futtasd a setup() függvényt.');
  if (String(k || '') !== key) {
    try { log_(String(by || '').trim() || 'ismeretlen', 'Sikertelen belépés', '', 'Hibás jelszóval próbálkozott', '', ''); } catch (_) {}
    throw new Error('Hibás jelszó');
  }
}

/* ---------- napló (history) ---------- */
function logSheet_() {
  let sh = ss_().getSheetByName(SH.L);
  if (!sh) {
    sh = ss_().insertSheet(SH.L);
    sh.appendRow(L_HEAD); sh.getRange(1, 1, 1, L_HEAD.length).setFontWeight('bold'); sh.setFrozenRows(1);
    sh.setColumnWidth(5, 420);
  }
  return sh;
}
function log_(by, action, member, details, before, after) {
  const sh = logSheet_();
  sh.appendRow([new Date(), by, action, member, details, before, after]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}
function readLog_(n) {
  const sh = logSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  const from = Math.max(2, last - n + 1), tz = ss_().getSpreadsheetTimeZone();
  return sh.getRange(from, 1, last - from + 1, 5).getValues().reverse().map(r => ({
    at: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm:ss') : String(r[0]),
    by: String(r[1]), action: String(r[2]), member: String(r[3]), details: String(r[4])
  }));
}
const ft_ = n => Math.round(Number(n) || 0).toLocaleString('hu-HU').replace(/,/g, ' ') + ' Ft';
const per_ = p => p.from === p.to ? HO[p.from - 1] : HO[p.from - 1] + '–' + HO[p.to - 1];
const payDesc_ = p => ft_(p.amount) + ', ' + p.year + ' ' + per_(p) + (p.date ? ', befizetve ' + p.date : '') + (p.method ? ', ' + p.method : '') + (p.note ? ', „' + p.note + '”' : '');
const LBL = Object.assign({ name: 'Név', active: 'Aktív' }, FIELDS);
const show_ = (k, v) => k === 'active' ? (v ? 'igen' : 'nem') : (v === '' || v == null ? '–' : String(v));
function logDiff_(by, A, B) {
  const nm = {}; A.members.concat(B.members).forEach(m => nm[m.id] = m.name);
  const J = o => JSON.stringify(o);
  // tagok
  const am = {}, bm = {}; A.members.forEach(m => am[m.id] = m); B.members.forEach(m => bm[m.id] = m);
  Object.keys(bm).forEach(id => {
    const a = am[id], b = bm[id];
    if (!a) {
      const d = Object.keys(LBL).filter(k => k !== 'name' && b[k] !== '' && b[k] != null).map(k => LBL[k] + ': ' + show_(k, b[k]));
      Object.keys(b.pkgs).forEach(y => d.push('Csomag ' + y + ': ' + b.pkgs[y]));
      log_(by, 'Új tag', b.name, d.join('; '), '', J(b)); return;
    }
    const d = [];
    Object.keys(LBL).forEach(k => { if (show_(k, a[k]) !== show_(k, b[k])) d.push(LBL[k] + ': ' + show_(k, a[k]) + ' → ' + show_(k, b[k])); });
    new Set(Object.keys(a.pkgs).concat(Object.keys(b.pkgs))).forEach(y => { if ((a.pkgs[y] || '') !== (b.pkgs[y] || '')) d.push('Csomag ' + y + ': ' + (a.pkgs[y] || '–') + ' → ' + (b.pkgs[y] || '–')); });
    if (d.length) log_(by, 'Tagadat módosítva', b.name, d.join('; '), J(a), J(b));
  });
  Object.keys(am).forEach(id => { if (!bm[id]) log_(by, 'Tag törölve', am[id].name, '', J(am[id]), ''); });
  // befizetések
  const ap = {}, bp = {}; A.payments.forEach(p => ap[p.id] = p); B.payments.forEach(p => bp[p.id] = p);
  Object.keys(bp).forEach(id => {
    const a = ap[id], b = bp[id];
    if (!a) log_(by, 'Új befizetés', nm[b.memberId] || b.memberId, payDesc_(b), '', J(b));
    else if (J(a) !== J(b)) log_(by, 'Befizetés módosítva', nm[b.memberId] || b.memberId, payDesc_(a) + ' → ' + payDesc_(b), J(a), J(b));
  });
  Object.keys(ap).forEach(id => { if (!bp[id]) log_(by, 'Befizetés törölve', nm[ap[id].memberId] || ap[id].memberId, payDesc_(ap[id]), J(ap[id]), ''); });
  // csomagok
  const ak = {}, bk = {}; A.packages.forEach(p => ak[p.name] = p.fee); B.packages.forEach(p => bk[p.name] = p.fee);
  Object.keys(bk).forEach(n => {
    if (!(n in ak)) log_(by, 'Új csomag', '', n + ': ' + ft_(bk[n]) + '/év', '', '');
    else if (ak[n] !== bk[n]) log_(by, 'Csomagdíj módosítva', '', n + ': ' + ft_(ak[n]) + ' → ' + ft_(bk[n]), '', '');
  });
  Object.keys(ak).forEach(n => { if (!(n in bk)) log_(by, 'Csomag törölve', '', n + ': ' + ft_(ak[n]) + '/év', '', ''); });
}

/* Kézi szerkesztés a táblázatban is a naplóba kerül */
function onEdit(e) {
  try {
    const sh = e.range.getSheet(), name = sh.getName();
    if (![SH.P, SH.M, SH.B].includes(name)) return;
    const who = (Session.getActiveUser().getEmail() || 'táblázat') + ' (kézi)';
    const cells = e.range.getNumRows() * e.range.getNumColumns();
    const head = String(sh.getRange(1, e.range.getColumn()).getValue());
    const member = name === SH.P ? '' : String(sh.getRange(e.range.getRow(), name === SH.M ? 2 : 3).getValue());
    const det = name + '!' + e.range.getA1Notation() + (cells === 1 ? ' (' + head + '): ' + (e.oldValue == null ? '–' : e.oldValue) + ' → ' + (e.value == null ? '–' : e.value) : ' (' + cells + ' cella)');
    log_(who, 'Kézi szerkesztés', member, det, '', '');
  } catch (_) {}
}

/* ---------- segédek ---------- */
const ss_ = () => SpreadsheetApp.getActiveSpreadsheet();
const sh_ = n => ss_().getSheetByName(n) || (() => { throw new Error('Hiányzó lap: ' + n); })();
const id_ = () => Utilities.getUuid().replace(/-/g, '').slice(0, 10);
function day_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}
function table_(name) {
  const sh = sh_(name), v = sh.getDataRange().getValues();
  return { sh, head: v[0].map(String), rows: v.slice(1) };
}
function rowOf_(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  throw new Error('Nem található: ' + id);
}
function yearCol_(sh, year) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  let c = head.indexOf(String(year));
  if (c < 0) { c = head.length; sh.getRange(1, c + 1).setValue(String(year)).setFontWeight('bold'); }
  return c + 1;
}
function col_(sh, name) {            // oszlop sorszáma fejléc alapján, ha nincs, létrehozza
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  let c = head.indexOf(name);
  if (c < 0) { c = head.length; sh.getRange(1, c + 1).setValue(name).setFontWeight('bold'); }
  return c + 1;
}
function setFields_(sh, r, f) {
  Object.keys(FIELDS).forEach(k => {
    if (!f || f[k] == null) return;
    const c = col_(sh, FIELDS[k]), cell = sh.getRange(r, c), v = String(f[k]).trim();
    if (DATE_FIELDS.includes(k)) cell.setNumberFormat('yyyy-mm-dd');
    else cell.setNumberFormat('@');                 // telefonszám, licencszám szövegként maradjon
    cell.setValue(v);
  });
}
const isYes_ = v => !(v === false || /^(nem|no|false|0|inaktív)$/i.test(String(v).trim()));
function memberName_(id) {
  const t = table_(SH.M), r = t.rows.find(x => String(x[0]) === String(id));
  if (!r) throw new Error('Ismeretlen tag');
  return String(r[1]);
}

/* ---------- olvasás ---------- */
function all_() {
  const p = table_(SH.P), m = table_(SH.M), b = table_(SH.B);
  const packages = p.rows.filter(r => r[0] !== '').map(r => ({ name: String(r[0]), fee: Number(r[1]) || 0 }));
  const yearCols = m.head.map((h, i) => /^\d{4}$/.test(h) ? i : -1).filter(i => i >= 0);
  const members = m.rows.filter(r => r[0] !== '').map(r => {
    const pkgs = {};
    yearCols.forEach(i => { if (r[i] !== '') pkgs[m.head[i]] = String(r[i]); });
    const f = {};
    Object.keys(FIELDS).forEach(k => { const i = m.head.indexOf(FIELDS[k]); f[k] = i < 0 ? '' : (DATE_FIELDS.includes(k) ? day_(r[i]) : String(r[i])); });
    return Object.assign({ id: String(r[0]), name: String(r[1]), active: isYes_(r[2]), pkgs }, f);
  });
  const payments = b.rows.filter(r => r[0] !== '').map(r => ({
    id: String(r[0]), memberId: String(r[1]), amount: Number(r[3]) || 0, date: day_(r[4]),
    year: Number(r[5]) || 0, from: Number(r[6]) || 1, to: Number(r[7]) || Number(r[6]) || 1,
    method: String(r[8] || ''), note: String(r[9] || '')
  }));
  return { packages, members, payments };
}

/* ---------- írás ---------- */
function payRow_(id, d) {
  const amount = Math.round(Number(d.amount)), from = Number(d.from), to = Number(d.to), year = Number(d.year);
  if (!(amount > 0)) throw new Error('Érvénytelen összeg');
  if (!(from >= 1 && to <= 12 && from <= to)) throw new Error('Érvénytelen időszak');
  if (!(year > 2000)) throw new Error('Érvénytelen év');
  return [id, d.memberId, memberName_(d.memberId), amount, String(d.date || ''), year, from, to,
          String(d.method || ''), String(d.note || ''), new Date()];
}
function act_(b) {
  switch (b.action) {
    case 'addPayment': {
      const id = id_(), sh = sh_(SH.B);
      sh.appendRow(payRow_(id, b.data || {}));
      sh.getRange(sh.getLastRow(), 5).setNumberFormat('yyyy-mm-dd');
      return id;
    }
    case 'updatePayment': {
      const sh = sh_(SH.B), r = rowOf_(sh, b.id);
      sh.getRange(r, 1, 1, B_HEAD.length).setValues([payRow_(b.id, b.data || {})]);
      return b.id;
    }
    case 'deletePayment': {
      const sh = sh_(SH.B); sh.deleteRow(rowOf_(sh, b.id)); return true;
    }
    case 'addMember': {
      const name = String(b.name || '').trim(); if (!name) throw new Error('Hiányzó név');
      const sh = sh_(SH.M), id = id_();
      if (table_(SH.M).rows.some(x => String(x[1]).trim().toLowerCase() === name.toLowerCase())) throw new Error('Már van ilyen nevű tag: ' + name);
      sh.appendRow([id, name, b.active === false ? 'nem' : 'igen']);
      const r = sh.getLastRow();
      if (b.pkg) sh.getRange(r, yearCol_(sh, b.year)).setValue(b.pkg);
      setFields_(sh, r, b.fields);
      return id;
    }
    case 'addMembers': {                              // tömeges felvitel
      const sh = sh_(SH.M), have = new Set(table_(SH.M).rows.map(x => String(x[1]).trim().toLowerCase()));
      let added = 0, skipped = [];
      (b.list || []).forEach(it => {
        const name = String(it.name || '').trim(); if (!name) return;
        if (have.has(name.toLowerCase())) { skipped.push(name); return; }
        have.add(name.toLowerCase());
        sh.appendRow([id_(), name, 'igen']);
        const r = sh.getLastRow();
        if (b.pkg) sh.getRange(r, yearCol_(sh, b.year)).setValue(b.pkg);
        setFields_(sh, r, it);
        added++;
      });
      return { added, skipped };
    }
    case 'updateMember': {
      const sh = sh_(SH.M), r = rowOf_(sh, b.id);
      if (b.name != null) {
        const name = String(b.name).trim(); if (!name) throw new Error('Hiányzó név');
        sh.getRange(r, 2).setValue(name);
        const bs = sh_(SH.B), v = bs.getDataRange().getValues();      // név frissítése a befizetéseknél is
        for (let i = 1; i < v.length; i++) if (String(v[i][1]) === String(b.id)) bs.getRange(i + 1, 3).setValue(name);
      }
      if (b.active != null) sh.getRange(r, 3).setValue(b.active ? 'igen' : 'nem');
      if (b.year) sh.getRange(r, yearCol_(sh, b.year)).setValue(b.pkg || '');
      setFields_(sh, r, b.fields);
      return true;
    }
    case 'deleteMember': {
      if (table_(SH.B).rows.some(x => String(x[1]) === String(b.id))) throw new Error('A tagnak van befizetése; állítsd inaktívra.');
      const sh = sh_(SH.M); sh.deleteRow(rowOf_(sh, b.id)); return true;
    }
    case 'savePackages': {
      const list = (b.packages || []).map(p => [String(p.name), Math.round(Number(p.fee) || 0)]);
      const sh = sh_(SH.P), last = sh.getLastRow();
      if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
      if (list.length) sh.getRange(2, 1, list.length, 2).setValues(list);
      return true;
    }
  }
  throw new Error('Ismeretlen művelet: ' + b.action);
}

/* ---------- egyszeri beállítás + 2026-os adatok az Excelből ---------- */
function setup() {
  const ss = ss_();
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', KEZDO_JELSZO);
  const mk = (name, head) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(head); sh.getRange(1, 1, 1, head.length).setFontWeight('bold'); sh.setFrozenRows(1); }
    return sh;
  };
  const p = mk(SH.P, ['Csomag', 'Éves díj']);
  const m = mk(SH.M, ['ID', 'Név', 'Aktív', '2026']);
  const b = mk(SH.B, B_HEAD);
  if (p.getLastRow() > 1 || m.getLastRow() > 1 || b.getLastRow() > 1) {
    Logger.log('A lapokon már vannak adatok – a kezdő adatokat nem töltöm be újra. Jelszó elmentve.');
    return;
  }
  p.getRange(2, 1, 2, 2).setValues([['VESC+', 60000], ['VESC', 15000]]);

  const M = [
    ['m01','Komáromi Balázs','VESC+'],['m02','Kovács Gergő','VESC+'],['m03','Gombos Róbert','VESC+'],
    ['m04','Horváth Kinga','VESC+'],['m05','Bachstedter Milán','VESC+'],['m06','Leé Diána','VESC+'],
    ['m07','Klein Mónika','VESC+'],['m08','Hamerli Péter','VESC+'],['m09','Gyulai Barnabás','VESC+'],
    ['m10','Sárossy Zsombor','VESC+'],['m11','Abonyi János','VESC+'],['m12','Saad','VESC+'],
    ['m13','Szegedi Máté','VESC+'],['m14','Morvai Balázs','VESC+'],['m15','Módra György','VESC+'],
    ['m16','Vasmatics Lajos','VESC+'],['m17','Pintér Mónika','VESC+'],['m18','Toldi Levente','VESC+'],
    ['m19','Szelesné Nagy Emília','VESC+'],['m20','Mihalko Benjamin Robin','VESC+'],
    ['m21','Pintér Krisztián','VESC'],['m22','Hajós Krisztián','VESC'],['m23','Hamerli Csenge','VESC'],
    ['m24','Simon Gergely','VESC'],['m25','Fülöp Zsófia','VESC']
  ];
  m.getRange(2, 1, M.length, 4).setValues(M.map(r => [r[0], r[1], 'igen', r[2]]));
  const name = {}; M.forEach(r => name[r[0]] = r[1]);

  // [tag, összeg, tól, ig]
  const P = [
    ['m01',60000,1,12],['m02',15000,1,3],['m02',15000,4,6],['m02',15000,7,9],
    ['m03',12000,1,3],['m03',18000,4,6],['m03',15000,7,9],
    ['m05',60000,1,12],['m06',15000,1,3],['m07',60000,1,12],['m08',30000,1,6],['m08',30000,7,12],
    ['m09',60000,1,12],['m10',15000,1,3],['m11',60000,1,12],['m12',15000,7,7],
    ['m14',60000,1,12],['m15',60000,1,12],['m16',10000,1,2],['m17',60000,1,12],
    ['m18',15000,1,3],['m18',15000,4,6],['m19',10000,7,8],['m19',10000,9,10],['m20',15000,7,9],
    ['m21',15000,1,12],['m22',15000,1,12],['m23',15000,1,12],['m24',15000,1,12],
    ['m25',15000,1,1],['m25',5000,5,5]
  ];
  for (let i = 1; i <= 10; i++) P.push(['m04', 5000, i, i]);
  [1,4,5,6,7,8,9].forEach(i => P.push(['m13', 5000, i, i]));
  const now = new Date();
  const rows = P.map(x => [id_(), x[0], name[x[0]], x[1], '2026-' + String(x[2]).padStart(2, '0') + '-01',
                           2026, x[2], x[3], '', 'Excelből átvéve', now]);
  b.getRange(2, 1, rows.length, B_HEAD.length).setValues(rows);
  b.getRange(2, 5, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  Logger.log('Kész: ' + M.length + ' tag, ' + rows.length + ' befizetés. Jelszó elmentve.');
}

/** Jelszócsere: írd át az új jelszót, és futtasd. */
function jelszoCsere() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', 'UJ-JELSZO-IDE');
}
