/*************************************************************
 *  SONICOD POINTAGE — Agadir Tilila — Code.gs (fichier wa7ed)
 *  Drari kayscaniw QR wa7ed → kaytsjjel sa3a li jaw fiha
 *  Admin: <lien>?admin  → liste dyal l7odor b tarikh, drari, QR
 *************************************************************/

const CONFIG = {
  ADMIN_PASSWORD: 'sonicod2026',   // ⚠️ BDDLO !
  AGENCE: 'Agadir Tilila',
  TZ: 'Africa/Casablanca',
  TOLERANCE_MIN: 5,                // d9ayé9 sma7 9bel ma ytkteb retard
  DEFAULT_START: '09:00',
  DEFAULT_RAYON: 150,              // metres
  SHEET_NAME: 'SONICOD POINTAGE'
};

const H_POINTAGE = ['Date', 'Heure d\'arrivée', 'Nom', 'Retard (min)', 'Distance (m)', 'GPS', 'Appareil'];
const H_DRARI = ['Nom', 'PIN', 'Actif', 'Appareil', 'Date ajout'];

/* ====================== PAGES ====================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const admin = Object.prototype.hasOwnProperty.call(p, 'admin');
  const code = String(p.a || '').replace(/[^A-Za-z0-9-]/g, '');
  const html = (admin ? ADMIN_HTML : SCAN_HTML)
    .split('__AGENCE_CODE__').join(code)
    .split('__AGENCE__').join(CONFIG.AGENCE)
    .split('__APP_URL__').join(ScriptApp.getService().getUrl());
  return HtmlService.createHtmlOutput(html)
    .setTitle(admin ? 'Sonicod Pointage — Admin' : 'Sonicod Pointage')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ====================== API (l site GitHub kayhder m3a had l9it3a) ====================== */

const API_FNS = { initScan: initScan, linkDevice: linkDevice, pointer: pointer,
  adminData: adminData, adminReport: adminReport, adminSaveDrari: adminSaveDrari,
  adminDrariAction: adminDrariAction, adminSettings: adminSettings };

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API_FNS[req.fn];
    if (!fn) throw new Error('Fonction ma kaynach');
    out = { r: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { err: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Dir Run 3la had la fonction mra wa7da bach t3ti l'autorisation o ytsawb l Sheet. */
function installer() {
  const ss = ss_();
  settings_();
  Logger.log('Sheet: ' + ss.getUrl());
}

/* ====================== REGLAGES (QR, heure, GPS) ====================== */

function props_() { return PropertiesService.getScriptProperties(); }
function settings_() {
  const pr = props_();
  let code = pr.getProperty('QR_CODE');
  if (!code) { code = newCode_(); pr.setProperty('QR_CODE', code); }
  const lat = pr.getProperty('LAT'), lng = pr.getProperty('LNG');
  return {
    code: code,
    url: pr.getProperty('APP_URL') || '',
    start: pr.getProperty('START') || CONFIG.DEFAULT_START,
    rayon: Number(pr.getProperty('RAYON')) || CONFIG.DEFAULT_RAYON,
    lat: lat ? Number(lat) : null,
    lng: lng ? Number(lng) : null
  };
}

/* ====================== SHEET ====================== */

function ss_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) { ensure_(active); return active; }
  const pr = props_();
  const id = pr.getProperty('SS_ID');
  if (id) {
    try { const s = SpreadsheetApp.openById(id); ensure_(s); return s; } catch (err) { /* tsayb jdid */ }
  }
  const ss = SpreadsheetApp.create(CONFIG.SHEET_NAME);
  pr.setProperty('SS_ID', ss.getId());
  ensure_(ss);
  return ss;
}

let _ensured = false;
function ensure_(ss) {
  if (_ensured) return;
  _ensured = true;
  const mk = (name, headers, color) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange('A:Z').setNumberFormat('@');
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setFontColor('#ffffff').setBackground(color);
      sh.setFrozenRows(1);
    }
    return sh;
  };
  mk('Pointage', H_POINTAGE, '#E30613');
  mk('Drari', H_DRARI, '#222222');
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Feuille 1') || ss.getSheetByName('Feuille1');
  if (def && ss.getSheets().length > 2 && def.getLastRow() === 0) ss.deleteSheet(def);
}

function sh_(name) { return ss_().getSheetByName(name); }

function rows_(name, maxRows) {
  const sh = sh_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const width = name === 'Pointage' ? H_POINTAGE.length : H_DRARI.length;
  let start = 2;
  if (maxRows && last - 1 > maxRows) start = last - maxRows + 1;
  return sh.getRange(start, 1, last - start + 1, width).getDisplayValues()
    .map((r, i) => { r._row = start + i; return r; });
}

/* ====================== OUTILS ====================== */

function now_() {
  const d = new Date();
  return {
    date: Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM-dd'),
    heure: Utilities.formatDate(d, CONFIG.TZ, 'HH:mm:ss')
  };
}
function mins_(hhmm) {
  const p = String(hhmm || '').split(':');
  return (parseInt(p[0], 10) || 0) * 60 + (parseInt(p[1], 10) || 0) + (parseInt(p[2], 10) || 0) / 60;
}
function newCode_() {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = 'TIL-';
  for (let i = 0; i < 6; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
  return s;
}
function dist_(lat1, lng1, lat2, lng2) {
  const R = 6371000, toR = x => x * Math.PI / 180;
  const dLat = toR(lat2 - lat1), dLng = toR(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toR(lat1)) * Math.cos(toR(lat2)) * Math.sin(dLng / 2) ** 2;
  return Math.round(2 * R * Math.asin(Math.sqrt(a)));
}
function norm_(s) { return String(s || '').trim().toLowerCase(); }

function drariObj_(r) {
  return { nom: r[0], pin: r[1], actif: r[2] !== 'Non', appareil: r[3], ajout: r[4], row: r._row };
}
function personByToken_(token) {
  if (!token) return null;
  const r = rows_('Drari').find(x => x[3] === token);
  return r ? drariObj_(r) : null;
}
function todayRec_(nom, date) {
  return rows_('Pointage', 3000).find(r => r[0] === date && r[2] === nom) || null;
}

/* ====================== DRARI (scan) ====================== */

function initScan(code, token) {
  const st = settings_();
  if (!code || code !== st.code) return { ok: false, msg: 'QR machi s7i7 wla tbddel. Scanni QR li f lagence.' };
  const me = personByToken_(token);
  const res = { ok: true, gps: st.lat !== null, me: null, drari: [] };
  if (me && me.actif) {
    const rec = todayRec_(me.nom, now_().date);
    res.me = { nom: me.nom, deja: rec ? rec[1].substring(0, 5) : null };
  } else {
    res.drari = rows_('Drari').map(drariObj_).filter(d => d.actif).map(d => d.nom).sort((a, b) => a.localeCompare(b));
    if (me && !me.actif) res.blocked = true;
  }
  return res;
}

function linkDevice(code, nom, pin, token) {
  if (code !== settings_().code) return { ok: false, msg: 'QR machi s7i7.' };
  if (!token || String(token).length < 10) return { ok: false, msg: 'Erreur dyal telephone, 3awd 7ell l page.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const list = rows_('Drari').map(drariObj_);
    const p = list.find(d => norm_(d.nom) === norm_(nom));
    if (!p || !p.actif) return { ok: false, msg: 'Had smiya ma kaynach wla m7bousa.' };
    if (String(p.pin).trim() !== String(pin).trim()) return { ok: false, msg: 'PIN ghalat.' };
    const other = list.find(d => d.appareil === token && d.nom !== p.nom);
    if (other) return { ok: false, msg: 'Had telephone deja mrbout b ' + other.nom + '. Kol wa7d b telephone dyalo.' };
    if (p.appareil && p.appareil !== token) {
      return { ok: false, msg: 'Nta deja mrbout b telephone akhor. Gol l admin ydir « Reset tel ».' };
    }
    sh_('Drari').getRange(p.row, 4).setValue(token);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function pointer(code, token, geo) {
  const st = settings_();
  if (code !== st.code) return { ok: false, msg: 'QR machi s7i7.' };
  const me = personByToken_(token);
  if (!me) return { ok: false, msg: 'Telephone machi mrbout. 7ell l page men jdid.', relink: true };
  if (!me.actif) return { ok: false, msg: 'Compte dyalk m7bous. Tsl b l admin.' };

  let distance = '', gpsTxt = 'Non';
  if (geo && geo.lat !== undefined && geo.lat !== null) {
    gpsTxt = 'Oui (±' + Math.round(geo.acc || 0) + 'm)';
    if (st.lat !== null) distance = dist_(st.lat, st.lng, Number(geo.lat), Number(geo.lng));
  }
  if (st.lat !== null) {
    if (distance === '') return { ok: false, msg: 'Khass tkhlli l position (GPS). Fte7 Localisation o 3awd.' };
    const marge = Math.min(Number(geo.acc) || 0, 100);
    if (distance > st.rayon + marge) {
      return { ok: false, msg: 'Nta b3id 3la lagence (' + distance + ' m). Khassk tkon f lagence bach tscanni.' };
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const t = now_();
    const rec = todayRec_(me.nom, t.date);
    if (rec) return { ok: true, deja: true, heure: rec[1].substring(0, 5), retard: rec[3], nom: me.nom };
    const late = Math.floor(mins_(t.heure) - mins_(st.start));
    const retard = late > CONFIG.TOLERANCE_MIN ? late : 0;
    sh_('Pointage').appendRow([t.date, t.heure, me.nom, String(retard), String(distance), gpsTxt, String(token).substring(0, 8)]);
    return { ok: true, heure: t.heure.substring(0, 5), retard: retard, nom: me.nom };
  } finally { lock.releaseLock(); }
}

/* ====================== ADMIN ====================== */

function check_(pwd) {
  if (String(pwd) !== CONFIG.ADMIN_PASSWORD) throw new Error('Mot de passe ghalat');
}

function adminData(pwd) {
  check_(pwd);
  const ss = ss_();
  return {
    ok: true,
    url: ScriptApp.getService().getUrl(),
    sheetUrl: ss.getUrl(),
    today: now_().date,
    settings: settings_(),
    drari: rows_('Drari').map(drariObj_).map(d => { d.lie = !!d.appareil; delete d.appareil; return d; })
  };
}

function adminReport(pwd, f) {
  check_(pwd);
  f = f || {};
  const from = f.from || now_().date, to = f.to || from;
  const q = norm_(f.nom);
  const days = rows_('Pointage')
    .filter(r => r[0] >= from && r[0] <= to && (!q || norm_(r[2]).indexOf(q) >= 0))
    .map(r => ({ date: r[0], heure: r[1].substring(0, 5), nom: r[2], retard: r[3], gps: r[5] }))
    .sort((a, b) => b.date.localeCompare(a.date) || a.heure.localeCompare(b.heure));

  let absents = [];
  if (from === to) {
    const presents = {};
    days.forEach(d => presents[d.nom] = 1);
    absents = rows_('Drari').map(drariObj_)
      .filter(d => d.actif && !presents[d.nom] && (!q || norm_(d.nom).indexOf(q) >= 0))
      .map(d => d.nom);
  }
  return { ok: true, from: from, to: to, days: days, absents: absents };
}

function adminSaveDrari(pwd, d) {
  check_(pwd);
  const nom = String(d.nom || '').trim();
  if (!nom) throw new Error('Kteb smiya');
  const list = rows_('Drari').map(drariObj_);
  const sh = sh_('Drari');
  const existing = d.oldNom ? list.find(x => x.nom === d.oldNom) : null;
  if (list.find(x => norm_(x.nom) === norm_(nom) && x !== existing)) throw new Error('Had smiya deja kayna');
  const pin = String(d.pin || '').trim() || String(Math.floor(1000 + Math.random() * 9000));
  if (existing) {
    sh.getRange(existing.row, 1, 1, 2).setValues([[nom, pin]]);
  } else {
    sh.appendRow([nom, pin, 'Oui', '', now_().date]);
  }
  return adminData(pwd);
}

function adminDrariAction(pwd, nom, action) {
  check_(pwd);
  const p = rows_('Drari').map(drariObj_).find(x => x.nom === nom);
  if (!p) throw new Error('Ma kaynch');
  const sh = sh_('Drari');
  if (action === 'reset') sh.getRange(p.row, 4).setValue('');
  if (action === 'toggle') sh.getRange(p.row, 3).setValue(p.actif ? 'Non' : 'Oui');
  if (action === 'pin') sh.getRange(p.row, 2).setValue(String(Math.floor(1000 + Math.random() * 9000)));
  if (action === 'delete') sh.deleteRow(p.row);
  return adminData(pwd);
}

function adminSettings(pwd, action, a) {
  check_(pwd);
  const pr = props_();
  a = a || {};
  if (action === 'save') {
    if (!/^\d{1,2}:\d{2}$/.test(String(a.start || ''))) throw new Error('Heure: b7al 09:00');
    pr.setProperty('START', a.start);
    pr.setProperty('RAYON', String(Number(a.rayon) || CONFIG.DEFAULT_RAYON));
  }
  if (action === 'pos') { pr.setProperty('LAT', String(a.lat)); pr.setProperty('LNG', String(a.lng)); }
  if (action === 'nopos') { pr.deleteProperty('LAT'); pr.deleteProperty('LNG'); }
  if (action === 'code') pr.setProperty('QR_CODE', newCode_());
  if (action === 'url') {
    let u = String(a.url || '').trim().split('?')[0].split('#')[0];
    if (!/^https:\/\/\S+$/.test(u)) throw new Error('Lien khass ybda b https://');
    pr.setProperty('APP_URL', u);
  }
  return adminData(pwd);
}
