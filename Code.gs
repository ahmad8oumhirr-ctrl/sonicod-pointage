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

/* ====================== HTML (ma tbddel walo hna) ====================== */
const SCAN_HTML = `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<meta charset="utf-8">
<style>
  *{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
  body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:#f4f4f5;color:#18181b}
  header{background:#E30613;color:#fff;padding:18px 20px 22px;text-align:center}
  header .logo{font-weight:900;font-size:22px;letter-spacing:1px}
  header .ag{opacity:.9;font-size:15px;margin-top:4px}
  main{padding:18px;max-width:460px;margin:0 auto}
  .card{background:#fff;border-radius:16px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:14px}
  h2{margin:0 0 6px;font-size:20px}
  p{margin:6px 0;color:#52525b;line-height:1.45}
  .btn{display:block;width:100%;border:0;border-radius:14px;padding:18px;font-size:19px;font-weight:800;color:#fff;background:#E30613;margin-top:14px}
  .btn.green{background:#16a34a}.btn.dark{background:#27272a}
  .btn:disabled{opacity:.55}
  input,select{width:100%;padding:14px;border:1.5px solid #d4d4d8;border-radius:12px;font-size:17px;margin-top:10px;background:#fff}
  .list{max-height:260px;overflow:auto;margin-top:10px;border:1.5px solid #e4e4e7;border-radius:12px}
  .item{padding:13px 14px;border-bottom:1px solid #f0f0f0;font-size:16px;display:flex;justify-content:space-between}
  .item small{color:#a1a1aa}
  .item.sel{background:#fee2e2;font-weight:700}
  .big{font-size:64px;text-align:center;line-height:1}
  .res{text-align:center}
  .res h1{margin:10px 0 2px;font-size:30px}
  .res .time{font-size:46px;font-weight:900;margin:6px 0}
  .ok{color:#16a34a}.out{color:#2563eb}.warn{color:#ea580c}.err{color:#dc2626}
  .today{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
  .chip{background:#f4f4f5;border-radius:20px;padding:6px 12px;font-size:14px}
  .spin{width:42px;height:42px;border:4px solid #fecaca;border-top-color:#E30613;border-radius:50%;animation:s 1s linear infinite;margin:30px auto}
  @keyframes s{to{transform:rotate(360deg)}}
</style>
</head>
<body>
<header><div class="logo">SONICOD</div><div class="ag" id="ag">📍 Agence __AGENCE__</div></header>
<main id="app"><div class="spin"></div></main>

<script>
const CODE = '__AGENCE_CODE__';
const $app = document.getElementById('app');
let INFO = null, selNom = null;

function token(){
  let t = null;
  try { t = localStorage.getItem('sonicod_pt_token'); } catch(e){}
  if (!t) {
    t = 'T' + Date.now().toString(36) + Math.random().toString(36).slice(2, 12);
    try { localStorage.setItem('sonicod_pt_token', t); } catch(e){}
  }
  return t;
}
function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function call(fn){
  const args = Array.prototype.slice.call(arguments, 1);
  return new Promise((ok, ko) => google.script.run.withSuccessHandler(ok).withFailureHandler(ko)[fn].apply(null, args));
}
function spin(){ $app.innerHTML = '<div class="spin"></div>'; }
function fail(msg, retry){
  $app.innerHTML = '<div class="card res"><div class="big">⚠️</div><h2 class="err" style="margin-top:12px">' + esc(msg) + '</h2>' +
    '<button class="btn dark" onclick="' + (retry || 'start()') + '">3awd</button></div>';
}

async function start(){
  if (!CODE || (typeof API_ID !== 'undefined' && !API_ID)) {
    $app.innerHTML = '<div class="card res"><div class="big">📷</div><h2 style="margin-top:12px">Scanni QR dyal lagence</h2>' +
      '<p>7ell camera dyal telephone o scanni l QR li m3l9 f lagence.</p></div>';
    return;
  }
  spin();
  try { INFO = await call('initScan', CODE, token()); } catch(e){ return fail('Mochkil f connexion. 3awd.'); }
  if (!INFO.ok) return fail(INFO.msg);
  
  if (INFO.me) showMe(); else showLink();
}

function showMe(){
  const me = INFO.me;
  if (me.deja) return showDone({ nom: me.nom, heure: me.deja, deja: true });
  $app.innerHTML = '<div class="card"><p>Salam 👋</p><h2>' + esc(me.nom) + '</h2>' +
    '<button class="btn green" id="go" onclick="doPoint()">✅ Ana jit — Sjjel</button>' +
    (INFO.gps ? '<p style="font-size:13px;margin-top:12px">📍 Ila tlb mnk l position, 9bel (Autoriser).</p>' : '') +
    '</div>';
  doPoint();
}

function showLink(){
  const blocked = INFO.blocked ? '<p class="err">Compte dyalk m7bous. Tsl b l admin.</p>' : '';
  $app.innerHTML = '<div class="card"><h2>Awel merra 👋</h2>' +
    '<p>Khtar smiytek o dkhel l PIN li 3tak l admin. Ghir merra wa7da f had telephone.</p>' + blocked +
    '<input id="q" placeholder="🔎 9leb 3la smiytek" oninput="renderList()">' +
    '<div class="list" id="list"></div>' +
    '<input id="pin" type="tel" inputmode="numeric" maxlength="6" placeholder="PIN (4 ar9am)">' +
    '<button class="btn" id="lk" onclick="doLink()">Rbet telephone dyali</button></div>';
  renderList();
}
function renderList(){
  const q = (document.getElementById('q').value || '').toLowerCase();
  const items = INFO.drari.filter(n => n.toLowerCase().indexOf(q) >= 0);
  document.getElementById('list').innerHTML = items.length ? items.map(n =>
    '<div class="item' + (n === selNom ? ' sel' : '') + '" data-n="' + esc(n) + '" onclick="pick(this)">' +
    '<span>' + esc(n) + '</span></div>').join('')
    : '<div class="item"><small>Ma kayn 7ta wa7ed. Gol l admin yzidek.</small></div>';
}
function pick(el){ selNom = el.getAttribute('data-n'); renderList(); document.getElementById('pin').focus(); }

async function doLink(){
  const pin = document.getElementById('pin').value.trim();
  if (!selNom) return alertBox('Khtar smiytek men la liste.');
  if (!pin) return alertBox('Dkhel l PIN.');
  const b = document.getElementById('lk'); b.disabled = true; b.textContent = '...';
  try {
    const r = await call('linkDevice', CODE, selNom, pin, token());
    if (!r.ok) { b.disabled = false; b.textContent = 'Rbet telephone dyali'; return alertBox(r.msg); }
    start();
  } catch(e){ b.disabled = false; b.textContent = 'Rbet telephone dyali'; alertBox('Mochkil f connexion.'); }
}
function alertBox(m){
  let el = document.getElementById('msg');
  if (!el) { el = document.createElement('p'); el.id = 'msg'; el.className = 'err'; el.style.fontWeight = '700'; document.querySelector('.card').appendChild(el); }
  el.textContent = m;
}

function getPos(){
  return new Promise(res => {
    if (!navigator.geolocation) return res(null);
    navigator.geolocation.getCurrentPosition(
      p => res({ lat: p.coords.latitude, lng: p.coords.longitude, acc: p.coords.accuracy }),
      () => res(null), { enableHighAccuracy: true, timeout: 12000, maximumAge: 30000 });
  });
}

async function doPoint(){
  const b = document.getElementById('go'); b.disabled = true; b.textContent = INFO.gps ? '📍 Kanchouf l position...' : '...';
  const geo = INFO.gps ? await getPos() : await Promise.race([getPos(), new Promise(r => setTimeout(() => r(null), 4000))]);
  b.textContent = '...';
  let r;
  try { r = await call('pointer', CODE, token(), geo); } catch(e){ return fail('Mochkil f connexion. 3awd.', 'start()'); }
  if (!r.ok) { if (r.relink) { try{localStorage.removeItem('sonicod_pt_token');}catch(e){} } return fail(r.msg, 'start()'); }
  showDone(r);
}

function showDone(r){
  let extra = '';
  if (r.deja) extra = '<p class="ok" style="font-weight:700">Deja tsjjelti lyoum ✔️</p>';
  else if (Number(r.retard) > 0) extra = '<p class="warn" style="font-weight:700">⏰ Retard ' + r.retard + ' d9i9a</p>';
  else extra = '<p class="ok" style="font-weight:700">F lwe9t 👏</p>';
  $app.innerHTML = '<div class="card res"><div class="big">✅</div>' +
    '<h1 class="ok">' + (r.deja ? 'MSJJEL' : 'MR7BA') + '</h1>' +
    '<p style="margin-top:8px">Jiti m3a</p><div class="time">' + esc(r.heure) + '</div>' +
    '<p><b>' + esc(r.nom) + '</b></p>' + extra +
    '<p style="font-size:13px;margin-top:16px">Nhar mbrouk 💪</p></div>';
}

start();
</script>
</body>
</html>
`;

const ADMIN_HTML = `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<meta charset="utf-8">
<script>
// Copyright (c) 2009 Kazuhiko Arase
// Licensed under the MIT license:
var qrcode=function(){var t=function(t,r){var e=t,n=g[r],o=null,i=0,a=null,u=[],f={},c=function(t,r){o=function(t){for(var r=new Array(t),e=0;e<t;e+=1){r[e]=new Array(t);for(var n=0;n<t;n+=1)r[e][n]=null}return r}(i=4*e+17),l(0,0),l(i-7,0),l(0,i-7),s(),h(),d(t,r),e>=7&&v(t),null==a&&(a=p(e,n,u)),w(a,r)},l=function(t,r){for(var e=-1;e<=7;e+=1)if(!(t+e<=-1||i<=t+e))for(var n=-1;n<=7;n+=1)r+n<=-1||i<=r+n||(o[t+e][r+n]=0<=e&&e<=6&&(0==n||6==n)||0<=n&&n<=6&&(0==e||6==e)||2<=e&&e<=4&&2<=n&&n<=4)},h=function(){for(var t=8;t<i-8;t+=1)null==o[t][6]&&(o[t][6]=t%2==0);for(var r=8;r<i-8;r+=1)null==o[6][r]&&(o[6][r]=r%2==0)},s=function(){for(var t=B.getPatternPosition(e),r=0;r<t.length;r+=1)for(var n=0;n<t.length;n+=1){var i=t[r],a=t[n];if(null==o[i][a])for(var u=-2;u<=2;u+=1)for(var f=-2;f<=2;f+=1)o[i+u][a+f]=-2==u||2==u||-2==f||2==f||0==u&&0==f}},v=function(t){for(var r=B.getBCHTypeNumber(e),n=0;n<18;n+=1){var a=!t&&1==(r>>n&1);o[Math.floor(n/3)][n%3+i-8-3]=a}for(n=0;n<18;n+=1){a=!t&&1==(r>>n&1);o[n%3+i-8-3][Math.floor(n/3)]=a}},d=function(t,r){for(var e=n<<3|r,a=B.getBCHTypeInfo(e),u=0;u<15;u+=1){var f=!t&&1==(a>>u&1);u<6?o[u][8]=f:u<8?o[u+1][8]=f:o[i-15+u][8]=f}for(u=0;u<15;u+=1){f=!t&&1==(a>>u&1);u<8?o[8][i-u-1]=f:u<9?o[8][15-u-1+1]=f:o[8][15-u-1]=f}o[i-8][8]=!t},w=function(t,r){for(var e=-1,n=i-1,a=7,u=0,f=B.getMaskFunction(r),c=i-1;c>0;c-=2)for(6==c&&(c-=1);;){for(var g=0;g<2;g+=1)if(null==o[n][c-g]){var l=!1;u<t.length&&(l=1==(t[u]>>>a&1)),f(n,c-g)&&(l=!l),o[n][c-g]=l,-1==(a-=1)&&(u+=1,a=7)}if((n+=e)<0||i<=n){n-=e,e=-e;break}}},p=function(t,r,e){for(var n=A.getRSBlocks(t,r),o=b(),i=0;i<e.length;i+=1){var a=e[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var u=0;for(i=0;i<n.length;i+=1)u+=n[i].dataCount;if(o.getLengthInBits()>8*u)throw"code length overflow. ("+o.getLengthInBits()+">"+8*u+")";for(o.getLengthInBits()+4<=8*u&&o.put(0,4);o.getLengthInBits()%8!=0;)o.putBit(!1);for(;!(o.getLengthInBits()>=8*u||(o.put(236,8),o.getLengthInBits()>=8*u));)o.put(17,8);return function(t,r){for(var e=0,n=0,o=0,i=new Array(r.length),a=new Array(r.length),u=0;u<r.length;u+=1){var f=r[u].dataCount,c=r[u].totalCount-f;n=Math.max(n,f),o=Math.max(o,c),i[u]=new Array(f);for(var g=0;g<i[u].length;g+=1)i[u][g]=255&t.getBuffer()[g+e];e+=f;var l=B.getErrorCorrectPolynomial(c),h=k(i[u],l.getLength()-1).mod(l);for(a[u]=new Array(l.getLength()-1),g=0;g<a[u].length;g+=1){var s=g+h.getLength()-a[u].length;a[u][g]=s>=0?h.getAt(s):0}}var v=0;for(g=0;g<r.length;g+=1)v+=r[g].totalCount;var d=new Array(v),w=0;for(g=0;g<n;g+=1)for(u=0;u<r.length;u+=1)g<i[u].length&&(d[w]=i[u][g],w+=1);for(g=0;g<o;g+=1)for(u=0;u<r.length;u+=1)g<a[u].length&&(d[w]=a[u][g],w+=1);return d}(o,n)};f.addData=function(t,r){var e=null;switch(r=r||"Byte"){case"Numeric":e=M(t);break;case"Alphanumeric":e=x(t);break;case"Byte":e=m(t);break;case"Kanji":e=L(t);break;default:throw"mode:"+r}u.push(e),a=null},f.isDark=function(t,r){if(t<0||i<=t||r<0||i<=r)throw t+","+r;return o[t][r]},f.getModuleCount=function(){return i},f.make=function(){if(e<1){for(var t=1;t<40;t++){for(var r=A.getRSBlocks(t,n),o=b(),i=0;i<u.length;i++){var a=u[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var g=0;for(i=0;i<r.length;i++)g+=r[i].dataCount;if(o.getLengthInBits()<=8*g)break}e=t}c(!1,function(){for(var t=0,r=0,e=0;e<8;e+=1){c(!0,e);var n=B.getLostPoint(f);(0==e||t>n)&&(t=n,r=e)}return r}())},f.createTableTag=function(t,r){t=t||2;var e="";e+='<table style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: "+(r=void 0===r?4*t:r)+"px;",e+='">',e+="<tbody>";for(var n=0;n<f.getModuleCount();n+=1){e+="<tr>";for(var o=0;o<f.getModuleCount();o+=1)e+='<td style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: 0px;",e+=" width: "+t+"px;",e+=" height: "+t+"px;",e+=" background-color: ",e+=f.isDark(n,o)?"#000000":"#ffffff",e+=";",e+='"/>';e+="</tr>"}return e+="</tbody>",e+="</table>"},f.createSvgTag=function(t,r,e,n){var o={};"object"==typeof arguments[0]&&(t=(o=arguments[0]).cellSize,r=o.margin,e=o.alt,n=o.title),t=t||2,r=void 0===r?4*t:r,(e="string"==typeof e?{text:e}:e||{}).text=e.text||null,e.id=e.text?e.id||"qrcode-description":null,(n="string"==typeof n?{text:n}:n||{}).text=n.text||null,n.id=n.text?n.id||"qrcode-title":null;var i,a,u,c,g=f.getModuleCount()*t+2*r,l="";for(c="l"+t+",0 0,"+t+" -"+t+",0 0,-"+t+"z ",l+='<svg version="1.1" xmlns="http://www.w3.org/2000/svg"',l+=o.scalable?"":' width="'+g+'px" height="'+g+'px"',l+=' viewBox="0 0 '+g+" "+g+'" ',l+=' preserveAspectRatio="xMinYMin meet"',l+=n.text||e.text?' role="img" aria-labelledby="'+y([n.id,e.id].join(" ").trim())+'"':"",l+=">",l+=n.text?'<title id="'+y(n.id)+'">'+y(n.text)+"</title>":"",l+=e.text?'<description id="'+y(e.id)+'">'+y(e.text)+"</description>":"",l+='<rect width="100%" height="100%" fill="white" cx="0" cy="0"/>',l+='<path d="',a=0;a<f.getModuleCount();a+=1)for(u=a*t+r,i=0;i<f.getModuleCount();i+=1)f.isDark(a,i)&&(l+="M"+(i*t+r)+","+u+c);return l+='" stroke="transparent" fill="black"/>',l+="</svg>"},f.createDataURL=function(t,r){t=t||2,r=void 0===r?4*t:r;var e=f.getModuleCount()*t+2*r,n=r,o=e-r;return I(e,e,function(r,e){if(n<=r&&r<o&&n<=e&&e<o){var i=Math.floor((r-n)/t),a=Math.floor((e-n)/t);return f.isDark(a,i)?0:1}return 1})},f.createImgTag=function(t,r,e){t=t||2,r=void 0===r?4*t:r;var n=f.getModuleCount()*t+2*r,o="";return o+="<img",o+=' src="',o+=f.createDataURL(t,r),o+='"',o+=' width="',o+=n,o+='"',o+=' height="',o+=n,o+='"',e&&(o+=' alt="',o+=y(e),o+='"'),o+="/>"};var y=function(t){for(var r="",e=0;e<t.length;e+=1){var n=t.charAt(e);switch(n){case"<":r+="&lt;";break;case">":r+="&gt;";break;case"&":r+="&amp;";break;case'"':r+="&quot;";break;default:r+=n}}return r};return f.createASCII=function(t,r){if((t=t||1)<2)return function(t){t=void 0===t?2:t;var r,e,n,o,i,a=1*f.getModuleCount()+2*t,u=t,c=a-t,g={"██":"█","█ ":"▀"," █":"▄","  ":" "},l={"██":"▀","█ ":"▀"," █":" ","  ":" "},h="";for(r=0;r<a;r+=2){for(n=Math.floor((r-u)/1),o=Math.floor((r+1-u)/1),e=0;e<a;e+=1)i="█",u<=e&&e<c&&u<=r&&r<c&&f.isDark(n,Math.floor((e-u)/1))&&(i=" "),u<=e&&e<c&&u<=r+1&&r+1<c&&f.isDark(o,Math.floor((e-u)/1))?i+=" ":i+="█",h+=t<1&&r+1>=c?l[i]:g[i];h+="\\n"}return a%2&&t>0?h.substring(0,h.length-a-1)+Array(a+1).join("▀"):h.substring(0,h.length-1)}(r);t-=1,r=void 0===r?2*t:r;var e,n,o,i,a=f.getModuleCount()*t+2*r,u=r,c=a-r,g=Array(t+1).join("██"),l=Array(t+1).join("  "),h="",s="";for(e=0;e<a;e+=1){for(o=Math.floor((e-u)/t),s="",n=0;n<a;n+=1)i=1,u<=n&&n<c&&u<=e&&e<c&&f.isDark(o,Math.floor((n-u)/t))&&(i=0),s+=i?g:l;for(o=0;o<t;o+=1)h+=s+"\\n"}return h.substring(0,h.length-1)},f.renderTo2dContext=function(t,r){r=r||2;for(var e=f.getModuleCount(),n=0;n<e;n++)for(var o=0;o<e;o++)t.fillStyle=f.isDark(n,o)?"black":"white",t.fillRect(n*r,o*r,r,r)},f};t.stringToBytes=(t.stringToBytesFuncs={default:function(t){for(var r=[],e=0;e<t.length;e+=1){var n=t.charCodeAt(e);r.push(255&n)}return r}}).default,t.createStringToBytes=function(t,r){var e=function(){for(var e=S(t),n=function(){var t=e.read();if(-1==t)throw"eof";return t},o=0,i={};;){var a=e.read();if(-1==a)break;var u=n(),f=n()<<8|n();i[String.fromCharCode(a<<8|u)]=f,o+=1}if(o!=r)throw o+" != "+r;return i}(),n="?".charCodeAt(0);return function(t){for(var r=[],o=0;o<t.length;o+=1){var i=t.charCodeAt(o);if(i<128)r.push(i);else{var a=e[t.charAt(o)];"number"==typeof a?(255&a)==a?r.push(a):(r.push(a>>>8),r.push(255&a)):r.push(n)}}return r}};var r,e,n,o,i,a=1,u=2,f=4,c=8,g={L:1,M:0,Q:3,H:2},l=0,h=1,s=2,v=3,d=4,w=5,p=6,y=7,B=(r=[[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70],[6,26,50,74],[6,30,54,78],[6,30,56,82],[6,30,58,86],[6,34,62,90],[6,28,50,72,94],[6,26,50,74,98],[6,30,54,78,102],[6,28,54,80,106],[6,32,58,84,110],[6,30,58,86,114],[6,34,62,90,118],[6,26,50,74,98,122],[6,30,54,78,102,126],[6,26,52,78,104,130],[6,30,56,82,108,134],[6,34,60,86,112,138],[6,30,58,86,114,142],[6,34,62,90,118,146],[6,30,54,78,102,126,150],[6,24,50,76,102,128,154],[6,28,54,80,106,132,158],[6,32,58,84,110,136,162],[6,26,54,82,110,138,166],[6,30,58,86,114,142,170]],e=1335,n=7973,i=function(t){for(var r=0;0!=t;)r+=1,t>>>=1;return r},(o={}).getBCHTypeInfo=function(t){for(var r=t<<10;i(r)-i(e)>=0;)r^=e<<i(r)-i(e);return 21522^(t<<10|r)},o.getBCHTypeNumber=function(t){for(var r=t<<12;i(r)-i(n)>=0;)r^=n<<i(r)-i(n);return t<<12|r},o.getPatternPosition=function(t){return r[t-1]},o.getMaskFunction=function(t){switch(t){case l:return function(t,r){return(t+r)%2==0};case h:return function(t,r){return t%2==0};case s:return function(t,r){return r%3==0};case v:return function(t,r){return(t+r)%3==0};case d:return function(t,r){return(Math.floor(t/2)+Math.floor(r/3))%2==0};case w:return function(t,r){return t*r%2+t*r%3==0};case p:return function(t,r){return(t*r%2+t*r%3)%2==0};case y:return function(t,r){return(t*r%3+(t+r)%2)%2==0};default:throw"bad maskPattern:"+t}},o.getErrorCorrectPolynomial=function(t){for(var r=k([1],0),e=0;e<t;e+=1)r=r.multiply(k([1,C.gexp(e)],0));return r},o.getLengthInBits=function(t,r){if(1<=r&&r<10)switch(t){case a:return 10;case u:return 9;case f:case c:return 8;default:throw"mode:"+t}else if(r<27)switch(t){case a:return 12;case u:return 11;case f:return 16;case c:return 10;default:throw"mode:"+t}else{if(!(r<41))throw"type:"+r;switch(t){case a:return 14;case u:return 13;case f:return 16;case c:return 12;default:throw"mode:"+t}}},o.getLostPoint=function(t){for(var r=t.getModuleCount(),e=0,n=0;n<r;n+=1)for(var o=0;o<r;o+=1){for(var i=0,a=t.isDark(n,o),u=-1;u<=1;u+=1)if(!(n+u<0||r<=n+u))for(var f=-1;f<=1;f+=1)o+f<0||r<=o+f||0==u&&0==f||a==t.isDark(n+u,o+f)&&(i+=1);i>5&&(e+=3+i-5)}for(n=0;n<r-1;n+=1)for(o=0;o<r-1;o+=1){var c=0;t.isDark(n,o)&&(c+=1),t.isDark(n+1,o)&&(c+=1),t.isDark(n,o+1)&&(c+=1),t.isDark(n+1,o+1)&&(c+=1),0!=c&&4!=c||(e+=3)}for(n=0;n<r;n+=1)for(o=0;o<r-6;o+=1)t.isDark(n,o)&&!t.isDark(n,o+1)&&t.isDark(n,o+2)&&t.isDark(n,o+3)&&t.isDark(n,o+4)&&!t.isDark(n,o+5)&&t.isDark(n,o+6)&&(e+=40);for(o=0;o<r;o+=1)for(n=0;n<r-6;n+=1)t.isDark(n,o)&&!t.isDark(n+1,o)&&t.isDark(n+2,o)&&t.isDark(n+3,o)&&t.isDark(n+4,o)&&!t.isDark(n+5,o)&&t.isDark(n+6,o)&&(e+=40);var g=0;for(o=0;o<r;o+=1)for(n=0;n<r;n+=1)t.isDark(n,o)&&(g+=1);return e+=Math.abs(100*g/r/r-50)/5*10},o),C=function(){for(var t=new Array(256),r=new Array(256),e=0;e<8;e+=1)t[e]=1<<e;for(e=8;e<256;e+=1)t[e]=t[e-4]^t[e-5]^t[e-6]^t[e-8];for(e=0;e<255;e+=1)r[t[e]]=e;var n={glog:function(t){if(t<1)throw"glog("+t+")";return r[t]},gexp:function(r){for(;r<0;)r+=255;for(;r>=256;)r-=255;return t[r]}};return n}();function k(t,r){if(void 0===t.length)throw t.length+"/"+r;var e=function(){for(var e=0;e<t.length&&0==t[e];)e+=1;for(var n=new Array(t.length-e+r),o=0;o<t.length-e;o+=1)n[o]=t[o+e];return n}(),n={getAt:function(t){return e[t]},getLength:function(){return e.length},multiply:function(t){for(var r=new Array(n.getLength()+t.getLength()-1),e=0;e<n.getLength();e+=1)for(var o=0;o<t.getLength();o+=1)r[e+o]^=C.gexp(C.glog(n.getAt(e))+C.glog(t.getAt(o)));return k(r,0)},mod:function(t){if(n.getLength()-t.getLength()<0)return n;for(var r=C.glog(n.getAt(0))-C.glog(t.getAt(0)),e=new Array(n.getLength()),o=0;o<n.getLength();o+=1)e[o]=n.getAt(o);for(o=0;o<t.getLength();o+=1)e[o]^=C.gexp(C.glog(t.getAt(o))+r);return k(e,0).mod(t)}};return n}var A=function(){var t=[[1,26,19],[1,26,16],[1,26,13],[1,26,9],[1,44,34],[1,44,28],[1,44,22],[1,44,16],[1,70,55],[1,70,44],[2,35,17],[2,35,13],[1,100,80],[2,50,32],[2,50,24],[4,25,9],[1,134,108],[2,67,43],[2,33,15,2,34,16],[2,33,11,2,34,12],[2,86,68],[4,43,27],[4,43,19],[4,43,15],[2,98,78],[4,49,31],[2,32,14,4,33,15],[4,39,13,1,40,14],[2,121,97],[2,60,38,2,61,39],[4,40,18,2,41,19],[4,40,14,2,41,15],[2,146,116],[3,58,36,2,59,37],[4,36,16,4,37,17],[4,36,12,4,37,13],[2,86,68,2,87,69],[4,69,43,1,70,44],[6,43,19,2,44,20],[6,43,15,2,44,16],[4,101,81],[1,80,50,4,81,51],[4,50,22,4,51,23],[3,36,12,8,37,13],[2,116,92,2,117,93],[6,58,36,2,59,37],[4,46,20,6,47,21],[7,42,14,4,43,15],[4,133,107],[8,59,37,1,60,38],[8,44,20,4,45,21],[12,33,11,4,34,12],[3,145,115,1,146,116],[4,64,40,5,65,41],[11,36,16,5,37,17],[11,36,12,5,37,13],[5,109,87,1,110,88],[5,65,41,5,66,42],[5,54,24,7,55,25],[11,36,12,7,37,13],[5,122,98,1,123,99],[7,73,45,3,74,46],[15,43,19,2,44,20],[3,45,15,13,46,16],[1,135,107,5,136,108],[10,74,46,1,75,47],[1,50,22,15,51,23],[2,42,14,17,43,15],[5,150,120,1,151,121],[9,69,43,4,70,44],[17,50,22,1,51,23],[2,42,14,19,43,15],[3,141,113,4,142,114],[3,70,44,11,71,45],[17,47,21,4,48,22],[9,39,13,16,40,14],[3,135,107,5,136,108],[3,67,41,13,68,42],[15,54,24,5,55,25],[15,43,15,10,44,16],[4,144,116,4,145,117],[17,68,42],[17,50,22,6,51,23],[19,46,16,6,47,17],[2,139,111,7,140,112],[17,74,46],[7,54,24,16,55,25],[34,37,13],[4,151,121,5,152,122],[4,75,47,14,76,48],[11,54,24,14,55,25],[16,45,15,14,46,16],[6,147,117,4,148,118],[6,73,45,14,74,46],[11,54,24,16,55,25],[30,46,16,2,47,17],[8,132,106,4,133,107],[8,75,47,13,76,48],[7,54,24,22,55,25],[22,45,15,13,46,16],[10,142,114,2,143,115],[19,74,46,4,75,47],[28,50,22,6,51,23],[33,46,16,4,47,17],[8,152,122,4,153,123],[22,73,45,3,74,46],[8,53,23,26,54,24],[12,45,15,28,46,16],[3,147,117,10,148,118],[3,73,45,23,74,46],[4,54,24,31,55,25],[11,45,15,31,46,16],[7,146,116,7,147,117],[21,73,45,7,74,46],[1,53,23,37,54,24],[19,45,15,26,46,16],[5,145,115,10,146,116],[19,75,47,10,76,48],[15,54,24,25,55,25],[23,45,15,25,46,16],[13,145,115,3,146,116],[2,74,46,29,75,47],[42,54,24,1,55,25],[23,45,15,28,46,16],[17,145,115],[10,74,46,23,75,47],[10,54,24,35,55,25],[19,45,15,35,46,16],[17,145,115,1,146,116],[14,74,46,21,75,47],[29,54,24,19,55,25],[11,45,15,46,46,16],[13,145,115,6,146,116],[14,74,46,23,75,47],[44,54,24,7,55,25],[59,46,16,1,47,17],[12,151,121,7,152,122],[12,75,47,26,76,48],[39,54,24,14,55,25],[22,45,15,41,46,16],[6,151,121,14,152,122],[6,75,47,34,76,48],[46,54,24,10,55,25],[2,45,15,64,46,16],[17,152,122,4,153,123],[29,74,46,14,75,47],[49,54,24,10,55,25],[24,45,15,46,46,16],[4,152,122,18,153,123],[13,74,46,32,75,47],[48,54,24,14,55,25],[42,45,15,32,46,16],[20,147,117,4,148,118],[40,75,47,7,76,48],[43,54,24,22,55,25],[10,45,15,67,46,16],[19,148,118,6,149,119],[18,75,47,31,76,48],[34,54,24,34,55,25],[20,45,15,61,46,16]],r=function(t,r){var e={};return e.totalCount=t,e.dataCount=r,e},e={};return e.getRSBlocks=function(e,n){var o=function(r,e){switch(e){case g.L:return t[4*(r-1)+0];case g.M:return t[4*(r-1)+1];case g.Q:return t[4*(r-1)+2];case g.H:return t[4*(r-1)+3];default:return}}(e,n);if(void 0===o)throw"bad rs block @ typeNumber:"+e+"/errorCorrectionLevel:"+n;for(var i=o.length/3,a=[],u=0;u<i;u+=1)for(var f=o[3*u+0],c=o[3*u+1],l=o[3*u+2],h=0;h<f;h+=1)a.push(r(c,l));return a},e}(),b=function(){var t=[],r=0,e={getBuffer:function(){return t},getAt:function(r){var e=Math.floor(r/8);return 1==(t[e]>>>7-r%8&1)},put:function(t,r){for(var n=0;n<r;n+=1)e.putBit(1==(t>>>r-n-1&1))},getLengthInBits:function(){return r},putBit:function(e){var n=Math.floor(r/8);t.length<=n&&t.push(0),e&&(t[n]|=128>>>r%8),r+=1}};return e},M=function(t){var r=a,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+2<r.length;)t.put(o(r.substring(n,n+3)),10),n+=3;n<r.length&&(r.length-n==1?t.put(o(r.substring(n,n+1)),4):r.length-n==2&&t.put(o(r.substring(n,n+2)),7))}},o=function(t){for(var r=0,e=0;e<t.length;e+=1)r=10*r+i(t.charAt(e));return r},i=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);throw"illegal char :"+t};return n},x=function(t){var r=u,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+1<r.length;)t.put(45*o(r.charAt(n))+o(r.charAt(n+1)),11),n+=2;n<r.length&&t.put(o(r.charAt(n)),6)}},o=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);if("A"<=t&&t<="Z")return t.charCodeAt(0)-"A".charCodeAt(0)+10;switch(t){case" ":return 36;case"$":return 37;case"%":return 38;case"*":return 39;case"+":return 40;case"-":return 41;case".":return 42;case"/":return 43;case":":return 44;default:throw"illegal char :"+t}};return n},m=function(r){var e=f,n=t.stringToBytes(r),o={getMode:function(){return e},getLength:function(t){return n.length},write:function(t){for(var r=0;r<n.length;r+=1)t.put(n[r],8)}};return o},L=function(r){var e=c,n=t.stringToBytesFuncs.SJIS;if(!n)throw"sjis not supported.";!function(){var t=n("友");if(2!=t.length||38726!=(t[0]<<8|t[1]))throw"sjis not supported."}();var o=n(r),i={getMode:function(){return e},getLength:function(t){return~~(o.length/2)},write:function(t){for(var r=o,e=0;e+1<r.length;){var n=(255&r[e])<<8|255&r[e+1];if(33088<=n&&n<=40956)n-=33088;else{if(!(57408<=n&&n<=60351))throw"illegal char at "+(e+1)+"/"+n;n-=49472}n=192*(n>>>8&255)+(255&n),t.put(n,13),e+=2}if(e<r.length)throw"illegal char at "+(e+1)}};return i},D=function(){var t=[],r={writeByte:function(r){t.push(255&r)},writeShort:function(t){r.writeByte(t),r.writeByte(t>>>8)},writeBytes:function(t,e,n){e=e||0,n=n||t.length;for(var o=0;o<n;o+=1)r.writeByte(t[o+e])},writeString:function(t){for(var e=0;e<t.length;e+=1)r.writeByte(t.charCodeAt(e))},toByteArray:function(){return t},toString:function(){var r="";r+="[";for(var e=0;e<t.length;e+=1)e>0&&(r+=","),r+=t[e];return r+="]"}};return r},S=function(t){var r=t,e=0,n=0,o=0,i={read:function(){for(;o<8;){if(e>=r.length){if(0==o)return-1;throw"unexpected end of file./"+o}var t=r.charAt(e);if(e+=1,"="==t)return o=0,-1;t.match(/^\\s$/)||(n=n<<6|a(t.charCodeAt(0)),o+=6)}var i=n>>>o-8&255;return o-=8,i}},a=function(t){if(65<=t&&t<=90)return t-65;if(97<=t&&t<=122)return t-97+26;if(48<=t&&t<=57)return t-48+52;if(43==t)return 62;if(47==t)return 63;throw"c:"+t};return i},I=function(t,r,e){for(var n=function(t,r){var e=t,n=r,o=new Array(t*r),i={setPixel:function(t,r,n){o[r*e+t]=n},write:function(t){t.writeString("GIF87a"),t.writeShort(e),t.writeShort(n),t.writeByte(128),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(255),t.writeByte(255),t.writeByte(255),t.writeString(","),t.writeShort(0),t.writeShort(0),t.writeShort(e),t.writeShort(n),t.writeByte(0);var r=a(2);t.writeByte(2);for(var o=0;r.length-o>255;)t.writeByte(255),t.writeBytes(r,o,255),o+=255;t.writeByte(r.length-o),t.writeBytes(r,o,r.length-o),t.writeByte(0),t.writeString(";")}},a=function(t){for(var r=1<<t,e=1+(1<<t),n=t+1,i=u(),a=0;a<r;a+=1)i.add(String.fromCharCode(a));i.add(String.fromCharCode(r)),i.add(String.fromCharCode(e));var f,c,g,l=D(),h=(f=l,c=0,g=0,{write:function(t,r){if(t>>>r!=0)throw"length over";for(;c+r>=8;)f.writeByte(255&(t<<c|g)),r-=8-c,t>>>=8-c,g=0,c=0;g|=t<<c,c+=r},flush:function(){c>0&&f.writeByte(g)}});h.write(r,n);var s=0,v=String.fromCharCode(o[s]);for(s+=1;s<o.length;){var d=String.fromCharCode(o[s]);s+=1,i.contains(v+d)?v+=d:(h.write(i.indexOf(v),n),i.size()<4095&&(i.size()==1<<n&&(n+=1),i.add(v+d)),v=d)}return h.write(i.indexOf(v),n),h.write(e,n),h.flush(),l.toByteArray()},u=function(){var t={},r=0,e={add:function(n){if(e.contains(n))throw"dup key:"+n;t[n]=r,r+=1},size:function(){return r},indexOf:function(r){return t[r]},contains:function(r){return void 0!==t[r]}};return e};return i}(t,r),o=0;o<r;o+=1)for(var i=0;i<t;i+=1)n.setPixel(i,o,e(i,o));var a=D();n.write(a);for(var u=function(){var t=0,r=0,e=0,n="",o={},i=function(t){n+=String.fromCharCode(a(63&t))},a=function(t){if(t<0);else{if(t<26)return 65+t;if(t<52)return t-26+97;if(t<62)return t-52+48;if(62==t)return 43;if(63==t)return 47}throw"n:"+t};return o.writeByte=function(n){for(t=t<<8|255&n,r+=8,e+=1;r>=6;)i(t>>>r-6),r-=6},o.flush=function(){if(r>0&&(i(t<<6-r),t=0,r=0),e%3!=0)for(var o=3-e%3,a=0;a<o;a+=1)n+="="},o.toString=function(){return n},o}(),f=a.toByteArray(),c=0;c<f.length;c+=1)u.writeByte(f[c]);return u.flush(),"data:image/gif;base64,"+u};return t}();qrcode.stringToBytesFuncs["UTF-8"]=function(t){return function(t){for(var r=[],e=0;e<t.length;e++){var n=t.charCodeAt(e);n<128?r.push(n):n<2048?r.push(192|n>>6,128|63&n):n<55296||n>=57344?r.push(224|n>>12,128|n>>6&63,128|63&n):(e++,n=65536+((1023&n)<<10|1023&t.charCodeAt(e)),r.push(240|n>>18,128|n>>12&63,128|n>>6&63,128|63&n))}return r}(t)},function(t){"function"==typeof define&&define.amd?define([],t):"object"==typeof exports&&(module.exports=t())}(function(){return qrcode});
</script>
<style>
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:#f4f4f5;color:#18181b;font-size:15px}
  header{background:#E30613;color:#fff;padding:14px 16px;display:flex;align-items:center;justify-content:space-between}
  header b{font-size:19px;letter-spacing:1px}
  header a{color:#fff;font-size:13px;opacity:.9}
  nav{display:flex;background:#fff;border-bottom:1px solid #e4e4e7;position:sticky;top:0;z-index:5}
  nav button{flex:1;border:0;background:none;padding:14px 6px;font-size:15px;font-weight:700;color:#71717a;border-bottom:3px solid transparent}
  nav button.on{color:#E30613;border-color:#E30613}
  main{padding:14px;max-width:1100px;margin:0 auto}
  .card{background:#fff;border-radius:14px;padding:14px;box-shadow:0 1px 3px rgba(0,0,0,.07);margin-bottom:12px}
  .row{display:flex;gap:8px;flex-wrap:wrap;align-items:end}
  .row > *{flex:1;min-width:130px}
  label{font-size:12px;color:#71717a;font-weight:600;display:block;margin-bottom:3px}
  input,select{width:100%;padding:10px;border:1.5px solid #d4d4d8;border-radius:10px;font-size:15px;background:#fff}
  .btn{border:0;border-radius:10px;padding:11px 14px;font-size:15px;font-weight:700;color:#fff;background:#E30613;cursor:pointer}
  .btn.light{background:#f4f4f5;color:#27272a}.btn.dark{background:#27272a}.btn.sm{padding:7px 10px;font-size:13px}
  .quick{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px}
  .quick button{border:1.5px solid #e4e4e7;background:#fff;border-radius:20px;padding:7px 13px;font-weight:600;font-size:14px}
  .quick button.on{background:#E30613;color:#fff;border-color:#E30613}
  .stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px}
  .stat{background:#fff;border-radius:14px;padding:12px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.07)}
  .stat b{display:block;font-size:26px}
  .stat span{font-size:12px;color:#71717a;font-weight:600}
  .tw{overflow-x:auto}
  table{width:100%;border-collapse:collapse;background:#fff}
  th{text-align:left;font-size:12px;color:#71717a;padding:9px 8px;border-bottom:2px solid #f0f0f0;white-space:nowrap}
  td{padding:10px 8px;border-bottom:1px solid #f4f4f5;white-space:nowrap;vertical-align:top}
  td.sm{font-size:12px;color:#71717a;white-space:normal;min-width:120px}
  .late{color:#ea580c;font-weight:700}.okc{color:#16a34a;font-weight:700}.muted{color:#a1a1aa}
  .tag{display:inline-block;border-radius:12px;padding:2px 9px;font-size:12px;font-weight:700}
  .tag.g{background:#dcfce7;color:#15803d}.tag.r{background:#fee2e2;color:#b91c1c}.tag.y{background:#fef3c7;color:#a16207}
  .acts{display:flex;gap:5px;flex-wrap:wrap}
  h3{margin:4px 0 10px;font-size:16px}
  .li{padding:12px 4px;border-bottom:1px solid #f0f0f0}
  .li:last-child{border-bottom:0}
  .li .top{display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap}
  .li .nm{font-weight:800;font-size:16px}
  .li .meta{display:flex;gap:12px;flex-wrap:wrap;margin-top:6px;font-size:15px}
  .li .rm{font-size:12px;color:#71717a;margin-top:5px}
  .ag{background:#f4f4f5;border-radius:10px;padding:2px 8px;font-size:12px;font-weight:600;color:#52525b}
  .hide{display:none!important}
  #toast{position:fixed;bottom:18px;left:50%;transform:translateX(-50%);background:#27272a;color:#fff;padding:11px 18px;border-radius:12px;font-weight:600;z-index:50;max-width:90%}
  #modal{position:fixed;inset:0;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;z-index:40}
  #modal .box{background:#fff;border-radius:16px;padding:18px;max-width:420px;width:92%;text-align:center}
  .print{border:3px solid #E30613;border-radius:16px;padding:16px}
  .print .t1{font-weight:900;color:#E30613;font-size:24px;letter-spacing:1px}
  .print .t2{font-size:20px;font-weight:800;margin:4px 0 8px}
  .print img{width:100%;max-width:320px;image-rendering:pixelated}
  .print .t3{font-size:15px;margin-top:6px}
  .login{max-width:360px;margin:60px auto}
  @media print{ body *{visibility:hidden} #printZone, #printZone *{visibility:visible} #printZone{position:absolute;left:0;top:0;width:100%} .noprint{display:none} }
</style>
</head>
<body>
<header><b>SONICOD · POINTAGE</b><a id="sheetLink" class="hide" href="#" target="_blank">📄 Google Sheet</a></header>

<div id="loginView" class="login card">
  <h3>Admin — __AGENCE__</h3>
  <div id="apiBox" class="hide"><label>Lien /exec dyal Google Script (merra wa7da)</label>
  <input id="apiIn" placeholder="https://script.google.com/macros/s/.../exec" style="margin-bottom:10px"></div>
  <input id="pwd" type="password" placeholder="Mot de passe" onkeydown="if(event.key==='Enter')login()">
  <button class="btn" style="width:100%;margin-top:10px" onclick="login()">Dkhol</button>
  <p id="loginErr" style="color:#dc2626"></p>
</div>

<div id="appView" class="hide">
  <nav>
    <button class="on" data-t="rep" onclick="tab('rep')">🕘 L7odor</button>
    <button data-t="dr" onclick="tab('dr')">👥 Drari</button>
    <button data-t="qr" onclick="tab('qr')">🔳 QR</button>
  </nav>
  <main>
    <!-- L7ODOR -->
    <section id="t-rep">
      <div class="card">
        <div class="quick" id="quick">
          <button data-q="0" onclick="quick(0)">Lyoum</button>
          <button data-q="1" onclick="quick(1)">Lbare7</button>
          <button data-q="7" onclick="quick(7)">7 iyam</button>
          <button data-q="30" onclick="quick(30)">Had chhar</button>
        </div>
        <div class="row">
          <div><label>Men</label><input type="date" id="from"></div>
          <div><label>7tal</label><input type="date" id="to"></div>
          <div><label>Smiya</label><input id="fNom" placeholder="Kolchi"></div>
          <div style="flex:0 0 auto"><button class="btn" onclick="report()">Filtrer</button></div>
        </div>
      </div>
      <div class="stats">
        <div class="stat"><b id="sP" class="okc">0</b><span>Jaw</span></div>
        <div class="stat"><b id="sR" class="late">0</b><span>Retard</span></div>
        <div class="stat"><b id="sA" style="color:#dc2626">–</b><span>Ma jawch</span></div>
      </div>
      <div class="card" id="repBody"></div>
      <div class="card" id="absCard">
        <h3>❌ Ma jawch</h3><div id="absList" class="muted"></div>
      </div>
    </section>

    <!-- DRARI -->
    <section id="t-dr" class="hide">
      <div class="card">
        <h3 id="drTitle">➕ Zid wa7ed</h3>
        <div class="row">
          <div><label>Smiya</label><input id="dNom"></div>
          <div><label>PIN (khawi = automatique)</label><input id="dPin" inputmode="numeric" maxlength="6"></div>
          <div style="flex:0 0 auto"><button class="btn" onclick="saveDrari()">Enregistrer</button>
            <button class="btn light hide" id="dCancel" onclick="resetDrForm()">Annuler</button></div>
        </div>
      </div>
      <div class="card" id="drBody"></div>
    </section>

    <!-- QR -->
    <section id="t-qr" class="hide">
      <div class="card" id="urlCard">
        <h3>🔗 Lien dyal l'app</h3>
        <p class="muted" style="font-size:13px;margin:0 0 8px">Copier l lien li f barre d'adresse lfo9 (li kaysali b <b>/exec</b>, bla ?admin) o colli hna.</p>
        <div class="row">
          <div><input id="sUrl" placeholder="https://script.google.com/macros/s/.../exec"></div>
          <div style="flex:0 0 auto"><button class="btn" onclick="setAct('url', { url: document.getElementById('sUrl').value })">Enregistrer</button></div>
        </div>
        <p id="urlWarn" class="late" style="font-size:13px"></p>
      </div>
      <div class="card" style="text-align:center">
        <div id="printZone"></div>
        <div class="acts noprint" style="justify-content:center;margin-top:12px">
          <button class="btn" onclick="window.print()">🖨️ Imprimer QR</button>
          <button class="btn light" onclick="setAct('code')">🔄 QR jdid</button>
        </div>
        <p class="muted noprint" style="font-size:12px;word-break:break-all" id="qrLink"></p>
      </div>
      <div class="card">
        <h3>⏰ Sa3a dyal lkhdma</h3>
        <div class="row">
          <div><label>Khasshom ykouno hna m3a</label><input id="sStart" type="time"></div>
          <div><label>Rayon GPS (m)</label><input id="sRayon" type="number"></div>
          <div style="flex:0 0 auto"><button class="btn" onclick="saveSettings()">Enregistrer</button></div>
        </div>
        <p class="muted" style="font-size:13px">Li ja ba3d had sa3a b ktar men 5 d9ayé9 kaytkteb lih retard.</p>
      </div>
      <div class="card">
        <h3>📍 Position dyal lagence</h3>
        <p id="gpsState"></p>
        <div class="acts">
          <button class="btn" onclick="setPos()">📍 Ana daba f lagence</button>
          <button class="btn light" id="noPosBtn" onclick="setAct('nopos')">Ms7 GPS</button>
        </div>
        <p class="muted" style="font-size:13px">Mli katkon f lagence, dghet had l bouton. Men b3d ma y9der 7ta wa7ed yscanni men dar (ila sowro l QR).</p>
      </div>
    </section>
  </main>
</div>

<div id="toast" class="hide"></div>

<script>
const APP_URL = '__APP_URL__';
const AGENCE = '__AGENCE__';
let PWD = '', D = null, editNom = null;

function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function attr(s){ return esc(String(s == null ? '' : s).replace(/\\\\/g,'\\\\\\\\').replace(/'/g,"\\\\'")); }
function call(fn){
  const args = Array.prototype.slice.call(arguments, 1);
  return new Promise((ok, ko) => google.script.run.withSuccessHandler(ok).withFailureHandler(ko)[fn].apply(null, args));
}
function toast(m){ const t = document.getElementById('toast'); t.textContent = m; t.classList.remove('hide'); clearTimeout(t._h); t._h = setTimeout(() => t.classList.add('hide'), 2800); }
function errMsg(e){ return (e && e.message ? e.message : String(e)).replace(/^Error:\\s*/, ''); }

async function login(){
  PWD = document.getElementById('pwd').value;
  if (typeof STATIC_SITE !== 'undefined') {
    const v = document.getElementById('apiIn').value.trim();
    const m = v.match(/\\/s\\/([A-Za-z0-9_-]+)\\/(exec|dev)/);
    if (v && !m) { document.getElementById('loginErr').textContent = 'Lien machi s7i7: khasso ykon fih /s/.../exec'; return; }
    if (m) { API_ID = m[1]; API_URL = 'https://script.google.com/macros/s/' + API_ID + '/exec'; try { localStorage.setItem('sonicod_pt_api', API_ID); } catch(e){} }
    if (!API_ID) { document.getElementById('loginErr').textContent = 'Colli lien /exec dyal Google Script.'; return; }
  }
  try {
    D = await call('adminData', PWD);
    try { localStorage.setItem('sonicod_pt_admin', PWD); } catch(e){}
    document.getElementById('loginView').classList.add('hide');
    document.getElementById('appView').classList.remove('hide');
    const sl = document.getElementById('sheetLink'); sl.href = D.sheetUrl; sl.classList.remove('hide');
    renderAll(); quick(0);
  } catch(e){ document.getElementById('loginErr').textContent = errMsg(e); try{localStorage.removeItem('sonicod_pt_admin');}catch(x){} }
}

function tab(t){
  document.querySelectorAll('nav button').forEach(b => b.classList.toggle('on', b.dataset.t === t));
  ['rep','dr','qr'].forEach(x => document.getElementById('t-' + x).classList.toggle('hide', x !== t));
}
function renderAll(){ renderDrari(); renderQR(); }

/* ---------- L7ODOR ---------- */
function addDays(iso, n){ const d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() + n); return d.toISOString().slice(0,10); }
function quick(n){
  const t = D.today;
  let from = t, to = t;
  if (n === 1) { from = to = addDays(t, -1); }
  if (n === 7) { from = addDays(t, -6); }
  if (n === 30) { from = t.slice(0, 8) + '01'; }
  document.getElementById('from').value = from; document.getElementById('to').value = to;
  document.querySelectorAll('#quick button').forEach(b => b.classList.toggle('on', Number(b.dataset.q) === n));
  report(true);
}
async function report(fromQuick){
  if (fromQuick !== true) document.querySelectorAll('#quick button').forEach(b => b.classList.remove('on'));
  const f = { from: document.getElementById('from').value, to: document.getElementById('to').value, nom: document.getElementById('fNom').value };
  if (f.to && f.from && f.to < f.from) { const x = f.from; f.from = f.to; f.to = x; }
  const body = document.getElementById('repBody');
  body.innerHTML = '<div class="muted">Kaytchargé...</div>';
  let r;
  try { r = await call('adminReport', PWD, f); } catch(e){ body.innerHTML = ''; return toast(errMsg(e)); }
  let late = 0; r.days.forEach(d => { if (Number(d.retard) > 0) late++; });
  document.getElementById('sP').textContent = r.days.length;
  document.getElementById('sR').textContent = late;
  const single = r.from === r.to;
  document.getElementById('sA').textContent = single ? r.absents.length : '–';
  document.getElementById('absCard').classList.toggle('hide', !single);
  document.getElementById('absList').innerHTML = r.absents.length
    ? r.absents.map(n => '<span class="tag r" style="margin:3px">' + esc(n) + '</span>').join('')
    : '<span class="okc">Kolchi ja ✔️</span>';
  body.innerHTML = r.days.length ? r.days.map(d => {
    const ret = Number(d.retard) > 0 ? '<span class="late">⏰ Retard ' + esc(d.retard) + ' min</span>' : '<span class="okc">✔ F lwe9t</span>';
    return '<div class="li"><div class="top"><span class="nm">' + esc(d.nom) + '</span>' +
      '<span class="muted" style="font-size:13px">' + esc(fmtDate(d.date)) + '</span></div>' +
      '<div class="meta"><span>Ja m3a <b>' + esc(d.heure) + '</b></span>' + ret +
      (d.gps === 'Non' ? '<span class="muted" style="font-size:12px">bla GPS</span>' : '') + '</div></div>';
  }).join('') : '<div class="muted">7ta wa7ed ma tsjjel f had tarikh.</div>';
}
function fmtDate(iso){
  const j = ['7ad','Tnin','Tlat','Larb3','Lkhmis','Jem3a','Sebt'];
  const d = new Date(iso + 'T12:00:00');
  return j[d.getDay()] + ' ' + iso.slice(8,10) + '/' + iso.slice(5,7);
}

/* ---------- DRARI ---------- */
function renderDrari(){
  document.getElementById('drBody').innerHTML = D.drari.length ? D.drari.map(d =>
    '<div class="li"><div class="top"><span class="nm">' + esc(d.nom) + '</span>' +
    '<span>PIN <b>' + esc(d.pin) + '</b></span></div>' +
    '<div class="meta" style="font-size:13px">' + (d.lie ? '<span class="tag g">📱 Mrbout</span>' : '<span class="tag y">📱 Mazal ma rbet</span>') +
    (d.actif ? '<span class="tag g">Actif</span>' : '<span class="tag r">M7bous</span>') + '</div>' +
    '<div class="acts" style="margin-top:8px">' +
      '<button class="btn sm light" onclick="editDr(\\'' + attr(d.nom) + '\\')">✏️</button>' +
      (d.lie ? '<button class="btn sm light" onclick="drAct(\\'' + attr(d.nom) + '\\',\\'reset\\')">Reset tel</button>' : '') +
      '<button class="btn sm light" onclick="drAct(\\'' + attr(d.nom) + '\\',\\'pin\\')">PIN jdid</button>' +
      '<button class="btn sm light" onclick="drAct(\\'' + attr(d.nom) + '\\',\\'toggle\\')">' + (d.actif ? 'Bloquer' : 'Activer') + '</button>' +
      '<button class="btn sm light" onclick="drAct(\\'' + attr(d.nom) + '\\',\\'delete\\')">🗑️</button>' +
    '</div></div>').join('') : '<div class="muted">Zid drari lfo9 👆</div>';
}
function editDr(nom){
  const d = D.drari.find(x => x.nom === nom); if (!d) return;
  editNom = nom;
  document.getElementById('dNom').value = d.nom; document.getElementById('dPin').value = d.pin;
  document.getElementById('drTitle').textContent = '✏️ Bddel ' + d.nom;
  document.getElementById('dCancel').classList.remove('hide');
  window.scrollTo(0, 0);
}
function resetDrForm(){
  editNom = null; ['dNom','dPin'].forEach(i => document.getElementById(i).value = '');
  document.getElementById('drTitle').textContent = '➕ Zid wa7ed'; document.getElementById('dCancel').classList.add('hide');
}
async function saveDrari(){
  const d = { oldNom: editNom, nom: document.getElementById('dNom').value, pin: document.getElementById('dPin').value };
  try { D = await call('adminSaveDrari', PWD, d); renderAll(); resetDrForm(); toast('Tsjjel ✔️'); } catch(e){ toast(errMsg(e)); }
}
async function drAct(nom, a){
  const q = { delete: 'Tms7 ' + nom + ' ?', reset: 'Reset telephone dyal ' + nom + ' ? (ghadi yrbet telephone jdid)', pin: 'PIN jdid l ' + nom + ' ?' }[a];
  if (q && !confirm(q)) return;
  try { D = await call('adminDrariAction', PWD, nom, a); renderAll(); toast('Safi ✔️'); } catch(e){ toast(errMsg(e)); }
}

/* ---------- QR & REGLAGES ---------- */
function qrImg(text){
  if (typeof qrcode === 'undefined') return '';
  const q = qrcode(0, 'M'); q.addData(text); q.make();
  return q.createDataURL(10, 2);
}
function renderQR(){
  const s = D.settings;
  const base = (typeof STATIC_SITE !== 'undefined') ? STATIC_SITE : (s.url || D.url || APP_URL);
  const link = base + '?a=' + encodeURIComponent(s.code) + ((typeof STATIC_SITE !== 'undefined') ? '&s=' + API_ID : '');
  if (typeof STATIC_SITE !== 'undefined') document.getElementById('urlCard').classList.add('hide');
  document.getElementById('sUrl').value = s.url || '';
  document.getElementById('urlWarn').textContent = s.url ? '' : '⚠️ Colli l lien 9bel ma t-imprimi l QR.';
  document.getElementById('printZone').innerHTML =
    '<div class="print"><div class="t1">SONICOD</div><div class="t2">Pointage · ' + esc(AGENCE) + '</div>' +
    '<img src="' + qrImg(link) + '">' +
    '<div class="t3">📷 Scanni b camera dyal telephone<br>mli tji l lagence</div></div>';
  document.getElementById('qrLink').textContent = link;
  document.getElementById('sStart').value = String(s.start).padStart(5, '0');
  document.getElementById('sRayon').value = s.rayon;
  document.getElementById('gpsState').innerHTML = s.lat !== null
    ? '<span class="tag g">📍 GPS mfa3el — ' + s.rayon + ' m</span>'
    : '<span class="tag y">GPS mazal ma tsjjel</span>';
  document.getElementById('noPosBtn').classList.toggle('hide', s.lat === null);
}
async function setAct(a, extra){
  const q = { code: 'QR jdid ? L QR l9dim ma yb9ach ykhdem, khassk t-imprimi jdid.', nopos: 'Tms7 GPS ? (ay wa7ed y9der yscanni men ay blassa)' }[a];
  if (q && !confirm(q)) return;
  try { D = await call('adminSettings', PWD, a, extra || {}); renderAll(); toast('Safi ✔️'); } catch(e){ toast(errMsg(e)); }
}
function saveSettings(){ setAct('save', { start: document.getElementById('sStart').value, rayon: document.getElementById('sRayon').value }); }
function setPos(){
  if (!navigator.geolocation) return toast('GPS ma khdamch f had navigateur');
  toast('📍 Kanchouf l position...');
  navigator.geolocation.getCurrentPosition(p => {
    if (confirm('Nsjjel had position (±' + Math.round(p.coords.accuracy) + 'm) l lagence ?')) setAct('pos', { lat: p.coords.latitude, lng: p.coords.longitude });
  }, () => toast('Ma 9derch njib l position. 3ti l autorisation dyal Localisation.'), { enableHighAccuracy: true, timeout: 15000 });
}

(function(){
  if (typeof STATIC_SITE !== 'undefined') {
    const box = document.getElementById('apiBox'); box.classList.remove('hide');
    if (API_ID) document.getElementById('apiIn').value = 'https://script.google.com/macros/s/' + API_ID + '/exec';
  }
  let p = null; try { p = localStorage.getItem('sonicod_pt_admin'); } catch(e){}
  if (p) { document.getElementById('pwd').value = p; login(); }
})();
</script>
</body>
</html>
`;
