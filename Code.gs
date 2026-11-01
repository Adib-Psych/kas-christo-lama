/**
 * Kas Christo API — Google Apps Script (web app, Execute as: Adib, Access: Anyone)
 * Data: Google Sheet "Kas Christo — Data" di folder Drive "Kas Christo".
 * Akses: link khusus (#k=<kunci>) — satu kunci untuk Adib (pemilik), satu untuk Christo.
 * Cadangan: JSON mingguan ke folder "Kas Christo/Cadangan" (pemicu Senin pagi).
 *
 * Fungsi yang dijalankan sekali dari editor: siapkan()  → membuat folder, Sheet, kunci, pemicu.
 * Fungsi bantu dari editor: tampilkanLink() · gantiKunciChristo() · cadanganMingguan()
 */
const PAGE_URL = 'https://adib-psych.github.io/kas-christo/';
const POS_KELUAR = ['bonsari', 'kebun', 'rumah', 'pribadi', 'pindah', 'bank', 'lain'];
const POS_MASUK = ['bonsari', 'kebun', 'rumah', 'pribadi', 'pindah', 'lain'];
const COLS = ['tx', 'config'];
const MAX_NOTA = 6 * 1024 * 1024;

function props_() { return PropertiesService.getScriptProperties(); }

function siapkan() {
  const p = props_();
  let folder = p.getProperty('FOLDER_ID') ? DriveApp.getFolderById(p.getProperty('FOLDER_ID')) : DriveApp.createFolder('Kas Christo');
  p.setProperty('FOLDER_ID', folder.getId());
  if (!p.getProperty('SHEET_ID')) {
    const ss = SpreadsheetApp.create('Kas Christo — Data');
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    const d = ss.getSheets()[0]; d.setName('Data');
    d.getRange(1, 1, 1, 4).setValues([['koleksi', 'id', 'json', 'diubah']]).setFontWeight('bold');
    d.setFrozenRows(1);
    ss.insertSheet('Transaksi');
    ss.insertSheet('Log');
    ss.getSheetByName('Log').getRange(1, 1, 1, 5).setValues([['waktu', 'peran', 'op', 'koleksi', 'id']]).setFontWeight('bold');
    p.setProperty('SHEET_ID', ss.getId());
  }
  ['Cadangan', 'Nota'].forEach(n => {
    const k = 'FOLDER_' + n.toUpperCase();
    if (!p.getProperty(k)) { const it = folder.getFoldersByName(n); p.setProperty(k, (it.hasNext() ? it.next() : folder.createFolder(n)).getId()); }
  });
  if (!p.getProperty('KEY_ADIB')) p.setProperty('KEY_ADIB', kunciBaru_());
  if (!p.getProperty('KEY_CHRISTO')) p.setProperty('KEY_CHRISTO', kunciBaru_());
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'cadanganMingguan')) {
    ScriptApp.newTrigger('cadanganMingguan').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).inTimezone('Asia/Jakarta').create();
  }
  tampilkanLink();
}

function kunciBaru_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function tampilkanLink() {
  const p = props_();
  Logger.log('Link Adib   : ' + PAGE_URL + '#k=' + p.getProperty('KEY_ADIB'));
  Logger.log('Link Christo: ' + PAGE_URL + '#k=' + p.getProperty('KEY_CHRISTO'));
  Logger.log('Sheet       : https://docs.google.com/spreadsheets/d/' + p.getProperty('SHEET_ID'));
}
function gantiKunciChristo() { props_().setProperty('KEY_CHRISTO', kunciBaru_()); tampilkanLink(); }

// ---------------------------------------------------------------- storage
function sheet_() { return SpreadsheetApp.openById(props_().getProperty('SHEET_ID')).getSheetByName('Data'); }
function readAll_() {
  const sh = sheet_(); const n = sh.getLastRow(); const store = {}; const index = {};
  COLS.forEach(c => store[c] = {});
  if (n > 1) sh.getRange(2, 1, n - 1, 3).getValues().forEach((r, i) => {
    if (!r[0] || !r[1]) return;
    try { (store[r[0]] = store[r[0]] || {})[r[1]] = JSON.parse(r[2]); index[r[0] + '/' + r[1]] = i + 2; } catch (e) {}
  });
  return { store, index };
}
function put_(ctx, col, id, obj) {
  const sh = sheet_(); const key = col + '/' + id; const now = new Date().toISOString();
  const json = JSON.stringify(obj);
  if (json.length > 45000) throw new Error('too_large');
  if (ctx.index[key]) sh.getRange(ctx.index[key], 3, 1, 2).setValues([[json, now]]);
  else { sh.appendRow([col, id, json, now]); ctx.index[key] = sh.getLastRow(); }
  ctx.store[col] = ctx.store[col] || {}; ctx.store[col][id] = obj;
}
function del_(ctx, col, id) {
  const key = col + '/' + id; const row = ctx.index[key];
  if (row) { sheet_().deleteRow(row); Object.keys(ctx.index).forEach(k => { if (ctx.index[k] > row) ctx.index[k]--; }); delete ctx.index[key]; }
  if (ctx.store[col]) delete ctx.store[col][id];
}
function log_(role, op, col, id) {
  try { SpreadsheetApp.openById(props_().getProperty('SHEET_ID')).getSheetByName('Log').appendRow([new Date(), role, op, col || '', id || '']); } catch (e) {}
}

// Sheet "Transaksi" = tampilan yang mudah dibaca Adib (dibangun ulang tiap ada perubahan)
function rebuildView_(store) {
  const ss = SpreadsheetApp.openById(props_().getProperty('SHEET_ID'));
  const sh = ss.getSheetByName('Transaksi') || ss.insertSheet('Transaksi');
  const rows = Object.values(store.tx || {}).sort((a, b) => (a.rek + a.tanggal + (a.jam || '')).localeCompare(b.rek + b.tanggal + (b.jam || '')));
  const out = [['Rekening', 'Tanggal', 'Jam', 'Arah', 'Jumlah', 'Biaya', 'Total', 'Uraian bank', 'Kantong', 'Catatan', 'Dicek Adib', 'Sumber', 'Saldo bank', 'ID']];
  rows.forEach(t => {
    const tot = (t.jumlah || 0) + (t.biaya || 0);
    const kant = (t.alokasi || []).map(a => a.pos + (t.alokasi.length > 1 ? ' ' + a.jumlah : '')).join(' + ');
    out.push([t.rek, t.tanggal, t.jam || '', t.arah, t.jumlah || 0, t.biaya || 0, t.arah === 'keluar' ? -tot : tot, t.uraian || '', kant, t.ket || '', t.dikunci ? 'ya' : '', t.sumber || '', t.saldoBank == null ? '' : t.saldoBank, t.id]);
  });
  sh.clearContents();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(1, 1, 1, out[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

// ---------------------------------------------------------------- web app
function doGet() { return json_({ ok: true, app: 'Kas Christo API' }); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function roleOf_(key) {
  const p = props_();
  if (key && key === p.getProperty('KEY_ADIB')) return 'adib';
  if (key && key === p.getProperty('KEY_CHRISTO')) return 'christo';
  return null;
}

function doPost(e) {
  let body; try { body = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'bad_json' }); }
  const role = roleOf_(body.key);
  if (!role) return json_({ ok: false, error: 'kunci' });
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (x) { return json_({ ok: false, error: 'sibuk' }); }
  try {
    const ctx = readAll_();
    const res = handle_(role, body, ctx);
    if (res.changed) { try { rebuildView_(ctx.store); } catch (x) {} }
    const out = Object.assign({ ok: true, role, data: ctx.store }, res.extra || {});
    if (role === 'adib') out.linkChristo = PAGE_URL + '#k=' + props_().getProperty('KEY_CHRISTO');
    return json_(out);
  } catch (x) {
    return json_({ ok: false, error: String(x && x.message || x) });
  } finally { lock.releaseLock(); }
}

function deny_() { throw new Error('tidak diizinkan'); }
function totalOf_(t) { return Math.round((+t.jumlah || 0) + (+t.biaya || 0)); }
function cleanAlok_(t, alok) {
  if (!Array.isArray(alok)) return [];
  const allowed = t.arah === 'masuk' ? POS_MASUK : POS_KELUAR;
  const a = alok.filter(x => x && allowed.indexOf(x.pos) >= 0).map(x => ({ pos: x.pos, jumlah: Math.round(+x.jumlah || 0), ket: String(x.ket || '').slice(0, 160) }));
  if (!a.length) return [];
  const sum = a.reduce((s, x) => s + x.jumlah, 0);
  if (sum !== totalOf_(t)) throw new Error('alokasi_tidak_cocok');
  return a;
}
function str_(v, n) { return String(v == null ? '' : v).slice(0, n || 300); }

function handle_(role, b, ctx) {
  const op = b.op, col = b.col, id = b.id;
  const isAdib = role === 'adib';
  const now = new Date().toISOString();
  if (op === 'load') return {};

  if (op === 'export') { if (!isAdib) deny_(); return { extra: { exportedAt: now } }; }

  if (op === 'upload') {
    const bytes = Utilities.base64Decode(b.base64 || '');
    if (bytes.length > MAX_NOTA) throw new Error('too_large');
    const type = /^(image\/(jpeg|png|webp|heic|heif)|application\/pdf)$/.test(b.type) ? b.type : 'image/jpeg';
    const f = DriveApp.getFolderById(props_().getProperty('FOLDER_NOTA')).createFile(Utilities.newBlob(bytes, type, 'nota-' + now.replace(/[:.]/g, '-')));
    log_(role, 'upload', 'nota', f.getId());
    return { extra: { id: f.getId(), contentType: type } };
  }
  if (op === 'nota') {
    const f = DriveApp.getFileById(String(id));
    const parents = f.getParents(); let okFolder = false;
    while (parents.hasNext()) if (parents.next().getId() === props_().getProperty('FOLDER_NOTA')) okFolder = true;
    if (!okFolder) deny_();
    const blob = f.getBlob();
    return { extra: { contentType: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) } };
  }

  if (op === 'import') {
    if (!isAdib) deny_();
    const list = (b.data && b.data.transaksi) || [];
    let baru = 0, diperbarui = 0, digabung = 0;
    const manual = Object.values(ctx.store.tx || {}).filter(t => t.sumber === 'manual');
    list.forEach(n => {
      if (!n || !n.id || !/^(BRI|BCA)$/.test(n.rek)) return;
      const ex = ctx.store.tx[n.id];
      const bank = { id: n.id, rek: n.rek, tanggal: str_(n.tanggal, 10), jam: str_(n.jam, 5), arah: n.arah === 'masuk' ? 'masuk' : 'keluar', jumlah: Math.round(+n.jumlah || 0), biaya: Math.round(+n.biaya || 0), uraian: str_(n.uraian, 200), saldoBank: n.saldoBank == null ? null : Math.round(+n.saldoBank), sumber: 'mutasi' };
      if (ex) { put_(ctx, 'tx', n.id, Object.assign({}, ex, bank)); diperbarui++; return; }
      const m = manual.find(t => t.rek === bank.rek && t.tanggal === bank.tanggal && t.arah === bank.arah && totalOf_(t) === totalOf_(bank) && !t._dipakai);
      let rec = Object.assign({ alokasi: [], ket: '', nota: [], dibuat: now }, bank);
      if (m) { m._dipakai = true; rec = Object.assign(rec, { alokasi: m.alokasi || [], ket: m.ket || '', nota: m.nota || [], dipilahOleh: m.dipilahOleh || '', dipilahPada: m.dipilahPada || '', dikunci: !!m.dikunci, dariManual: m.id }); del_(ctx, 'tx', m.id); digabung++; }
      else if (Array.isArray(n.alokasi) && n.alokasi.length) { rec.alokasi = cleanAlok_(bank, n.alokasi); rec.dipilahOleh = 'otomatis'; }
      put_(ctx, 'tx', n.id, rec); baru++;
    });
    if (b.data && b.data.saldoAwal) {
      const cfg = ctx.store.config.main || {}; cfg.saldoAwal = Object.assign({}, cfg.saldoAwal || {}, b.data.saldoAwal); put_(ctx, 'config', 'main', cfg);
    }
    log_(role, 'import', 'tx', baru + ' baru · ' + diperbarui + ' diperbarui · ' + digabung + ' digabung');
    return { changed: true, extra: { hasil: { baru, diperbarui, digabung } } };
  }

  if (COLS.indexOf(col) < 0) throw new Error('koleksi');
  if (col === 'config') {
    if (!isAdib) deny_();
    if (op !== 'set' && op !== 'update') throw new Error('op');
    const cur = ctx.store.config[id] || {};
    const v = op === 'set' ? b.data : Object.assign({}, cur, b.data);
    put_(ctx, 'config', id, v); log_(role, op, col, id); return { changed: false };
  }

  // ---- transaksi
  if (op === 'add') {
    const d = b.data || {};
    const nid = 'man-' + Utilities.getUuid().slice(0, 12);
    const t = { id: nid, rek: d.rek === 'BCA' ? 'BCA' : 'BRI', tanggal: str_(d.tanggal, 10), jam: str_(d.jam, 5), arah: d.arah === 'masuk' ? 'masuk' : 'keluar', jumlah: Math.round(+d.jumlah || 0), biaya: Math.round(+d.biaya || 0), uraian: str_(d.uraian, 200), sumber: 'manual', ket: str_(d.ket, 300), nota: Array.isArray(d.nota) ? d.nota.slice(0, 6).map(String) : [], dibuat: now, oleh: role };
    if (!/^\d{4}-\d{2}-\d{2}$/.test(t.tanggal) || t.jumlah <= 0) throw new Error('isian');
    if (t.arah === 'masuk' && !isAdib) deny_();
    t.alokasi = cleanAlok_(t, d.alokasi);
    if (t.alokasi.length) { t.dipilahOleh = role; t.dipilahPada = now; }
    put_(ctx, 'tx', nid, t); log_(role, 'add', 'tx', nid);
    return { changed: true, extra: { id: nid } };
  }
  const cur = ctx.store.tx[id];
  if (!cur) throw new Error('tidak_ada');
  if (op === 'delete') {
    if (!isAdib && !(cur.sumber === 'manual' && cur.oleh === 'christo' && !cur.dikunci)) deny_();
    if (cur.sumber === 'mutasi' && !isAdib) deny_();
    del_(ctx, 'tx', id); log_(role, 'delete', 'tx', id); return { changed: true };
  }
  if (op !== 'update') throw new Error('op');
  const d = b.data || {}; const next = Object.assign({}, cur);
  if (!isAdib) {
    if (cur.dikunci) deny_();
    if ('ket' in d) next.ket = str_(d.ket, 300);
    if ('nota' in d) next.nota = Array.isArray(d.nota) ? d.nota.slice(0, 6).map(String) : [];
    if ('alokasi' in d) {
      if (cur.arah === 'masuk') deny_();
      next.alokasi = cleanAlok_(cur, d.alokasi); next.dipilahOleh = 'christo'; next.dipilahPada = now;
    }
    if (cur.sumber === 'manual' && cur.oleh === 'christo') {
      ['tanggal', 'uraian'].forEach(k => { if (k in d) next[k] = str_(d[k], k === 'tanggal' ? 10 : 200); });
      if ('jumlah' in d) next.jumlah = Math.round(+d.jumlah || 0);
      if ('rek' in d) next.rek = d.rek === 'BCA' ? 'BCA' : 'BRI';
      if ('jumlah' in d && !('alokasi' in d) && (next.alokasi || []).length === 1) next.alokasi = [{ pos: next.alokasi[0].pos, jumlah: totalOf_(next), ket: '' }];
    }
  } else {
    ['ket', 'uraian', 'tanggal', 'jam'].forEach(k => { if (k in d) next[k] = str_(d[k], 300); });
    if ('nota' in d) next.nota = Array.isArray(d.nota) ? d.nota.slice(0, 6).map(String) : [];
    if ('catatanAdib' in d) next.catatanAdib = str_(d.catatanAdib, 300);
    if ('dikunci' in d) { next.dikunci = !!d.dikunci; next.dikunciPada = d.dikunci ? now : ''; }
    if (cur.sumber === 'manual') {
      if ('jumlah' in d) next.jumlah = Math.round(+d.jumlah || 0);
      if ('rek' in d) next.rek = d.rek === 'BCA' ? 'BCA' : 'BRI';
    }
    if ('alokasi' in d) { next.alokasi = cleanAlok_(next, d.alokasi); next.dipilahOleh = 'adib'; next.dipilahPada = now; }
  }
  put_(ctx, 'tx', id, next); log_(role, 'update', 'tx', id);
  return { changed: true };
}

// ---------------------------------------------------------------- cadangan
function cadanganMingguan() {
  const ctx = readAll_();
  const stamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const payload = JSON.stringify({ dibuat: new Date().toISOString(), sumber: 'Kas Christo API', data: ctx.store }, null, 1);
  DriveApp.getFolderById(props_().getProperty('FOLDER_CADANGAN')).createFile('Kas Christo ' + stamp + '.json', payload, 'application/json');
}
