// ============================================================
//  Dedi Vocal Academy (DiVA) – Google Apps Script Backend
//
//  Sheet struktur:
//    guru   → id, nama, fee_per_sesi, kode_login, aktif
//    murid  → id, nama, no_hp, program, paket_sesi, fee_per_sesi,
//              guru_id, link_id, nominal_spp, tgl_bayar_spp, aktif
//    sesi   → id, guru_id, murid_id, tanggal, bulan,
//              today_lesson, foto_url, links, created_at
//
//  Deploy sebagai Web App:
//    Execute as: Me
//    Who has access: Anyone
// ============================================================

const S = {
  GURU:     'guru',
  MURID:    'murid',
  SESI:     'sesi',
  SETTINGS: 'settings',
};

const FOTO_FOLDER_NAME = 'DiVA KBM Photos';

// ── HELPERS ──────────────────────────────────────────────────

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet(name) {
  return ss().getSheetByName(name) || ss().insertSheet(name);
}

function sheetRows(name) {
  const sh   = getSheet(name);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0].map(h => String(h).trim());
  const tz = Session.getScriptTimeZone();
  return vals.slice(1)
    .filter(r => r[0] !== '')
    .map(r => Object.fromEntries(headers.map((h, i) => {
      let v = r[i];
      if (v instanceof Date) {
        // Google Sheets sering auto-convert string ke Date object
        if (h === 'bulan') v = Utilities.formatDate(v, tz, 'yyyy-MM');
        else               v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      }
      return [h, v == null ? '' : String(v)];
    })));
}

function hexId() {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    new Date().toISOString() + Math.random()
  ).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2,'0'))
   .join('').substring(0, 16);
}

function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getBulan(tanggal) {
  // tanggal: 'yyyy-MM-dd' → return 'yyyy-MM'
  return String(tanggal).substring(0, 7);
}

function ok(data)  { return respond({ status: 'ok',    data }); }
function err(msg)  { return respond({ status: 'error', message: msg }); }
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SETUP ────────────────────────────────────────────────────

function setupSheets() {
  const guruSh = getSheet(S.GURU);
  if (guruSh.getLastRow() === 0) {
    guruSh.appendRow(['id','nama','fee_per_sesi','kode_login','aktif']);
    guruSh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1C1026').setFontColor('#FFFFFF');
  }

  const muridSh = getSheet(S.MURID);
  if (muridSh.getLastRow() === 0) {
    muridSh.appendRow(['id','nama','no_hp','program','paket_sesi','fee_per_sesi','guru_id','link_id','nominal_spp','tgl_bayar_spp','aktif']);
    muridSh.getRange(1,1,1,11).setFontWeight('bold').setBackground('#1C1026').setFontColor('#FFFFFF');
  }

  const sesiSh = getSheet(S.SESI);
  if (sesiSh.getLastRow() === 0) {
    sesiSh.appendRow(['id','guru_id','murid_id','tanggal','bulan','today_lesson','foto_url','links','created_at']);
    sesiSh.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1C1026').setFontColor('#FFFFFF');
  }

  const settingsSh = getSheet(S.SETTINGS);
  if (settingsSh.getLastRow() === 0) {
    settingsSh.appendRow(['key', 'value']);
    settingsSh.appendRow(['admin_user', '123456']);
    settingsSh.appendRow(['admin_pass', '123456']);
    settingsSh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1C1026').setFontColor('#FFFFFF');
  }

  // Buat folder foto di Drive
  getFotoFolder();
  Logger.log('Setup selesai!');
}

function getFotoFolder() {
  const folders = DriveApp.getFoldersByName(FOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOTO_FOLDER_NAME);
}

// ── ROUTING ──────────────────────────────────────────────────

function doGet(e) {
  try {
    const p      = e.parameter || {};
    const action = p.action;
    if (action === 'getMuridByLink') return getMuridByLink(p.link_id);
    return err('Unknown GET action: ' + action);
  } catch(ex) { return err(ex.toString()); }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    switch (action) {
      case 'getGuru':         return getGuru();
      case 'getGuruByKode':   return getGuruByKode(body.kode_login);
      case 'addGuru':         return addGuru(body);
      case 'updateGuru':      return updateGuru(body);
      case 'deleteGuru':      return deleteGuru(body.id);

      case 'getMurid':        return getMurid(body);
      case 'addMurid':        return addMurid(body);
      case 'updateMurid':     return updateMurid(body);
      case 'deleteMurid':     return deleteMurid(body.id);

      case 'getSesi':         return getSesi(body);
      case 'addKBM':          return addKBM(body);
      case 'updateKBM':       return updateKBM(body);
      case 'deleteSesi':      return deleteSesi(body.id);
      case 'uploadFotoKBM':   return uploadFotoKBM(body);

      // legacy compat
      case 'saveSesiGuru':    return saveSesiGuru(body);

      case 'getSettings':    return getSettings();
      case 'updateSettings': return updateSettings(body);

      default:                return err('Unknown action: ' + action);
    }
  } catch(ex) { return err(ex.toString()); }
}

// ── GURU CRUD ────────────────────────────────────────────────

function getGuru() { return ok(sheetRows(S.GURU)); }

function getGuruByKode(kode) {
  if (!kode) return err('Kode diperlukan');
  const rows = sheetRows(S.GURU);
  const guru = rows.find(g => String(g.kode_login).toUpperCase() === String(kode).toUpperCase());
  if (!guru) return err('Kode guru tidak ditemukan');
  return ok(guru);
}

function addGuru(d) {
  if (!d.nama || !d.kode_login) return err('Nama dan kode login wajib diisi');
  const existing = sheetRows(S.GURU);
  if (existing.find(g => String(g.kode_login).toUpperCase() === String(d.kode_login).toUpperCase()))
    return err('Kode login sudah digunakan guru lain');
  const id = hexId();
  getSheet(S.GURU).appendRow([id, d.nama, d.fee_per_sesi||0, String(d.kode_login).toUpperCase(), d.aktif||'aktif']);
  return ok({ id });
}

function updateGuru(d) {
  if (!d.id) return err('ID guru diperlukan');
  const sh = getSheet(S.GURU);
  const vals = sh.getDataRange().getValues();
  const h = vals[0].map(x => String(x).trim());
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][h.indexOf('id')]) === String(d.id)) {
      const set = (col, val) => { const i = h.indexOf(col); if (i>=0) sh.getRange(r+1,i+1).setValue(val); };
      set('nama', d.nama); set('fee_per_sesi', d.fee_per_sesi||0);
      set('kode_login', String(d.kode_login||'').toUpperCase()); set('aktif', d.aktif||'aktif');
      return ok({ updated: true });
    }
  }
  return err('Guru tidak ditemukan');
}

function deleteGuru(id) {
  if (!id) return err('ID diperlukan');
  const sh = getSheet(S.GURU);
  const vals = sh.getDataRange().getValues();
  const idx = vals[0].map(x => String(x).trim()).indexOf('id');
  for (let r = vals.length-1; r >= 1; r--) {
    if (String(vals[r][idx]) === String(id)) { sh.deleteRow(r+1); return ok({ deleted: true }); }
  }
  return err('Guru tidak ditemukan');
}

// ── MURID CRUD ───────────────────────────────────────────────

function getMurid(opts) {
  let rows = sheetRows(S.MURID);
  if (opts.guru_id) rows = rows.filter(m => String(m.guru_id) === String(opts.guru_id));
  if (opts.aktif)   rows = rows.filter(m => m.aktif === opts.aktif);

  // Hitung sisa sesi tiap murid
  const sesiAll = sheetRows(S.SESI);
  rows = rows.map(m => {
    const terpakai = sesiAll.filter(s => String(s.murid_id) === String(m.id)).length;
    const paket    = Number(m.paket_sesi || 0);
    return { ...m, sesi_terpakai: terpakai, sisa_sesi: Math.max(0, paket - terpakai) };
  });
  return ok(rows);
}

function getMuridByLink(link_id) {
  if (!link_id) return err('Link ID diperlukan');
  const rows  = sheetRows(S.MURID);
  const murid = rows.find(m => String(m.link_id) === String(link_id));
  if (!murid) return err('Murid tidak ditemukan');
  const terpakai = sheetRows(S.SESI).filter(s => String(s.murid_id) === String(murid.id)).length;
  const paket    = Number(murid.paket_sesi || 0);
  return ok({ ...murid, sesi_terpakai: terpakai, sisa_sesi: Math.max(0, paket - terpakai) });
}

function addMurid(d) {
  if (!d.nama) return err('Nama murid wajib diisi');
  const id = hexId(); const link_id = hexId();
  getSheet(S.MURID).appendRow([
    id, d.nama, d.no_hp||'', d.program||'', d.paket_sesi||0,
    d.fee_per_sesi||0, d.guru_id||'', link_id,
    d.nominal_spp||0, d.tgl_bayar_spp||'', d.aktif||'aktif',
  ]);
  return ok({ id, link_id });
}

function updateMurid(d) {
  if (!d.id) return err('ID murid diperlukan');
  const sh = getSheet(S.MURID);
  const vals = sh.getDataRange().getValues();
  const h = vals[0].map(x => String(x).trim());
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][h.indexOf('id')]) === String(d.id)) {
      const set = (col, val) => { const i = h.indexOf(col); if (i>=0) sh.getRange(r+1,i+1).setValue(val); };
      set('nama', d.nama); set('no_hp', d.no_hp||'');
      set('program', d.program||''); set('paket_sesi', d.paket_sesi||0);
      set('fee_per_sesi', d.fee_per_sesi||0);
      set('guru_id', d.guru_id||''); set('nominal_spp', d.nominal_spp||0);
      set('tgl_bayar_spp', d.tgl_bayar_spp||''); set('aktif', d.aktif||'aktif');
      return ok({ updated: true });
    }
  }
  return err('Murid tidak ditemukan');
}

function deleteMurid(id) {
  if (!id) return err('ID diperlukan');
  const sh = getSheet(S.MURID);
  const vals = sh.getDataRange().getValues();
  const idx = vals[0].map(x => String(x).trim()).indexOf('id');
  for (let r = vals.length-1; r >= 1; r--) {
    if (String(vals[r][idx]) === String(id)) { sh.deleteRow(r+1); return ok({ deleted: true }); }
  }
  return err('Murid tidak ditemukan');
}

// ── SETTINGS ─────────────────────────────────────────────

function getSettings() {
  const rows = sheetRows(S.SETTINGS);
  const obj  = {};
  rows.forEach(r => { if (r.key) obj[String(r.key)] = String(r.value || ''); });
  return ok(obj);
}

function updateSettings(d) {
  if (!d.settings || typeof d.settings !== 'object') return err('settings object diperlukan');
  const sh   = getSheet(S.SETTINGS);
  const vals = sh.getDataRange().getValues();
  const keys = vals.map(r => String(r[0]));

  Object.entries(d.settings).forEach(([key, val]) => {
    const rowIdx = keys.indexOf(key);
    if (rowIdx >= 1) {
      sh.getRange(rowIdx + 1, 2).setValue(val); // update existing
    } else {
      sh.appendRow([key, val]); // insert baru
      keys.push(key);
    }
  });
  return ok({ updated: true });
}

// ── SESI / KBM ───────────────────────────────────────────────

function getSesi(opts) {
  let rows = sheetRows(S.SESI);
  if (opts.guru_id)  rows = rows.filter(s => String(s.guru_id)  === String(opts.guru_id));
  if (opts.murid_id) rows = rows.filter(s => String(s.murid_id) === String(opts.murid_id));
  if (opts.bulan)    rows = rows.filter(s => String(s.bulan)    === String(opts.bulan));
  rows.sort((a,b) => String(b.tanggal).localeCompare(String(a.tanggal)));
  return ok(rows);
}

/** Guru tambah satu sesi KBM */
function addKBM(d) {
  if (!d.guru_id || !d.murid_id || !d.tanggal) return err('guru_id, murid_id, dan tanggal wajib diisi');
  const id    = hexId();
  const bulan = getBulan(d.tanggal);
  getSheet(S.SESI).appendRow([
    id, d.guru_id, d.murid_id, d.tanggal, bulan,
    d.today_lesson || '', d.foto_url || '', d.links || '', today(),
  ]);
  return ok({ id, bulan });
}

/** Guru edit sesi KBM yang sudah ada */
function updateKBM(d) {
  if (!d.id) return err('ID sesi diperlukan');
  const sh   = getSheet(S.SESI);
  const vals = sh.getDataRange().getValues();
  const h    = vals[0].map(x => String(x).trim());
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][h.indexOf('id')]) === String(d.id)) {
      const set = (col, val) => { const i = h.indexOf(col); if (i >= 0) sh.getRange(r+1, i+1).setValue(val); };
      if (d.tanggal      !== undefined) { set('tanggal', d.tanggal); set('bulan', getBulan(d.tanggal)); }
      if (d.today_lesson !== undefined) set('today_lesson', d.today_lesson || '');
      if (d.foto_url     !== undefined) set('foto_url', d.foto_url || '');
      if (d.links        !== undefined) set('links', d.links || '');
      return ok({ updated: true });
    }
  }
  return err('Sesi tidak ditemukan');
}

function deleteSesi(id) {
  if (!id) return err('ID diperlukan');
  const sh = getSheet(S.SESI);
  const vals = sh.getDataRange().getValues();
  const idx = vals[0].map(x => String(x).trim()).indexOf('id');
  for (let r = vals.length-1; r >= 1; r--) {
    if (String(vals[r][idx]) === String(id)) { sh.deleteRow(r+1); return ok({ deleted: true }); }
  }
  return err('Sesi tidak ditemukan');
}

/** Upload foto KBM ke Google Drive, return URL */
function uploadFotoKBM(d) {
  if (!d.base64 || !d.filename) return err('base64 dan filename diperlukan');
  try {
    const decoded  = Utilities.base64Decode(d.base64.split(',').pop());
    const blob     = Utilities.newBlob(decoded, d.mimeType || 'image/jpeg', d.filename);
    const folder   = getFotoFolder();
    const file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId   = file.getId();
    const url      = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w600';
    return ok({ url, fileId });
  } catch(e) {
    return err('Upload gagal: ' + e.toString());
  }
}

/** Legacy: saveSesiGuru (batch) — tetap support untuk backward compat */
function saveSesiGuru(body) {
  if (!body.guru_id || !body.bulan) return err('guru_id dan bulan diperlukan');
  const sesiArr = body.sesi || [];
  const sh   = getSheet(S.SESI);
  const vals = sh.getDataRange().getValues();
  const h    = vals[0].map(x => String(x).trim());
  const gi   = h.indexOf('guru_id'); const bi = h.indexOf('bulan');
  for (let r = vals.length-1; r >= 1; r--) {
    if (String(vals[r][gi]) === String(body.guru_id) && String(vals[r][bi]) === String(body.bulan))
      sh.deleteRow(r+1);
  }
  const now = today();
  sesiArr.forEach(s => {
    if (!s.tanggal) return;
    sh.appendRow([hexId(), body.guru_id, s.murid_id||'', s.tanggal, body.bulan,
                  s.today_lesson||'', s.foto_url||'', now]);
  });
  return ok({ saved: sesiArr.filter(s => s.tanggal).length });
}
