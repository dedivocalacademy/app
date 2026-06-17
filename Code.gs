// ============================================================
//  Dedi Vocal Academy (DiVA) – Google Apps Script Backend
//
//  Sheet struktur:
//    guru   → id, nama, fee_per_sesi, kode_login, aktif
//    murid  → id, nama, no_hp, program, rate_per_sesi, guru_id,
//              link_id, nominal_spp, tgl_bayar_spp, aktif
//    sesi   → id, guru_id, murid_id, tanggal, bulan, catatan, created_at
//
//  Deploy sebagai Web App:
//    Execute as: Me
//    Who has access: Anyone
// ============================================================

const S = {
  GURU:  'guru',
  MURID: 'murid',
  SESI:  'sesi',
};

// ── HELPERS ──────────────────────────────────────────────────

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  return ss().getSheetByName(name) || ss().insertSheet(name);
}

function sheetRows(name) {
  const sh   = getSheet(name);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0].map(h => String(h).trim());
  return vals.slice(1)
    .filter(r => r[0] !== '')
    .map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
}

function hexId() {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    new Date().toISOString() + Math.random()
  ).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2,'0'))
   .join('')
   .substring(0, 16);
}

function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
  // GURU sheet headers
  const guruSh = getSheet(S.GURU);
  if (guruSh.getLastRow() === 0) {
    guruSh.appendRow(['id','nama','fee_per_sesi','kode_login','aktif']);
    guruSh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1C1026').setFontColor('#FFFFFF');
  }

  // MURID sheet headers
  const muridSh = getSheet(S.MURID);
  if (muridSh.getLastRow() === 0) {
    muridSh.appendRow(['id','nama','no_hp','program','rate_per_sesi','guru_id','link_id','nominal_spp','tgl_bayar_spp','aktif']);
    muridSh.getRange(1,1,1,10).setFontWeight('bold').setBackground('#1C1026').setFontColor('#FFFFFF');
  }

  // SESI sheet headers
  const sesiSh = getSheet(S.SESI);
  if (sesiSh.getLastRow() === 0) {
    sesiSh.appendRow(['id','guru_id','murid_id','tanggal','bulan','catatan','created_at']);
    sesiSh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1C1026').setFontColor('#FFFFFF');
  }
}

// ── ROUTING ──────────────────────────────────────────────────

function doGet(e) {
  try {
    const p      = e.parameter || {};
    const action = p.action;
    if (action === 'getMuridByLink') return getMuridByLink(p.link_id);
    return err('Unknown GET action: ' + action);
  } catch(ex) {
    return err(ex.toString());
  }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    switch (action) {
      // GURU
      case 'getGuru':       return getGuru();
      case 'getGuruByKode': return getGuruByKode(body.kode_login);
      case 'addGuru':       return addGuru(body);
      case 'updateGuru':    return updateGuru(body);
      case 'deleteGuru':    return deleteGuru(body.id);

      // MURID
      case 'getMurid':      return getMurid(body);
      case 'addMurid':      return addMurid(body);
      case 'updateMurid':   return updateMurid(body);
      case 'deleteMurid':   return deleteMurid(body.id);

      // SESI
      case 'getSesi':       return getSesi(body);
      case 'saveSesiGuru':  return saveSesiGuru(body);
      case 'deleteSesi':    return deleteSesi(body.id);

      default:              return err('Unknown action: ' + action);
    }
  } catch(ex) {
    return err(ex.toString());
  }
}

// ── GURU CRUD ────────────────────────────────────────────────

function getGuru() {
  return ok(sheetRows(S.GURU));
}

function getGuruByKode(kode) {
  if (!kode) return err('Kode diperlukan');
  const rows = sheetRows(S.GURU);
  const guru = rows.find(g => String(g.kode_login).toUpperCase() === String(kode).toUpperCase());
  if (!guru) return err('Kode guru tidak ditemukan');
  return ok(guru);
}

function addGuru(d) {
  if (!d.nama || !d.kode_login) return err('Nama dan kode login wajib diisi');

  // Cek duplikat kode
  const existing = sheetRows(S.GURU);
  const dupKode  = existing.find(g => String(g.kode_login).toUpperCase() === String(d.kode_login).toUpperCase());
  if (dupKode) return err('Kode login sudah digunakan oleh guru lain');

  const id = hexId();
  getSheet(S.GURU).appendRow([
    id,
    d.nama,
    d.fee_per_sesi || 0,
    String(d.kode_login).toUpperCase(),
    d.aktif || 'aktif',
  ]);
  return ok({ id });
}

function updateGuru(d) {
  if (!d.id) return err('ID guru diperlukan');
  const sh   = getSheet(S.GURU);
  const vals = sh.getDataRange().getValues();
  const headers = vals[0].map(h => String(h).trim());
  const idIdx   = headers.indexOf('id');
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][idIdx]) === String(d.id)) {
      const setCol = (col, val) => {
        const i = headers.indexOf(col);
        if (i >= 0) sh.getRange(r+1, i+1).setValue(val);
      };
      setCol('nama',          d.nama);
      setCol('fee_per_sesi',  d.fee_per_sesi || 0);
      setCol('kode_login',    String(d.kode_login || '').toUpperCase());
      setCol('aktif',         d.aktif || 'aktif');
      return ok({ updated: true });
    }
  }
  return err('Guru tidak ditemukan');
}

function deleteGuru(id) {
  if (!id) return err('ID diperlukan');
  const sh   = getSheet(S.GURU);
  const vals = sh.getDataRange().getValues();
  const idIdx = vals[0].map(h => String(h).trim()).indexOf('id');
  for (let r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][idIdx]) === String(id)) {
      sh.deleteRow(r + 1);
      return ok({ deleted: true });
    }
  }
  return err('Guru tidak ditemukan');
}

// ── MURID CRUD ───────────────────────────────────────────────

function getMurid(opts) {
  let rows = sheetRows(S.MURID);
  if (opts.guru_id)  rows = rows.filter(m => String(m.guru_id) === String(opts.guru_id));
  if (opts.aktif)    rows = rows.filter(m => m.aktif === opts.aktif);
  return ok(rows);
}

function getMuridByLink(link_id) {
  if (!link_id) return err('Link ID diperlukan');
  const rows = sheetRows(S.MURID);
  const murid = rows.find(m => String(m.link_id) === String(link_id));
  if (!murid) return err('Murid tidak ditemukan');
  return ok(murid);
}

function addMurid(d) {
  if (!d.nama) return err('Nama murid wajib diisi');
  const id      = hexId();
  const link_id = hexId();
  getSheet(S.MURID).appendRow([
    id,
    d.nama,
    d.no_hp          || '',
    d.program        || '',
    d.rate_per_sesi  || 0,
    d.guru_id        || '',
    link_id,
    d.nominal_spp    || 0,
    d.tgl_bayar_spp  || '',
    d.aktif          || 'aktif',
  ]);
  return ok({ id, link_id });
}

function updateMurid(d) {
  if (!d.id) return err('ID murid diperlukan');
  const sh      = getSheet(S.MURID);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0].map(h => String(h).trim());
  const idIdx   = headers.indexOf('id');
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][idIdx]) === String(d.id)) {
      const setCol = (col, val) => {
        const i = headers.indexOf(col);
        if (i >= 0) sh.getRange(r+1, i+1).setValue(val);
      };
      setCol('nama',          d.nama);
      setCol('no_hp',         d.no_hp         || '');
      setCol('program',       d.program        || '');
      setCol('rate_per_sesi', d.rate_per_sesi  || 0);
      setCol('guru_id',       d.guru_id        || '');
      setCol('nominal_spp',   d.nominal_spp    || 0);
      setCol('tgl_bayar_spp', d.tgl_bayar_spp  || '');
      setCol('aktif',         d.aktif          || 'aktif');
      return ok({ updated: true });
    }
  }
  return err('Murid tidak ditemukan');
}

function deleteMurid(id) {
  if (!id) return err('ID diperlukan');
  const sh    = getSheet(S.MURID);
  const vals  = sh.getDataRange().getValues();
  const idIdx = vals[0].map(h => String(h).trim()).indexOf('id');
  for (let r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][idIdx]) === String(id)) {
      sh.deleteRow(r + 1);
      return ok({ deleted: true });
    }
  }
  return err('Murid tidak ditemukan');
}

// ── SESI ─────────────────────────────────────────────────────

function getSesi(opts) {
  let rows = sheetRows(S.SESI);
  if (opts.guru_id)  rows = rows.filter(s => String(s.guru_id)  === String(opts.guru_id));
  if (opts.murid_id) rows = rows.filter(s => String(s.murid_id) === String(opts.murid_id));
  if (opts.bulan)    rows = rows.filter(s => String(s.bulan)    === String(opts.bulan));
  // Sort by tanggal desc
  rows.sort((a,b) => String(b.tanggal).localeCompare(String(a.tanggal)));
  return ok(rows);
}

/**
 * Guru kirim laporan sesi sekaligus untuk satu bulan.
 * Body: { guru_id, bulan, sesi: [{ murid_id, tanggal, catatan }] }
 * Behavior: hapus dulu semua sesi guru di bulan itu, lalu insert baru.
 */
function saveSesiGuru(body) {
  if (!body.guru_id || !body.bulan) return err('guru_id dan bulan diperlukan');
  const sesiArr = body.sesi || [];

  const sh      = getSheet(S.SESI);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0].map(h => String(h).trim());
  const guruIdx = headers.indexOf('guru_id');
  const bulanIdx= headers.indexOf('bulan');

  // Hapus baris lama (dari belakang)
  for (let r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][guruIdx]) === String(body.guru_id) &&
        String(vals[r][bulanIdx]) === String(body.bulan)) {
      sh.deleteRow(r + 1);
    }
  }

  // Insert baru
  const now = today();
  sesiArr.forEach(s => {
    if (!s.tanggal) return; // skip jika tanggal kosong
    sh.appendRow([
      hexId(),
      body.guru_id,
      s.murid_id  || '',
      s.tanggal,
      body.bulan,
      s.catatan   || '',
      now,
    ]);
  });

  return ok({ saved: sesiArr.filter(s => s.tanggal).length });
}

function deleteSesi(id) {
  if (!id) return err('ID diperlukan');
  const sh    = getSheet(S.SESI);
  const vals  = sh.getDataRange().getValues();
  const idIdx = vals[0].map(h => String(h).trim()).indexOf('id');
  for (let r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][idIdx]) === String(id)) {
      sh.deleteRow(r + 1);
      return ok({ deleted: true });
    }
  }
  return err('Sesi tidak ditemukan');
}
