// ============================================================
//  Dedi Vocal Academy (DiVA) – Client API Helper
//  Paste your Google Apps Script Web App URL below:
// ============================================================

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx-MxgjthcefPSYuvlciKmI9_6WbohkzEZvj3RUNoG7YgScfCQWWOTlqDgQRYyDndJH/exec';

const API = {

  // ── Core ─────────────────────────────────────────────────
  async post(payload) {
    const res  = await fetch(SCRIPT_URL, {
      method:  'POST',
      headers: { 'Content-Type': 'text/plain' },
      body:    JSON.stringify(payload),
    });
    const json = await res.json();
    if (json.status === 'error') throw new Error(json.message);
    return json.data;
  },

  async get(params) {
    const qs  = new URLSearchParams(params).toString();
    const res = await fetch(SCRIPT_URL + '?' + qs);
    const json = await res.json();
    if (json.status === 'error') throw new Error(json.message);
    return json.data;
  },

  // ── GURU ─────────────────────────────────────────────────
  // getGuru()                     → all guru
  // getGuruByKode('GURU01')       → satu guru by kode_login
  getGuru:       ()       => API.post({ action: 'getGuru' }),
  getGuruByKode: (kode)   => API.post({ action: 'getGuruByKode', kode_login: kode }),
  addGuru:       (d)      => API.post({ action: 'addGuru', ...d }),
  updateGuru:    (d)      => API.post({ action: 'updateGuru', ...d }),
  deleteGuru:    (id)     => API.post({ action: 'deleteGuru', id }),

  // ── MURID ────────────────────────────────────────────────
  // getMurid()                    → all murid
  // getMurid({ guru_id })         → by guru
  // getMuridByLink(link_id)       → portal murid via URL ?id=
  getMurid:       (opts={})  => API.post({ action: 'getMurid', ...opts }),
  getMuridByLink: (link_id)  => API.get({ action: 'getMuridByLink', link_id }),
  addMurid:       (d)        => API.post({ action: 'addMurid', ...d }),
  updateMurid:    (d)        => API.post({ action: 'updateMurid', ...d }),
  deleteMurid:    (id)       => API.post({ action: 'deleteMurid', id }),

  // ── SESI ─────────────────────────────────────────────────
  // getSesi({ bulan, guru_id, murid_id })   → filter sesi
  // saveSesiGuru({ guru_id, bulan, sesi })  → guru submit laporan bulan
  // deleteSesi(id)                          → admin hapus satu sesi
  getSesi:      (opts={})  => API.post({ action: 'getSesi', ...opts }),
  saveSesiGuru: (d)        => API.post({ action: 'saveSesiGuru', ...d }),
  deleteSesi:   (id)       => API.post({ action: 'deleteSesi', id }),

};
