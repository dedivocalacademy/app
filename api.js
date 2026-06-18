// ============================================================
//  Dedi Vocal Academy (DiVA) – Client API Helper
// ============================================================

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx-MxgjthcefPSYuvlciKmI9_6WbohkzEZvj3RUNoG7YgScfCQWWOTlqDgQRYyDndJH/exec';

const API = {

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
  getGuru:       ()      => API.post({ action: 'getGuru' }),
  getGuruByKode: (kode)  => API.post({ action: 'getGuruByKode', kode_login: kode }),
  addGuru:       (d)     => API.post({ action: 'addGuru', ...d }),
  updateGuru:    (d)     => API.post({ action: 'updateGuru', ...d }),
  deleteGuru:    (id)    => API.post({ action: 'deleteGuru', id }),

  // ── MURID ────────────────────────────────────────────────
  // Response murid sudah include: sesi_terpakai, sisa_sesi
  getMurid:       (opts={}) => API.post({ action: 'getMurid', ...opts }),
  getMuridByLink: (link_id) => API.get({ action: 'getMuridByLink', link_id }),
  addMurid:       (d)       => API.post({ action: 'addMurid', ...d }),
  updateMurid:    (d)       => API.post({ action: 'updateMurid', ...d }),
  deleteMurid:    (id)      => API.post({ action: 'deleteMurid', id }),

  // ── SESI / KBM ───────────────────────────────────────────
  // getSesi({ bulan, guru_id, murid_id })
  // addKBM({ guru_id, murid_id, tanggal, today_lesson, foto_url })
  getSesi:    (opts={}) => API.post({ action: 'getSesi', ...opts }),
  addKBM:     (d)       => API.post({ action: 'addKBM', ...d }),
  deleteSesi: (id)      => API.post({ action: 'deleteSesi', id }),

  // ── UPLOAD FOTO KBM ──────────────────────────────────────
  // uploadFotoKBM(file: File) → { url, fileId }
  uploadFotoKBM: (file) =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const result = await API.post({
            action:   'uploadFotoKBM',
            base64:   e.target.result,
            filename: file.name,
            mimeType: file.type,
          });
          resolve(result);
        } catch(err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    }),
};
