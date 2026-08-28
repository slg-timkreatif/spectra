/* =========================================================================
   SPECTRA UI CORE v2 — klien, sesi, tema, toast, modal, lucide
   ========================================================================= */
const SPECTRA_URL = 'https://qdgtfhknuvncbzsbcbwo.supabase.co';
const SPECTRA_KEY = 'sb_publishable_dYm2x9Juh6lHFRMgGsXB7g_eP557huM';
const SESI_KEY = 'spectra_sesi';

let _sb = null;
async function sb() {
  if (!_sb) {
    const { createClient } = await import('https://cdn.jsdelivr.net/npm/@supabase/supabase-js/+esm');
    _sb = createClient(SPECTRA_URL, SPECTRA_KEY);
  }
  return _sb;
}

/* ---------- Tema: light default, dark opsional; 'lock' untuk halaman broadcast ---------- */
function initTheme(defaultT, lock) {
  let t = defaultT || 'light';
  if (!lock) t = localStorage.getItem('spectra_theme') || t;
  document.documentElement.setAttribute('data-theme', t);
}
function toggleTheme() {
  const cur = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', cur);
  localStorage.setItem('spectra_theme', cur);
  icons();
  return cur;
}
function themeIcon() {
  return document.documentElement.getAttribute('data-theme') === 'dark' ? 'sun' : 'moon';
}

/* ---------- Sesi ---------- */
const bacaSesi = () => { try { return JSON.parse(localStorage.getItem(SESI_KEY)); } catch (e) { return null; } };
const simpanSesi = s => localStorage.setItem(SESI_KEY, JSON.stringify(s));
const keluarSpectra = () => { localStorage.removeItem(SESI_KEY); location.href = 'spectra-login.html'; };
const roleHome = p => ({ JURI: 'spectra-juri.html', OPERATOR: 'spectra-operator.html', ADMIN: 'spectra-admin.html' }[p] || 'spectra-login.html');

/* ---------- Toast ---------- */
function toast(msg, tipe) {
  let t = document.getElementById('spectraToast');
  if (!t) { t = document.createElement('div'); t.id = 'spectraToast'; document.body.appendChild(t); }
  const ic = tipe === 'ok' ? 'check-circle-2' : tipe === 'warn' ? 'alert-triangle' : tipe === 'err' ? 'x-circle' : 'info';
  t.innerHTML = '<i data-lucide="' + ic + '"></i> ' + msg;
  t.classList.add('show');
  if (window.lucide) lucide.createIcons();
  clearTimeout(t._h); t._h = setTimeout(() => t.classList.remove('show'), 3200);
}

/* ---------- Modal konfirmasi ---------- */
function spectraConfirm({ judul, pesan, teksYa = 'Ya, lanjutkan', bahaya = false }) {
  return new Promise(res => {
    const m = document.createElement('div');
    m.id = 'spectraModal';
    m.innerHTML = '<div class="card"><h4><i data-lucide="' + (bahaya ? 'alert-triangle' : 'help-circle') + '"></i>' + judul +
      '</h4><p>' + pesan + '</p><div class="modal-actions"><button class="btn btn-ghost" id="mBatal">Batal</button>' +
      '<button class="btn ' + (bahaya ? 'btn-danger' : 'btn-primary') + '" id="mYa">' + teksYa + '</button></div></div>';
    document.body.appendChild(m);
    if (window.lucide) lucide.createIcons();
    const tutup = v => { m.remove(); res(v); };
    m.querySelector('#mBatal').onclick = () => tutup(false);
    m.querySelector('#mYa').onclick = () => tutup(true);
    m.addEventListener('click', e => { if (e.target === m) tutup(false); });
  });
}

/* ---------- Indikator koneksi ---------- */
function initConn(elId) {
  const el = document.getElementById(elId);
  if (!el) return;
  const set = () => {
    el.innerHTML = navigator.onLine
      ? '<span class="badge ok"><i data-lucide="wifi"></i> Online</span>'
      : '<span class="badge danger"><i data-lucide="wifi-off"></i> Offline</span>';
    if (window.lucide) lucide.createIcons();
  };
  window.addEventListener('online', set); window.addEventListener('offline', set); set();
}

/* ---------- Lucide ---------- */
function icons() { if (window.lucide) lucide.createIcons(); }
document.addEventListener('DOMContentLoaded', icons);