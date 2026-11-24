/**
 * code.gs — Ledger Gabungan Peringkat Umum Kelas X
 * SMK Negeri 1 Maluku Tengah, Semester Ganjil TP 2026/2027
 *
 * Fungsi:
 *  1. Menu "Sertifikat" di Google Sheet (cek kelengkapan ledger).
 *  2. Web App yang mengeluarkan data ledger sebagai JSON, dilindungi token,
 *     dipakai oleh generate_sertifikat.py  (python generate_sertifikat.py --json "URL?token=...").
 *
 * Cara pasang:
 *  - Buka Google Sheet > Extensions > Apps Script, tempel seluruh kode ini, simpan.
 *  - Jalankan fungsi aturToken() sekali (setujui izin), lalu lihat token di View > Logs.
 *  - Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone > Deploy.
 *  - Salin URL /exec, tambahkan  ?token=TOKEN_ANDA  di belakangnya.
 *
 * Catatan privasi: data berisi nilai siswa. Jangan sebarkan URL beserta token.
 */

const GID_LEDGER = 705250680;          // gid tab "Ledger Gabungan"
const NAMA_TAB_CADANGAN = 'Ledger Gabungan'; // dipakai bila gid tidak ditemukan

/* ------------------------------ MENU ------------------------------ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sertifikat')
    .addItem('Cek kelengkapan ledger', 'cekKelengkapan')
    .addItem('Buat / ganti token Web App', 'aturToken')
    .addToUi();
}

function cekKelengkapan() {
  const data = bacaLedger_();
  const per = {};
  data.siswa.forEach(s => {
    per[s.kelas] = per[s.kelas] || { total: 0, dinilai: 0 };
    per[s.kelas].total++;
    if (s.rata !== null && s.jumlah !== null) per[s.kelas].dinilai++;
  });
  const baris = Object.keys(per).map(k =>
    (per[k].dinilai === per[k].total ? 'OK  ' : '!!  ') + k + ': ' + per[k].dinilai + ' / ' + per[k].total);
  const ada = data.siswa.filter(s => s.pu !== null).map(s => s.pu).sort((a, b) => a - b);
  const hilang = [];
  for (let n = 1; ada.length && n <= ada[ada.length - 1]; n++) if (ada.indexOf(n) < 0) hilang.push(n);
  baris.push('', 'Peringkat umum terbaca: ' + (ada.join(', ') || '-'));
  if (hilang.length) baris.push('Peringkat tidak ditemukan: ' + hilang.join(', ') + '  (data belum lengkap)');
  SpreadsheetApp.getUi().alert('Kelengkapan ledger (terisi / total siswa)', baris.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function aturToken() {
  const token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TOKEN', token);
  Logger.log('TOKEN BARU: ' + token);
  try { SpreadsheetApp.getUi().alert('Token baru (salin):\n\n' + token); } catch (e) { /* dijalankan dari editor */ }
  return token;
}

/* ------------------------------ WEB APP ------------------------------ */
function doGet(e) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  let out;
  if (!token) {
    out = { ok: false, error: 'Token belum dibuat. Jalankan aturToken() terlebih dahulu.' };
  } else if (!e || !e.parameter || e.parameter.token !== token) {
    out = { ok: false, error: 'Token tidak valid.' };
  } else {
    try {
      out = Object.assign({ ok: true }, bacaLedger_());
    } catch (err) {
      out = { ok: false, error: String(err) };
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------ PARSER ------------------------------ */
function norm_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim().toLowerCase();
}

function angka_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  let s = String(v).trim().replace(/\u00a0/g, '');
  if (!s || s.charAt(0) === '#') return null;
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
    s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  } else if (s.indexOf(',') >= 0) {
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function ambilSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheets().filter(s => s.getSheetId() === GID_LEDGER)[0] || ss.getSheetByName(NAMA_TAB_CADANGAN);
  if (!sh) throw new Error('Tab ledger tidak ditemukan (cek GID_LEDGER / NAMA_TAB_CADANGAN).');
  return sh;
}

function bacaLedger_() {
  const sh = ambilSheet_();
  const rows = sh.getDataRange().getValues();

  // 1. cari baris header lewat kolom "Nama Siswa"
  let h = -1;
  for (let i = 0; i < rows.length && h < 0; i++) {
    if (rows[i].some(c => norm_(c) === 'nama siswa')) h = i;
  }
  if (h < 0) throw new Error("Header 'Nama Siswa' tidak ditemukan.");
  const head = rows[h].map(norm_);

  const cari = (cocok) => head.findIndex(cocok);
  const cKelas = cari(c => c === 'kelas');
  const cNis = cari(c => c === 'nis');
  const cNama = cari(c => c === 'nama siswa');
  const cJml = cari(c => c.indexOf('jumlah') === 0);
  const cRata = cari(c => c.indexOf('rata') === 0);
  const cPu = cari(c => c.indexOf('peringkat umum') === 0);
  const cPk = cari(c => c.indexOf('peringkat kelas') === 0);
  if ([cKelas, cNama, cJml, cRata, cPu].some(x => x < 0)) throw new Error('Kolom wajib ledger tidak lengkap.');

  // 2. nama mapel ada di baris di bawah header
  const barisMapel = rows[h + 1] || [];
  const idxMapel = [];
  for (let i = cNama + 1; i < cJml; i++) idxMapel.push(i);
  const mapel = idxMapel.map(i => String(barisMapel[i] || ('Mapel ' + i)).replace(/\s+/g, ' ').trim());

  // 3. baca siswa
  const siswa = [];
  for (let r = h + 2; r < rows.length; r++) {
    const row = rows[r];
    const nama = String(row[cNama] === undefined ? '' : row[cNama]).trim();
    if (!nama) continue;
    const nilai = {};
    idxMapel.forEach((i, k) => { nilai[mapel[k]] = angka_(row[i]); });
    siswa.push({
      kelas: String(row[cKelas] === undefined ? '' : row[cKelas]).trim(),
      nis: cNis >= 0 ? String(row[cNis] === undefined ? '' : row[cNis]).trim() : '',
      nama: nama,
      nilai: nilai,
      jumlah: angka_(row[cJml]),
      rata: angka_(row[cRata]),
      pu: angka_(row[cPu]),
      pk: cPk >= 0 ? angka_(row[cPk]) : null
    });
  }
  return { sumber: sh.getParent().getName() + ' / ' + sh.getName(), diambil: new Date().toISOString(), mapel: mapel, siswa: siswa };
}
