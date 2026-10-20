/**
 * APLIKASI PENUGASAN & PEMBUATAN SOAL UJIAN MADRASAH
 * Google Apps Script (container-bound ke Google Spreadsheet)
 *
 * Cara pasang: buka Spreadsheet > Ekstensi > Apps Script,
 * tempel file ini sebagai Code.gs, lalu buat file HTML bernama "Index"
 * dan tempel isi Index.html. Jalankan menu "Penugasan Soal > Siapkan Sheet".
 */

// ====================== KONFIGURASI ======================
const TINGKAT  = ['Aliyah', 'Tsanawiyah', 'Ibtidaiyah', 'Idadiyah PK', 'Idadiyah Reguler'];
const KATEGORI = ['Seluruh Fan', 'Mapel Utama', 'Faroid', 'Mapel Tambahan'];
const UJIAN    = ['IMDA I', 'IMDA II', 'IMDA III', 'IMNI'];
const UNIT     = ['Induk', 'Ranting'];
const STATUS   = ['Belum Diisi', 'Process', 'Submitted', 'Approved'];
const TIPE_NAMA = {
  A: 'Tipe A (Pilihan Ganda)',
  B: 'Tipe B (Melengkapi / Isian)',
  C: 'Tipe C (Essay / Uraian)'
};
const ROMAWI = { '1': 1, '2': 2, '3': 3, I: 1, II: 2, III: 3 };

const SCHEMA = {
  ustadz:    { sheet: 'Ustadz',    cols: ['ID', 'Nama', 'NIP', 'Tingkat', 'Kontak'], req: ['Nama', 'Tingkat'], text: ['NIP', 'Kontak'] },
  mapel:     { sheet: 'Mapel',     cols: ['ID', 'Tingkat', 'Kelas', 'Mapel', 'Kategori'], req: ['Tingkat', 'Kelas', 'Mapel', 'Kategori'], text: ['Kelas'] },
  tugas:     { sheet: 'Tugas',     cols: ['ID', 'Ustadz', 'Tingkat', 'Kelas', 'Mapel', 'Ujian', 'Unit', 'Deadline', 'Status', 'Diperbarui', 'Sonik', 'Korektor1', 'Korektor2', 'Korektor3'] /* 'Sonik' tidak dipakai lagi (pembuat soal = kolom Ustadz); dipertahankan agar kolom lama tetap sejajar */, req: ['Ustadz', 'Tingkat', 'Kelas', 'Mapel', 'Ujian', 'Unit'], text: ['Kelas', 'Deadline', 'Diperbarui'] },
  murid:     { sheet: 'Murid',     cols: ['ID', 'Nama', 'NIS', 'NISN', 'Kelas'], req: ['Nama'], text: ['NIS', 'NISN', 'Kelas'] },
  soal:      { sheet: 'Soal',      cols: ['TugasID', 'Tipe', 'Materi', 'No', 'Pertanyaan', 'A', 'B', 'C', 'D', 'Kunci', 'Kisi'], req: [], text: ['Materi', 'No', 'A', 'B', 'C', 'D'] },
  komposisi: { sheet: 'Komposisi', cols: ['Tingkat', 'Kategori', 'Kelas', 'Ujian', 'Tipe', 'Materi', 'Jumlah'], req: [], text: ['Kelas', 'Materi'] },
  bobot:     { sheet: 'Bobot',     cols: ['Tingkat', 'Kategori', 'Ujian', 'Tipe', 'Bobot', 'Kelas'], req: [], text: [] },
  kisi:      { sheet: 'KisiKisi',  cols: ['Tingkat', 'Kelas', 'Imda', 'Fan', 'Materi', 'ID'], req: ['Tingkat', 'Kelas', 'Fan', 'Imda', 'Materi'], text: ['Kelas', 'Imda'] },
  petunjuk:  { sheet: 'Petunjuk',  cols: ['Tipe', 'Teks'], req: [], text: [] },
  pengaturan:{ sheet: 'Pengaturan', cols: ['Kunci', 'Nilai'], req: [], text: [] }
};

// Teks petunjuk soal (Arab). {skor} diganti otomatis dengan bobot, contoh: ثلاث نتائج
const PETUNJUK_AWAL = {
  A: 'أ ~ إختر أصح الاجوبة مما بين القوسين ! (لكل رقم {skor})',
  B: 'ب ~ أكمل الفراغات الآتية بما يناسب ! (لكل رقم {skor})',
  C: 'ج ~ أجب عن الأسئلة الآتية إجابة صحيحة ! (لكل رقم {skor})',
  D: 'د ~ أجب عن الأسئلة الآتية إجابة صحيحة ! (لكل رقم {skor})'
};
// Pesan pengingat. Placeholder: {ustadz} {daftar} | baris tugas: {mapel} {kelas} {tingkat} {ujian} {unit} {deadline} {status}
const PENGATURAN_AWAL = {
  PesanWA: "Assalamu'alaikum {ustadz}.\nPengingat penulisan soal ujian:\n{daftar}\nMohon diselesaikan sebelum tenggat. Jazakumullahu khairan.",
  BarisTugas: '• {mapel} kelas {kelas} ({tingkat}) {ujian} {unit} - tenggat {deadline} [{status}]',
  WarnaIsian: 'FF0000' // warna teks isian otomatis di halaman kunci & kisi-kisi (hex tanpa #; 000000 = hitam)
};

// Nama kolom Excel yang dikenali saat impor (selain nama kolom aslinya)
const ALIAS = {
  ustadz: {
    Nama: ['nama', 'namalengkap', 'namaustadz', 'namalengkapustadz'],
    NIP: ['nip', 'nipid', 'id'],
    Tingkat: ['tingkat'],
    Kontak: ['kontak', 'nomorwhatsapp', 'whatsapp', 'wa', 'email', 'nohp', 'hp']
  },
  mapel: {
    Tingkat: ['tingkat'], Kelas: ['kelas'],
    Mapel: ['mapel', 'fan', 'namafan', 'matapelajaran', 'namafanmatapelajaran'],
    Kategori: ['kategorisoal', 'kategori', 'kategoritipesoal']
  },
  tugas: {
    Ustadz: ['ustadz', 'namaustadz'], Tingkat: ['tingkat'], Kelas: ['kelas'],
    Mapel: ['mapel', 'fan', 'fanmapel', 'matapelajaran'],
    Ujian: ['ujian', 'tahapujian', 'tahap'],
    Unit: ['unit', 'kategoriunit', 'kategori'],
    Deadline: ['deadline', 'tenggat', 'tenggatwaktu'],
    Korektor1: ['korektor1', 'korektor'], Korektor2: ['korektor2'], Korektor3: ['korektor3']
  },
  murid: {
    Nama: ['nama', 'namamurid'], NIS: ['nis'], NISN: ['nisn'], Kelas: ['kelas']
  },
  kisi: {
    Tingkat: ['tingkat'], Kelas: ['kelas'], Fan: ['fan', 'mapel', 'matapelajaran', 'materi'],
    Imda: ['imda', 'tahapujian', 'tahap'], Materi: ['materibabfasal', 'materi', 'bab', 'fasal']
  }
};

// ====================== TITIK MASUK ======================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('LABSOMA PPS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('LABSOMA PPS')
    .addItem('Buka aplikasi', 'bukaAplikasi')
    .addItem('Siapkan sheet', 'setup')
    .addItem('Aktifkan pengingat email harian', 'aktifkanEmail')
    .addItem('Matikan pengingat email harian', 'matikanEmail')
    .addItem('Tambahkan komposisi IMNI', 'seedKomposisiIMNI_')
    .addToUi();
}

// Fungsi baru:
function seedKomposisiIMNI_() {
  const rows = readAll_('komposisi');
  const adaImni = rows.some(r => ujianCanon_(r.Ujian) === 'IMNI');
  if (adaImni) {
    SpreadsheetApp.getUi().alert('Baris IMNI sudah ada di sheet Komposisi. Tidak ditambahkan lagi.');
    return;
  }
  // Sesuaikan jumlah di bawah dengan panduan madrasah Anda
  const DEF = [
    // [Kategori, Tipe, Materi(kelas sumber), Jumlah]
    ['Seluruh Fan', 'A', '1', 10], ['Seluruh Fan', 'A', '2', 15], ['Seluruh Fan', 'A', '3', 20],
    ['Seluruh Fan', 'B', '1',  5], ['Seluruh Fan', 'B', '2',  5], ['Seluruh Fan', 'B', '3',  5],
    ['Seluruh Fan', 'C', '1',  3], ['Seluruh Fan', 'C', '2',  4], ['Seluruh Fan', 'C', '3',  5],
    ['Seluruh Fan', 'D', '1',  2], ['Seluruh Fan', 'D', '2',  3], ['Seluruh Fan', 'D', '3',  4]
  ];
  const out = DEF.map(([kat, tipe, materi, jml]) =>
    toRow_('komposisi', { Tingkat: 'Tsanawiyah', Kategori: kat, Kelas: '3', Ujian: 'IMNI', Tipe: tipe, Materi: materi, Jumlah: jml }));
  append_('komposisi', out);
  SpreadsheetApp.getUi().alert(out.length + ' baris komposisi IMNI ditambahkan. Sesuaikan jumlahnya di sheet Komposisi bila perlu.');
}

function aktifkanEmail() { api_setEmailTrigger(true); SpreadsheetApp.getUi().alert('Pengingat email harian aktif (setiap pukul 07.00).'); }
function matikanEmail() { api_setEmailTrigger(false); SpreadsheetApp.getUi().alert('Pengingat email harian dimatikan.'); }

function bukaAplikasi() {
  const html = HtmlService.createHtmlOutputFromFile('Index').setWidth(1280).setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'LABSOMA PPS - Pembuatan Soal');
}

// ====================== UTILITAS ======================
const TZ_ = () => Session.getScriptTimeZone();
const norm_ = s => String(s == null ? '' : s).toLowerCase().replace(/[^\p{L}\p{N}]/gu, ''); // huruf Latin & Arab
const nowStr_ = () => Utilities.formatDate(new Date(), TZ_(), 'yyyy-MM-dd HH:mm');

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(key) {
  const def = SCHEMA[key];
  let sh = ss_().getSheetByName(def.sheet);
  if (!sh) { setup(); sh = ss_().getSheetByName(def.sheet); }
  return sh;
}

function newId_(key) {
  const p = { ustadz: 'U', mapel: 'M', tugas: 'T', murid: 'S', kisi: 'K' }[key] || 'X';
  return p + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
}

function readAll_(key) {
  const def = SCHEMA[key];
  const sh = ss_().getSheetByName(def.sheet);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, def.cols.length).getValues();
  return vals
    .filter(r => r.some(c => c !== ''))
    .map(r => {
      const o = {};
      def.cols.forEach((c, i) => {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, TZ_(), 'yyyy-MM-dd');
        o[c] = (v === null || v === undefined) ? '' : String(v);
      });
      return o;
    });
}

function toRow_(key, o) {
  return SCHEMA[key].cols.map(c => (o[c] === undefined || o[c] === null) ? '' : o[c]);
}

function canon_(list, val) {
  const n = norm_(val);
  return list.find(x => norm_(x) === n) || '';
}

function kelasNum_(v) {
  const m = String(v == null ? '' : v).match(/\d+/);
  return m ? m[0] : String(v || '').trim();
}

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ_(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  const d = new Date(s);
  return isNaN(d) ? s : Utilities.formatDate(d, TZ_(), 'yyyy-MM-dd');
}

const BULAN_ = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
function tglId_(iso) {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? (Number(m[3]) + ' ' + BULAN_[Number(m[2]) - 1] + ' ' + m[1]) : (iso || '-');
}

// ====================== SETUP & DATA AWAL ======================
function setup(formatTextColumns) {
  const shouldFormatText = formatTextColumns !== false;
  const ss = ss_();
  Object.keys(SCHEMA).forEach(key => {
    const def = SCHEMA[key];
    let sh = ss.getSheetByName(def.sheet);
    if (!sh) sh = ss.insertSheet(def.sheet);
    const head = sh.getRange(1, 1, 1, def.cols.length).getValues()[0];
    if (head.join('|') !== def.cols.join('|')) {
      sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols])
        .setFontWeight('bold').setBackground('#0f2e2a').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    // Formatting entire columns on every page load is expensive. Format only
    // existing rows during an explicit setup; api_init skips this recurring work.
    if (shouldFormatText) {
      const rowsToFormat = Math.max(sh.getLastRow(), 1);
      def.text.forEach(c => {
        const i = def.cols.indexOf(c) + 1;
        if (i > 0) sh.getRange(1, i, rowsToFormat, 1).setNumberFormat('@');
      });
    }
  });
  if (readAll_('mapel').length === 0) seedMapel_();
  if (readAll_('komposisi').length === 0) seedKomposisi_();
  if (readAll_('bobot').length === 0) seedBobot_();
  if (readAll_('petunjuk').length === 0) append_('petunjuk', Object.keys(PETUNJUK_AWAL).map(k => [k, PETUNJUK_AWAL[k]]));
  if (readAll_('pengaturan').length === 0) append_('pengaturan', Object.keys(PENGATURAN_AWAL).map(k => [k, PENGATURAN_AWAL[k]]));
  return true;
}

function append_(key, rows) {
  if (!rows.length) return;
  const sh = ss_().getSheetByName(SCHEMA[key].sheet);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, SCHEMA[key].cols.length).setValues(rows);
}

function seedMapel_() {
  const T = 'Tsanawiyah', rows = [];
  const add = (nama, kelasList, kat) => kelasList.forEach(k => rows.push(toRow_('mapel', { ID: newId_('mapel'), Tingkat: T, Kelas: String(k), Mapel: nama, Kategori: kat })));
  add('Seluruh Fan', [1, 2, 3], 'Seluruh Fan');
  ['Fiqh', 'Nahwu', 'Tauhid'].forEach(n => add(n, [1, 2, 3], 'Mapel Utama'));
  add('Faroid', [1], 'Faroid');
  ['Musthalah', 'Mantiq', 'Ilmu Tafsir'].forEach(n => add(n, [2, 3], 'Mapel Tambahan'));
  append_('mapel', rows);
}

// Komposisi soal DIBUAT per Tingkat-Kategori-Kelas-Ujian-Tipe-Materi (sumber: panduan MMU Tsanawiyah).
// Soal jadi = dibuat - 1 per tipe. Untuk IMNI atau tingkat lain, tambahkan baris di sheet "Komposisi".
function seedKomposisi_() {
  const T = 'Tsanawiyah';
  const DEF = [
    ['Seluruh Fan', '1,2,3', {
      'IMDA I':   { A: { I: 11 }, B: { I: 7 }, C: { I: 6 } },
      'IMDA II':  { A: { I: 5, II: 8 }, B: { I: 3, II: 5 }, C: { I: 3, II: 4 } },
      'IMDA III': { A: { I: 4, II: 5, III: 8 }, B: { I: 2, II: 3, III: 4 }, C: { I: 1, II: 2, III: 4 } } }],
    ['Mapel Utama', '1,2,3', {
      'IMDA I':   { A: { I: 16 }, C: { I: 9 } },
      'IMDA II':  { A: { I: 8, II: 13 }, C: { I: 4, II: 5 } },
      'IMDA III': { A: { I: 4, II: 8, III: 14 }, C: { I: 1, II: 2, III: 3 } } }],
    ['Faroid', '1', {
      'IMDA I':   { A: { I: 11 }, C: { I: 11 } },
      'IMDA II':  { A: { I: 6, II: 10 }, C: { I: 5, II: 7 } },
      'IMDA III': { A: { I: 4, II: 5, III: 8 }, C: { I: 3, II: 5, III: 6 } } }],
    ['Mapel Tambahan', '2,3', {
      'IMDA I':  { A: { I: 8 }, B: { I: 5 }, C: { I: 5 } },
      'IMDA II': { A: { I: 4, II: 7 }, B: { I: 3, II: 4 }, C: { I: 2, II: 4 } } }],
    // IMDA III kelas 3 Mapel Tambahan memakai Ujian Akhir (tidak ada komposisi)
    ['Mapel Tambahan', '2', {
      'IMDA III': { A: { I: 2, II: 4, III: 7 }, B: { I: 2, II: 2, III: 4 }, C: { I: 1, II: 2, III: 4 } } }]
  ];
  const rows = [];
  DEF.forEach(([kat, kelas, ujians]) => {
    Object.keys(ujians).forEach(uj => {
      Object.keys(ujians[uj]).forEach(tipe => {
        Object.keys(ujians[uj][tipe]).forEach(m => {
          rows.push(toRow_('komposisi', { Tingkat: T, Kategori: kat, Kelas: kelas, Ujian: uj, Tipe: tipe, Materi: m, Jumlah: ujians[uj][tipe][m] }));
        });
      });
    });
  });
  append_('komposisi', rows);
}

// Bobot skor per butir soal benar
function seedBobot_() {
  const T = 'Tsanawiyah', rows = [];
  const U = ['IMDA I', 'IMDA II', 'IMDA III'];
  const put = (kat, tipe, vals) => U.forEach((uj, i) => {
    if (vals[i] != null) rows.push(toRow_('bobot', { Tingkat: T, Kategori: kat, Ujian: uj, Tipe: tipe, Bobot: vals[i] }));
  });
  KATEGORI.forEach(k => put(k, 'A', [4, 3, 2]));
  put('Seluruh Fan', 'B', [5, 4, 4]);
  put('Mapel Tambahan', 'B', [8, 5, 4]);
  put('Seluruh Fan', 'C', [6, 6, 6]);
  put('Faroid', 'C', [6, 5, 4]);
  put('Mapel Utama', 'C', [5, 5, 5]);
  put('Mapel Tambahan', 'C', [10, 6, 6]);
  append_('bobot', rows);
}

// ====================== API UNTUK ANTARMUKA ======================
function api_all() {
  return {
    ustadz: readAll_('ustadz'),
    mapel: readAll_('mapel'),
    tugas: readAll_('tugas'),
    murid: readAll_('murid'),
    kisi: (ensureKisiIds_(), readAll_('kisi')),
    komposisi: readAll_('komposisi'),
    bobot: readAll_('bobot'),
    petunjuk: readAll_('petunjuk'),
    pengaturan: pengaturan_(),
    emailAktif: emailTriggerAktif_(),
    enums: { TINGKAT: TINGKAT, KATEGORI: semuaKategori_(), UJIAN: UJIAN, UNIT: UNIT, STATUS: STATUS }
  };
}

function api_init() {
  // Ensure schemas and seed data without reformatting every sheet at each launch.
  setup(false);
  return api_all();
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function dupKey_(key, o) {
  const n = norm_;
  switch (key) {
    case 'ustadz': return n(o.Nama);
    case 'mapel':  return [n(o.Tingkat), n(o.Kelas), n(o.Mapel)].join('|');
    case 'tugas':  return [n(o.Ustadz), n(o.Tingkat), n(o.Kelas), n(o.Mapel), n(o.Ujian), n(o.Unit)].join('|');
    case 'murid':  return o.NIS ? n(o.NIS) : n(o.Nama);
    case 'kisi':   return [n(o.Tingkat), n(o.Kelas), n(o.Fan), n(o.Imda), n(o.Materi)].join('|');
  }
  return '';
}

// "IMDA 2", "2", "II", "imda II" -> 'II' ; tidak dikenal -> ''
function imdaRom_(v) {
  if (/IMNI/i.test(String(v == null ? '' : v))) return 'IMNI'; // ujian akhir, tanpa tahap 1-3
  const s = String(v == null ? '' : v).toUpperCase().replace(/IMDA|[^A-Z0-9]/g, '');
  return ({ '1': 'I', 'I': 'I', '2': 'II', 'II': 'II', '3': 'III', 'III': 'III' })[s] || '';
}

function cleanRecord_(key, rec) {
  const o = Object.assign({}, rec);
  if (o.Tingkat !== undefined) o.Tingkat = canon_(TINGKAT, o.Tingkat) || o.Tingkat;
  if (key === 'mapel' || key === 'tugas' || key === 'kisi') o.Kelas = kelasNum_(o.Kelas);
  if (key === 'kisi') o.Imda = imdaRom_(o.Imda);
  if (key === 'mapel') o.Kategori = canon_(KATEGORI, o.Kategori) || o.Kategori;
  if (key === 'tugas') {
    o.Ujian = canon_(UJIAN, o.Ujian) || o.Ujian;
    o.Unit = canon_(UNIT, o.Unit) || o.Unit;
    o.Deadline = normDate_(o.Deadline);
  }
  Object.keys(o).forEach(k => { if (typeof o[k] === 'string') o[k] = o[k].trim(); });
  return o;
}

function api_save(key, rec) {
  if (!SCHEMA[key] || ['soal', 'komposisi', 'bobot', 'petunjuk', 'pengaturan'].includes(key)) throw new Error('Data tidak dikenal.');
  return withLock_(() => {
    const def = SCHEMA[key];
    const o = cleanRecord_(key, rec);
    def.req.forEach(f => { if (!o[f]) throw new Error('Kolom "' + f + '" wajib diisi.'); });
    if (key === 'kisi' && (!TINGKAT.includes(o.Tingkat) || !o.Imda)) throw new Error('Tingkat atau IMDA tidak valid (IMDA harus I/II/III atau IMNI).');
    if (key === 'tugas') {
      if (!TINGKAT.includes(o.Tingkat) || !UJIAN.includes(o.Ujian) || !UNIT.includes(o.Unit)) throw new Error('Tingkat, Ujian, atau Unit tidak valid.');
      if (!readAll_('ustadz').some(u => norm_(u.Nama) === norm_(o.Ustadz))) throw new Error('Ustadz belum ada di Data Ustadz.');
      if (!readAll_('mapel').some(m => m.Tingkat === o.Tingkat && m.Kelas === o.Kelas && norm_(m.Mapel) === norm_(o.Mapel))) throw new Error('Mapel tidak ditemukan untuk Tingkat & Kelas tersebut.');
    }
    const all = readAll_(key);
    const dup = all.find(r => r.ID !== o.ID && dupKey_(key, r) === dupKey_(key, o));
    if (dup) throw new Error('Data yang sama sudah ada.');
    const sh = sheet_(key);
    if (o.ID) {
      const idx = all.findIndex(r => r.ID === o.ID);
      if (idx < 0) throw new Error('Data tidak ditemukan.');
      const merged = Object.assign({}, all[idx], o);
      if (key === 'tugas') merged.Diperbarui = nowStr_();
      // baris data ke-idx berada di baris sheet idx+2 hanya jika tidak ada baris kosong; cari lewat ID
      const row = rowOfId_(sh, o.ID);
      sh.getRange(row, 1, 1, def.cols.length).setValues([toRow_(key, merged)]);
    } else {
      o.ID = newId_(key);
      if (key === 'tugas') { o.Status = 'Belum Diisi'; o.Diperbarui = nowStr_(); }
      append_(key, [toRow_(key, o)]);
    }
    return api_all();
  });
}

function rowOfId_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) throw new Error('Data tidak ditemukan.');
  const col = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf('ID') + 1;
  if (col < 1) throw new Error('Kolom ID tidak ada.');
  const ids = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  throw new Error('Data tidak ditemukan.');
}

// Baris kisi-kisi yang ditempel langsung ke sheet (tanpa ID) diberi ID otomatis
function ensureKisiIds_() {
  const sh = ss_().getSheetByName(SCHEMA.kisi.sheet);
  if (!sh || sh.getLastRow() < 2) return;
  const cols = SCHEMA.kisi.cols, idCol = cols.indexOf('ID');
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getValues();
  let changed = false;
  const ids = vals.map(r => {
    if (r[idCol] === '' && r.slice(0, idCol).some(v => v !== '')) { changed = true; return [newId_('kisi')]; }
    return [r[idCol]];
  });
  if (changed) sh.getRange(2, idCol + 1, ids.length, 1).setValues(ids);
}

function api_delete(key, id) {
  if (!['ustadz', 'mapel', 'tugas', 'murid', 'kisi'].includes(key)) throw new Error('Data tidak dikenal.');
  return withLock_(() => {
    sheet_(key).deleteRow(rowOfId_(sheet_(key), id));
    if (key === 'tugas') clearSoal_(id);
    return api_all();
  });
}

// ---------- Impor Excel (data sudah dibaca di browser, dikirim sebagai array objek) ----------
function api_import(key, rows) {
  if (!ALIAS[key]) throw new Error('Data tidak dikenal.');
  return withLock_(() => {
    const def = SCHEMA[key];
    const fields = def.cols.filter(c => c !== 'ID' && c !== 'Status' && c !== 'Diperbarui');
    const ali = ALIAS[key];
    const result = { ditambah: 0, dilewati: 0, galat: [] };
    if (!rows || !rows.length) throw new Error('File tidak berisi baris data.');

    // petakan header Excel -> kolom
    const headers = Object.keys(rows[0]);
    const map = {};
    const used = {};
    const aliasOf = f => (ali[f] || [norm_(f)]).concat([norm_(f)]);
    fields.forEach(f => {
      let h;
      for (const a of aliasOf(f)) { h = headers.find(x => !used[x] && norm_(x) === a); if (h) break; }
      if (h) { map[f] = h; used[h] = true; }
    });
    fields.forEach(f => {
      if (map[f]) return;
      const h = headers.find(x => !used[x] && aliasOf(f).some(a => norm_(x).indexOf(a) === 0));
      if (h) { map[f] = h; used[h] = true; }
    });
    const missing = def.req.filter(f => !map[f]);
    if (missing.length) throw new Error('Kolom tidak ditemukan di Excel: ' + missing.join(', ') + '. Gunakan template yang disediakan.');

    const ustadzNames = {};
    readAll_('ustadz').forEach(u => { ustadzNames[norm_(u.Nama)] = u.Nama; });
    const mapels = readAll_('mapel');
    const seen = {};
    readAll_(key).forEach(r => { seen[dupKey_(key, r)] = true; });
    const out = [];

    rows.forEach((raw, i) => {
      const line = i + 2;
      const o = {};
      fields.forEach(f => { o[f] = map[f] ? raw[map[f]] : ''; });
      if (fields.every(f => String(o[f] == null ? '' : o[f]).trim() === '')) return; // baris kosong
      let c = cleanRecord_(key, o);
      if (key === 'murid') c.Kelas = String(o.Kelas || '').trim();
      const miss = def.req.find(f => !c[f]);
      if (miss) { result.galat.push('Baris ' + line + ': kolom "' + miss + '" kosong.'); return; }

      if (key === 'kisi') {
        if (!TINGKAT.includes(c.Tingkat)) { result.galat.push('Baris ' + line + ': Tingkat "' + c.Tingkat + '" tidak dikenal.'); return; }
        if (!c.Imda) { result.galat.push('Baris ' + line + ': IMDA harus 1, 2, atau 3.'); return; }
      }
      if (key === 'ustadz' && c.Tingkat && !TINGKAT.includes(c.Tingkat)) { result.galat.push('Baris ' + line + ': Tingkat "' + c.Tingkat + '" tidak dikenal.'); return; }

      if (key === 'mapel') {
        if (!TINGKAT.includes(c.Tingkat)) { result.galat.push('Baris ' + line + ': Tingkat "' + c.Tingkat + '" tidak dikenal.'); return; }
        if (!String(c.Kategori || '').trim()) { result.galat.push('Baris ' + line + ': Kategori Soal kosong.'); return; }
      }

      if (key === 'tugas') {
        if (!TINGKAT.includes(c.Tingkat)) { result.galat.push('Baris ' + line + ': Tingkat tidak dikenal.'); return; }
        if (!UJIAN.includes(c.Ujian)) { result.galat.push('Baris ' + line + ': Tahap Ujian harus salah satu dari ' + UJIAN.join(', ') + '.'); return; }
        if (!UNIT.includes(c.Unit)) { result.galat.push('Baris ' + line + ': Kategori Unit harus Induk atau Ranting.'); return; }
        const nm = ustadzNames[norm_(c.Ustadz)];
        if (!nm) { result.galat.push('Baris ' + line + ': Ustadz "' + c.Ustadz + '" belum ada di Data Ustadz.'); return; }
        c.Ustadz = nm;
        const mp = mapels.find(m => m.Tingkat === c.Tingkat && m.Kelas === c.Kelas && norm_(m.Mapel) === norm_(c.Mapel));
        if (!mp) { result.galat.push('Baris ' + line + ': Mapel "' + c.Mapel + '" kelas ' + c.Kelas + ' (' + c.Tingkat + ') belum ada di Data Mapel.'); return; }
        c.Mapel = mp.Mapel;
      }
      const dk = dupKey_(key, c);
      if (seen[dk]) { result.dilewati++; return; }
      seen[dk] = true;
      c.ID = newId_(key);
      if (key === 'tugas') { c.Status = 'Belum Diisi'; c.Diperbarui = nowStr_(); }
      out.push(toRow_(key, c));
      result.ditambah++;
    });
    append_(key, out);
    result.data = api_all();
    return result;
  });
}

// ====================== KOMPOSISI & KATEGORI ======================
function kategoriOf_(t) {
  const m = readAll_('mapel').find(x => x.Tingkat === t.Tingkat && String(x.Kelas) === String(t.Kelas) && norm_(x.Mapel) === norm_(t.Mapel));
  return m ? m.Kategori : '';
}

function kategoriCanon_(v) {
  const s = norm_(v);
  if (!s || s === 'seluruhfan' || s === 'semuafan' || s === 'semuamapel') return 'seluruhfan';
  if (s.indexOf('mapelutama') >= 0 || s.indexOf('fanutama') >= 0) return 'mapelutama';
  if (s.indexOf('faroid') >= 0 || s.indexOf('faraid') >= 0) return 'faroid';
  if (s.indexOf('mapeltambahan') >= 0 || s.indexOf('fantambahan') >= 0) return 'mapeltambahan';
  return s;
}

// Kumpulkan semua kategori dari Mapel + Komposisi + Bobot + 4 standar
function semuaKategori_() {
  const set = new Set(KATEGORI);
  ['mapel', 'komposisi', 'bobot'].forEach(key => {
    readAll_(key).forEach(r => {
      const v = String(r.Kategori || '').trim();
      if (v) set.add(v);
    });
  });
  return [...set].sort((a, b) => a.localeCompare(b, 'id', { numeric: true, sensitivity: 'base' }));
}

function compRows_(tingkat, kat, kelas, ujian) {
  const targetUjian = ujianCanon_(ujian) || String(ujian || '').trim();
  const targetKelas = String(kelas == null ? '' : kelas).trim();
  const scoped = readAll_('komposisi').filter(r =>
    norm_(r.Tingkat) === norm_(tingkat) &&
    (ujianCanon_(r.Ujian) || String(r.Ujian || '').trim()) === targetUjian &&
    String(r.Kelas).split(/[,;]+/).map(s => s.trim()).includes(targetKelas));
  const exact = scoped.filter(r => kategoriCanon_(r.Kategori) === kategoriCanon_(kat));
  if (exact.length) return exact;
  return scoped.filter(r => kategoriCanon_(r.Kategori) === 'seluruhfan');
}

function summarize_(rows) {
  const types = {};
  rows.forEach(r => {
    const t = r.Tipe;
    types[t] = types[t] || { tipe: t, total: 0, materi: [] };
    const n = Number(r.Jumlah) || 0;
    types[t].total += n;
    types[t].materi.push({ m: r.Materi, n: n });
  });
  const list = ['A', 'B', 'C', 'D'].filter(t => types[t]).map(t => {
    types[t].materi.sort((a, b) => (ROMAWI[a.m] || 9) - (ROMAWI[b.m] || 9));
    return types[t];
  });
  const dibuat = list.reduce((s, x) => s + x.total, 0);
  return { types: list, dibuat: dibuat, jadi: dibuat - list.length };
}

// ====================== SOAL ======================
function api_getSoal(tugasId) {
  return readAll_('soal').filter(s => s.TugasID === tugasId);
}

function clearSoal_(tugasId) {
  const sh = ss_().getSheetByName(SCHEMA.soal.sheet);
  if (!sh || sh.getLastRow() < 2) return;
  const keep = readAll_('soal').filter(s => s.TugasID !== tugasId);
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), SCHEMA.soal.cols.length).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, SCHEMA.soal.cols.length).setValues(keep.map(s => toRow_('soal', s)));
}

function setStatusRow_(id, status) {
  const sh = sheet_('tugas');
  const row = rowOfId_(sh, id);
  const cols = SCHEMA.tugas.cols;
  sh.getRange(row, cols.indexOf('Status') + 1).setValue(status);
  sh.getRange(row, cols.indexOf('Diperbarui') + 1).setValue(nowStr_());
}

function api_saveSoal(tugasId, items, submit) {
  return withLock_(() => {
    const t = readAll_('tugas').find(x => x.ID === tugasId);
    if (!t) throw new Error('Tugas tidak ditemukan.');
    if (t.Status === 'Approved') throw new Error('Tugas sudah disetujui. Kembalikan statusnya dulu jika ingin mengubah soal.');
    const filled = (items || []).filter(i => String(i.Pertanyaan || '').trim() !== '');
    if (submit) {
      const kat = kategoriOf_(t);
      const sum = summarize_(compRows_(t.Tingkat, kat, t.Kelas, t.Ujian));
      if (sum.dibuat && filled.length < sum.dibuat) {
        throw new Error('Soal belum lengkap: terisi ' + filled.length + ' dari ' + sum.dibuat + ' soal.');
      }
    }
    clearSoal_(tugasId);
    append_('soal', filled.map(i => toRow_('soal', {
      TugasID: tugasId, Tipe: i.Tipe, Materi: i.Materi, No: i.No, Pertanyaan: i.Pertanyaan,
      A: i.A || '', B: i.B || '', C: i.C || '', D: i.D || '', Kunci: i.Kunci || '', Kisi: i.Kisi || ''
    })));
    if (submit) setStatusRow_(tugasId, 'Submitted');
    else if (t.Status === 'Belum Diisi' && filled.length) setStatusRow_(tugasId, 'Process');
    else setStatusRow_(tugasId, t.Status);
    return api_all();
  });
}

function api_setStatus(tugasId, status) {
  if (!STATUS.includes(status)) throw new Error('Status tidak valid.');
  return withLock_(() => { setStatusRow_(tugasId, status); return api_all(); });
}

// ====================== CETAK TEMPLATE (Word .docx asli & PDF) ======================
// File Word dibuat langsung sebagai .docx (kanan-ke-kiri, penomoran otomatis dengan format "( ) 1").
// Model data disusun di browser (templateModel). PDF dibuat sebisanya dari HTML pratinjau.
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

const W_NS_ = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const AR_RPR_ = '<w:rFonts w:ascii="Traditional Arabic" w:hAnsi="Traditional Arabic" w:cs="Traditional Arabic"/><w:sz w:val="30"/><w:szCs w:val="30"/><w:rtl/><w:lang w:bidi="ar-SA"/>';
const NB_ = '\u00A0';

function arRun_(t) { return '<w:r><w:rPr>' + AR_RPR_ + '</w:rPr><w:t xml:space="preserve">' + esc_(t) + '</w:t></w:r>'; }
function arPara_(text, numId, before) {
  return '<w:p><w:pPr>' + (numId ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + numId + '"/></w:numPr>' : '') +
    '<w:bidi/><w:spacing w:before="' + (before || 0) + '" w:after="40"/></w:pPr>' + arRun_(text) + '</w:p>';
}

function tc_(paras, o) {
  o = o || {};
  const pr = '<w:tcW w:w="' + o.w + '" w:type="dxa"/>' + (o.span ? '<w:gridSpan w:val="' + o.span + '"/>' : '') +
    (o.vm ? '<w:vMerge' + (o.vm === 'restart' ? ' w:val="restart"' : '') + '/>' : '') +
    (o.fill ? '<w:shd w:val="clear" w:color="auto" w:fill="' + o.fill + '"/>' : '') + 
    (o.border ? o.border : '') + 
    '<w:vAlign w:val="' + (o.valign || 'center') + '"/>';
  return '<w:tc><w:tcPr>' + pr + '</w:tcPr>' + (paras || '<w:p/>') + '</w:tc>';
}


function cp_(text, o) {
  o = o || {};
  return '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="' + (o.jc || 'left') + '"/></w:pPr>' +
    (text === '' || text == null ? '' : '<w:r><w:rPr><w:rFonts w:ascii="' + (o.font || 'Calibri') + '" w:hAnsi="' + (o.font || 'Calibri') + '" w:cs="' + (o.font || 'Calibri') + '"/>' + (o.bold ? '<w:b/>' : '') +
      '<w:color w:val="' + (o.color || '12302C') + '"/><w:sz w:val="' + (o.sz || 21) + '"/><w:szCs w:val="' + (o.sz || 21) + '"/></w:rPr><w:t xml:space="preserve">' + esc_(text) + '</w:t></w:r>') + '</w:p>';
}
function tbl_(widths, rows, opt) {
  opt = opt || {};
  const B = opt.none ? '<w:{k} w:val="nil"/>' : '<w:{k} w:val="single" w:sz="4" w:space="0" w:color="' + (opt.color || 'B7C8C3') + '"/>';
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(k => B.replace('{k}', k)).join('');
  return '<w:tbl><w:tblPr><w:tblW w:w="' + widths.reduce((a, b) => a + b, 0) + '" w:type="dxa"/><w:tblBorders>' + borders + '</w:tblBorders>' +
    '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="110" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="110" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
    '<w:tblGrid>' + widths.map(w => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid>' +
    rows.map((r, i) => '<w:tr><w:trPr><w:trHeight w:val="' + (opt.h || 400) + '"/>' + (i < (opt.head || 0) ? '<w:tblHeader/>' : '') + '</w:trPr>' + r + '</w:tr>').join('') + '</w:tbl>';
}

function para_(text, o) {
  o = o || {};
  return '<w:p><w:pPr>' + (o.pbb ? '<w:pageBreakBefore/>' : '') + '<w:spacing w:before="' + (o.before || 0) + '" w:after="' + (o.after == null ? 0 : o.after) + '"/><w:jc w:val="' + (o.jc || 'left') + '"/></w:pPr>' +
    (text ? '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>' + (o.bold ? '<w:b/>' : '') + '<w:color w:val="' + (o.color || '12302C') + '"/><w:sz w:val="' + (o.sz || 22) + '"/><w:szCs w:val="' + (o.sz || 22) + '"/></w:rPr><w:t xml:space="preserve">' + esc_(text) + '</w:t></w:r>' : '') + '</w:p>';
}

function runs_(list) {
  return list.map(r => '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' + (r.bold ? '<w:b/>' : '') + (r.color ? '<w:color w:val="' + r.color + '"/>' : '') +
    '<w:sz w:val="' + (r.sz || 28) + '"/><w:szCs w:val="' + (r.sz || 28) + '"/></w:rPr><w:t xml:space="preserve">' + esc_(r.t) + '</w:t></w:r>').join('');
}
function paraRuns_(list, o) {
  o = o || {};
  return '<w:p><w:pPr>' + (o.pbb ? '<w:pageBreakBefore/>' : '') + '<w:spacing w:before="' + (o.before || 0) + '" w:after="' + (o.after || 0) + '"/><w:jc w:val="' + (o.jc || 'center') + '"/></w:pPr>' + runs_(list) + '</w:p>';
}
function arHead_(text) {
  return '<w:p><w:pPr><w:bidi/><w:spacing w:before="240" w:after="60"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Traditional Arabic" w:hAnsi="Traditional Arabic" w:cs="Traditional Arabic"/><w:b/><w:bCs/><w:sz w:val="36"/><w:szCs w:val="36"/><w:rtl/><w:lang w:bidi="ar-SA"/></w:rPr><w:t xml:space="preserve">' + esc_(text) + '</w:t></w:r></w:p>';
}
const arNum = n => String(n).replace(/[0-9]/g, d => '٠١٢٣٤٥٦٧٨٩'[d]);
// Halaman kunci jawaban (mengikuti file contoh): judul, baris rincian merah, lalu أ ~ ~ ~ / ب ~ ~ ~ / ج ~ ~ ~
// dengan daftar bernomor "N (      )" dan isi "jawaban / halaman".

function kunciPage_(m) {
  // Menggunakan numId 5, 6, 7, 8 yang sudah disetting restart dari 1
  const red = m.warna || 'FF0000', L = { A: 'أ', B: 'ب', C: 'ج', D: 'د' }, numOf = { A: 5, B: 6, C: 7, D: 8 };
  
  let x = paraRuns_([{ t: 'LEMBAR KUNCI JAWABAN', bold: true, sz: 36 }], { pbb: true, after: 60 }) +
    paraRuns_([{ t: m.head.rincian, bold: true, color: red, sz: 30 }], { after: 120 });
  
  m.kunci.tipe.forEach(k => {
    x += arHead_(L[k.tipe] + ' ~ ~ ~');
    for (let i = 0; i < k.n; i++) {
      // Hanya spasi dan garis miring, tanpa kurung
      x += arPara_(NB_.repeat(8) + '/' + NB_.repeat(8), numOf[k.tipe], 0);
    }
  });
  return x;
}

// Halaman kisi-kisi (mengikuti file contoh): NO | IMDA | MATERI / BAB / FASAL | MODEL SOAL (3 kolom) | JUMLAH
function kisiPage_(m) {
  const red = m.warna || 'FF0000', GRAY = 'D9D9D9', W = [500, 900, 3300, 1200, 1200, 1200, 1200, 1606], K = '000000';
  const c = (t, o) => cp_(t, Object.assign({ font: 'Arial', sz: 22, jc: 'center', color: K }, o || {}));
  const hd = (t, w, o) => tc_(c(t, { bold: true }), Object.assign({ w: w, fill: GRAY }, o || {}));
  const cont = w => tc_('<w:p/>', { w: w, vm: 'continue', fill: GRAY });
  const rows = [
    hd('NO', W[0], { vm: 'restart' }) + hd(m.detail.ujian === 'IMNI' ? 'KELAS SUMBER' : 'IMDA', W[1], { vm: 'restart' }) + hd('MATERI / BAB / FASAL', W[2], { vm: 'restart' }) +
      hd('MODEL SOAL', W[3] + W[4] + W[5] + W[6], { span: 4 }) + hd('JUMLAH', W[7], { vm: 'restart' }),
    cont(W[0]) + cont(W[1]) + cont(W[2]) + hd('PILIHAN GANDA', W[3]) + hd('MENJODOHKAN / ISIAN', W[4]) + hd('MELENGKAPI', W[5]) + hd('ESSAY', W[6]) + cont(W[7])
  ];
  const data = m.kisi.rows.slice();
  data.push({}, {});
  data.forEach(r => {
    rows.push(tc_(c(r.no || ''), { w: W[0] }) + tc_(c(r.imda || ''), { w: W[1] }) +
      tc_(c(r.materi || '', { font: 'Traditional Arabic', sz: 26 }), { w: W[2] }) + tc_(c(r.A || ''), { w: W[3] }) +
      tc_(c(r.B || ''), { w: W[4] }) + tc_(c(r.C || ''), { w: W[5] }) + tc_(c(r.D || ''), { w: W[6] }) + tc_(c(r.jml == null ? '' : r.jml), { w: W[7] }));
  });
  const big = (t, o) => paraRuns_([{ t: t, sz: 28 }], o);
  let x = big('KISI-KISI SOAL', { pbb: true }) +
    paraRuns_([{ t: 'MADRASAH MIFTAHUL ULUM ', sz: 28 }, { t: m.head.tingkatUnit, sz: 28, bold: true, color: red }]) +
    big('PONDOK PESANTREN SIDOGIRI') +
    paraRuns_([{ t: 'KATEGORI ', sz: 28 }, { t: m.head.kategori, sz: 28, bold: true, color: red }], { after: 200 }) +
    paraRuns_([{ t: m.head.rincian, bold: true, color: red, sz: 30 }], { after: 160 });
  // judul kolom hanya di awal tabel (tidak diulang di halaman berikutnya)
  x += tbl_(W, rows, { color: K, h: 300 });
  return x;
}

function buildDocx_(m) {
  const D = m.detail, DARK = '0F4C43', TINT = 'E6F0ED', SOFT = 'F6FAF9';
  const lab = (t, w) => tc_(cp_(t, { bold: true, color: DARK, sz: 19 }), { w: w, fill: TINT });
  const val = (t, w) => tc_(cp_(t), { w: w });
  const isImni = /IMNI|UJIAN AKHIR/i.test(String(D.ujian || ''));
  const kors = (D.kor || [D.kor1, D.kor2]).slice(0, isImni ? 3 : 2);
  const roles = isImni ? ['PIMPINAN LABSOMA', 'PIMPINAN MADRASAH', 'KETUA 1'] : ['PIMPINAN MADRASAH', 'PIMPINAN LABSOMA'];
  const kor = (t, v, w, vm) => {
    const content = vm === 'continue' ? '<w:p/>' :
      cp_(t, { bold: true, color: DARK, sz: 20, jc: 'center', after: 260 }) +
      cp_('(_______________)', { jc: 'center', sz: 22, before: 1500, after: 100 }) +
      cp_('Tgl: ______________', { jc: 'center', sz: 19, before: 0 });
    return tc_(content, { w: w, fill: SOFT, valign: 'top', vm: vm });
  };
  const W1 = isImni ? [800, 1150, 1150, 1550, 2140, 2140, 2130] : [1000, 1350, 1350, 1900, 2660, 2660];
  const korCells = vm => kors.map((v, i) => kor(roles[i], v, W1[4 + i], vm)).join('');
  const t1 = tbl_(W1, [
    lab('TAHAP', W1[0]) + val(D.ujian, W1[1]) + lab('TINGKAT', W1[2]) + val(D.tingkat, W1[3]) + korCells('restart'),
    lab('KELAS', W1[0]) + val(D.kelas, W1[1]) + lab('MADRASAH', W1[2]) + val(D.unit, W1[3]) + korCells('continue'),
    lab('MATERI', W1[0]) + val(D.mapel, W1[1]) + lab('SONIK', W1[2]) + val(D.sonik, W1[3]) + korCells('continue')
  ]);

  const stages = m.mek.stages, nS = stages.length;
  const WT = 2700, WJ = 2100, WS = 1500, rest = 11106 - WT - WJ - WS, wc = Math.floor(rest / (nS + 1));
  const W2 = [WT].concat(stages.map(() => wc), [rest - wc * nS, WJ, WS]);
  const hd = (t, w, o) => tc_(cp_(t, { bold: true, color: 'FFFFFF', sz: 19, jc: 'center' }), Object.assign({ w: w, fill: DARK }, o || {}));
  const hs = (t, w) => tc_(cp_(t, { bold: true, color: DARK, sz: 19, jc: 'center' }), { w: w, fill: TINT });
  const dark = w => tc_('<w:p/>', { w: w, vm: 'continue', fill: DARK });
  const rows2 = [
    hd('TIPE', W2[0], { vm: 'restart' }) + hd('KOMPOSISI SOAL', rest, { span: nS + 1 }) + hd('JUMLAH SOAL JADI', WJ, { vm: 'restart' }) + hd('SKOR', WS, { vm: 'restart' }),
    dark(W2[0]) + stages.map((s, k) => hs(s, W2[1 + k])).join('') + hs('JML', W2[1 + nS]) + dark(WJ) + dark(WS)
  ];
  m.mek.rows.forEach((r, i) => {
    const f = i % 2 ? SOFT : 'FFFFFF';
    const c = (v, w, b) => tc_(cp_(v, { jc: 'center', bold: b }), { w: w, fill: f });
    rows2.push(c(r.tipe, W2[0], true) + r.vals.map((v, k) => c(v, W2[1 + k])).join('') + c(r.jml, W2[1 + nS]) + c(r.jadi, WJ, true) + c(r.skor, WS));
  });
  const t2 = tbl_(W2, rows2);

  const hasComp = !!(m.mek && m.mek.rows && m.mek.rows.length);
  let body = t1;
  if (hasComp) {
    body += '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="160" w:lineRule="exact"/></w:pPr></w:p>' + t2;
    const numIdOf = { A: 1, B: 2, C: 3, D: 4 };
    
    m.sections.forEach(sec => {
      body += arPara_(sec.heading, 0, 320);
      
      sec.items.forEach(it => {
        let t;
        if (sec.tipe === 'A') {
          const o = (hr, v, isLast) => hr + '.' + (v ? NB_ + v : '') + (isLast ? '' : NB_.repeat(3));
          t = (it.q ? it.q + NB_ : '') + '(' + o('أ', it.A) + o('ب', it.B) + o('ج', it.C) + o('د', it.D, true) + ')';
        } else if (sec.tipe === 'B' || (m.detail.ujian === 'IMNI' && sec.tipe === 'C')) {
          t = it.q || '.'.repeat(40);
        } else {
          t = it.q || '؟';
        }
        body += arPara_(t, numIdOf[sec.tipe] || 3, 0);
      });
    });
  }
  
  const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="18720"/><w:pgMar w:top="567" w:right="567" w:bottom="533" w:left="567" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>';
  if (hasComp && m.kunci && m.kisi) body += kunciPage_(m) + kisiPage_(m);

  const documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W_NS_ + '><w:body>' + body + sect + '</w:body></w:document>';

  const lvl = (text, ind, format) => '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="' + (format || 'decimal') + '"/><w:suff w:val="tab"/><w:lvlText w:val="' + text + '"/><w:lvlJc w:val="left"/>' +
    '<w:pPr><w:ind w:left="' + ind + '" w:hanging="' + ind + '"/></w:pPr><w:rPr>' + AR_RPR_.replace('<w:rtl/>', '<w:rtl/><w:cs/>') + '</w:rPr></w:lvl>';
  
  const numberingXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + W_NS_ + '>' +
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>' + lvl('%1' + NB_ + NB_ + '(' + NB_.repeat(6) + ')', 1300, 'decimal') + '</w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>' + lvl('%1.', 567, 'decimal') + '</w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="singleLevel"/>' + lvl('%1.', 567, 'hindiNumbers') + '</w:abstractNum>' +
    [[1, 0], [2, 1], [3, 1], [4, 2], [5, 2], [6, 2], [7, 2]].map(p => '<w:num w:numId="' + p[0] + '"><w:abstractNumId w:val="' + p[1] + '"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>').join('') +
    
        
    // KOTAK KIRI TIPE B (Huruf Arab + Garis Miring)
    '<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="singleLevel"/>' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="arabicAbjad"/><w:suff w:val="nothing"/><w:lvlText w:val="%1. /"/><w:lvlJc w:val="center"/>' +
    '<w:pPr><w:bidi/><w:ind w:left="0" w:hanging="0"/></w:pPr><w:rPr>' + AR_RPR_ + '</w:rPr></w:lvl></w:abstractNum>' +

    // KOLOM KANAN TIPE B (Angka 1,2,3 RTL)
    '<w:abstractNum w:abstractNumId="4"><w:multiLevelType w:val="singleLevel"/>' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="nothing"/><w:lvlText w:val="%1."/><w:lvlJc w:val="right"/>' +
    '<w:pPr><w:bidi/><w:ind w:right="0" w:left="0" w:hanging="0" w:firstLine="0"/></w:pPr><w:rPr><w:rtl/><w:sz w:val="24"/></w:rPr></w:lvl></w:abstractNum>' +


    // Daftar Penomoran
    [[1, 0], [2, 1], [3, 1], [4, 2], [5, 2], [6, 2], [7, 2]].map(p => '<w:num w:numId="' + p[0] + '"><w:abstractNumId w:val="' + p[1] + '"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>').join('') +
    '<w:num w:numId="9"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' +
    '<w:num w:numId="10"><w:abstractNumId w:val="4"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' +
    '</w:numbering>';

  const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + W_NS_ + '><w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Traditional Arabic"/><w:sz w:val="22"/><w:szCs w:val="30"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>';
  const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>';
  const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
  const nb = (xml, name) => Utilities.newBlob(xml, 'application/xml', name);
  return [nb(ct, '[Content_Types].xml'), nb(rels, '_rels/.rels'), nb(documentXml, 'word/document.xml'),
    nb(numberingXml, 'word/numbering.xml'), nb(stylesXml, 'word/styles.xml'), nb(docRels, 'word/_rels/document.xml.rels')];
}

const DOCX_MIME_ = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const aman_ = s => String(s || 'Lembar Soal').replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim().slice(0, 120);

function docxBlob_(m, nama) {
  return Utilities.zip(buildDocx_(m), nama + '.docx').setContentType(DOCX_MIME_);
}

// PDF: bila layanan lanjutan "Drive API" diaktifkan, PDF dibuat dari file Word (hasil terbaik).
// Bila tidak, dibuat dari HTML pratinjau (lebih sederhana).
function pdfBlob_(html, m, nama) {
  try {
    if (typeof Drive !== 'undefined') {
      const blob = docxBlob_(m, nama);
      const id = Drive.Files.create
        ? Drive.Files.create({ name: nama, mimeType: 'application/vnd.google-apps.document' }, blob).id
        : Drive.Files.insert({ title: nama, mimeType: 'application/vnd.google-apps.document' }, blob, { convert: true }).id;
      const f = DriveApp.getFileById(id);
      const pdf = f.getAs('application/pdf').setName(nama + '.pdf');
      f.setTrashed(true);
      return { blob: pdf, via: 'docx' };
    }
  } catch (e) { /* lanjut ke cadangan HTML */ }
  const page = '<html><head><meta charset="utf-8"><style>@page{size:8.5in 13in;margin:1cm}</style></head><body>' + (html || '') + '</body></html>';
  return { blob: Utilities.newBlob(page, 'text/html', nama + '.html').getAs('application/pdf').setName(nama + '.pdf'), via: 'html' };
}

function berkas_(blob, via) {
  return { name: blob.getName(), mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()), via: via || '' };
}

function api_exportWord(model, nama) {
  if (!model || !model.sections) throw new Error('Data template kosong.');
  return berkas_(docxBlob_(model, aman_(nama)));
}

function api_exportPdf(html, model, nama) {
  if (!model || !model.sections) throw new Error('Data template kosong.');
  const r = pdfBlob_(html, model, aman_(nama));
  return berkas_(r.blob, r.via);
}

// Semua template satu ustadz dalam satu ZIP. items: [{nama, model, html}], kind: 'word' | 'pdf'
function api_exportZip(items, kind, zipNama) {
  if (!items || !items.length) throw new Error('Tidak ada template.');
  const pakai = {}, blobs = [];
  let via = '';
  items.forEach(it => {
    let nama = aman_(it.nama);
    pakai[nama] = (pakai[nama] || 0) + 1;
    if (pakai[nama] > 1) nama += ' (' + pakai[nama] + ')';
    if (kind === 'pdf') { const r = pdfBlob_(it.html, it.model, nama); blobs.push(r.blob); via = r.via; }
    else blobs.push(docxBlob_(it.model, nama));
  });
  return berkas_(Utilities.zip(blobs, aman_(zipNama || 'Template Soal') + '.zip'), via);
}

// ====================== IMPOR KOMPOSISI & BOBOT (menu Mekanisme Soal) ======================
function ujianCanon_(v) {
  const s = String(v == null ? '' : v).toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (s.indexOf('IMNI') === 0 || s === 'UJIANAKHIR' || s === 'UJIANKHIR' || s === 'AKHIR') return 'IMNI';
  const m = s.match(/^IMDA(1|2|3|I|II|III)$/);
  return m ? 'IMDA ' + ({ '1': 'I', '2': 'II', '3': 'III' }[m[1]] || m[1]) : '';
}
function materiCanon_(v, ujian) {
  const s = String(v == null ? '' : v).toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (ujian === 'IMNI') return ({ '1': '1', 'I': '1', 'KELAS1': '1', '2': '2', 'II': '2', 'KELAS2': '2', '3': '3', 'III': '3', 'KELAS3': '3', '4': '4', 'IV': '4', 'KELAS4': '4', '5': '5', 'V': '5', 'KELAS5': '5', '6': '6', 'VI': '6', 'KELAS6': '6' })[s] || '';
  return ({ '1': 'I', 'I': 'I', '2': 'II', 'II': 'II', '3': 'III', 'III': 'III' })[s] || '';
}
const MEK_DEF_ = {
  komposisi: { alias: { Tingkat: ['tingkat'], Kategori: ['kategori', 'kelompok', 'kategorisoal'], Kelas: ['kelas'], Ujian: ['ujian', 'tahapujian', 'imda'], Tipe: ['tipe', 'tipesoal'], Materi: ['materi', 'tahap', 'cakupan'], Jumlah: ['jumlah', 'jumlahsoal'] },
    key: r => [r.Tingkat, r.Kategori, r.Kelas, r.Ujian, r.Tipe, r.Materi].join('|') },
  bobot: { alias: { Tingkat: ['tingkat'], Kategori: ['kategori', 'kelompok', 'kategorisoal'], Ujian: ['ujian', 'tahapujian', 'imda'], Tipe: ['tipe', 'tipesoal'], Bobot: ['bobot', 'skor', 'bobotskor'] },
    key: r => [r.Tingkat, r.Kategori, r.Ujian, r.Tipe].join('|') }
};

// mode: 'tambah' (gabungkan, baris dengan kunci sama diperbarui) atau 'ganti' (hapus semua lalu isi dari file)
function api_importMekanisme(kind, rows, mode) {
  const def = MEK_DEF_[kind];
  if (!def) throw new Error('Data tidak dikenal.');
  if (!rows || !rows.length) throw new Error('File tidak berisi baris data.');
  return withLock_(() => {
    const cols = SCHEMA[kind].cols, headers = Object.keys(rows[0]), map = {}, used = {};
    cols.forEach(f => {
      let h;
      for (const a of def.alias[f]) { h = headers.find(x => !used[x] && norm_(x) === a); if (h) break; }
      if (h) { map[f] = h; used[h] = true; }
    });
    const hilang = cols.filter(f => !map[f]);
    if (hilang.length) throw new Error('Kolom tidak ditemukan di Excel: ' + hilang.join(', ') + '. Gunakan template yang disediakan.');
    const res = { ditambah: 0, dilewati: 0, galat: [] }, baru = {}, urut = [];
    rows.forEach((raw, i) => {
      const line = i + 2, o = {};
      cols.forEach(f => { o[f] = raw[map[f]]; });
      if (cols.every(f => String(o[f] == null ? '' : o[f]).trim() === '')) return;
      
      o.Tingkat = canon_(TINGKAT, o.Tingkat);
      // Kategori: pakai canonical kalau cocok, kalau tidak pakai apa adanya
      o.Kategori = canon_(KATEGORI, o.Kategori) || String(o.Kategori || '').trim();
      o.Ujian = ujianCanon_(o.Ujian);

      o.Tipe = String(o.Tipe || '').trim().toUpperCase().charAt(0);
      if (!o.Tingkat) { res.galat.push('Baris ' + line + ': Tingkat tidak dikenal.'); return; }
      
      if (!o.Kategori) { res.galat.push('Baris ' + line + ': Kategori kosong.'); return; }
      if (!o.Ujian) { res.galat.push('Baris ' + line + ': Ujian harus IMDA 1/2/3 atau IMNI.'); return; }
      const tipeDiizinkan = o.Ujian === 'IMNI' && o.Tingkat !== 'Aliyah' ? ['A', 'B', 'C', 'D'] : ['A', 'B', 'C'];
      if (!tipeDiizinkan.includes(o.Tipe)) { res.galat.push('Baris ' + line + ': tipe tidak sesuai tingkat/ujian. IMDA dan IMNI Aliyah memakai A, B, C; IMNI tingkat lainnya memakai A, B, C, D.'); return; }
      if (kind === 'komposisi') {
        o.Materi = materiCanon_(o.Materi, o.Ujian);
        const kl = String(o.Kelas == null ? '' : o.Kelas).split(/[,;\s]+/).filter(Boolean).map(kelasNum_);
        o.Kelas = kl.join(',');
        const n = Number(String(o.Jumlah).replace(',', '.'));
        if (!o.Materi) { res.galat.push('Baris ' + line + ': Materi IMDA harus I/II/III; Materi IMNI diisi nomor kelas sumber (1, 2, 3, dst.).'); return; }
        if (o.Ujian === 'IMNI' && !/^\d+$/.test(String(o.Materi))) { res.galat.push('Baris ' + line + ': untuk IMNI, Materi harus nomor kelas sumber (1, 2, 3, dst.).'); return; }
        if (o.Ujian !== 'IMNI' && !['I', 'II', 'III'].includes(o.Materi)) { res.galat.push('Baris ' + line + ': Materi IMDA harus I, II, atau III.'); return; }
        if (!o.Kelas) { res.galat.push('Baris ' + line + ': Kelas kosong.'); return; }
        if (!isFinite(n) || n < 0) { res.galat.push('Baris ' + line + ': Jumlah bukan angka.'); return; }
        o.Jumlah = n;
      } else {
        const n = Number(String(o.Bobot).replace(',', '.'));
        if (!isFinite(n) || n < 0) { res.galat.push('Baris ' + line + ': Bobot bukan angka.'); return; }
        o.Bobot = n;
      }
      const k = def.key(o);
      if (!(k in baru)) urut.push(k);
      baru[k] = o;
      res.ditambah++;
    });
    if (!urut.length) { res.data = api_all(); return res; }
    const gabung = {}, order = [];
    if (mode !== 'ganti') readAll_(kind).forEach(r => {
      if (kind === 'komposisi') r.Jumlah = Number(r.Jumlah); else r.Bobot = Number(r.Bobot);
      const k = def.key(r); if (!(k in gabung)) order.push(k); gabung[k] = r;
    });
    urut.forEach(k => { if (!(k in gabung)) order.push(k); gabung[k] = baru[k]; });
    const sh = sheet_(kind);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).clearContent();
    sh.getRange(2, 1, order.length, cols.length).setValues(order.map(k => toRow_(kind, gabung[k])));
    res.data = api_all();
    return res;
  });
}

// ====================== PENGATURAN & PENGINGAT ======================
function pengaturan_() {
  const o = Object.assign({}, PENGATURAN_AWAL);
  readAll_('pengaturan').forEach(r => { if (r.Kunci) o[r.Kunci] = r.Nilai; });
  return o;
}

function api_savePengaturan(obj) {
  return withLock_(() => {
    const sh = sheet_('pengaturan');
    const cur = pengaturan_();
    const rows = Object.keys(PENGATURAN_AWAL).map(k => [k, (obj && obj[k] != null && String(obj[k]).trim()) ? String(obj[k]) : cur[k]]);
    sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).clearContent();
    sh.getRange(2, 1, rows.length, 2).setValues(rows);
    return api_all();
  });
}

function isiPesan_(tpl, data) {
  return String(tpl).replace(/\{(\w+)\}/g, (m, k) => data[k] != null ? data[k] : m);
}

function pesanUstadz_(peng, nama, tugas) {
  const daftar = tugas.map(t => isiPesan_(peng.BarisTugas, {
    mapel: t.Mapel, kelas: t.Kelas, tingkat: t.Tingkat, ujian: t.Ujian, unit: t.Unit,
    deadline: tglId_(t.Deadline), status: t.Status || 'Belum Diisi'
  })).join('\n');
  return isiPesan_(peng.PesanWA, { ustadz: nama, daftar: daftar });
}

// Pengingat email otomatis: ustadz yang kolom Kontak-nya berisi email, tugas belum Submitted
// dan tenggatnya sudah lewat atau tinggal 3 hari. WhatsApp tidak bisa terkirim otomatis (lihat menu Pengingat WA).
function kirimPengingatEmail() {
  const peng = pengaturan_();
  const batas = Utilities.formatDate(new Date(Date.now() + 3 * 86400000), TZ_(), 'yyyy-MM-dd');
  const tugas = readAll_('tugas').filter(t => !['Submitted', 'Approved'].includes(t.Status) && t.Deadline && t.Deadline <= batas);
  let terkirim = 0;
  readAll_('ustadz').forEach(u => {
    if (!/@/.test(u.Kontak || '')) return;
    const mine = tugas.filter(t => norm_(t.Ustadz) === norm_(u.Nama));
    if (!mine.length) return;
    MailApp.sendEmail(u.Kontak.trim(), 'Pengingat penulisan soal ujian', pesanUstadz_(peng, u.Nama, mine));
    terkirim++;
  });
  return terkirim;
}

function emailTriggerAktif_() {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'kirimPengingatEmail');
}

function api_setEmailTrigger(aktif) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'kirimPengingatEmail').forEach(t => ScriptApp.deleteTrigger(t));
  if (aktif) ScriptApp.newTrigger('kirimPengingatEmail').timeBased().everyDays(1).atHour(7).create();
  return api_all();
}
