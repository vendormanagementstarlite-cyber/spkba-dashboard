/**
 * ======================================================================
 * AUTO-GENERATE SPK (WORD + PDF) DARI GOOGLE SHEET
 * Menangani 2 jenis dokumen dalam 1 spreadsheet:
 *   A. SPK TAKEOVER  -> tab "TEMPLATE TAKEOVER PDF WORD"
 *   B. SPK NEW       -> tab "TEMPLATE NEW PDF WORD"
 * ======================================================================
 * Cara pasang:
 * 1. Buka Google Sheet Anda -> Ekstensi -> Apps Script.
 * 2. Hapus SEMUA isi kode lama di Code.gs, lalu paste SELURUH isi file ini.
 * 3. Simpan (Ctrl+S).
 * 4. Jalankan sekali fungsi `setupTrigger` dari dropdown fungsi di toolbar
 *    Apps Script (hanya perlu sekali, aman dijalankan ulang kapan saja).
 * 5. Kembali ke Sheet -> refresh halaman. Menu "SPK Take Over" akan muncul.
 * ======================================================================
 */

const FIRST_DATA_ROW = 2; // baris 1 = header

// ----------------------------------------------------------------------
// RETRY HELPER — DriveApp/DocumentApp kadang melempar error sesaat
// ("Access denied: DriveApp", "Service invoked too many times", dll)
// yang sifatnya SEMENTARA (transient), terutama tepat setelah izin baru
// di-Allow atau saat Drive lagi sibuk. Fungsi ini mengulang panggilan
// beberapa kali dengan jeda sebelum benar-benar menyerah.
// ----------------------------------------------------------------------
function withRetry_(fn, attempts, label) {
  attempts = attempts || 3;
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      return fn();
    } catch (err) {
      lastErr = err;
      Logger.log('withRetry_ percobaan ' + i + '/' + attempts + ' gagal' + (label ? ' (' + label + ')' : '') + ': ' + err);
      if (i < attempts) Utilities.sleep(1000 * i); // 1s, lalu 2s, dst.
    }
  }
  throw lastErr;
}

// ----------------------------------------------------------------------
// GANTI PLACEHOLDER LEBIH CEPAT — sebelumnya, tiap nama alternatif untuk
// 1 field (mis. NO_PO bisa ditulis "NO_PO", "NO PO", atau "NOMOR PO" di
// template) memanggil body.replaceText() TERPISAH untuk masing-masing
// variasi -- padahal itu artinya 1 field bisa makan 2-3 panggilan ke
// Google Docs. Setiap panggilan replaceText() adalah satu kali
// bolak-balik jaringan ke server Google, jadi makin sedikit panggilan
// = makin cepat. Fungsi ini menggabungkan semua variasi nama jadi SATU
// pola regex (pakai "atau"/"|"), sehingga cukup SATU panggilan
// replaceText() per field, walau field itu punya banyak nama alternatif.
// Hasil akhirnya identik -- cuma jalannya lebih cepat.
// ----------------------------------------------------------------------
function replaceAliasesMerged_(body, aliases, data, escapeRegex) {
  Object.keys(aliases).forEach(function (dataKey) {
    const value = data[dataKey];
    if (value === undefined) return;
    const patterns = aliases[dataKey].map(function (name) {
      return '\\{\\{?' + escapeRegex(name) + '\\}?\\}';
    });
    const combined = patterns.length > 1 ? '(?:' + patterns.join('|') + ')' : patterns[0];
    body.replaceText(combined, String(value));
  });
}


const TAKEOVER = {
  SHEET_NAME: 'TEMPLATE TAKEOVER PDF WORD',
  TEMPLATE_ID: '16bOel7w5-Pz5jpdp7f7wqsT-yJKxCs-B3VvpOOgUtWg',
  OUTPUT_FOLDER_WORD_ID: '1rp2LUd3r7yLOCdnpRlEwSaYPUyVT1tK5',
  OUTPUT_FOLDER_PDF_ID: '12-iJvU7N14bhbjnNHNUd1zXynBo3dRwP',
  COL: {
    MITRA_LAMA: 1,   // A
    MITRA_PENGGANTI: 2, // B
    REGION: 3,       // C
    STASIUN: 4,      // D
    NO_PO: 5,        // E
    KLASIFIKASI_CANCEL: 6, // F
    HP_PO: 7,        // G
    HP_AKTUAL: 8,    // H
    HC_AKTUAL: 9,    // I
    NO_SPK: 10,      // J
    NAMA_MITRA_PT_CV: 11, // K
    DIREKSI: 12,     // L
    ALAMAT: 13,      // M
    DOKUMEN_WORD: 14, // N
    DOKUMEN_PDF: 15   // O
  },
  LAST_DATA_COL: 13, // sampai kolom ALAMAT (M)
  FILE_PREFIX: 'SPK TAKE OVER'
};

// ----------------------------------------------------------------------
// KONFIGURASI B: SPK NEW (tab "TEMPLATE NEW PDF WORD")
// ----------------------------------------------------------------------
const SPKNEW = {
  SHEET_NAME: 'TEMPLATE NEW PDF WORD',
  TEMPLATE_ID: '1eRO0r9S0KrBsTyJof98YXsNxsqol31qcG3Fprl16jEM',
  OUTPUT_FOLDER_WORD_ID: '1elu-fmZVVfjNLnXDnWKSv62FOrTE5r7g',
  OUTPUT_FOLDER_PDF_ID: '1dfQBUP0f8kqss9sOLuJhziN9jgRNey9K',
  COL: {
    MITRA_TANPA_PTCV: 1, // A
    REGION: 2,           // B
    STASIUN: 3,           // C
    HP_PENGAJUAN: 4,      // D
    NAMA_MITRA_PT_CV: 5,  // E
    DIREKSI: 6,           // F
    ALAMAT: 7,            // G
    NO_SPK: 8             // H
  },
  LAST_DATA_COL: 8, // sampai kolom NO SPK (H)
  FILE_PREFIX: 'SPK NEW'
  // Tidak ada kolom output link di sheet ini (sesuai permintaan) -
  // dokumen tetap dibuat & disimpan ke folder Drive, hanya tidak
  // ditulis balik ke sel manapun.
};

// ----------------------------------------------------------------------
// KONFIGURASI C: SPK EKSPAND (tab "TEMPLATE EKSPAND PDF WORD")
// ----------------------------------------------------------------------
const EKSPAND = {
  SHEET_NAME: 'TEMPLATE EKSPAND PDF WORD',
  TEMPLATE_ID: '1IQd7mvMGScODp8dPIjQg0-977uzjZHJt2-GCncuamRg',
  OUTPUT_FOLDER_WORD_ID: '1TP2MMsilhYhjCIXWTduKY_SY7V-vLf7T',
  OUTPUT_FOLDER_PDF_ID: '1hcj9FBGjRBstxMWVexIsHL-YmndqOU3A',
  COL: {
    NO_PO: 1,              // A
    MITRA_LAMA: 2,        // B (nama mitra tanpa PT/CV)
    REGION: 3,            // C
    STASIUN: 4,            // D
    HP_PENGAJUAN: 5,      // E
    HP_PO: 6,             // F
    HP_AKTUAL: 7,         // G
    HC_AKTUAL: 8,         // H
    NO_SPK: 9,             // I
    NAMA_MITRA_PT_CV: 10,  // J
    DIREKSI: 11,           // K
    ALAMAT: 12             // L
  },
  LAST_DATA_COL: 12,
  FILE_PREFIX: 'SPK EKSPAND'
  // Sama seperti SPK New: tidak ada kolom output link di sheet ini -
  // dokumen tetap dibuat & disimpan ke folder Drive saja.
};


// ----------------------------------------------------------------------
// KONFIGURASI D: PEMBATALAN PO (tab "TEMPLATE PEMBATALAN PO")
// ----------------------------------------------------------------------
const PEMBATALAN = {
  SHEET_NAME: 'TEMPLATE PEMBATALAN PO',
  TEMPLATE_ID: '1WfPwsohFvVdNODMdsyCzdvlmUmEGl-xg1DvaKqtZpKc',
  OUTPUT_FOLDER_WORD_ID: '1toI2Qd6lrEumeN-XxJ2LwTrZz1bIoWT6',
  OUTPUT_FOLDER_PDF_ID: '1Elsuy_ygXwZ8fAa3nRZxMZCbP4aI6Sud',
  COL: {
    NO_PO: 1,             // A
    MITRA_LAMA: 2,        // B (nama mitra tanpa PT/CV)
    REGION: 3,            // C
    STASIUN: 4,            // D
    NO_SPK: 5,             // E
    NAMA_MITRA_PT_CV: 6,   // F (Kepada, pakai PT/CV)
    DIREKSI: 7,            // G
    ALAMAT: 8              // H
  },
  LAST_DATA_COL: 8,
  FILE_PREFIX: 'PEMBATALAN PO'
  // Sama seperti SPK New & Ekspand: tidak ada kolom output link di sheet ini.
};


// ----------------------------------------------------------------------
// KONFIGURASI F: BA CANCEL ONE ON ONE (tab "BA CANCEL ONE ON ONE")
// Struktur kolomnya sama persis dengan PEMBATALAN PO.
// ----------------------------------------------------------------------
const BACANCEL = {
  SHEET_NAME: 'BA CANCEL ONE ON ONE',
  TEMPLATE_ID: '1-VLA2aqPTON_9dm75m1NQ18TQ7iE4vfryEISd8gTDcA',
  OUTPUT_FOLDER_WORD_ID: '1czLMcsJJzlh2M6at0_GyDeofEvJvwtNJ',
  OUTPUT_FOLDER_PDF_ID: '1_Q22OC85EazDOO1q71pfVuAtuKG3grnN',
  COL: {
    NO_PO: 1,             // A
    MITRA_LAMA: 2,        // B (nama mitra tanpa PT/CV)
    REGION: 3,            // C
    STASIUN: 4,            // D
    NO_SPK: 5,             // E
    NAMA_MITRA_PT_CV: 6,   // F (Kepada, pakai PT/CV)
    DIREKSI: 7,            // G
    ALAMAT: 8              // H
  },
  LAST_DATA_COL: 8,
  FILE_PREFIX: 'BA CANCEL ONE ON ONE'
  // Tidak ada kolom output link di sheet ini, sama seperti SPK New/Ekspand/Pembatalan.
};

// ----------------------------------------------------------------------
// KONFIGURASI E: BB PERCEPATAN (tab "TEMPLATE BB PERCEPATAN")
// ----------------------------------------------------------------------
// Beda dengan 4 config di atas: SATU dokumen BB Percepatan mewakili
// SATU "MITRA PEMBANGUNAN", tapi bisa merangkum BANYAK baris (banyak
// SEGMENT & STASIUN sekaligus). Makanya config ini punya GROUPED: true
// dan GROUP_BY_COL menunjuk kolom yang jadi kunci pengelompokan baris
// (kolom A / MITRA PEMBANGUNAN).
//
// Di dalam template Google Docs-nya, pada baris data tabel MITRA/SEGMENT/
// STASIUN, isi placeholder berikut:
//   Kolom MITRA    -> {{KEPADA}}
//   Kolom SEGMENT  -> {{SEGMENT_LIST}}
//   Kolom STASIUN  -> {{STASIUN_LIST}}
// SEGMENT_LIST & STASIUN_LIST otomatis berisi daftar tiap baris (satu per
// baris di dalam sel yang sama), sejajar urutannya satu sama lain.
// Placeholder lain yang juga tersedia: {{MITRA_PEMBANGUNAN}}, {{REGION}},
// {{DIREKSI}}, {{ALAMAT}}, {{MITRA_LAMA}}, {{MITRA_PENGGANTI}},
// {{NO_SPK_SURVEY_LIST}}, {{NO_SPK_PATCHING_LIST}}, {{NO_SPK_HC_LIST}},
// plus placeholder tanggal seperti config lain.
// ----------------------------------------------------------------------
const BBPERCEPATAN = {
  SHEET_NAME: 'TEMPLATE BB PERCEPATAN',
  TEMPLATE_ID: '13u7yTIVcHB8sr064pqGI9hb5bfGS1rdZphB-GdL8zkM',
  OUTPUT_FOLDER_WORD_ID: '1K7rT6UIEAfZvqDdglGpz701jme09h-4c',
  OUTPUT_FOLDER_PDF_ID: '1o7g_U-A_lnCRrjCxPd5ArwZ_X6c0vkwX',
  COL: {
    MITRA_PEMBANGUNAN: 1,  // A (kunci pengelompokan)
    REGION: 2,             // B
    SEGMENT: 3,            // C
    STASIUN: 4,            // D
    MITRA_LAMA: 5,         // E
    MITRA_PENGGANTI: 6,    // F
    NO_SPK_SURVEY: 7,      // G
    NO_SPK_PATCHING: 8,    // H
    NO_SPK_HC: 9,          // I
    NAMA_MITRA_PT_CV: 10,  // J (Kepada -> mitra pelaksana yang sebenarnya)
    DIREKSI: 11,           // K
    ALAMAT: 12             // L
  },
  LAST_DATA_COL: 12,
  FILE_PREFIX: 'BB PERCEPATAN',
  GROUPED: true,
  // Kunci grup dokumen = MITRA_PEMBANGUNAN (kolom A, isi FAMIKA/KOPINDOSAT)
  // DIGABUNG dengan NAMA_MITRA_PT_CV (kolom J, mitra pelaksana sebenarnya).
  // Ini WAJIB digabung dua-duanya supaya:
  //  - baris dengan MITRA_PEMBANGUNAN sama ("FAMIKA") tapi mitra pelaksana
  //    beda (PT J&JC vs PT WANWAN DIKDIK) TIDAK tercampur jadi 1 dokumen.
  //  - 1 mitra pelaksana yang sama (misal PT WANWAN DIKDIK) yang muncul di
  //    FAMIKA maupun KOPINDOSAT tetap dianggap 2 dokumen terpisah.
  // Lihat buildBBPercepatanGroups_().
  GROUP_BY_COL: 1 // = COL.MITRA_PEMBANGUNAN
  // Tidak ada kolom output link di sheet ini, sama seperti SPK New/Ekspand/Pembatalan.
};

// ----------------------------------------------------------------------
// VALIDASI: NO SPK TIDAK BOLEH SAMA / DOBEL
// Dipakai di generateForRow() (Takeover/New/Ekspand/Pembatalan), di
// generateBBPercepatanForGroup_() (NO SPK SURVEY/PATCHING/HC), dan saat
// user mengetik langsung di kolom NO SPK (onEditInstallable) supaya
// langsung ketahuan sebelum dokumen sempat dibuat.
// ----------------------------------------------------------------------
const NO_SPK_DUPLICATE_COLOR = '#f4cccc'; // merah muda -> tanda sel ini dobel

function normalizeNoSpk_(value) {
  return String(value === null || value === undefined ? '' : value).trim().toUpperCase();
}

// Cari baris lain (selain currentRow) pada kolom colIndex di sheet yang
// nilainya sama persis dengan value (setelah dirapikan/normalize).
// Kembalikan nomor baris pertama yang bentrok, atau null kalau aman.
// ----------------------------------------------------------------------
// CACHE KOLOM NO SPK — dibaca SEKALI per sheet per proses (bukan sekali
// per dokumen). Sebelumnya, findDuplicateNoSpkRow_ /
// findDuplicateNoSpkRowBBPercepatan_ membaca ULANG seluruh kolom NO SPK
// dari Google Sheets setiap kali dipanggil -- kalau generate massal 20
// dokumen dari tab yang sama, itu 20x baca ulang data yang ISINYA SAMA
// PERSIS. Ini salah satu penyebab generate massal terasa lambat. Cache
// ini otomatis segar lagi tiap kali ada permintaan baru (buka dashboard
// lagi / klik Generate lagi), jadi tidak pernah basi antar sesi.
// ----------------------------------------------------------------------
const _colValuesCache_ = {};
function getColumnValuesCached_(sheet, colIndex) {
  const key = sheet.getSheetId() + ':' + colIndex;
  if (!_colValuesCache_[key]) {
    const lastRow = sheet.getLastRow();
    _colValuesCache_[key] = lastRow < FIRST_DATA_ROW
      ? []
      : sheet.getRange(FIRST_DATA_ROW, colIndex, lastRow - FIRST_DATA_ROW + 1, 1).getValues().map(function (r) { return r[0]; });
  }
  return _colValuesCache_[key];
}

function findDuplicateNoSpkRow_(sheet, colIndex, currentRow, value) {
  const normalized = normalizeNoSpk_(value);
  if (!normalized) return null;

  const colValues = getColumnValuesCached_(sheet, colIndex);
  for (let i = 0; i < colValues.length; i++) {
    const r = FIRST_DATA_ROW + i;
    if (r === currentRow) continue;
    if (normalizeNoSpk_(colValues[i]) === normalized) return r;
  }
  return null;
}

// Dipanggil dari onEditInstallable supaya begitu user selesai mengetik
// NO SPK, langsung ketahuan kalau nomor itu sudah dipakai baris lain —
// sel ditandai merah muda + muncul notifikasi toast di pojok sheet.
// Kalau ternyata tidak/tidak lagi dobel, warna peringatan dibersihkan.
function warnIfDuplicateNoSpkOnEdit_(sheet, row, colIndex, label) {
  if (!colIndex) return;
  const cell = sheet.getRange(row, colIndex);
  const value = cell.getValue();
  const dupRow = findDuplicateNoSpkRow_(sheet, colIndex, row, value);
  if (dupRow) {
    cell.setBackground(NO_SPK_DUPLICATE_COLOR);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      label + ' "' + value + '" di baris ' + row + ' sudah dipakai di baris ' + dupRow +
      '. Tidak boleh sama/dobel, harap diganti.',
      '⚠️ ' + label + ' DOBEL', 8
    );
  } else if (cell.getBackground() === NO_SPK_DUPLICATE_COLOR) {
    cell.setBackground(null);
  }
}

// Versi khusus BB Percepatan dari warnIfDuplicateNoSpkOnEdit_ di atas:
// nomor SPK yang sama TIDAK diperingatkan selama baris-barisnya masih
// 1 mitra pembangunan + mitra pelaksana yang sama (satu SPK payung
// untuk beberapa lokasi). Baru muncul warning kalau bentrok dengan
// mitra pembangunan/pelaksana yang beda. Fungsi ini butuh 'config'
// (BBPERCEPATAN) supaya tahu kolom MITRA_PEMBANGUNAN & NAMA_MITRA_PT/CV.
function warnIfDuplicateNoSpkOnEditBBPercepatan_(sheet, config, row, colIndex, label) {
  if (!colIndex) return;
  const cell = sheet.getRange(row, colIndex);
  const value = cell.getValue();
  const dupRow = findDuplicateNoSpkRowBBPercepatan_(sheet, config, colIndex, row, value);
  if (dupRow) {
    cell.setBackground(NO_SPK_DUPLICATE_COLOR);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      label + ' "' + value + '" di baris ' + row + ' sudah dipakai di baris ' + dupRow +
      ' (mitra pembangunan/pelaksana berbeda). Tidak boleh sama/dobel, harap diganti.',
      '⚠️ ' + label + ' DOBEL', 8
    );
  } else if (cell.getBackground() === NO_SPK_DUPLICATE_COLOR) {
    cell.setBackground(null);
  }
}

// Khusus BB Percepatan: satu grup berisi banyak baris (NO SPK SURVEY /
// PATCHING / HC masing-masing per baris). Cek ketiga kolom itu untuk
// SEMUA baris dalam grup, dibandingkan ke SELURUH tab (bukan cuma
// sesama anggota grup). Kembalikan pesan error kalau ada yang dobel,
// atau '' kalau semua aman.
// Kunci grup BB Percepatan untuk 1 baris tertentu di sheet (dibaca
// langsung dari sel, bukan dari group object) — dipakai untuk
// menentukan apakah 2 baris yang punya NO SPK sama itu masih 1
// dokumen/mitra yang sama (boleh sama) atau beda dokumen (berarti
// dobel beneran).
function bbPercepatanGroupKeyForRow_(sheet, config, row) {
  const mitraValue = sheet.getRange(row, config.COL.MITRA_PEMBANGUNAN).getValue();
  if (!mitraValue) return '';
  const kepadaValue = sheet.getRange(row, config.COL.NAMA_MITRA_PT_CV).getValue();
  return String(mitraValue).trim() + '||' + String(kepadaValue).trim().toUpperCase();
}

// Sama seperti findDuplicateNoSpkRow_, tapi khusus BB Percepatan:
// 1 nomor SPK SURVEY/PATCHING/HC MEMANG BOLEH dipakai untuk lebih dari
// 1 lokasi/STASIUN, selama baris-baris itu masih 1 mitra pembangunan +
// mitra pelaksana yang sama (1 dokumen/SPK payung untuk beberapa
// lokasi). Baru dianggap dobel/error kalau nomor yang sama dipakai di
// baris dari mitra pembangunan/pelaksana yang BEDA.
function findDuplicateNoSpkRowBBPercepatan_(sheet, config, colIndex, currentRow, value) {
  const normalized = normalizeNoSpk_(value);
  if (!normalized) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return null;

  const currentGroupKey = bbPercepatanGroupKeyForRow_(sheet, config, currentRow);

  const colValues = getColumnValuesCached_(sheet, colIndex);
  for (let i = 0; i < colValues.length; i++) {
    const r = FIRST_DATA_ROW + i;
    if (r === currentRow) continue;
    if (normalizeNoSpk_(colValues[i]) !== normalized) continue;
    const otherGroupKey = bbPercepatanGroupKeyForRow_(sheet, config, r);
    if (otherGroupKey && otherGroupKey === currentGroupKey) continue; // 1 dokumen yang sama -> boleh sama
    return r;
  }
  return null;
}

function checkBBPercepatanNoSpkDuplicates_(sheet, config, group) {
  const cols = [
    { field: 'noSpkSurvey', col: config.COL.NO_SPK_SURVEY, label: 'NO SPK SURVEY' },
    { field: 'noSpkPatching', col: config.COL.NO_SPK_PATCHING, label: 'NO SPK PATCHING' },
    { field: 'noSpkHc', col: config.COL.NO_SPK_HC, label: 'NO SPK HC' }
  ];
  for (let i = 0; i < group.rows.length; i++) {
    const r = group.rows[i];
    for (let j = 0; j < cols.length; j++) {
      const c = cols[j];
      const value = r[c.field];
      if (!value) continue;
      const dupRow = findDuplicateNoSpkRowBBPercepatan_(sheet, config, c.col, r.rowNumber, value);
      if (dupRow) {
        sheet.getRange(r.rowNumber, c.col).setBackground(NO_SPK_DUPLICATE_COLOR);
        return c.label + ' "' + value + '" di baris ' + r.rowNumber + ' sudah dipakai di baris ' + dupRow +
          ' (mitra pembangunan/pelaksana berbeda). Ganti dulu (tidak boleh sama/dobel) sebelum dokumen bisa dibuat.';
      }
    }
  }
  return '';
}

// ----------------------------------------------------------------------
// TABEL LAMPIRAN BB PERCEPATAN — 3 kolom: MITRA | SEGMENT | STASIUN
// Isi "MITRA" = MITRA PAKAI PT/CV (Kepada), diulang di tiap baris;
// "SEGMENT" & "STASIUN" mengikuti tiap baris data di dalam grup.
//
// Cara kerja: cari tabel di dalam dokumen yang baris headernya (baris
// ke-1) berisi teks "MITRA", "SEGMENT", "STASIUN" (3 sel pertama, tidak
// case-sensitive). Baris data lama di bawah header (kalau ada, misal
// contoh/placeholder) dihapus semua, lalu diisi ulang satu baris per
// SEGMENT+STASIUN sesuai data grup ini.
//
// Syarat di template Google Docs: buat 1 tabel 3 kolom di halaman
// Lampiran, baris pertama = header "MITRA" | "SEGMENT" | "STASIUN".
// Boleh ada 1 baris kosong/contoh di bawahnya atau tidak sama sekali —
// keduanya otomatis ditangani.
// ----------------------------------------------------------------------
function fillLampiranTable_(doc, group) {
  const body = doc.getBody();
  // Kumpulkan SEMUA tabel yang header-nya cocok (MITRA + SEGMENT + STASIUN),
  // karena template BB Percepatan bisa punya beberapa halaman/section
  // (Survey, Patching, HC) yang masing-masing punya tabel Lampiran sendiri.
  // Sebelumnya cuma tabel PERTAMA yang diisi (pakai break) -> tabel di
  // halaman lain tetap kosong. Sekarang semua yang cocok diisi.
  const targets = []; // { table, col: { mitra, segment, stasiun, total } }

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
    const table = child.asTable();
    if (table.getNumRows() === 0) continue;
    const headerRow = table.getRow(0);
    const numCells = headerRow.getNumCells();
    if (numCells < 3) continue;

    // Cari kolom MITRA/SEGMENT/STASIUN di posisi manapun di baris header
    // (bukan cuma kolom 0/1/2), supaya template yang punya kolom
    // tambahan (mis. "No") di depan tetap terdeteksi.
    let mitraCol = -1, segmentCol = -1, stasiunCol = -1;
    for (let c = 0; c < numCells; c++) {
      const text = headerRow.getCell(c).getText().trim().toUpperCase();
      if (mitraCol === -1 && text.indexOf('MITRA') !== -1) mitraCol = c;
      else if (segmentCol === -1 && text.indexOf('SEGMENT') !== -1) segmentCol = c;
      else if (stasiunCol === -1 && text.indexOf('STASIUN') !== -1) stasiunCol = c;
    }
    if (mitraCol !== -1 && segmentCol !== -1 && stasiunCol !== -1) {
      targets.push({ table: table, col: { mitra: mitraCol, segment: segmentCol, stasiun: stasiunCol, total: numCells } });
    }
  }

  if (targets.length === 0) {
    // Template belum punya tabel lampiran yang cocok -> lewati, jangan
    // gagal total, tapi catat di log supaya kelihatan kenapa tabel
    // lampiran kosong. Jalankan debugLampiranTemplateBBPercepatan_()
    // dari editor Apps Script untuk lihat header tabel apa saja yang
    // ada di template saat ini.
    Logger.log('fillLampiranTable_: tidak ketemu tabel dengan header MITRA + SEGMENT + STASIUN di template "' +
      BBPERCEPATAN.TEMPLATE_ID + '". Lampiran tidak terisi. Jalankan debugLampiranTemplateBBPercepatan_() untuk detail.');
    return;
  }

  targets.forEach(function (target) {
    const targetTable = target.table;
    const col = target.col;

    // Buang semua baris data lama (selain header baris ke-0).
    while (targetTable.getNumRows() > 1) {
      targetTable.removeRow(1);
    }

    group.rows.forEach(function (r) {
      if (!r.segment && !r.stasiun) return; // lewati baris kosong
      const newRow = targetTable.appendTableRow();
      for (let c = 0; c < col.total; c++) {
        const text = c === col.mitra ? String(group.kepada || '')
          : c === col.segment ? String(r.segment || '')
          : c === col.stasiun ? String(r.stasiun || '')
          : '';
        newRow.appendTableCell(text);
      }
    });
  });

  Logger.log('fillLampiranTable_: berhasil mengisi ' + targets.length + ' tabel Lampiran.');
}

// ----------------------------------------------------------------------
// Bantuan diagnosa: jalankan fungsi ini SEKALI secara manual dari editor
// Apps Script (pilih fungsi ini di dropdown, lalu klik Jalankan/Run),
// lalu buka View > Logs (atau Executions) untuk lihat semua tabel yang
// ada di template BB Percepatan beserta isi baris headernya. Pakai ini
// untuk cari tahu kenapa tabel Lampiran tidak kedeteksi/tidak terisi --
// biasanya karena teks header di template tidak persis "MITRA",
// "SEGMENT", "STASIUN" (mis. typo, bahasa lain, atau tabelnya belum ada
// sama sekali di template).
// ----------------------------------------------------------------------
function debugLampiranTemplateBBPercepatan_() {
  const doc = DocumentApp.openById(BBPERCEPATAN.TEMPLATE_ID);
  const body = doc.getBody();
  let tableCount = 0;
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
    tableCount++;
    const table = child.asTable();
    if (table.getNumRows() === 0) {
      Logger.log('Tabel #' + tableCount + ': tidak punya baris sama sekali.');
      continue;
    }
    const headerRow = table.getRow(0);
    const headers = [];
    for (let c = 0; c < headerRow.getNumCells(); c++) {
      headers.push(headerRow.getCell(c).getText().trim());
    }
    Logger.log('Tabel #' + tableCount + ' (' + table.getNumRows() + ' baris) - header: [' + headers.join(' | ') + ']');
  }
  if (tableCount === 0) {
    Logger.log('Tidak ada tabel sama sekali di dalam template ini (ID: ' + BBPERCEPATAN.TEMPLATE_ID + ').');
  }
}

// ----------------------------------------------------------------------
// Pembungkus sementara supaya fungsi debug di atas (yang diakhiri garis
// bawah, sehingga disembunyikan dari dropdown Run oleh Apps Script) bisa
// muncul & dipilih langsung di dropdown Run. Cukup pilih
// "jalankanDebugLampiran" di dropdown lalu klik Run, kemudian buka
// Executions untuk baca hasil Logs-nya.
// ----------------------------------------------------------------------
function jalankanDebugLampiran() {
  debugLampiranTemplateBBPercepatan_();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SPK Takeover')
    .addItem('Generate baris ini (baris aktif)', 'generateActiveRow_Takeover')
    .addItem('Generate semua baris yang belum ada dokumen', 'generateAllMissing_Takeover')
    .addToUi();
  ui.createMenu('SPK New')
    .addItem('Generate baris ini (baris aktif)', 'generateActiveRow_SPKNew')
    .addItem('Generate semua baris yang belum ada dokumen', 'generateAllMissing_SPKNew')
    .addToUi();
  ui.createMenu('SPK Ekspand')
    .addItem('Generate baris ini (baris aktif)', 'generateActiveRow_Ekspand')
    .addItem('Generate semua baris yang belum ada dokumen', 'generateAllMissing_Ekspand')
    .addToUi();
  ui.createMenu('Pembatalan PO')
    .addItem('Generate baris ini (baris aktif)', 'generateActiveRow_Pembatalan')
    .addItem('Generate semua baris yang belum ada dokumen', 'generateAllMissing_Pembatalan')
    .addToUi();
  ui.createMenu('BB Percepatan')
    .addItem('Generate mitra pembangunan aktif', 'generateActiveRow_BBPercepatan')
    .addItem('Generate semua mitra yang belum ada dokumen', 'generateAllMissing_BBPercepatan')
    .addToUi();
  ui.createMenu('BA Cancel One on One')
    .addItem('Generate baris ini (baris aktif)', 'generateActiveRow_BACancel')
    .addItem('Generate semua baris yang belum ada dokumen', 'generateAllMissing_BACancel')
    .addToUi();
  ui.createMenu('Pengaturan Otomatis')
    .addItem('Aktifkan auto-generate saat input data', 'setupTrigger')
    .addItem('Nonaktifkan auto-generate saat input data', 'removeAutoGenerateTrigger')
    .addItem('Bersihkan cache dashboard', 'bersihkanCacheManual')
    .addToUi();
}

// Jalankan fungsi ini SEKALI secara manual dari editor Apps Script
// untuk memberi izin akses Drive/Docs dan memasang trigger otomatis.
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(t);
    }
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('Auto-generate aktif untuk kelima tab (Takeover, New, Ekspand, Pembatalan PO, & BB Percepatan).');
}

// ----------------------------------------------------------------------
// Hapus trigger onEdit yang sudah terpasang (mis. dari setupTrigger()
// yang pernah dijalankan sebelumnya), supaya dokumen Word/PDF TIDAK
// lagi dibuat otomatis saat sel diedit. Dokumen tetap bisa dibuat kapan
// saja lewat menu manual di atas atau lewat dashboard web.
// ----------------------------------------------------------------------
function removeAutoGenerateTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  const ui = SpreadsheetApp.getUi();
  if (removed > 0) {
    ui.alert('Auto-generate sudah dinonaktifkan. Dokumen Word/PDF tidak lagi dibuat otomatis saat sel diedit — gunakan menu generate manual atau dashboard.');
  } else {
    ui.alert('Auto-generate memang belum aktif (tidak ada trigger yang perlu dihapus).');
  }
}

// ======================================================================
// TRIGGER OTOMATIS: jalan tiap kali ada perubahan sel di sheet manapun
// CATATAN: auto-generate dokumen (Word/PDF) SUDAH DIMATIKAN di sini.
// Fungsi ini sekarang hanya menjalankan validasi NO SPK dobel (memberi
// warna merah muda di sel yang bentrok) supaya user tetap dapat
// peringatan langsung saat mengetik, tanpa dokumen ikut dibuat otomatis.
// Untuk membuat dokumen, gunakan menu generate manual di atas atau
// dashboard web.
// ======================================================================
function onEditInstallable(e) {
  try {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    if (row < FIRST_DATA_ROW) return;
    const editedCol = e.range.getColumn();

    if (sheet.getName() === TAKEOVER.SHEET_NAME) {
      if (editedCol > TAKEOVER.LAST_DATA_COL) return;
      if (editedCol === TAKEOVER.COL.NO_SPK) warnIfDuplicateNoSpkOnEdit_(sheet, row, TAKEOVER.COL.NO_SPK, 'NO SPK');
    } else if (sheet.getName() === SPKNEW.SHEET_NAME) {
      if (editedCol > SPKNEW.LAST_DATA_COL) return;
      if (editedCol === SPKNEW.COL.NO_SPK) warnIfDuplicateNoSpkOnEdit_(sheet, row, SPKNEW.COL.NO_SPK, 'NO SPK');
    } else if (sheet.getName() === EKSPAND.SHEET_NAME) {
      if (editedCol > EKSPAND.LAST_DATA_COL) return;
      if (editedCol === EKSPAND.COL.NO_SPK) warnIfDuplicateNoSpkOnEdit_(sheet, row, EKSPAND.COL.NO_SPK, 'NO SPK');
    } else if (sheet.getName() === PEMBATALAN.SHEET_NAME) {
      if (editedCol > PEMBATALAN.LAST_DATA_COL) return;
      if (editedCol === PEMBATALAN.COL.NO_SPK) warnIfDuplicateNoSpkOnEdit_(sheet, row, PEMBATALAN.COL.NO_SPK, 'NO SPK');
    } else if (sheet.getName() === BACANCEL.SHEET_NAME) {
      if (editedCol > BACANCEL.LAST_DATA_COL) return;
      if (editedCol === BACANCEL.COL.NO_SPK) warnIfDuplicateNoSpkOnEdit_(sheet, row, BACANCEL.COL.NO_SPK, 'NO SPK');
    } else if (sheet.getName() === BBPERCEPATAN.SHEET_NAME) {
      if (editedCol > BBPERCEPATAN.LAST_DATA_COL) return;
      if (editedCol === BBPERCEPATAN.COL.NO_SPK_SURVEY) warnIfDuplicateNoSpkOnEditBBPercepatan_(sheet, BBPERCEPATAN, row, BBPERCEPATAN.COL.NO_SPK_SURVEY, 'NO SPK SURVEY');
      if (editedCol === BBPERCEPATAN.COL.NO_SPK_PATCHING) warnIfDuplicateNoSpkOnEditBBPercepatan_(sheet, BBPERCEPATAN, row, BBPERCEPATAN.COL.NO_SPK_PATCHING, 'NO SPK PATCHING');
      if (editedCol === BBPERCEPATAN.COL.NO_SPK_HC) warnIfDuplicateNoSpkOnEditBBPercepatan_(sheet, BBPERCEPATAN, row, BBPERCEPATAN.COL.NO_SPK_HC, 'NO SPK HC');
    }
  } catch (err) {
    Logger.log('onEditInstallable error: ' + err);
  }
}

function logIfError(result) {
  if (result !== true && result !== false) {
    Logger.log('generateForRow error: ' + result);
  }
}

// ======================================================================
// MENU MANUAL — SPK TAKEOVER
// ======================================================================
function generateActiveRow_Takeover() {
  runGenerateActiveRow(TAKEOVER);
}
function generateAllMissing_Takeover() {
  runGenerateAllMissing(TAKEOVER);
}

// ======================================================================
// MENU MANUAL — SPK NEW
// ======================================================================
function generateActiveRow_SPKNew() {
  runGenerateActiveRow(SPKNEW);
}
function generateAllMissing_SPKNew() {
  runGenerateAllMissing(SPKNEW);
}

// ======================================================================
// MENU MANUAL — SPK EKSPAND
// ======================================================================
function generateActiveRow_Ekspand() {
  runGenerateActiveRow(EKSPAND);
}
function generateAllMissing_Ekspand() {
  runGenerateAllMissing(EKSPAND);
}

// ======================================================================
// MENU MANUAL — PEMBATALAN PO
// ======================================================================
function generateActiveRow_Pembatalan() {
  runGenerateActiveRow(PEMBATALAN);
}
function generateAllMissing_Pembatalan() {
  runGenerateAllMissing(PEMBATALAN);
}

// ======================================================================
// MENU MANUAL — BA CANCEL ONE ON ONE
// ======================================================================
function generateActiveRow_BACancel() {
  runGenerateActiveRow(BACANCEL);
}
function generateAllMissing_BACancel() {
  runGenerateAllMissing(BACANCEL);
}

// ======================================================================
// MENU MANUAL — BB PERCEPATAN
// Beda dari 4 tipe lain: "baris aktif" di sini berarti "grup mitra
// pembangunan dari baris aktif" — semua baris dengan MITRA PEMBANGUNAN
// yang sama akan digabung jadi SATU dokumen.
// ======================================================================
function generateActiveRow_BBPercepatan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BBPERCEPATAN.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Tab "' + BBPERCEPATAN.SHEET_NAME + '" tidak ditemukan.');
    return;
  }
  const activeSheet = SpreadsheetApp.getActiveSheet();
  let row;
  if (activeSheet.getSheetId() === sheet.getSheetId()) {
    row = sheet.getActiveCell().getRow();
  } else {
    SpreadsheetApp.getUi().alert('Pindah dulu ke tab "' + BBPERCEPATAN.SHEET_NAME + '", klik baris datanya, baru jalankan menu ini lagi.');
    return;
  }
  if (row < FIRST_DATA_ROW) {
    SpreadsheetApp.getUi().alert('Pilih salah satu baris data (bukan header).');
    return;
  }
  const mitraValue = sheet.getRange(row, BBPERCEPATAN.COL.MITRA_PEMBANGUNAN).getValue();
  if (!mitraValue) {
    SpreadsheetApp.getUi().alert('Baris ' + row + ' belum punya nilai MITRA PEMBANGUNAN.');
    return;
  }
  const kepadaValue = sheet.getRange(row, BBPERCEPATAN.COL.NAMA_MITRA_PT_CV).getValue();
  const group = findBBPercepatanGroupByRow_(sheet, BBPERCEPATAN, row);
  if (!group) {
    SpreadsheetApp.getUi().alert('Baris ' + row + ' tidak ditemukan dalam daftar dokumen BB Percepatan.');
    return;
  }
  const label = mitraValue + (kepadaValue ? ' - ' + kepadaValue : '') +
    (group.firstNoSpk ? ' (SPK ' + group.firstNoSpk + ')' : '');
  const ok = generateBBPercepatanForGroup_(sheet, BBPERCEPATAN, group.key);
  if (ok === true) {
    invalidateDocsCache_();
    SpreadsheetApp.getUi().alert('Dokumen berhasil dibuat untuk mitra pembangunan "' + label + '".');
  } else if (ok === false) {
    SpreadsheetApp.getUi().alert('Data untuk mitra pembangunan "' + label + '" belum lengkap, dokumen belum dibuat.');
  } else {
    SpreadsheetApp.getUi().alert('Gagal membuat dokumen untuk mitra pembangunan "' + label + '":\n\n' + ok);
  }
}

function generateAllMissing_BBPercepatan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BBPERCEPATAN.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Tab "' + BBPERCEPATAN.SHEET_NAME + '" tidak ditemukan.');
    return;
  }
  const groups = buildBBPercepatanGroups_(sheet, BBPERCEPATAN);

  let wordMap = {};
  try {
    wordMap = buildFileMap_(DriveApp.getFolderById(BBPERCEPATAN.OUTPUT_FOLDER_WORD_ID));
  } catch (err) { /* biarkan kosong, semua akan dianggap belum dibuat */ }

  let count = 0;
  groups.forEach(function (group) {
    if (!group.isComplete) return;
    if (lookupBBPercepatanFiles_(group, wordMap, {}).wordFile) return; // sudah ada -> lewati
    const ok = generateBBPercepatanForGroup_(sheet, BBPERCEPATAN, group.key);
    if (ok === true) count++;
  });

  if (count > 0) invalidateDocsCache_();
  SpreadsheetApp.getUi().alert(count + ' dokumen BB Percepatan berhasil dibuat.');
}

// ======================================================================
// Implementasi bersama (dipanggil oleh fungsi menu di atas)
// ======================================================================
function runGenerateActiveRow(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Tab "' + config.SHEET_NAME + '" tidak ditemukan.');
    return;
  }
  // Ambil baris aktif HANYA kalau memang sedang berada di tab yang benar,
  // kalau tidak, pakai baris terakhir yang ada data sebagai fallback aman.
  const activeSheet = SpreadsheetApp.getActiveSheet();
  let row;
  if (activeSheet.getSheetId() === sheet.getSheetId()) {
    row = sheet.getActiveCell().getRow();
  } else {
    SpreadsheetApp.getUi().alert('Pindah dulu ke tab "' + config.SHEET_NAME + '", klik baris datanya, baru jalankan menu ini lagi.');
    return;
  }
  if (row < FIRST_DATA_ROW) {
    SpreadsheetApp.getUi().alert('Pilih salah satu baris data (bukan header).');
    return;
  }
  const ok = generateForRow(sheet, row, config);
  if (ok === true) {
    invalidateDocsCache_();
    SpreadsheetApp.getUi().alert('Dokumen berhasil dibuat untuk baris ' + row);
  } else if (ok === false) {
    SpreadsheetApp.getUi().alert('Data di baris ' + row + ' belum lengkap, dokumen belum dibuat.');
  } else {
    // ok berisi pesan error (string)
    SpreadsheetApp.getUi().alert('Gagal membuat dokumen untuk baris ' + row + ':\n\n' + ok);
  }
}

function runGenerateAllMissing(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Tab "' + config.SHEET_NAME + '" tidak ditemukan.');
    return;
  }
  const lastRow = sheet.getLastRow();
  let count = 0;
  for (let row = FIRST_DATA_ROW; row <= lastRow; row++) {
    let alreadyDone = false;
    if (config.COL.DOKUMEN_WORD) {
      alreadyDone = !!sheet.getRange(row, config.COL.DOKUMEN_WORD).getValue();
    }
    if (!alreadyDone) {
      const ok = generateForRow(sheet, row, config);
      if (ok === true) count++;
    }
  }
  if (count > 0) invalidateDocsCache_();
  SpreadsheetApp.getUi().alert(count + ' dokumen berhasil dibuat.');
}

// ======================================================================
// INTI: baca 1 baris data -> isi template -> simpan Word & PDF -> (opsional) tulis link
// ======================================================================
function generateForRow(sheet, row, config) {
  const values = sheet.getRange(row, 1, 1, config.LAST_DATA_COL).getValues()[0];
  const COL = config.COL;
  const get = function (colIndex) {
    return colIndex ? values[colIndex - 1] : '';
  };

  const data = {
    MITRA_LAMA: get(COL.MITRA_LAMA) || get(COL.MITRA_TANPA_PTCV),
    REGION: get(COL.REGION),
    STASIUN: get(COL.STASIUN),
    NO_PO: get(COL.NO_PO),
    HP_PO: get(COL.HP_PO),
    HP_AKTUAL: get(COL.HP_AKTUAL),
    HC_AKTUAL: get(COL.HC_AKTUAL),
    HP_PENGAJUAN: get(COL.HP_PENGAJUAN),
    NO_SPK: get(COL.NO_SPK),
    KEPADA: get(COL.NAMA_MITRA_PT_CV),
    DIREKSI: get(COL.DIREKSI),
    ALAMAT: get(COL.ALAMAT)
  };

  // Field wajib minimal (yang selalu ada di kedua template)
  const required = [data.NO_SPK, data.KEPADA, data.STASIUN, data.DIREKSI, data.ALAMAT];
  if (required.some(function (v) { return v === '' || v === null || v === undefined; })) {
    return false; // belum lengkap, skip dulu
  }

  // Validasi: NO SPK tidak boleh sama/dobel dengan baris lain di tab ini.
  if (COL.NO_SPK) {
    const dupRow = findDuplicateNoSpkRow_(sheet, COL.NO_SPK, row, data.NO_SPK);
    if (dupRow) {
      sheet.getRange(row, COL.NO_SPK).setBackground(NO_SPK_DUPLICATE_COLOR);
      return 'NO SPK "' + data.NO_SPK + '" di baris ' + row + ' sudah dipakai di baris ' + dupRow +
        '. Ganti dulu dengan nomor SPK yang berbeda (tidak boleh sama/dobel) sebelum dokumen bisa dibuat.';
    }
    // Bersihkan warna peringatan kalau sebelumnya pernah ditandai lalu sudah diperbaiki.
    if (sheet.getRange(row, COL.NO_SPK).getBackground() === NO_SPK_DUPLICATE_COLOR) {
      sheet.getRange(row, COL.NO_SPK).setBackground(null);
    }
  }

  const wordFolder = withRetry_(function () { return DriveApp.getFolderById(config.OUTPUT_FOLDER_WORD_ID); }, 3, 'getFolderById WORD');
  const pdfFolder = withRetry_(function () { return DriveApp.getFolderById(config.OUTPUT_FOLDER_PDF_ID); }, 3, 'getFolderById PDF');

  // Ambil angka/teks sebelum tanda "/" pertama di NO SPK
  const noSpkStr = String(data.NO_SPK);
  const noSpkAwal = noSpkStr.indexOf('/') !== -1
    ? noSpkStr.substring(0, noSpkStr.indexOf('/'))
    : noSpkStr;

  const fileName = config.FILE_PREFIX + '_' + noSpkAwal + '_' + data.KEPADA + '_' + data.STASIUN;

  // 1. Duplikat template Google Docs ke folder Word
  const templateFile = withRetry_(function () { return DriveApp.getFileById(config.TEMPLATE_ID); }, 3, 'getFileById TEMPLATE');
  const newDocFile = withRetry_(function () { return templateFile.makeCopy(fileName, wordFolder); }, 3, 'makeCopy');

  let doc, body;
  try {
    doc = DocumentApp.openById(newDocFile.getId());
    body = doc.getBody();
  } catch (err) {
    // Gagal dibuka sebagai Google Docs -> kemungkinan besar template masih
    // format Word (.docx) asli, bukan native Google Docs. Hapus file
    // setengah jadi supaya tidak menumpuk sampah di folder Output.
    newDocFile.setTrashed(true);
    return 'Template "' + config.TEMPLATE_ID + '" tidak bisa dibuka sebagai Google Docs. ' +
      'Kemungkinan besar file template masih format Word (.docx) asli. ' +
      'Buka template-nya, lalu File > Simpan sebagai Google Dokumen, dan ganti TEMPLATE_ID dengan ID hasil konversi itu. ' +
      '(Pesan asli: ' + err + ')';
  }

  // 2. Ganti semua kemungkinan placeholder dengan data baris ini.
  //    ALIASES di bawah mencakup semua variasi nama yang pernah dipakai
  //    di kedua template (garis bawah / spasi, kurung tunggal / ganda),
  //    supaya cocok otomatis tanpa perlu tahu persis format tiap template.
  const ALIASES = {
    NO_SPK: ['NO_SPK', 'NO SPK'],
    KEPADA: ['KEPADA', 'NAMA MITRA PAKAI PT/CV'],
    NO_PO: ['NO_PO', 'NO PO', 'NOMOR PO'],
    STASIUN: ['STASIUN'],
    REGION: ['REGION'],
    DIREKSI: ['DIREKSI'],
    ALAMAT: ['ALAMAT'],
    MITRA_LAMA: ['MITRA_LAMA', 'MITRA LAMA', 'MITRA TANPA PT/CV'],
    HP_PO: ['HP_PO', 'HP PO'],
    HP_AKTUAL: ['HP_AKTUAL', 'HP AKTUAL', 'HP ACTUAL'],
    HC_AKTUAL: ['HC_AKTUAL', 'HC AKTUAL', 'HC ACTUAL'],
    HP_PENGAJUAN: ['HP_PENGAJUAN', 'HP PENGAJUAN']
  };

  const escapeRegex = function (s) {
    return s.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&');
  };

  replaceAliasesMerged_(body, ALIASES, data, escapeRegex);

  // 2b. Isi placeholder tanggal hari ini (Bahasa Indonesia) — dihitung
  //     saat dokumen dibuat, bukan dari data sheet. Aman untuk semua
  //     template; kalau template tidak punya placeholder ini, dilewati.
  const HARI_ID = ['Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu', 'Minggu'];
  const BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const dateAliases = {
    HARI_SEKARANG: ['HARI SEKARANG'],
    TANGGAL_SEKARANG: ['TANGGAL SEKARANG'],
    BULAN_SEKARANG: ['BULAN SEKARANG'],
    TAHUN_SEKARANG: ['TAHUN SEKARANG']
  };
  const dateValues = {
    HARI_SEKARANG: HARI_ID[Number(Utilities.formatDate(now, tz, 'u')) - 1],
    TANGGAL_SEKARANG: Utilities.formatDate(now, tz, 'd'),
    BULAN_SEKARANG: BULAN_ID[Number(Utilities.formatDate(now, tz, 'M')) - 1],
    TAHUN_SEKARANG: Utilities.formatDate(now, tz, 'yyyy')
  };
  Object.keys(dateAliases).forEach(function (key) {
    dateAliases[key].forEach(function (name) {
      const pattern = '\\{\\{?' + escapeRegex(name) + '\\}?\\}';
      body.replaceText(pattern, String(dateValues[key]));
    });
  });

  // Rekatkan "TANGGAL BULAN TAHUN" (mis. "2 September 2026") pakai spasi
  // non-breaking, supaya Google Docs tidak pernah memotongnya jadi 2
  // baris gara-gara lebar paragraf tanda tangan yang sempit.
  const dateGluedFrom_ = dateValues.TANGGAL_SEKARANG + ' ' + dateValues.BULAN_SEKARANG + ' ' + dateValues.TAHUN_SEKARANG;
  const dateGluedTo_ = dateValues.TANGGAL_SEKARANG + '\u00A0' + dateValues.BULAN_SEKARANG + '\u00A0' + dateValues.TAHUN_SEKARANG;
  body.replaceText(escapeRegex(dateGluedFrom_), dateGluedTo_);

  doc.saveAndClose();

  // 3. Export ke PDF, simpan di folder PDF
  const pdfFile = withRetry_(function () {
    const pdfBlob = newDocFile.getAs(MimeType.PDF).setName(fileName + '.pdf');
    return pdfFolder.createFile(pdfBlob);
  }, 3, 'export PDF');

  // 4. Set akses "siapa saja yang punya link bisa lihat" (OPSIONAL).
  //    Di beberapa organisasi Google Workspace, admin MEMBLOKIR sharing
  //    "siapa saja yang punya link" ke luar domain lewat kebijakan di
  //    Admin Console (Apps > Google Workspace > Drive dan Dokumen >
  //    Setelan berbagi). Kalau itu terjadi, setSharing() SELALU
  //    melempar "Akses ditolak: DriveApp" walau dokumennya SENDIRI
  //    (Word & PDF) sudah berhasil dibuat sempurna. Sebelumnya, error
  //    ini menghentikan SELURUH proses generateForRow di tengah jalan
  //    -- link hyperlink tidak sempat ditulis ke sheet dan cache tidak
  //    sempat dibersihkan, padahal dokumennya sudah ada di Drive
  //    (berisiko dokumen dobel kalau baris itu di-generate ulang).
  //    Makanya langkah ini SEKARANG dibungkus try/catch TERSENDIRI:
  //    kalau setSharing gagal, cukup dicatat di Log lalu dilewati --
  //    dokumen tetap dianggap berhasil dibuat. File itu nanti tetap
  //    bisa diakses oleh siapa pun yang sudah punya akses ke folder/
  //    Shared Drive tempat file itu disimpan, hanya saja tidak otomatis
  //    "Anyone with link".
  try {
    withRetry_(function () {
      newDocFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }, 3, 'setSharing');
  } catch (sharingErr) {
    Logger.log('generateForRow: setSharing gagal (kemungkinan dibatasi kebijakan Google Workspace), dilewati: ' + sharingErr);
  }

  // 5. Tulis link ke sheet HANYA kalau tab ini punya kolom output (Takeover).
  //    Tab SPK New sengaja tidak punya kolom output -> dilewati otomatis.
  if (COL.DOKUMEN_WORD) {
    sheet.getRange(row, COL.DOKUMEN_WORD).setFormula(
      '=HYPERLINK("' + newDocFile.getUrl() + '", "Buka Word")'
    );
  }
  if (COL.DOKUMEN_PDF) {
    sheet.getRange(row, COL.DOKUMEN_PDF).setFormula(
      '=HYPERLINK("' + pdfFile.getUrl() + '", "Buka PDF")'
    );
  }

  // ---- OTOMATIS BERSIHKAN CACHE DASHBOARD ----
  // Dipanggil di sini (bukan cuma di fungsi menu/dashboard) supaya cache
  // SELALU terhapus begitu dokumen berhasil dibuat, lewat jalur apa pun
  // (menu Sheet, dashboard web, atau dijalankan manual dari editor Apps
  // Script). Tidak perlu lagi klik "Bersihkan cache dashboard" manual.
  invalidateDocsCache_();

  return true;
}

// ======================================================================
// INTI BB PERCEPATAN: kelompokkan baris per MITRA PEMBANGUNAN, lalu buat
// SATU dokumen per grup yang merangkum semua SEGMENT & STASIUN di
// dalamnya. Terpisah dari generateForRow() karena satu dokumen di sini
// mewakili banyak baris, bukan satu baris seperti config lainnya.
// ======================================================================

// Kumpulkan seluruh baris tab BB Percepatan menjadi DOKUMEN-DOKUMEN (grup).
//
// Aturan pengelompokan (2 tahap):
//  1. Baris dikelompokkan dulu per MITRA: MITRA PEMBANGUNAN (kolom A) +
//     mitra pelaksana / NAMA MITRA PAKAI PT/CV (kolom J).
//  2. Di dalam 1 mitra, baris DIPECAH lagi jadi dokumen terpisah berdasarkan
//     NOMOR SPK: baris-baris hanya digabung jadi 1 dokumen kalau mereka
//     berbagi nomor SPK yang sama (SURVEY / PATCHING / HC) -- itulah kasus
//     "1 SPK payung untuk beberapa STASIUN". Baris dengan nomor SPK yang
//     sama sekali beda dianggap PENGAJUAN BARU -> dokumen sendiri.
//
// Sebelumnya cuma tahap 1, sehingga baris baru untuk mitra yang sama
// (dengan nomor SPK baru) ikut menyatu dengan dokumen yang sudah dibuat.
// Baris yang belum punya nomor SPK sama sekali dikumpulkan jadi 1 dokumen
// "TANPA NO SPK" per mitra (sampai nomornya diisi).
//
// Data KEPADA/DIREKSI/ALAMAT diambil dari baris PERTAMA pada tiap dokumen.
function buildBBPercepatanGroups_(sheet, config) {
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return [];

  const values = sheet.getRange(
    FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, config.LAST_DATA_COL
  ).getValues();

  const COL = config.COL;
  const get = function (row, colIndex) { return colIndex ? row[colIndex - 1] : ''; };

  // ---- Tahap 1: kelompokkan per mitra ----
  const rowsByMitra = {};
  const mitraOrder = [];

  values.forEach(function (row, idx) {
    const mitraPembangunan = get(row, COL.MITRA_PEMBANGUNAN);
    if (!mitraPembangunan) return; // baris kosong / bukan data -> lewati
    const kepadaRow = get(row, COL.NAMA_MITRA_PT_CV);
    const mitraKey = String(mitraPembangunan).trim() + '||' + String(kepadaRow).trim().toUpperCase();
    if (!rowsByMitra[mitraKey]) {
      rowsByMitra[mitraKey] = [];
      mitraOrder.push(mitraKey);
    }
    rowsByMitra[mitraKey].push({
      rowNumber: FIRST_DATA_ROW + idx,
      mitraPembangunan: mitraPembangunan,
      region: get(row, COL.REGION),
      kepada: kepadaRow,
      direksi: get(row, COL.DIREKSI),
      alamat: get(row, COL.ALAMAT),
      mitraLama: get(row, COL.MITRA_LAMA),
      mitraPengganti: get(row, COL.MITRA_PENGGANTI),
      segment: get(row, COL.SEGMENT),
      stasiun: get(row, COL.STASIUN),
      noSpkSurvey: get(row, COL.NO_SPK_SURVEY),
      noSpkPatching: get(row, COL.NO_SPK_PATCHING),
      noSpkHc: get(row, COL.NO_SPK_HC)
    });
  });

  // ---- Tahap 2: pecah tiap mitra per keluarga nomor SPK ----
  const groups = [];
  mitraOrder.forEach(function (mitraKey) {
    clusterRowsBySpk_(rowsByMitra[mitraKey]).forEach(function (cluster, clusterIdx) {
      const first = cluster.rows[0];
      const g = {
        key: mitraKey + '||' + normalizeNoSpk_(cluster.firstNoSpk),
        mitraKey: mitraKey,
        firstNoSpk: cluster.firstNoSpk,
        isFirstOfMitra: clusterIdx === 0, // dipakai utk mengenali file lama (tanpa suffix nomor SPK)
        mitraPembangunan: first.mitraPembangunan,
        region: first.region,
        kepada: first.kepada,
        direksi: first.direksi,
        alamat: first.alamat,
        mitraLama: first.mitraLama,
        mitraPengganti: first.mitraPengganti,
        rows: [],
        rowNumbers: []
      };
      cluster.rows.forEach(function (r) {
        g.rows.push({
          segment: r.segment,
          stasiun: r.stasiun,
          noSpkSurvey: r.noSpkSurvey,
          noSpkPatching: r.noSpkPatching,
          noSpkHc: r.noSpkHc,
          rowNumber: r.rowNumber
        });
        g.rowNumbers.push(r.rowNumber);
      });
      const hasSegmentStasiun = g.rows.some(function (r) { return r.segment && r.stasiun; });
      g.isComplete = !!(g.kepada && g.direksi && g.alamat && hasSegmentStasiun);
      groups.push(g);
    });
  });

  return groups;
}

// Pecah daftar baris (yang sudah 1 mitra) jadi kelompok-kelompok yang
// terhubung lewat nomor SPK yang SAMA (union-find). Dua baris masuk 1
// kelompok kalau salah satu nomor SURVEY/PATCHING/HC-nya sama persis
// (setelah dirapikan). Baris tanpa nomor SPK sama sekali digabung jadi 1
// kelompok tersendiri. Urutan kelompok mengikuti urutan baris di sheet.
function clusterRowsBySpk_(rows) {
  const parent = rows.map(function (_, i) { return i; });
  const find = function (i) {
    while (parent[i] !== i) { parent[i] = parent[parent[i]]; i = parent[i]; }
    return i;
  };
  const union = function (a, b) {
    const ra = find(a), rb = find(b);
    if (ra !== rb) parent[Math.max(ra, rb)] = Math.min(ra, rb);
  };

  const ownerByNumber = {};
  let firstRowWithoutNumber = -1;
  rows.forEach(function (r, i) {
    let hasNumber = false;
    [r.noSpkSurvey, r.noSpkPatching, r.noSpkHc].forEach(function (v) {
      const n = normalizeNoSpk_(v);
      if (!n) return;
      hasNumber = true;
      if (ownerByNumber[n] === undefined) ownerByNumber[n] = i;
      else union(i, ownerByNumber[n]);
    });
    if (!hasNumber) {
      if (firstRowWithoutNumber === -1) firstRowWithoutNumber = i;
      else union(i, firstRowWithoutNumber);
    }
  });

  const byRoot = {};
  const order = [];
  rows.forEach(function (r, i) {
    const root = find(i);
    if (!byRoot[root]) { byRoot[root] = []; order.push(root); }
    byRoot[root].push(r);
  });

  return order.map(function (root) {
    const list = byRoot[root];
    let firstNoSpk = '';
    for (let i = 0; i < list.length && !firstNoSpk; i++) {
      const cands = [list[i].noSpkSurvey, list[i].noSpkPatching, list[i].noSpkHc]
        .map(function (v) { return String(v === null || v === undefined ? '' : v).trim(); })
        .filter(function (v) { return v !== ''; });
      if (cands.length) firstNoSpk = cands[0];
    }
    return { rows: list, firstNoSpk: firstNoSpk };
  });
}

// Nama file HARUS SAMA PERSIS antara generateBBPercepatanForGroup_() dan
// getDocumentsForBBPercepatan_() supaya dashboard bisa mendeteksi dokumen
// yang sudah dibuat. Sekarang ditambah NOMOR SPK di belakang nama, supaya
// 2 dokumen untuk mitra yang sama (nomor SPK beda) tidak saling menimpa/
// tertukar.
function buildBBPercepatanFileName_(group) {
  const spkPart = group.firstNoSpk
    ? String(group.firstNoSpk).replace(/[\/\\]+/g, '-')
    : 'TANPA NO SPK';
  return buildBBPercepatanLegacyFileName_(group) + '_' + spkPart;
}

// Format nama file LAMA (sebelum ada suffix nomor SPK) -- hanya dipakai
// untuk mengenali dokumen yang sudah terlanjur dibuat dengan format lama.
function buildBBPercepatanLegacyFileName_(group) {
  return BBPERCEPATAN.FILE_PREFIX + '_' + group.mitraPembangunan + '_' + group.kepada;
}

// Cari file Word/PDF milik 1 grup. Kalau file dengan nama baru belum ada,
// dan grup ini adalah dokumen PERTAMA untuk mitra tsb, coba nama format
// lama: dokumen lama itu dibuat sebelum ada baris-baris "baru", jadi paling
// mungkin milik dokumen pertama. Dokumen berikutnya (nomor SPK baru) tidak
// pernah dicocokkan ke nama lama -> otomatis tampil "Belum Dibuat".
function lookupBBPercepatanFiles_(group, wordMap, pdfMap) {
  const fileName = buildBBPercepatanFileName_(group);
  let wordFile = wordMap[fileName] || null;
  let pdfFile = pdfMap[fileName + '.pdf'] || null;
  if (!wordFile && !pdfFile && group.isFirstOfMitra) {
    const legacy = buildBBPercepatanLegacyFileName_(group);
    wordFile = wordMap[legacy] || null;
    pdfFile = pdfMap[legacy + '.pdf'] || null;
  }
  return { fileName: fileName, wordFile: wordFile, pdfFile: pdfFile };
}

// groupKeyValue = "MITRA_PEMBANGUNAN||NAMA_MITRA_PT_CV||NO_SPK_PERTAMA"
// (lihat buildBBPercepatanGroups_). Dipakai di semua pemanggil (menu,
// dashboard) supaya dokumen yang tepat yang ditemukan.
function findBBPercepatanGroup_(sheet, config, groupKeyValue) {
  const groups = buildBBPercepatanGroups_(sheet, config);
  const key = String(groupKeyValue);
  for (let i = 0; i < groups.length; i++) {
    if (groups[i].key === key) return groups[i];
  }
  return null;
}

// Cari grup/dokumen yang memuat baris tertentu di sheet.
function findBBPercepatanGroupByRow_(sheet, config, row) {
  const groups = buildBBPercepatanGroups_(sheet, config);
  for (let i = 0; i < groups.length; i++) {
    if (groups[i].rowNumbers.indexOf(row) !== -1) return groups[i];
  }
  return null;
}

function generateBBPercepatanForGroup_(sheet, config, groupKeyValue) {
  const group = findBBPercepatanGroup_(sheet, config, groupKeyValue);
  if (!group) return false;
  if (!group.isComplete) return false; // KEPADA/DIREKSI/ALAMAT/SEGMENT/STASIUN belum lengkap

  // Validasi: NO SPK SURVEY/PATCHING/HC tidak boleh dobel dengan baris lain.
  const noSpkDupError = checkBBPercepatanNoSpkDuplicates_(sheet, config, group);
  if (noSpkDupError) return noSpkDupError;

  const wordFolder = withRetry_(function () { return DriveApp.getFolderById(config.OUTPUT_FOLDER_WORD_ID); }, 3, 'getFolderById WORD (BB Percepatan)');
  const pdfFolder = withRetry_(function () { return DriveApp.getFolderById(config.OUTPUT_FOLDER_PDF_ID); }, 3, 'getFolderById PDF (BB Percepatan)');

  const fileName = buildBBPercepatanFileName_(group);

  const templateFile = withRetry_(function () { return DriveApp.getFileById(config.TEMPLATE_ID); }, 3, 'getFileById TEMPLATE (BB Percepatan)');
  const newDocFile = withRetry_(function () { return templateFile.makeCopy(fileName, wordFolder); }, 3, 'makeCopy (BB Percepatan)');

  let doc, body;
  try {
    doc = DocumentApp.openById(newDocFile.getId());
    body = doc.getBody();
  } catch (err) {
    newDocFile.setTrashed(true);
    return 'Template "' + config.TEMPLATE_ID + '" tidak bisa dibuka sebagai Google Docs. ' +
      'Kemungkinan besar file template masih format Word (.docx) asli. ' +
      'Buka template-nya, lalu File > Simpan sebagai Google Dokumen, dan ganti TEMPLATE_ID dengan ID hasil konversi itu. ' +
      '(Pesan asli: ' + err + ')';
  }

  // Daftar SEGMENT & STASIUN — satu baris teks per pasangan, sejajar urutannya.
  const segmentList = group.rows.map(function (r) { return r.segment; }).join('\n');
  const stasiunList = group.rows.map(function (r) { return r.stasiun; }).join('\n');

  // NO SPK SURVEY/PATCHING/HC per lokasi: kalau untuk 1 lokasi nomor
  // PATCHING atau HC-nya kebetulan SAMA dengan SURVEY (1 SPK payung
  // untuk beberapa jenis pekerjaan), jangan ditulis ulang di baris
  // placeholder yang lain. Juga kalau nomor yang SAMA dipakai di
  // beberapa lokasi sekaligus (1 SPK payung untuk beberapa STASIUN,
  // seperti kasus baris 3 & 4 di sheet), jangan sampai nomor itu
  // ditulis berulang-ulang di dokumen -> field "No." di Word/PDF cukup
  // muncul sekali per nomor SPK.
  const noSpkSurveyArr = [];
  const noSpkPatchingArr = [];
  const noSpkHcArr = [];
  const seenSurveyAcrossRows_ = {};
  const seenPatchingAcrossRows_ = {};
  const seenHcAcrossRows_ = {};
  group.rows.forEach(function (r) {
    const survey = String(r.noSpkSurvey || '').trim();
    const patchingRaw = String(r.noSpkPatching || '').trim();
    const hcRaw = String(r.noSpkHc || '').trim();

    const usedInRow = {};
    if (survey) usedInRow[normalizeNoSpk_(survey)] = true;

    const patching = (patchingRaw && !usedInRow[normalizeNoSpk_(patchingRaw)]) ? patchingRaw : '';
    if (patching) usedInRow[normalizeNoSpk_(patching)] = true;

    const hc = (hcRaw && !usedInRow[normalizeNoSpk_(hcRaw)]) ? hcRaw : '';

    if (survey) {
      const n = normalizeNoSpk_(survey);
      if (!seenSurveyAcrossRows_[n]) { seenSurveyAcrossRows_[n] = true; noSpkSurveyArr.push(survey); }
    }
    if (patching) {
      const n = normalizeNoSpk_(patching);
      if (!seenPatchingAcrossRows_[n]) { seenPatchingAcrossRows_[n] = true; noSpkPatchingArr.push(patching); }
    }
    if (hc) {
      const n = normalizeNoSpk_(hc);
      if (!seenHcAcrossRows_[n]) { seenHcAcrossRows_[n] = true; noSpkHcArr.push(hc); }
    }
  });
  const noSpkSurveyList = noSpkSurveyArr.join('\n');
  const noSpkPatchingList = noSpkPatchingArr.join('\n');
  const noSpkHcList = noSpkHcArr.join('\n');

  const data = {
    KEPADA: group.kepada,
    MITRA_PEMBANGUNAN: group.mitraPembangunan,
    REGION: group.region,
    DIREKSI: group.direksi,
    ALAMAT: group.alamat,
    MITRA_LAMA: group.mitraLama,
    MITRA_PENGGANTI: group.mitraPengganti,
    SEGMENT_LIST: segmentList,
    STASIUN_LIST: stasiunList,
    NO_SPK_SURVEY_LIST: noSpkSurveyList,
    NO_SPK_PATCHING_LIST: noSpkPatchingList,
    NO_SPK_HC_LIST: noSpkHcList
  };

  const ALIASES = {
    KEPADA: ['KEPADA', 'MITRA', 'NAMA MITRA PAKAI PT/CV'],
    MITRA_PEMBANGUNAN: ['MITRA_PEMBANGUNAN', 'MITRA PEMBANGUNAN'],
    REGION: ['REGION', 'REGIONAL'],
    DIREKSI: ['DIREKSI'],
    ALAMAT: ['ALAMAT'],
    MITRA_LAMA: ['MITRA_LAMA', 'MITRA LAMA'],
    MITRA_PENGGANTI: ['MITRA_PENGGANTI', 'MITRA PENGGANTI'],
    SEGMENT_LIST: ['SEGMENT_LIST', 'SEGMENT LIST', 'LIST SEGMENT', 'SEGMENT'],
    STASIUN_LIST: ['STASIUN_LIST', 'STASIUN LIST', 'LIST STASIUN', 'STASIUN'],
    NO_SPK_SURVEY_LIST: ['NO_SPK_SURVEY_LIST', 'NO SPK SURVEY'],
    NO_SPK_PATCHING_LIST: ['NO_SPK_PATCHING_LIST', 'NO SPK PATCHING'],
    NO_SPK_HC_LIST: ['NO_SPK_HC_LIST', 'NO SPK HC']
  };

  const escapeRegex = function (s) {
    return s.replace(/[.*+?^\${}()|[\]\\\/]/g, '\\\$&');
  };

  replaceAliasesMerged_(body, ALIASES, data, escapeRegex);

  // Placeholder tanggal hari ini, sama seperti config lainnya.
  const HARI_ID = ['Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu', 'Minggu'];
  const BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const dateAliases = {
    HARI_SEKARANG: ['HARI SEKARANG'],
    TANGGAL_SEKARANG: ['TANGGAL SEKARANG'],
    BULAN_SEKARANG: ['BULAN SEKARANG'],
    TAHUN_SEKARANG: ['TAHUN SEKARANG']
  };
  const dateValues = {
    HARI_SEKARANG: HARI_ID[Number(Utilities.formatDate(now, tz, 'u')) - 1],
    TANGGAL_SEKARANG: Utilities.formatDate(now, tz, 'd'),
    BULAN_SEKARANG: BULAN_ID[Number(Utilities.formatDate(now, tz, 'M')) - 1],
    TAHUN_SEKARANG: Utilities.formatDate(now, tz, 'yyyy')
  };
  Object.keys(dateAliases).forEach(function (key) {
    dateAliases[key].forEach(function (name) {
      const pattern = '\\{\\{?' + escapeRegex(name) + '\\}?\\}';
      body.replaceText(pattern, String(dateValues[key]));
    });
  });

  // Rekatkan "TANGGAL BULAN TAHUN" (mis. "2 September 2026") pakai spasi
  // non-breaking, supaya Google Docs tidak pernah memotongnya jadi 2
  // baris (mis. "2 September" di baris 1, "2026" sendirian di baris 2)
  // gara-gara lebar paragraf tanda tangan yang sempit.
  const dateGluedFrom_ = dateValues.TANGGAL_SEKARANG + ' ' + dateValues.BULAN_SEKARANG + ' ' + dateValues.TAHUN_SEKARANG;
  const dateGluedTo_ = dateValues.TANGGAL_SEKARANG + '\u00A0' + dateValues.BULAN_SEKARANG + '\u00A0' + dateValues.TAHUN_SEKARANG;
  body.replaceText(escapeRegex(dateGluedFrom_), dateGluedTo_);



  doc.saveAndClose();

  const pdfFile = withRetry_(function () {
    const pdfBlob = newDocFile.getAs(MimeType.PDF).setName(fileName + '.pdf');
    return pdfFolder.createFile(pdfBlob);
  }, 3, 'export PDF (BB Percepatan)');

  // Sama seperti di generateForRow(): setSharing bisa gagal kalau
  // kebijakan Google Workspace organisasi memblokir sharing "Anyone
  // with link" ke luar domain. Dibungkus try/catch tersendiri supaya
  // kegagalan di sini tidak menggagalkan seluruh proses -- dokumen
  // (Word & PDF) tetap dianggap berhasil dibuat.
  try {
    withRetry_(function () {
      newDocFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }, 3, 'setSharing (BB Percepatan)');
  } catch (sharingErr) {
    Logger.log('generateBBPercepatanForGroup_: setSharing gagal (kemungkinan dibatasi kebijakan Google Workspace), dilewati: ' + sharingErr);
  }

  // Tidak ada kolom output link di tab ini (sama seperti SPK New/Ekspand/Pembatalan).

  // ---- OTOMATIS BERSIHKAN CACHE DASHBOARD ----
  // Sama seperti di generateForRow(): supaya dashboard web selalu tampil
  // data terbaru begitu dokumen BB Percepatan berhasil dibuat, tanpa
  // perlu klik "Bersihkan cache dashboard" secara manual.
  invalidateDocsCache_();

  return true;
}

/**
 * ======================================================================
 * TAMBAHAN UNTUK WEB APP — SPK DASHBOARD (LENGKAP)
 * ======================================================================
 */

// ----------------------------------------------------------------------
// KONTROL AKSES — silakan disesuaikan
// ----------------------------------------------------------------------
// Dashboard sekarang di-hosting di GitHub Pages (di luar Apps Script), jadi
// email/identitas orang yang membuka dashboard TIDAK bisa lagi dideteksi
// otomatis oleh Session.getActiveUser(). Sesuai keputusan: akses dibuka
// untuk semua orang -- daftar tetap dikosongkan supaya semua email
// dianggap admin (bisa generate) dan tidak ada pembatasan region.
const ACCESS_CONFIG = {
  ADMIN_EMAILS: [],
  REGION_RESTRICTIONS: {}
};

const LOG_SHEET_NAME = 'LOG GENERATE WEB';

// ----------------------------------------------------------------------
// CACHE — supaya dashboard tidak selalu memindai ulang seluruh folder
// Drive tiap kali dibuka/di-refresh. Cache otomatis dihapus begitu ada
// dokumen baru yang berhasil digenerate (lihat invalidateDocsCache_()
// yang dipanggil langsung di dalam generateForRow() dan
// generateBBPercepatanForGroup_(), bukan cuma di fungsi pembungkus).
// ----------------------------------------------------------------------
// PENTING (diperbaiki): sebelumnya SEMUA dokumen dari 6 jenis digabung
// jadi SATU entri cache. CacheService punya batas keras ~100KB PER KEY
// -- begitu jumlah dokumen banyak, ukuran gabungan itu gampang lewat
// 100KB, cache.put() diam-diam GAGAL (dibungkus try/catch tanpa
// notifikasi), dan akibatnya dashboard SELALU scan ulang ke-12 folder
// Drive dari nol di SETIAP pemuatan halaman -- ini penyebab utama
// loading terasa lama terus-menerus.
//
// Sekarang tiap jenis dokumen (Takeover/New/Ekspand/dst) punya KEY
// CACHE SENDIRI-SENDIRI yang jauh lebih kecil, jauh lebih jarang lewat
// batas 100KB.
//
// SOAL DATA YANG DIHAPUS LANGSUNG DARI GOOGLE DRIVE (bukan lewat menu/
// dashboard): Apps Script TIDAK punya trigger otomatis untuk "file
// dihapus dari folder" (beda dengan onEdit di Sheet), jadi
// invalidateDocsCache_() tidak bisa langsung tahu kejadian itu. Untuk
// mengatasi ini dipakai 2 lapis:
//   1. Durasi cache diperpendek jadi 5 menit (bukan 6 jam lagi) --
//      jadi paling lama 5 menit dashboard otomatis "menyembuhkan diri"
//      sendiri tanpa siapa pun perlu klik apa pun.
//   2. Tombol "Muat Ulang" di halaman web (BUKAN di Apps Script) selalu
//      melewati cache dan scan ulang Drive dari nol -- lihat
//      getAllDocumentsFresh() di bawah -- jadi begitu ditekan, hasilnya
//      pasti akurat saat itu juga.
const DOCS_CACHE_PREFIX = 'spk_dashboard_docs_v2_';
const DOCS_CACHE_SECONDS = 300; // 5 menit -- lihat catatan di atas

// Daftar jenis dokumen + fungsi pengambil datanya masing-masing, dipakai
// bersama oleh getAllDocsCached_() & invalidateDocsCache_() supaya
// keduanya selalu sinkron mengacu ke jenis yang sama.
const DOC_TYPE_FETCHERS_ = [
  { label: 'Takeover', fetch: function () { return getDocumentsForConfig_(TAKEOVER, 'Takeover'); } },
  { label: 'New', fetch: function () { return getDocumentsForConfig_(SPKNEW, 'New'); } },
  { label: 'Ekspand', fetch: function () { return getDocumentsForConfig_(EKSPAND, 'Ekspand'); } },
  { label: 'Pembatalan PO', fetch: function () { return getDocumentsForConfig_(PEMBATALAN, 'Pembatalan PO'); } },
  { label: 'BB Percepatan', fetch: function () { return getDocumentsForBBPercepatan_(BBPERCEPATAN, 'BB Percepatan'); } },
  { label: 'BA Cancel One on One', fetch: function () { return getDocumentsForConfig_(BACANCEL, 'BA Cancel One on One'); } }
];

function invalidateDocsCache_() {
  try {
    const keys = DOC_TYPE_FETCHERS_.map(function (t) { return DOCS_CACHE_PREFIX + t.label; });
    CacheService.getScriptCache().removeAll(keys);
  } catch (err) { /* abaikan */ }
}

// ----------------------------------------------------------------------
// Bersihkan cache dashboard secara manual — sekarang jarang dibutuhkan
// karena invalidateDocsCache_() sudah otomatis dipanggil setiap kali
// dokumen berhasil digenerate (lewat menu Sheet, dashboard web, ATAU
// dijalankan langsung dari editor Apps Script). Tetap dipertahankan
// sebagai tombol darurat kalau suatu saat CacheService berperilaku
// aneh atau ingin memaksa refresh manual.
// ----------------------------------------------------------------------
function bersihkanCacheManual() {
  invalidateDocsCache_();
  SpreadsheetApp.getUi().alert('Cache dashboard sudah dibersihkan. Silakan refresh/Muat Ulang halaman web app sekarang.');
}

function getConfigByType_(typeLabel) {
  switch (typeLabel) {
    case 'Takeover': return TAKEOVER;
    case 'New': return SPKNEW;
    case 'Ekspand': return EKSPAND;
    case 'Pembatalan PO': return PEMBATALAN;
    case 'BB Percepatan': return BBPERCEPATAN;
    case 'BA Cancel One on One': return BACANCEL;
    default: return null;
  }
}

// ----------------------------------------------------------------------
// Web app entry point
// ----------------------------------------------------------------------
function doGet(e) {
  // Kalau ada parameter ?action=..., ini adalah panggilan API JSON dari
  // dashboard yang di-hosting di luar Apps Script (mis. GitHub Pages),
  // bukan permintaan untuk membuka halaman HTML. Dipisah supaya dashboard
  // Apps Script (Index.html) yang lama tetap bisa jalan seperti biasa.
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest_(e.parameter.action, e.parameter);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SPK Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------------------------
// Endpoint POST — dipakai untuk aksi yang MENULIS data (generate dokumen,
// generate massal, download ZIP), supaya tidak bisa "diulang" cuma dengan
// membuka ulang sebuah URL seperti halnya GET.
// PENTING: body dikirim sebagai JSON, tapi header Content-Type dari client
// HARUS 'text/plain' (bukan 'application/json') supaya browser tidak
// mengirim preflight OPTIONS -- Apps Script Web App tidak punya doOptions()
// dan akan menolak preflight itu.
// ----------------------------------------------------------------------
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ error: 'Body request tidak valid (bukan JSON): ' + err });
  }
  return handleApiRequest_(body.action, body);
}

// ----------------------------------------------------------------------
// Router sederhana: memetakan nama "action" dari client ke fungsi backend
// yang sudah ada, lalu membungkus hasilnya (atau error-nya) jadi JSON.
// Ini pengganti google.script.run untuk dashboard yang di-hosting di luar
// Apps Script (mis. GitHub Pages) dan hanya bisa memanggil lewat fetch().
// ----------------------------------------------------------------------
function handleApiRequest_(action, params) {
  params = params || {};
  try {
    let result;
    switch (action) {
      case 'getAllDocuments':
        result = getAllDocuments();
        break;
      case 'getAllDocumentsFresh':
        result = getAllDocumentsFresh();
        break;
      case 'getRecentLogs':
        result = getRecentLogs(Number(params.limit) || 50);
        break;
      case 'generateDashboardRow':
        result = generateDashboardRow(params.type, Number(params.rowNumber), params.groupKey || null);
        break;
      case 'generateDashboardBatch':
        result = generateDashboardBatch(params.items || []);
        break;
      case 'downloadSelectedZip':
        result = downloadSelectedZip(params.items || [], params.format || 'both');
        break;
      case 'deleteDashboardDocument':
        result = deleteDashboardDocument(params.items || []);
        break;
      default:
        return jsonOutput_({ error: 'Action tidak dikenal: ' + action });
    }
    return jsonOutput_({ result: result });
  } catch (err) {
    return jsonOutput_({ error: (err && err.message) ? err.message : String(err) });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getCurrentUserAccess_() {
  let email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (err) { email = ''; }

  const adminList = ACCESS_CONFIG.ADMIN_EMAILS;
  const canGenerate = adminList.length === 0 || adminList.indexOf(email) !== -1;
  const allowedRegions = ACCESS_CONFIG.REGION_RESTRICTIONS[email] || null; // null = semua region

  return { email: email || '(tidak diketahui)', canGenerate: canGenerate, allowedRegions: allowedRegions };
}

function buildFileName_(data, config) {
  const noSpkStr = String(data.NO_SPK || '');
  const noSpkAwal = noSpkStr.indexOf('/') !== -1
    ? noSpkStr.substring(0, noSpkStr.indexOf('/'))
    : noSpkStr;
  return config.FILE_PREFIX + '_' + noSpkAwal + '_' + data.KEPADA + '_' + data.STASIUN;
}

function buildFileMap_(folder) {
  const map = {};
  try {
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      const name = f.getName();
      const existing = map[name];
      if (!existing || f.getDateCreated().getTime() > existing.getDateCreated().getTime()) {
        map[name] = f;
      }
    }
  } catch (err) {
    // folder tidak bisa diakses -> kembalikan kamus kosong, jangan gagal total
  }
  return map;
}

function getDocumentsForConfig_(config, typeLabel) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.SHEET_NAME);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return [];

  const values = sheet.getRange(
    FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, config.LAST_DATA_COL
  ).getValues();

  const COL = config.COL;
  const get = function (row, colIndex) { return colIndex ? row[colIndex - 1] : ''; };

  let wordMap, pdfMap;
  try {
    const wordFolder = DriveApp.getFolderById(config.OUTPUT_FOLDER_WORD_ID);
    const pdfFolder = DriveApp.getFolderById(config.OUTPUT_FOLDER_PDF_ID);
    wordMap = buildFileMap_(wordFolder);
    pdfMap = buildFileMap_(pdfFolder);
  } catch (err) {
    return [];
  }

  const tz = Session.getScriptTimeZone();
  const results = [];

  values.forEach(function (row, idx) {
    const data = {
      REGION: get(row, COL.REGION),
      STASIUN: get(row, COL.STASIUN),
      NO_PO: get(row, COL.NO_PO),
      NO_SPK: get(row, COL.NO_SPK),
      KEPADA: get(row, COL.NAMA_MITRA_PT_CV)
    };

    const rowNumber = FIRST_DATA_ROW + idx;
    const isComplete = !!(data.NO_SPK && data.KEPADA && data.STASIUN);
    if (!isComplete) return; // lewati baris yang datanya belum lengkap sama sekali

    const fileName = buildFileName_(data, config);
    const wordFile = wordMap[fileName] || null;
    const pdfFile = pdfMap[fileName + '.pdf'] || null;

    const createdDateObj = wordFile ? wordFile.getDateCreated() : (pdfFile ? pdfFile.getDateCreated() : null);
    const createdDateText = createdDateObj ? Utilities.formatDate(createdDateObj, tz, 'dd MMM yyyy, HH:mm') : null;

    results.push({
      type: typeLabel,
      rowNumber: rowNumber,
      noSpk: data.NO_SPK,
      noPo: data.NO_PO,
      kepada: data.KEPADA,
      stasiun: data.STASIUN,
      region: data.REGION,
      fileName: fileName,
      status: (wordFile || pdfFile) ? 'Sudah Dibuat' : 'Belum Dibuat',
      createdDate: createdDateText,
      createdDateSort: createdDateObj ? createdDateObj.getTime() : 0,

      wordFileId: wordFile ? wordFile.getId() : null,
      pdfFileId: pdfFile ? pdfFile.getId() : null,

      wordViewUrl: wordFile ? wordFile.getUrl() : null,
      wordPreviewUrl: wordFile ? ('https://docs.google.com/document/d/' + wordFile.getId() + '/preview') : null,
      wordDownloadUrl: wordFile ? ('https://docs.google.com/document/d/' + wordFile.getId() + '/export?format=docx') : null,

      pdfViewUrl: pdfFile ? pdfFile.getUrl() : null,
      pdfPreviewUrl: pdfFile ? ('https://drive.google.com/file/d/' + pdfFile.getId() + '/preview') : null,
      pdfDownloadUrl: pdfFile ? ('https://drive.google.com/uc?export=download&id=' + pdfFile.getId()) : null
    });
  });

  return results;
}

function getDocumentsForBBPercepatan_(config, typeLabel) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.SHEET_NAME);
  if (!sheet) return [];

  const groups = buildBBPercepatanGroups_(sheet, config);
  if (groups.length === 0) return [];

  let wordMap, pdfMap;
  try {
    const wordFolder = DriveApp.getFolderById(config.OUTPUT_FOLDER_WORD_ID);
    const pdfFolder = DriveApp.getFolderById(config.OUTPUT_FOLDER_PDF_ID);
    wordMap = buildFileMap_(wordFolder);
    pdfMap = buildFileMap_(pdfFolder);
  } catch (err) {
    return [];
  }

  const tz = Session.getScriptTimeZone();
  const results = [];

  groups.forEach(function (group) {
    if (!group.isComplete) return; // lewati grup yang datanya belum lengkap

    const found = lookupBBPercepatanFiles_(group, wordMap, pdfMap);
    const fileName = found.fileName;
    const wordFile = found.wordFile;
    const pdfFile = found.pdfFile;

    const createdDateObj = wordFile ? wordFile.getDateCreated() : (pdfFile ? pdfFile.getDateCreated() : null);
    const createdDateText = createdDateObj ? Utilities.formatDate(createdDateObj, tz, 'dd MMM yyyy, HH:mm') : null;

    const stasiunList = group.rows.map(function (r) { return r.stasiun; }).join(', ');

    const noSpkDisplayList = [];
    const seenNoSpk_ = {};
    group.rows.forEach(function (r) {
      [r.noSpkSurvey, r.noSpkPatching, r.noSpkHc].forEach(function (v) {
        const raw = String(v || '').trim();
        if (!raw) return;
        const norm = normalizeNoSpk_(raw);
        if (seenNoSpk_[norm]) return;
        seenNoSpk_[norm] = true;
        noSpkDisplayList.push(raw);
      });
    });
    const noSpkDisplay = noSpkDisplayList.join(', ');

    results.push({
      type: typeLabel,
      groupKey: group.key,
      rowNumber: group.rowNumbers[0],
      rowNumbers: group.rowNumbers,
      noSpk: noSpkDisplay || '-',
      noPo: group.rows.length + ' lokasi',
      kepada: group.kepada,
      mitraPembangunan: group.mitraPembangunan || '',
      stasiun: stasiunList,
      region: group.region,
      fileName: fileName,
      status: (wordFile || pdfFile) ? 'Sudah Dibuat' : 'Belum Dibuat',
      createdDate: createdDateText,
      createdDateSort: createdDateObj ? createdDateObj.getTime() : 0,

      wordFileId: wordFile ? wordFile.getId() : null,
      pdfFileId: pdfFile ? pdfFile.getId() : null,

      wordViewUrl: wordFile ? wordFile.getUrl() : null,
      wordPreviewUrl: wordFile ? ('https://docs.google.com/document/d/' + wordFile.getId() + '/preview') : null,
      wordDownloadUrl: wordFile ? ('https://docs.google.com/document/d/' + wordFile.getId() + '/export?format=docx') : null,

      pdfViewUrl: pdfFile ? pdfFile.getUrl() : null,
      pdfPreviewUrl: pdfFile ? ('https://drive.google.com/file/d/' + pdfFile.getId() + '/preview') : null,
      pdfDownloadUrl: pdfFile ? ('https://drive.google.com/uc?export=download&id=' + pdfFile.getId()) : null
    });
  });

  return results;
}

// Dipertahankan untuk kompatibilitas kalau ada kode lain yang memanggil
// (mis. dari fungsi debug manual) -- tapi getAllDocsCached_ di bawah
// TIDAK memakai ini lagi untuk alur normal dashboard.
function buildAllDocsFresh_() {
  let all = [];
  DOC_TYPE_FETCHERS_.forEach(function (t) { all.push.apply(all, t.fetch()); });
  return all;
}

function getAllDocsCached_() {
  const cache = CacheService.getScriptCache();
  let all = [];

  DOC_TYPE_FETCHERS_.forEach(function (t) {
    const key = DOCS_CACHE_PREFIX + t.label;
    let docsForType = null;

    try {
      const cached = cache.get(key);
      if (cached) docsForType = JSON.parse(cached);
    } catch (err) {
      // cache rusak/tidak terbaca -> lanjut ambil fresh seperti biasa
    }

    if (!docsForType) {
      docsForType = t.fetch();
      try {
        cache.put(key, JSON.stringify(docsForType), DOCS_CACHE_SECONDS);
      } catch (err) {
        // Data jenis ini masih kebesaran untuk 1 key cache -> tidak
        // masalah, dashboard tetap jalan normal, hanya jenis ini yang
        // akan discan ulang tiap load (jenis lain tetap ke-cache).
      }
    }

    all = all.concat(docsForType);
  });

  return all;
}

// ----------------------------------------------------------------------
// Versi "paksa segar" dari getAllDocuments() -- dipanggil oleh tombol
// "Muat Ulang" di dashboard web. Membersihkan cache dulu, baru scan
// ulang seluruh folder Drive, sehingga hasilnya PASTI mencerminkan
// kondisi Drive saat ini juga -- termasuk file yang baru saja dihapus
// langsung dari Drive (yang tidak bisa terdeteksi otomatis oleh
// invalidateDocsCache_() di jalur generate).
// ----------------------------------------------------------------------
function getAllDocumentsFresh() {
  invalidateDocsCache_();
  return getAllDocuments();
}

function getAllDocuments() {
  const access = getCurrentUserAccess_();

  let all = getAllDocsCached_();

  if (access.allowedRegions) {
    all = all.filter(function (d) { return access.allowedRegions.indexOf(d.region) !== -1; });
  }

  all.forEach(function (d) { d.canGenerate = access.canGenerate; });

  all.sort(function (a, b) {
    if (a.createdDateSort !== b.createdDateSort) return b.createdDateSort - a.createdDateSort;
    return b.rowNumber - a.rowNumber;
  });

  return {
    docs: all,
    user: {
      email: access.email,
      canGenerate: access.canGenerate,
      allowedRegions: access.allowedRegions
    }
  };
}

// ----------------------------------------------------------------------
// Generate dokumen langsung dari dashboard (satu dokumen).
// ----------------------------------------------------------------------
function generateDashboardRow(typeLabel, rowNumber, groupKey) {
  const access = getCurrentUserAccess_();
  if (!access.canGenerate) {
    throw new Error('Anda tidak memiliki izin untuk generate dokumen. Hubungi admin.');
  }

  const config = getConfigByType_(typeLabel);
  if (!config) throw new Error('Jenis dokumen "' + typeLabel + '" tidak dikenal.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.SHEET_NAME);
  if (!sheet) throw new Error('Tab "' + config.SHEET_NAME + '" tidak ditemukan.');

  if (config.GROUPED) {
    const group = findBBPercepatanGroup_(sheet, config, groupKey);
    if (access.allowedRegions && group) {
      if (access.allowedRegions.indexOf(group.region) === -1) {
        throw new Error('Anda tidak memiliki akses ke region "' + group.region + '".');
      }
    }
    const result = generateBBPercepatanForGroup_(sheet, config, groupKey);
    if (result === true) {
      logGenerateAction_(typeLabel, groupKey, access.email);
      invalidateDocsCache_();
      return { success: true };
    } else if (result === false) {
      return { success: false, message: 'Data untuk mitra pembangunan ini belum lengkap, dokumen belum dibuat.' };
    } else {
      return { success: false, message: String(result) };
    }
  }

  if (access.allowedRegions) {
    const region = sheet.getRange(rowNumber, config.COL.REGION).getValue();
    if (access.allowedRegions.indexOf(region) === -1) {
      throw new Error('Anda tidak memiliki akses ke region "' + region + '".');
    }
  }

  const result = generateForRow(sheet, rowNumber, config); // fungsi asli dari kode generator

  if (result === true) {
    logGenerateAction_(typeLabel, rowNumber, access.email);
    invalidateDocsCache_();
    return { success: true };
  } else if (result === false) {
    return { success: false, message: 'Data di baris ini belum lengkap, dokumen belum dibuat.' };
  } else {
    return { success: false, message: String(result) };
  }
}

// ----------------------------------------------------------------------
// Generate BANYAK baris/grup sekaligus (dipilih via checklist di dashboard).
// items: array of { type, rowNumber, groupKey }
//   - Untuk Takeover/New/Ekspand/Pembatalan PO/BA Cancel: rowNumber dipakai, groupKey diabaikan.
//   - Untuk BB Percepatan: groupKey (nilai MITRA PEMBANGUNAN) dipakai, rowNumber diabaikan.
// Memanggil generateForRow() / generateBBPercepatanForGroup_() yang SUDAH
// ADA, satu per satu, supaya perilakunya identik dengan generate satuan.
// ----------------------------------------------------------------------
function generateDashboardBatch(items) {
  const access = getCurrentUserAccess_();
  if (!access.canGenerate) {
    throw new Error('Anda tidak memiliki izin untuk generate dokumen. Hubungi admin.');
  }
  if (!items || items.length === 0) {
    return { successCount: 0, failCount: 0, messages: [] };
  }

  let successCount = 0;
  let failCount = 0;
  const messages = [];
  const sheetCache = {};

  items.forEach(function (item) {
    // PENTING: seluruh body dibungkus try/catch. Tanpa ini, kalau satu
    // item melempar error tak terduga (mis. withRetry_ menyerah setelah
    // 3 percobaan, error kuota Drive, dsb), forEach akan langsung
    // berhenti total dan item-item SETELAHNYA tidak pernah diproses --
    // padahal item SEBELUM error itu sudah terlanjur membuat dokumen.
    // Ini penyebab gejala "generate massal cuma jadi 1 dokumen".
    try {
      const config = getConfigByType_(item.type);
      const label = item.groupKey ? (item.type + ' "' + item.groupKey + '"') : (item.type + ' baris ' + item.rowNumber);
      if (!config) {
        failCount++;
        messages.push(label + ': jenis dokumen tidak dikenal.');
        return;
      }

      let sheet = sheetCache[config.SHEET_NAME];
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.SHEET_NAME);
        sheetCache[config.SHEET_NAME] = sheet;
      }
      if (!sheet) {
        failCount++;
        messages.push(label + ': tab "' + config.SHEET_NAME + '" tidak ditemukan.');
        return;
      }

      if (config.GROUPED) {
        if (access.allowedRegions) {
          const group = findBBPercepatanGroup_(sheet, config, item.groupKey);
          if (group && access.allowedRegions.indexOf(group.region) === -1) {
            failCount++;
            messages.push(label + ': tidak ada akses region "' + group.region + '".');
            return;
          }
        }
        const result = generateBBPercepatanForGroup_(sheet, config, item.groupKey);
        if (result === true) {
          successCount++;
          logGenerateAction_(item.type, item.groupKey, access.email);
        } else {
          failCount++;
          const msg = (result === false) ? 'data belum lengkap.' : String(result);
          messages.push(label + ': ' + msg);
        }
        return;
      }

      if (access.allowedRegions) {
        const region = sheet.getRange(item.rowNumber, config.COL.REGION).getValue();
        if (access.allowedRegions.indexOf(region) === -1) {
          failCount++;
          messages.push(label + ': tidak ada akses region "' + region + '".');
          return;
        }
      }

      const result = generateForRow(sheet, item.rowNumber, config);
      if (result === true) {
        successCount++;
        logGenerateAction_(item.type, item.rowNumber, access.email);
      } else {
        failCount++;
        const msg = (result === false) ? 'data belum lengkap.' : String(result);
        messages.push(label + ': ' + msg);
      }
    } catch (err) {
      // Tangkap SEMUA error tak terduga per item supaya item lain di
      // batch tetap lanjut diproses, bukan ikut gagal semua.
      failCount++;
      const label = item.groupKey ? (item.type + ' "' + item.groupKey + '"') : (item.type + ' baris ' + item.rowNumber);
      messages.push(label + ': error tak terduga - ' + err);
      Logger.log('generateDashboardBatch item error (' + label + '): ' + err);
    }
  });

  if (successCount > 0) invalidateDocsCache_();

  return { successCount: successCount, failCount: failCount, messages: messages };
}

// ----------------------------------------------------------------------
// Download beberapa dokumen sekaligus sebagai satu file ZIP.
// items: array of { wordFileId, pdfFileId }
// format: 'word' | 'pdf' | 'both'
// ----------------------------------------------------------------------
// ----------------------------------------------------------------------
// Hapus dokumen (Word & PDF) langsung dari dashboard web, tanpa perlu buka
// Google Drive manual. File dipindah ke SAMPAH Drive (setTrashed), BUKAN
// dihapus permanen -- supaya masih bisa dipulihkan lewat Sampah Drive
// kalau ternyata salah hapus. Kalau tab-nya punya kolom link (cuma
// Takeover yang punya kolom DOKUMEN_WORD/DOKUMEN_PDF), link lama di sheet
// juga dibersihkan supaya tidak menunjuk ke file yang sudah tidak ada.
// items: array of { type, rowNumber, groupKey, wordFileId, pdfFileId }
// ----------------------------------------------------------------------
function deleteDashboardDocument(items) {
  const access = getCurrentUserAccess_();
  if (!access.canGenerate) {
    throw new Error('Anda tidak memiliki izin untuk menghapus dokumen. Hubungi admin.');
  }
  if (!items || items.length === 0) {
    return { successCount: 0, failCount: 0, messages: [] };
  }

  let successCount = 0;
  let failCount = 0;
  const messages = [];
  const sheetCache = {};

  items.forEach(function (item) {
    // Dibungkus try/catch per item, sama seperti generateDashboardBatch --
    // supaya 1 item gagal (mis. file sudah kadung dihapus manual dari Drive)
    // tidak menghentikan proses hapus item lain dalam batch yang sama.
    const label = item.groupKey ? (item.type + ' "' + item.groupKey + '"') : (item.type + ' baris ' + item.rowNumber);
    try {
      const config = getConfigByType_(item.type);
      if (!config) {
        failCount++;
        messages.push(label + ': jenis dokumen tidak dikenal.');
        return;
      }

      let deletedAny = false;
      if (item.wordFileId) {
        try { DriveApp.getFileById(item.wordFileId).setTrashed(true); deletedAny = true; }
        catch (e) { Logger.log('deleteDashboardDocument: gagal hapus Word (' + label + '): ' + e); }
      }
      if (item.pdfFileId) {
        try { DriveApp.getFileById(item.pdfFileId).setTrashed(true); deletedAny = true; }
        catch (e) { Logger.log('deleteDashboardDocument: gagal hapus PDF (' + label + '): ' + e); }
      }

      if (!deletedAny) {
        failCount++;
        messages.push(label + ': file tidak ditemukan di Drive (mungkin sudah dihapus sebelumnya).');
        return;
      }

      // Bersihkan link lama di sheet, kalau tab ini punya kolom link.
      if (!config.GROUPED && item.rowNumber && (config.COL.DOKUMEN_WORD || config.COL.DOKUMEN_PDF)) {
        let sheet = sheetCache[config.SHEET_NAME];
        if (!sheet) {
          sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.SHEET_NAME);
          sheetCache[config.SHEET_NAME] = sheet;
        }
        if (sheet) {
          if (config.COL.DOKUMEN_WORD) sheet.getRange(item.rowNumber, config.COL.DOKUMEN_WORD).clearContent();
          if (config.COL.DOKUMEN_PDF) sheet.getRange(item.rowNumber, config.COL.DOKUMEN_PDF).clearContent();
        }
      }

      successCount++;
      logGenerateAction_(item.type + ' (HAPUS)', item.groupKey || item.rowNumber, access.email);
    } catch (err) {
      failCount++;
      messages.push(label + ': error tak terduga - ' + err);
      Logger.log('deleteDashboardDocument item error (' + label + '): ' + err);
    }
  });

  if (successCount > 0) invalidateDocsCache_();

  return { successCount: successCount, failCount: failCount, messages: messages };
}

// ----------------------------------------------------------------------
// Ekspor Google Doc -> file .docx (Word).
// KENAPA TIDAK PAKAI DriveApp getAs(MimeType.MICROSOFT_WORD)?
// Untuk Google Docs asli, DriveApp.getAs() HANYA mendukung konversi ke
// PDF/gambar. Konversi ke Word melempar error "Converting from
// application/vnd.google-apps.document to ...wordprocessingml.document
// is not supported". Dulu error itu ditelan diam-diam (try/catch kosong),
// jadi file Word tidak pernah masuk ke ZIP -- itu penyebab "save Word
// tidak bisa" (PDF aman karena PDF diambil apa adanya lewat getBlob()).
// Solusi: minta Google mengekspor Doc-nya sendiri lewat URL export
// resmi memakai token OAuth script ini.
// ----------------------------------------------------------------------
function exportDocAsDocxBlob_(fileId, baseName) {
  const token = ScriptApp.getOAuthToken();
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const urls = [
    'https://docs.google.com/document/d/' + encodeURIComponent(fileId) + '/export?format=docx',
    'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '/export?mimeType=' + encodeURIComponent(DOCX_MIME)
  ];

  let lastErr = '';
  for (let i = 0; i < urls.length; i++) {
    try {
      const resp = UrlFetchApp.fetch(urls[i], {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
        followRedirects: true
      });
      const code = resp.getResponseCode();
      if (code === 200) {
        return resp.getBlob().setName(baseName + '.docx');
      }
      lastErr = 'HTTP ' + code + ' - ' + resp.getContentText().substring(0, 150);
    } catch (err) {
      lastErr = String(err);
    }
  }
  throw new Error('Export Word gagal (' + lastErr + ')');
}

function downloadSelectedZip(items, format) {
  if (!items || items.length === 0) {
    throw new Error('Tidak ada dokumen yang dipilih.');
  }
  format = format || 'both';

  const blobs = [];
  const usedNames = {};
  const failed = []; // pesan kegagalan per file -> dikirim balik ke dashboard

  items.forEach(function (item) {
    if ((format === 'word' || format === 'both') && item.wordFileId) {
      let wordName = item.wordFileId;
      try {
        const wordFile = DriveApp.getFileById(item.wordFileId);
        wordName = wordFile.getName();
        const blob = withRetry_(function () {
          return exportDocAsDocxBlob_(item.wordFileId, wordName);
        }, 2, 'export Word');
        blobs.push(uniqueBlobName_(blob, usedNames));
      } catch (err) {
        failed.push('Word "' + wordName + '": ' + err);
        Logger.log('downloadSelectedZip: gagal ambil Word ' + wordName + ' -> ' + err);
      }
    }
    if ((format === 'pdf' || format === 'both') && item.pdfFileId) {
      let pdfName = item.pdfFileId;
      try {
        const pdfFile = DriveApp.getFileById(item.pdfFileId);
        pdfName = pdfFile.getName();
        blobs.push(uniqueBlobName_(pdfFile.getBlob(), usedNames));
      } catch (err) {
        failed.push('PDF "' + pdfName + '": ' + err);
        Logger.log('downloadSelectedZip: gagal ambil PDF ' + pdfName + ' -> ' + err);
      }
    }
  });

  if (blobs.length === 0) {
    throw new Error('Tidak ada file yang berhasil diambil untuk di-zip.' +
      (failed.length ? ' Penyebab: ' + failed[0] : ''));
  }

  const zipName = 'SPK_Dokumen_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.zip';
  const zipBlob = Utilities.zip(blobs, zipName);

  return {
    base64: Utilities.base64Encode(zipBlob.getBytes()),
    filename: zipBlob.getName(),
    mimeType: 'application/zip',
    fileCount: blobs.length,
    failed: failed
  };
}

// Jalankan SEKALI dari editor Apps Script (pilih fungsi ini di dropdown ->
// klik Run) untuk memberi izin "UrlFetchApp" (script.external_request) yang
// dibutuhkan ekspor Word. Tidak perlu isi ID apa pun: fungsi ini otomatis
// memakai Google Doc pertama yang ditemukan di folder Word Takeover.
// Kalau muncul jendela "Authorization required" -> Review permissions ->
// pilih akun -> Advanced -> Go to ... (unsafe) -> Allow.
// Hasil tes bisa dilihat di Execution log: "Export Word OK ...".
function tesExportWord() {
  const folders = [TAKEOVER, SPKNEW, EKSPAND, PEMBATALAN, BACANCEL, BBPERCEPATAN];
  for (let i = 0; i < folders.length; i++) {
    const it = DriveApp.getFolderById(folders[i].OUTPUT_FOLDER_WORD_ID).getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (f.getMimeType() !== MimeType.GOOGLE_DOCS) continue;
      const blob = exportDocAsDocxBlob_(f.getId(), f.getName());
      Logger.log('Export Word OK: ' + blob.getName() + ' (' + blob.getBytes().length + ' bytes)');
      return;
    }
  }
  Logger.log('Izin sudah OK, tapi belum ada Google Doc hasil generate di folder Word untuk dites.');
}

function uniqueBlobName_(blob, usedNames) {
  let name = blob.getName();
  if (usedNames[name] === undefined) {
    usedNames[name] = 0;
    return blob;
  }
  usedNames[name]++;
  const dotIdx = name.lastIndexOf('.');
  const base = dotIdx !== -1 ? name.substring(0, dotIdx) : name;
  const ext = dotIdx !== -1 ? name.substring(dotIdx) : '';
  return blob.setName(base + ' (' + usedNames[name] + ')' + ext);
}

function logGenerateAction_(typeLabel, rowNumber, email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(['Waktu', 'Email', 'Jenis Dokumen', 'Baris']);
    logSheet.setFrozenRows(1);
  }
  logSheet.appendRow([new Date(), email, typeLabel, rowNumber]);
}

function getRecentLogs(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return [];
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return [];

  const n = Math.min(limit || 50, lastRow - 1);
  const startRow = lastRow - n + 1;
  const values = logSheet.getRange(startRow, 1, n, 4).getValues();
  const tz = Session.getScriptTimeZone();

  return values.reverse().map(function (r) {
    return {
      waktu: Utilities.formatDate(new Date(r[0]), tz, 'dd MMM yyyy, HH:mm'),
      email: r[1],
      jenis: r[2],
      baris: r[3]
    };
  });
}

function tesAksesDrive() {
  const folder = DriveApp.getFolderById('12-iJvU7N14bhbjnNHNUd1zXynBo3dRwP'); // folder output PDF Takeover
  Logger.log('Berhasil akses folder: ' + folder.getName());
}
