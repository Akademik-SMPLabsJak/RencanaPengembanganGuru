/**
 * Rencana Pengembangan Diri Guru: penerima isian form + laporan per guru.
 * Dipasang di akun Google Bagian Akademik (akademik@labschool.xyz).
 *
 * - Setiap kiriman menjadi satu baris di sheet "Rekap v3" (Sheet "Rekap Pengembangan Diri Guru").
 * - Setiap guru otomatis punya satu laporan Google Docs di folder
 *   "Laporan Pengembangan Diri Guru". Kiriman ulang dari email yang sama memperbarui laporannya.
 * - Bagikan Sheet dan folder laporan hanya ke kepala sekolah dan wakil kepala sekolah.
 */

const SHEET_REKAP = 'Rekap v3';
const NAMA_FOLDER = 'Laporan Pengembangan Diri Guru';
// Email ke Akademik menyusul. Ubah ke true bila sudah siap.
const KIRIM_EMAIL = false;
const EMAIL_AKADEMIK = 'akademik@labschool.xyz';

const HEADER = ['Waktu kirim', 'Nama', 'Email', 'Mapel/kelas', 'Data pendukung', 'Temuan utama',
  'Butir dipilih', 'Kompetensi', 'Bentuk pengembangan', 'Dukungan', 'Penjelasan dukungan', 'Laporan', 'Data lengkap (JSON)', 'Status laporan'];

function doGet() {
  return ContentService.createTextOutput('Penerima form Rencana Pengembangan Diri Guru aktif.');
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    // Halaman rekap.html memakai aksi 'rekap' dan 'status'; selain itu kiriman form guru.
    if (d.aksi === 'rekap') return json_(Object.assign({ ok: true }, ambilRekap(d.pin)));
    if (d.aksi === 'status') return json_(Object.assign({ ok: true }, simpanStatus(d.pin, d.email, d.status, d.catatan)));
    simpan_(d);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function simpan_(d) {
  if (!d || !d.nama || !d.email) throw new Error('Nama dan email wajib diisi');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ambilSheet_();
    const sh = siapkanSheet_(ss);
    const butir = d.butir || [];
    let urlLaporan = '', status = 'OK';
    try { urlLaporan = buatLaporan_(d); } catch (err) {
      status = 'GAGAL: ' + (err && err.message || err);
      console.error('Laporan gagal untuk ' + d.email + ': ' + (err && err.stack || err));
    }
    sh.appendRow([
      new Date(), d.nama, d.email, d.mapel,
      (d.dataRef || []).join(', '), d.supervisi || '',
      butir.map(b => b.no + '. ' + b.teks).join('\n'),
      Array.from(new Set(butir.map(b => b.kompetensi))).join(', '),
      (d.bentuk || []).join(', '),
      (d.dukungan || []).join(', '), d.dukunganKet || '',
      urlLaporan, JSON.stringify(d), status,
    ]);
    if (KIRIM_EMAIL) {
      MailApp.sendEmail({
        to: EMAIL_AKADEMIK,
        subject: 'Rencana Pengembangan Diri Guru: ' + d.nama,
        body: 'Isian baru dari ' + d.nama + ' (' + d.email + ').\n\nLaporan: ' + urlLaporan + '\nRekap: ' + ss.getUrl(),
      });
    }
  } finally {
    lock.releaseLock();
  }
}

// Satu dokumen laporan per guru (berdasarkan email). Kiriman ulang menimpa isinya.
function buatLaporan_(d) {
  const props = PropertiesService.getScriptProperties();
  const kunci = 'DOC_' + String(d.email).toLowerCase().trim();
  let doc = null;
  const id = props.getProperty(kunci);
  if (id) { try { doc = DocumentApp.openById(id); } catch (e) { doc = null; } }
  if (!doc) {
    doc = DocumentApp.create('Laporan Pengembangan Diri - ' + d.nama);
    DriveApp.getFileById(doc.getId()).moveTo(folderLaporan_());
    props.setProperty(kunci, doc.getId());
  }
  doc.setName('Laporan Pengembangan Diri - ' + d.nama);
  const body = doc.getBody();
  body.clear();
  const tz = Session.getScriptTimeZone();
  const tgl = Utilities.formatDate(new Date(), tz, 'd MMMM yyyy');

  body.appendParagraph('Rencana Pengembangan Diri Guru').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('SMP Labschool · Bagian Akademik · diperbarui ' + tgl);

  body.appendParagraph('Identitas').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendTable([['Nama', d.nama], ['Email', d.email], ['Mata pelajaran / kelas', d.mapel || '-']]);

  body.appendParagraph('Dasar refleksi').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Data pendukung: ' + ((d.dataRef || []).join(', ') || '-'));
  body.appendParagraph('Temuan utama: ' + (d.supervisi || '-'));

  body.appendParagraph('Kompetensi yang ingin dikembangkan').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const rows = [['No.', 'Kompetensi', 'Butir']];
  (d.butir || []).forEach(b => rows.push([String(b.no), b.kompetensi, b.teks]));
  body.appendTable(rows).getRow(0).editAsText().setBold(true);

  body.appendParagraph('Bentuk pengembangan yang diinginkan').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  (d.bentuk || []).forEach(x => body.appendListItem(x).setGlyphType(DocumentApp.GlyphType.BULLET));

  body.appendParagraph('Dukungan sekolah').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if ((d.dukungan || []).length) (d.dukungan || []).forEach(x => body.appendListItem(x).setGlyphType(DocumentApp.GlyphType.BULLET));
  else body.appendParagraph('-');
  if (d.dukunganKet) body.appendParagraph('Penjelasan: ' + d.dukunganKet);

  body.appendParagraph('Catatan pimpinan').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Persetujuan: ☐ Disetujui   ☐ Perlu revisi');
  body.appendParagraph('Komentar:');
  body.appendParagraph('');
  body.appendParagraph('');
  body.appendParagraph('Butir kompetensi mengacu pada Pasal 8 Permendikdasmen No. 21 Tahun 2025.').editAsText().setFontSize(9);
  doc.saveAndClose();
  return doc.getUrl();
}

function folderLaporan_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(NAMA_FOLDER);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function siapkanSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_REKAP);
  if (!sh) {
    sh = ss.insertSheet(SHEET_REKAP, 0);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
  } else if (sh.getRange(1, HEADER.length).getValue() !== HEADER[HEADER.length - 1]) {
    // Sheet dari versi sebelumnya: perbarui baris judul (kolom baru ditambahkan di kanan).
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
  }
  return sh;
}

// Skrip terikat Sheet memakai Sheet itu; skrip mandiri membuat Sheet sendiri sekali, lalu mengingatnya.
function ambilSheet_() {
  const aktif = SpreadsheetApp.getActiveSpreadsheet();
  if (aktif) return aktif;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const baru = SpreadsheetApp.create('Rekap Pengembangan Diri Guru');
  props.setProperty('SHEET_ID', baru.getId());
  return baru;
}

// Jalankan sekali dari editor (pilih "siapkan", klik Run) untuk memberi izin baru
// (Sheet, Docs, Drive) dan menyiapkan Sheet serta folder laporan. Tautannya muncul di Log eksekusi.
function siapkan() {
  const ss = ambilSheet_();
  siapkanSheet_(ss);
  const f = folderLaporan_();
  Logger.log('Sheet rekap: ' + ss.getUrl());
  Logger.log('Folder laporan: ' + f.getUrl());
}

// Buat (atau perbarui) laporan untuk SEMUA baris di "Rekap v3".
// Jalankan dari editor: pilih "buatUlangLaporan", klik Run. Sekaligus memicu izin Docs/Drive.
function buatUlangLaporan() {
  const sh = siapkanSheet_(ambilSheet_());
  const last = sh.getLastRow();
  if (last < 2) { Logger.log('Belum ada kiriman di ' + SHEET_REKAP); return; }
  const kolJson = HEADER.indexOf('Data lengkap (JSON)') + 1;
  const kolUrl = HEADER.indexOf('Laporan') + 1;
  const kolStatus = HEADER.indexOf('Status laporan') + 1;
  const data = sh.getRange(2, kolJson, last - 1, 1).getValues();
  data.forEach((r, i) => {
    let url = '', status = 'OK';
    try { url = buatLaporan_(JSON.parse(r[0])); } catch (err) { status = 'GAGAL: ' + (err && err.message || err); }
    sh.getRange(i + 2, kolUrl).setValue(url);
    sh.getRange(i + 2, kolStatus).setValue(status);
    Logger.log('Baris ' + (i + 2) + ': ' + status + ' ' + url);
  });
  Logger.log('Folder laporan: ' + folderLaporan_().getUrl());
}

// ================= REKAP UNTUK KEPALA SEKOLAH & WAKASEK =================
// Halaman rekap.html (GitHub Pages) memanggil doPost dengan aksi 'rekap' dan 'status'.
// PIN TIDAK ditulis di kode (repo publik). Atur di Project Settings > Script Properties:
// nama PIN_REKAP, nilai = PIN yang diinginkan.
const SHEET_STATUS = 'Status Persetujuan';

function cekPin_(pin) {
  const benar = PropertiesService.getScriptProperties().getProperty('PIN_REKAP');
  if (!benar) throw new Error('PIN rekap belum diatur (Script Property PIN_REKAP)');
  if (String(pin || '').trim() !== benar) throw new Error('PIN salah');
}

function sheetStatus_(ss) {
  let sh = ss.getSheetByName(SHEET_STATUS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_STATUS);
    sh.appendRow(['Email', 'Status', 'Catatan', 'Waktu']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  return sh;
}

function pisah_(s) {
  return String(s || '').split(',').map(x => x.trim()).filter(String);
}

// Satu baris per guru (kiriman terakhir), plus status persetujuan terakhir.
function ambilRekap(pin) {
  cekPin_(pin);
  const ss = ambilSheet_();
  const sh = siapkanSheet_(ss);
  const baris = sh.getDataRange().getValues().slice(1);
  const perGuru = {};
  baris.forEach(r => {
    const email = String(r[2] || '').toLowerCase().trim();
    if (!email) return;
    let d = {};
    try { d = JSON.parse(r[12] || '{}'); } catch (e) {}
    const waktu = r[0] instanceof Date ? r[0].getTime() : 0;
    const lama = perGuru[email];
    const jumlah = (lama ? lama.jumlahKirim : 0) + 1;
    if (lama && lama.waktu > waktu) { lama.jumlahKirim = jumlah; return; }
    const butir = d.butir || [];
    perGuru[email] = {
      email: email,
      nama: String(r[1] || ''),
      mapel: String(r[3] || ''),
      waktu: waktu,
      jumlahKirim: jumlah,
      temuan: String(r[5] || ''),
      butir: butir.map(b => b.no + '. ' + b.teks),
      kompetensi: butir.length ? Array.from(new Set(butir.map(b => String(b.kompetensi)))) : pisah_(r[7]),
      bentuk: d.bentuk || pisah_(r[8]),
      dukungan: d.dukungan || pisah_(r[9]),
      dukunganKet: String(r[10] || ''),
      laporan: String(r[11] || ''),
      status: 'Menunggu',
      catatan: '',
      diperbarui: false,
    };
  });
  // Kalau guru mengirim ulang setelah ditinjau, status kembali "Menunggu".
  sheetStatus_(ss).getDataRange().getValues().slice(1).forEach(r => {
    const g = perGuru[String(r[0] || '').toLowerCase().trim()];
    if (!g) return;
    const t = r[3] instanceof Date ? r[3].getTime() : 0;
    if (t >= g.waktu) { g.status = String(r[1]); g.catatan = String(r[2] || ''); g.diperbarui = false; }
    else { g.status = 'Menunggu'; g.catatan = ''; g.diperbarui = true; }
  });
  const guru = Object.keys(perGuru).map(k => perGuru[k]).sort((a, b) => a.nama.localeCompare(b.nama));
  return {
    guru: guru,
    sheetUrl: ss.getUrl(),
    dibuat: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMMM yyyy, HH:mm'),
  };
}

function simpanStatus(pin, email, status, catatan) {
  cekPin_(pin);
  if (['Disetujui', 'Perlu revisi'].indexOf(status) < 0) throw new Error('Status tidak dikenal');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sheetStatus_(ambilSheet_()).appendRow([String(email).toLowerCase().trim(), status, catatan || '', new Date()]);
  } finally {
    lock.releaseLock();
  }
  try {
    tulisStatusKeDoc_(email, status, catatan);
  } catch (err) {
    return { peringatan: 'Status tersimpan, tetapi laporan Docs gagal diperbarui: ' + err.message };
  }
  return {};
}

// Mencentang kotak persetujuan dan mengisi komentar di laporan Docs guru.
function tulisStatusKeDoc_(email, status, catatan) {
  const id = PropertiesService.getScriptProperties().getProperty('DOC_' + String(email).toLowerCase().trim());
  if (!id) return;
  const body = DocumentApp.openById(id).getBody();
  const baris = status === 'Disetujui'
    ? 'Persetujuan: ☑ Disetujui   ☐ Perlu revisi'
    : 'Persetujuan: ☐ Disetujui   ☑ Perlu revisi';
  body.replaceText('Persetujuan:.*', baris);
  const aman = String(catatan || '').replace(/[\\$]/g, '\\$&');
  body.replaceText('Komentar:.*', 'Komentar: ' + aman);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
