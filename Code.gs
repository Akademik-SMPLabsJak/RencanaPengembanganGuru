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
  'Butir dipilih', 'Kompetensi', 'Bentuk pengembangan', 'Dukungan', 'Penjelasan dukungan', 'Laporan', 'Data lengkap (JSON)'];

function doGet() {
  return ContentService.createTextOutput('Penerima form Rencana Pengembangan Diri Guru aktif.');
}

function doPost(e) {
  try {
    simpan_(JSON.parse(e.postData.contents));
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
    let urlLaporan = '';
    try { urlLaporan = buatLaporan_(d); } catch (err) { urlLaporan = 'Gagal membuat laporan: ' + err.message; }
    sh.appendRow([
      new Date(), d.nama, d.email, d.mapel,
      (d.dataRef || []).join(', '), d.supervisi || '',
      butir.map(b => b.no + '. ' + b.teks).join('\n'),
      Array.from(new Set(butir.map(b => b.kompetensi))).join(', '),
      (d.bentuk || []).join(', '),
      (d.dukungan || []).join(', '), d.dukunganKet || '',
      urlLaporan, JSON.stringify(d),
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
    const file = DriveApp.getFileById(doc.getId());
    folderLaporan_().addFile(file);
    DriveApp.getRootFolder().removeFile(file);
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
  body.appendParagraph('\n\n');
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

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
