/**
 * Rencana Pengembangan Diri Guru: penerima isian form.
 * Dipasang di akun Google Bagian Akademik (akademik@labschool.xyz),
 * menyimpan ke Google Sheet "Rekap Pengembangan Diri Guru".
 * Satu kiriman = satu baris di sheet "Rekap v2".
 */

const SHEET_REKAP = 'Rekap v2';
// Email ke Akademik menyusul. Ubah ke true bila sudah siap.
const KIRIM_EMAIL = false;
const EMAIL_AKADEMIK = 'akademik@labschool.xyz';

const HEADER = ['Waktu kirim', 'Nama', 'NIP/NUPTK', 'Email', 'Mapel/kelas', 'Periode', 'Data pendukung', 'Temuan utama'];
for (let i = 1; i <= 3; i++) {
  HEADER.push('Butir ' + i, 'Kompetensi ' + i, 'Skor diri ' + i, 'Bentuk ' + i, 'Pengembangan ' + i, 'Target ' + i, 'Waktu ' + i);
}
HEADER.push('Dukungan', 'Penjelasan dukungan', 'Data lengkap (JSON)');

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
  lock.waitLock(20000);
  try {
    const ss = ambilSheet_();
    const sh = siapkanSheet_(ss);
    const row = [new Date(), d.nama, d.nip || '', d.email, d.mapel, d.periode, (d.dataRef || []).join(', '), d.supervisi || ''];
    for (let i = 0; i < 3; i++) {
      const it = (d.items || [])[i];
      if (it) row.push(Number(it.butir) + '. ' + it.teks, it.kompetensi, it.skor, it.bentuk, it.pengembangan, it.target, it.waktu);
      else row.push('', '', '', '', '', '', '');
    }
    row.push((d.dukungan || []).join(', '), d.dukunganKet || '', JSON.stringify(d));
    sh.appendRow(row);

    if (KIRIM_EMAIL) {
      MailApp.sendEmail({
        to: EMAIL_AKADEMIK,
        subject: 'Rencana Pengembangan Diri Guru: ' + d.nama + ' (' + d.periode + ')',
        body: 'Isian baru dari ' + d.nama + ' (' + d.email + ') untuk ' + d.periode + '.\n\nBuka rekap: ' + ss.getUrl(),
      });
    }
  } finally {
    lock.releaseLock();
  }
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

// Jalankan sekali dari editor (pilih "siapkan", klik Run). Tautan Sheet muncul di Log eksekusi.
function siapkan() {
  const ss = ambilSheet_();
  siapkanSheet_(ss);
  Logger.log('Sheet rekap: ' + ss.getUrl());
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
