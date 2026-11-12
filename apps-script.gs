/**
 * Celotehku — menyimpan lead dari lendinga ke Google Sheets.
 *
 * Cara pasang (5 menit):
 * 1. Buat Google Sheet baru, beri nama "Celotehku Leads".
 * 2. Menu Extensions → Apps Script. Hapus isi default, tempel kode ini, Save.
 * 3. Deploy → New deployment → tipe "Web app".
 *    Execute as: Me. Who has access: Anyone. Klik Deploy, izinkan akses.
 * 4. Salin URL Web app (berakhiran /exec) dan tempel ke CONFIG.leadWebhook di index.html.
 */
const HEADERS = ['Waktu', 'Nama orang tua', 'Nama anak', 'WhatsApp', 'Usia', 'Skor %', 'Kategori',
                 'Tanda bahaya', 'Jawaban', 'Sumber (ref)', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'Halaman'];

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  const d = JSON.parse(e.postData.contents || '{}');
  const utm = d.utm || {};
  sheet.appendRow([
    new Date(), d.parent || '', d.child || '', "'" + (d.wa || ''), d.age || '',
    d.score, d.category || '', (d.flags || []).join('; '), (d.answers || []).join(','),
    d.ref || '', utm.utm_source || '', utm.utm_medium || '', utm.utm_campaign || '', utm.utm_content || '', d.page || ''
  ]);
  return ContentService.createTextOutput('ok');
}
