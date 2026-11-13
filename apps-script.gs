/** @OnlyCurrentDoc */
// Celotehku — приём лидов с сайта в эту таблицу.
// Разворачивание: Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy (URL не меняется).

const HEADERS = ['Waktu','Nama orang tua','Nama anak','WhatsApp','Usia','Skor %','Kategori','Tanda bahaya','Jawaban',
  'Makan','Kekhawatiran utama','Produk utama',
  'Sumber (ref)','utm_source','utm_medium','utm_campaign','utm_content','Halaman',
  'Pengingat WA','Status'];

function doPost(e){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  // Заголовки: пишем/обновляем первую строку, если она не совпадает с текущей схемой
  const cur = sheet.getLastRow() ? sheet.getRange(1,1,1,HEADERS.length).getValues()[0].join('|') : '';
  if (cur !== HEADERS.join('|')) {
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
    else sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }
  const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  const utm = d.utm || {};
  sheet.appendRow([
    new Date(), d.parent||'', d.child||'', "'"+(d.wa||''), d.age||'', d.score, d.category||'',
    (d.flags||[]).join('; '), (d.answers||[]).join(','),
    d.makan||'', d.concern||'', d.product||'',
    d.ref||'', utm.utm_source||'', utm.utm_medium||'', utm.utm_campaign||'', utm.utm_content||'', d.page||''
  ]);
  // Готовая ссылка-напоминание: клик → открывается WhatsApp с текстом, остаётся нажать «Отправить»
  const wa = String(d.wa||'').replace(/\D/g,'');
  if (wa) {
    const txt = 'Halo ' + (d.parent||'Bunda') + ' 👋 Ini dari Celotehku. Bunda sudah cek tumbuh kembang ' + (d.child||'si Kecil') +
      ' di website kami. Pendaftaran gelombang ini tutup malam ini pukul 00.00 WIB. Kalau mau ikut, balas "IKUT" saja, nanti kami bantu. Kalau belum sekarang, tidak apa-apa 🙏';
    const url = 'https://wa.me/' + wa + '?text=' + encodeURIComponent(txt);
    sheet.getRange(sheet.getLastRow(), HEADERS.indexOf('Pengingat WA') + 1).setFormula('=HYPERLINK("' + url + '","Kirim pengingat")');
  }
  return ContentService.createTextOutput('ok');
}
