// ============================================================
//  ВентФасад Проект — Google Apps Script
//  Вставить в: script.google.com → Новый проект → Code.gs
// ============================================================

const SHEET_NAME = 'Заявки';

function doGet(e) {
  try {
    const p = e.parameter;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Дата', 'Имя', 'Телефон', 'Телеграм', 'Площадь (м²)', 'Тип облицовки', 'Источник']);
      sheet.setFrozenRows(1);
      const header = sheet.getRange(1, 1, 1, 7);
      header.setFontWeight('bold');
      header.setBackground('#2E6A8F');
      header.setFontColor('#ffffff');
    }

    sheet.appendRow([
      new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' }),
      p.name  || '—',
      p.phone || '—',
      p.tg    || '—',
      p.area  || '—',
      p.mat   || '—',
      'Сайт'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
