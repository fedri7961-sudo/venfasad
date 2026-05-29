// ============================================================
//  ВентФасад Проект — Google Apps Script
//  Вставить в: script.google.com → Новый проект → Code.gs
// ============================================================

const SHEET_NAME = 'Заявки';  // название листа в таблице

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Создать лист и шапку если нет
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Дата', 'Площадь (м²)', 'Тип облицовки', 'Контакт', 'Источник']);
      sheet.setFrozenRows(1);

      // Форматирование шапки
      const header = sheet.getRange(1, 1, 1, 5);
      header.setFontWeight('bold');
      header.setBackground('#2E6A8F');
      header.setFontColor('#ffffff');
    }

    sheet.appendRow([
      new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' }),
      data.area    || '—',
      data.mat     || '—',
      data.contact || '—',
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
