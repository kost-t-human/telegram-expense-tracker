/*
Для работы нужно создать новый проект в https://script.google.com/home
И развернуть его.
Для простоты - запуск от имени владельца, доступ - у всех.

В html в форму в action подставить эндпоинт скрипта
саму html страничку загрузить на хостинх, на поддомен (или домен) обязательно нужен HTTPS.
в телеграмм создать бота, а затем прикрепить к нему мини приложение (можно и кнопку заодно) c сылкой на нашу страницу.

костыль с iframe нужен, т.к. по другому получишь ошибку CORS

*/

function doPost(e) {
  try {
    const data = {
      spreadsheetId: e.parameter.spreadsheetId,
      date: e.parameter.date,
      category: e.parameter.category,
      amount: e.parameter.amount,
      note: e.parameter.note || ""
    };

    const expectedHeaders = ["Отметка времени", "Дата покупки", "Стоимость", "Категория", "Описание"];

    const ss = SpreadsheetApp.openById(data.spreadsheetId);
    let sheet = ss.getSheetByName('Расходы');
    if (!sheet) {
      sheet = ss.insertSheet('Расходы');
    }

    // обеспечить достаточное количество колонок
    if (sheet.getMaxColumns() < expectedHeaders.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), expectedHeaders.length - sheet.getMaxColumns());
    }

    // прочитать первую строку безопасно (если пустые ячейки вернут пустые строки)
    const firstRow = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];

    const hasHeaders = expectedHeaders.every((h, i) => {
      return firstRow[i] && String(firstRow[i]).toString().trim() === h;
    });

    if (!hasHeaders) {
      // вставляем строку заголовков наверх, не трогая существующие данные
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    }

    // подготовить значения для записи
    // Преобразуем YYYY-MM-DD в дату
    let purchaseDate = "";
    if (data.date) {
      const parts = data.date.split("-"); // [YYYY, MM, DD]
      purchaseDate = new Date(parts[0], parts[1] - 1, parts[2]); // месяц 0–11
    }
    //const purchaseDate = data.date ? (isNaN(Date.parse(data.date)) ? data.date : new Date(data.date)) : '';
    const amountValue = data.amount ? Number(data.amount) : '';

    // порядок соответствует заголовкам: Отметка времени, Дата покупки, Стоимость, Категория, Описание
    sheet.appendRow([ new Date(), purchaseDate, amountValue, data.category, data.note ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
