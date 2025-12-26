function processReserveEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Автомат") || ss.insertSheet("Автомат");

  console.log("--- СТАРТ---");

  const threads = GmailApp.search('subject:"Новая запись в резерв" newer_than:2d');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const existingKeys = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
        .map(r => normalizeKey(`${r[2] || ""}_${r[11] || ""}_${r[7] || ""}`))
    : [];

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      try {
        const body = message.getPlainBody();
        const messageDate = Utilities.formatDate(message.getDate(), "GMT+3", "dd.MM.yyyy");

        const data = parseReserveEmail(body);

        if (data && data.firstName && data.firstName !== "*") {
          const uniqueKey = normalizeKey(`${data.firstName}_${data.productName}_${data.phone}`);
          if (existingKeys.includes(uniqueKey)) continue;

          sheet.appendRow([
            "РЕЗЕРВ",           // A
            messageDate,        // B
            data.firstName,     // C
            "",                 // D
            "",                 // E
            data.quantity,      // F
            "",                 // G
            data.phone,         // H
            "",                 // I
            "",                 // J
            "",                 // K
            data.productName,   // L
            data.tourDate,      // M
            data.quantity,      // N
            0,                  // O
            "www",              // P
            "В резерве",        // Q
            "",                 // R
            ""                  // S
          ]);

          console.log("ДОБАВЛЕНО: " + data.firstName + " на " + data.tourDate);
          existingKeys.push(uniqueKey);
        }
      } catch (e) {
        console.log("Ошибка: " + e.toString());
      }
    }
  }
  console.log("--- КОНЕЦ ---");
}

function parseReserveEmail(body) {
  if (!body) return null;

  const firstName = (safeMatch(body, /\*?Имя\*?\s*[\n\r]+([^\n\r\*]+)/i, 1) || "").trim();
  let quantity = (safeMatch(body, /\*?Количество человек\*?\s*[\n\r]+(?:[^\d]*(\d+))?/i, 1) || "1").toString();
  let phone = (safeMatch(body, /\*?Телефон\*?\s*[\n\r]+([^\n\r\*]+)/i, 1) || "").replace(/\D/g, "");
  phone = normalizePhone(phone);

  const productName = (safeMatch(body, /\*?Название экскурсии\*?\s*[\n\r]+([^\n\r\*]+)/i, 1) || "").trim();
  const tourDateRaw = (safeMatch(body, /\*?Дата и время\*?\s*[\n\r]+([^\n\r\*]+)/i, 1) || "").trim();
  const tourDate = formatReserveDate(tourDateRaw);

  return {
    firstName: firstName,
    quantity: quantity || "1",
    phone: phone,
    productName: productName,
    tourDate: tourDate
  };
}
function formatReserveDate(dateStr) {

  if (!dateStr) return "";



  const monthsMap = {

    "янв": "янв", "фев": "фев", "мар": "мар", "апр": "апр", "май": "май", "июн": "июн",

    "июл": "июл", "авг": "авг", "сен": "сен", "окт": "окт", "ноя": "ноя", "дек": "дек",

    "января": "янв", "февраля": "фев", "марта": "мар", "апреля": "апр", "мая": "май", "июня": "июн",

    "июля": "июл", "августа": "авг", "сентября": "сен", "октября": "окт", "ноября": "ноя", "декабря": "дек"

  };




  const dayMatch = dateStr.match(/\d+/);

  if (!dayMatch) return dateStr;

  const day = parseInt(dayMatch[0], 10);


  let foundMonth = "";

  const lowerDate = dateStr.toLowerCase();

  for (let key in monthsMap) {

    if (lowerDate.includes(key)) {

      foundMonth = monthsMap[key];

      break;

    }

  }



  return foundMonth ? `${day}-${foundMonth}` : dateStr;

}
function logError(logSheet, error, messageId) {

  const now = new Date();

  logSheet.appendRow([now, messageId, error.toString()]);

}
