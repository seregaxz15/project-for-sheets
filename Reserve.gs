/**
 * Reserve (предварительная запись) parser — добавляет "РЕЗЕРВ" записи
 */

function processReserveEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Автомат") || ss.insertSheet("Автомат");

  console.log("--- СТАРТ ПРОВЕРКИ РЕЗЕРВА С ДАТОЙ И EMAIL ---");

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
  console.log("--- КОНЕЦ ПРОВЕРКИ ---");
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