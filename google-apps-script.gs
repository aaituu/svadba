var RSVP_HEADERS = [
  "Дата",
  "Имя",
  "Статус",
  "Количество гостей",
  "Комментарий",
  "User Agent",
];

var ATTENDANCE_LABELS = {
  yes: "ИӘ, ӘРИНЕ",
  with_partner: "ЖҰБАЙЫММЕН КЕЛЕМІН",
  no: "ӨКІНІШКЕ ОРАЙ, КЕЛЕ АЛМАЙМЫН",
};

function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error(
        "Apps Script must be bound to the target Google Spreadsheet.",
      );
    }

    var sheet = spreadsheet.getActiveSheet();
    var data = {};

    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(RSVP_HEADERS);
    } else {
      sheet
        .getRange(1, 1, 1, RSVP_HEADERS.length)
        .setValues([RSVP_HEADERS]);

      if (sheet.getLastColumn() > RSVP_HEADERS.length) {
        sheet
          .getRange(
            1,
            RSVP_HEADERS.length + 1,
            1,
            sheet.getLastColumn() - RSVP_HEADERS.length,
          )
          .clearContent();
      }
    }

    sheet.appendRow([
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone() || "Asia/Almaty",
        "dd.MM.yyyy",
      ),
      data.name || "",
      ATTENDANCE_LABELS[data.attendance] || data.attendance || "",
      data.attendance === "no" ? "0" : data.guests || "1",
      data.comment || "",
      data.userAgent || "",
    ]);

    return jsonResponse({
      status: "success",
      message: "Saved",
    });
  } catch (error) {
    return jsonResponse({
      status: "error",
      message: String(error),
    });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function doGet() {
  return ContentService.createTextOutput(
    "Google Sheets RSVP Web App is working",
  ).setMimeType(ContentService.MimeType.TEXT);
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(
    JSON.stringify(payload),
  ).setMimeType(ContentService.MimeType.JSON);
}
