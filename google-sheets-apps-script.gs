const SPREADSHEET_NAME = "Habitos diarios";
const SHEET_NAME = "Historial";
const HABITS = ["ejercicio", "lectura", "dibujo"];

function doGet(event) {
  if (event && event.parameter && event.parameter.action === "read") {
    return readData_(event.parameter.callback || "");
  }

  const spreadsheet = getOrCreateSpreadsheet_();
  const email = Session.getActiveUser().getEmail() || "Google conectado";
  const html = `
    <main style="font-family:Arial,sans-serif;max-width:680px;margin:40px auto;line-height:1.45">
      <h1>Habitos diarios conectado</h1>
      <p>Login activo: ${escapeHtml_(email)}</p>
      <p>La hoja donde se guardan los datos es:</p>
      <p><a href="${spreadsheet.getUrl()}" target="_blank" rel="noopener">${spreadsheet.getUrl()}</a></p>
      <p>Ya podes volver a la pagina de habitos. Los cambios se guardan automaticamente.</p>
    </main>
  `;

  return HtmlService.createHtmlOutput(html)
    .setTitle("Habitos diarios conectado");
}

function readData_(callback) {
  const response = { ok: true, data: readDataObject_() };
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(response)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse_(response);
}

function readDataObject_() {
  const spreadsheet = getOrCreateSpreadsheet_();
  const sheet = getOrCreateSheet_(spreadsheet);
  const values = sheet.getDataRange().getValues();
  const data = {};

  if (values.length < 2) {
    return data;
  }

  const headers = values[0].map((header) => String(header));
  const column = (name) => headers.indexOf(name);
  const dateColumn = column("Fecha");

  values.slice(1).forEach((row) => {
    const date = toDateKey_(dateColumn >= 0 ? row[dateColumn] : "");
    if (!date) return;

    data[date] = {
      ejercicio: {
        done: toBoolean_(row[column("Ejercicio hecho")]),
        minutes: toMinutes_(row[column("Ejercicio minutos")])
      },
      lectura: {
        done: toBoolean_(row[column("Lectura hecha")]),
        minutes: toMinutes_(row[column("Lectura minutos")])
      },
      dibujo: {
        done: toBoolean_(row[column("Dibujo hecho")]),
        minutes: toMinutes_(row[column("Dibujo minutos")])
      },
      noYt: {
        done: toBoolean_(row[column("No YT hecho")])
      }
    };
  });

  return data;
}

function jsonResponse_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function toBoolean_(value) {
  return value === true || String(value).toLowerCase() === "true" || String(value).toUpperCase() === "OK";
}

function toMinutes_(value) {
  const minutes = Number(value);
  return Number.isFinite(minutes) ? minutes : 0;
}

function toDateKey_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value).slice(0, 10);
}

function doPost(event) {
  const payload = JSON.parse(event.parameter.payload || "{}");
  const spreadsheet = getOrCreateSpreadsheet_();
  const sheet = getOrCreateSheet_(spreadsheet);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = readDataObject_();
    Object.assign(data, payload.data || {});
    const rows = buildRows_(data);

    sheet.clearContents();
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, rows[0].length);
  } finally {
    lock.releaseLock();
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      spreadsheetUrl: spreadsheet.getUrl(),
      updatedAt: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSpreadsheet_() {
  const properties = PropertiesService.getUserProperties();
  const existingId = properties.getProperty("spreadsheetId");

  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (error) {
      properties.deleteProperty("spreadsheetId");
    }
  }

  const spreadsheet = SpreadsheetApp.create(SPREADSHEET_NAME);
  properties.setProperty("spreadsheetId", spreadsheet.getId());
  return spreadsheet;
}

function getOrCreateSheet_(spreadsheet) {
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function buildRows_(data) {
  const rows = [[
    "Fecha",
    "Ejercicio hecho",
    "Ejercicio minutos",
    "Lectura hecha",
    "Lectura minutos",
    "Dibujo hecho",
    "Dibujo minutos",
    "No YT hecho",
    "Total minutos"
  ]];

  Object.keys(data)
    .sort()
    .forEach((date) => {
      const day = data[date] || {};
      let total = 0;
      const values = [date];

      HABITS.forEach((habit) => {
        const item = day[habit] || {};
        const minutes = Number(item.minutes || 0);
        total += minutes;
        values.push(Boolean(item.done), minutes);
      });

      values.push(Boolean((day.noYt || {}).done));

      values.push(total);
      rows.push(values);
    });

  if (rows.length === 1) {
    rows.push(["", false, 0, false, 0, false, 0, false, 0]);
  }

  return rows;
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
