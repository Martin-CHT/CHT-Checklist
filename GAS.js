/**
 * Pomocná funkce: vrátí list podle názvu, pokud neexistuje, vytvoří ho a zapíše hlavičky.
 * @param {string} sheetName - Název listu (např. "Prověrka", "Kontrola", "Audit")
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['ID Formuláře', 'Datum uložení', 'Data (JSON)']);
  }

  return sheet;
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid JSON" })).setMimeType(ContentService.MimeType.JSON);
  }

  // Určení cílového listu – klient posílá "sheet" (Prověrka / Kontrola / Audit)
  const sheetName = data.sheet || 'Data';
  const sheet = getOrCreateSheet(sheetName);

  const formId = data.id || data.formId;
  const formData = data.data || data.formData;

  const range = sheet.getDataRange();
  const values = range.getValues();

  let headers = values.length > 0 ? values[0] : [];

  // Detekce sloupců na základě existující tabulky uživatele
  let idHeader = headers.find(h => h === 'ID Formuláře' || h === 'ID');
  let dateHeader = headers.find(h => h === 'Datum uložení' || h === 'Timestamp');
  let dataHeader = headers.find(h => h === 'Data (JSON)');

  if (headers.length === 0 || headers[0] === "") {
    headers = ['ID Formuláře', 'Datum uložení', 'Data (JSON)'];
    sheet.appendRow(headers);
    idHeader = 'ID Formuláře';
    dateHeader = 'Datum uložení';
    dataHeader = 'Data (JSON)';
  }

  if (!idHeader) { headers.push('ID Formuláře'); idHeader = 'ID Formuláře'; sheet.getRange(1, headers.length).setValue(idHeader); }
  if (!dateHeader) { headers.push('Datum uložení'); dateHeader = 'Datum uložení'; sheet.getRange(1, headers.length).setValue(dateHeader); }
  if (!dataHeader) { headers.push('Data (JSON)'); dataHeader = 'Data (JSON)'; sheet.getRange(1, headers.length).setValue(dataHeader); }

  const idColIndex = headers.indexOf(idHeader);
  const dateColIndex = headers.indexOf(dateHeader);
  const dataColIndex = headers.indexOf(dataHeader);

  let rowIndex = -1;
  if (idColIndex !== -1) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idColIndex]) === String(formId)) {
        rowIndex = i + 1;
        break;
      }
    }
  }

  const now = new Date().toLocaleString('cs-CZ');
  const jsonString = JSON.stringify(formData);

  if (rowIndex > -1) {
    if (dateColIndex !== -1) sheet.getRange(rowIndex, dateColIndex + 1).setValue(now);
    if (dataColIndex !== -1) sheet.getRange(rowIndex, dataColIndex + 1).setValue(jsonString);
  } else {
    const rowArray = new Array(headers.length).fill("");
    if (idColIndex !== -1) rowArray[idColIndex] = formId;
    if (dateColIndex !== -1) rowArray[dateColIndex] = now;
    if (dataColIndex !== -1) rowArray[dataColIndex] = jsonString;
    sheet.appendRow(rowArray);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // Určení cílového listu – klient posílá parametr "sheet"
  const sheetName = e.parameter.sheet || 'Data';
  const sheet = getOrCreateSheet(sheetName);

  const action = e.parameter.action;

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ items: [] })).setMimeType(ContentService.MimeType.JSON);
  }

  const headers = values[0];
  const idHeader = headers.find(h => h === 'ID Formuláře' || h === 'ID');
  const dateHeader = headers.find(h => h === 'Datum uložení' || h === 'Timestamp');
  const dataHeader = headers.find(h => h === 'Data (JSON)');

  const idIndex = headers.indexOf(idHeader);
  const timestampIndex = headers.indexOf(dateHeader);
  const dataIndex = headers.indexOf(dataHeader);

  if (action === 'list') {
    const query = e.parameter.query ? e.parameter.query.toLowerCase() : '';

    if (idIndex === -1) {
      return ContentService.createTextOutput(JSON.stringify({ items: [] })).setMimeType(ContentService.MimeType.JSON);
    }

    const results = [];
    for (let i = 1; i < values.length; i++) {
      const rowId = String(values[i][idIndex]);
      if (rowId && (!query || rowId.toLowerCase().includes(query))) {
        results.push({
          id: rowId,
          timestamp: timestampIndex !== -1 ? values[i][timestampIndex] : ''
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ items: results })).setMimeType(ContentService.MimeType.JSON);
  }

  // Zpracování načtení konkrétního ID
  const formId = e.parameter.id;
  if (!formId) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Missing ID" })).setMimeType(ContentService.MimeType.JSON);
  }

  if (idIndex === -1 || dataIndex === -1) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Nesprávné hlavičky v tabulce. Chybí ID Formuláře nebo Data (JSON)." })).setMimeType(ContentService.MimeType.JSON);
  }

  let rowDataStr = null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIndex]) === String(formId)) {
      rowDataStr = values[i][dataIndex];
      break;
    }
  }

  if (!rowDataStr) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Záznam nebyl nalezen" })).setMimeType(ContentService.MimeType.JSON);
  }

  let dataObj;
  try {
    dataObj = JSON.parse(rowDataStr);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Chyba při parsování dat" })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: dataObj }))
    .setMimeType(ContentService.MimeType.JSON);
}
