// ── CONFIGURACIÓ ─────────────────────────────────────────
const SHEET_ID = '1StZBnekRiVpmhIcFlNWGQHXVErilP-WAh0ejI9SrvqA';

// ── INICIALITZAR FULL (executa manualment una vegada) ─────
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Pestanya "actividades"
  let actSheet = ss.getSheetByName('actividades');
  if (!actSheet) {
    actSheet = ss.insertSheet('actividades');
  } else {
    actSheet.clearContents();
  }
  actSheet.getRange(1, 1, 1, 10).setValues([[
    'id', 'day', 'type', 'name', 'time', 'location', 'instructor', 'material', 'obs', 'image'
  ]]);
  actSheet.getRange(2, 1, 1, 10).setValues([[
    1, 0, 'esportiva', 'Ioga al jardí', '08:00', 'Jardí principal', 'Maria García', 'Tovallola', '', ''
  ]]);

  // Pestanya "settings"
  let setSheet = ss.getSheetByName('settings');
  if (!setSheet) {
    setSheet = ss.insertSheet('settings');
  } else {
    setSheet.clearContents();
  }
  setSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  setSheet.getRange(2, 1, 6, 2).setValues([
    ['logo',          ''],
    ['adminPassword', '1234'],
    ['day0',          'Diumenge, 27 d\'Abril 2027|27|ABR'],
    ['day1',          'Dilluns, 28 d\'Abril 2027|28|ABR'],
    ['day2',          'Dimarts, 29 d\'Abril 2027|29|ABR'],
    ['day3',          'Dimecres, 30 d\'Abril 2027|30|ABR'],
  ]);

  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1') || ss.getSheetByName('Fulla 1');
  if (defaultSheet && ss.getSheets().length > 2) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log('✓ Fulls inicialitzats correctament');
}

// ── LLEGIR DADES (GET) ────────────────────────────────────
function doGet() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    const actSheet = ss.getSheetByName('actividades');
    const actRows  = actSheet.getDataRange().getValues();
    const acts = actRows.slice(1)
      .filter(r => r[0] !== '' && r[0] !== null)
      .map(r => ({
        id: Number(r[0]), day: Number(r[1]), type: String(r[2] || 'cultural'),
        name: String(r[3] || ''), time: String(r[4] || ''),
        location: String(r[5] || ''), instructor: String(r[6] || ''),
        material: String(r[7] || ''), obs: String(r[8] || ''),
        image: String(r[9] || '')
      }));

    const setSheet = ss.getSheetByName('settings');
    const setRows  = setSheet.getDataRange().getValues();
    const settings = {};
    setRows.slice(1).forEach(r => { if (r[0]) settings[String(r[0])] = String(r[1] || ''); });

    return ContentService
      .createTextOutput(JSON.stringify({ activities: acts, settings }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GUARDAR DADES (POST) ──────────────────────────────────
function doPost(e) {
  try {
    const action = e.parameter.action;
    const data   = JSON.parse(e.parameter.data);
    const ss     = SpreadsheetApp.openById(SHEET_ID);

    if (action === 'updateActivity') {
      const sheet = ss.getSheetByName('actividades');
      const rows  = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          sheet.getRange(i + 1, 1, 1, 10).setValues([[
            data.id, data.day, data.type, data.name,
            data.time, data.location, data.instructor,
            data.material, data.obs, data.image
          ]]);
          return ok();
        }
      }
      sheet.appendRow([
        data.id, data.day, data.type, data.name,
        data.time, data.location, data.instructor,
        data.material, data.obs, data.image
      ]);
      return ok();
    }

    if (action === 'deleteActivity') {
      const sheet = ss.getSheetByName('actividades');
      const rows  = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          sheet.deleteRow(i + 1);
          return ok();
        }
      }
      return ok();
    }

    if (action === 'updateSetting') {
      const sheet = ss.getSheetByName('settings');
      const rows  = sheet.getDataRange().getValues();
      for (let i = 0; i < rows.length; i++) {
        if (rows[i][0] === data.key) {
          sheet.getRange(i + 1, 2).setValue(data.value);
          return ok();
        }
      }
      sheet.appendRow([data.key, data.value]);
      return ok();
    }

    return ok();
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function ok() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
