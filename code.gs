// ── CONFIGURACIÓ ─────────────────────────────────────────
// Posa aquí l'ID del teu Google Sheet (entre /d/ i /edit de la URL)
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
  // Fila d'exemple
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
  setSheet.getRange(2, 1, 2, 2).setValues([
    ['logo', ''],
    ['adminPassword', '1234']
  ]);

  // Eliminar la pestanya "Sheet1" per defecte si existeix
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1') || ss.getSheetByName('Fulla 1');
  if (defaultSheet && ss.getSheets().length > 2) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log('✓ Fulls inicialitzats correctament');
  Logger.log('  - actividades: capçaleres + 1 fila d\'exemple');
  Logger.log('  - settings: logo (buit) + adminPassword (1234)');
}

// ── GESTIONAR PETICIONS POST (no modificar) ───────────────
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
      // Si no existeix, afegir fila nova
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
