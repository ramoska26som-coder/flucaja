// ================================================================
// FLUJO DE CAJA — Google Apps Script Backend
// Vincular a una Google Sheet antes de desplegar.
// Instrucciones completas en README.md
// ================================================================

var DENOMINACIONES = [500, 200, 100, 50, 20, 10, 5, 2, 1];

// ----------------------------------------------------------------
// Punto de entrada único (todo por GET con ?payload=JSON)
// ----------------------------------------------------------------
function doGet(e) {
  try {
    var raw = (e && e.parameter && e.parameter.payload) ? e.parameter.payload : '{"action":"getData"}';
    var payload = JSON.parse(raw);
    var result = ejecutar(payload.action, payload.data || null);
    return responder(result);
  } catch (err) {
    return responder({ ok: false, error: err.message });
  }
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// Router de acciones
// ----------------------------------------------------------------
function ejecutar(action, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  inicializarHojas(ss);

  switch (action) {
    case 'getData':          return getData(ss);
    case 'addMovimiento':    return addMovimiento(ss, data);
    case 'deleteMovimiento': return deleteMovimiento(ss, data.id);
    case 'updateConfig':     return updateConfig(ss, data);
    case 'addPersona':       return addPersona(ss, data.nombre);
    default:                 return { ok: false, error: 'Acción desconocida: ' + action };
  }
}

// ----------------------------------------------------------------
// Inicializar hojas si no existen
// ----------------------------------------------------------------
function inicializarHojas(ss) {
  // Hoja Config
  if (!ss.getSheetByName('Config')) {
    var cfg = ss.insertSheet('Config');
    cfg.getRange('A1:B4').setValues([
      ['saldoInicial',      0],
      ['efectivoPendiente', 0],
      ['saldoExcel',        ''],
      ['saldoSistema',      '']
    ]);
    cfg.hideSheet();
  }

  // Hoja Movimientos
  if (!ss.getSheetByName('Movimientos')) {
    var mov = ss.insertSheet('Movimientos');
    var headers = ['id','fecha','tipo','responsable','entregadoA','descripcion','monto',
                   'd500','d200','d100','d50','d20','d10','d5','d2','d1'];
    mov.getRange(1, 1, 1, headers.length).setValues([headers]);
    mov.setFrozenRows(1);
    // Formato moneda en columna monto (G)
    mov.getRange('G2:G').setNumberFormat('#,##0.00');
  }

  // Hoja Personas
  if (!ss.getSheetByName('Personas')) {
    var per = ss.insertSheet('Personas');
    per.getRange('A1').setValue('nombre');
    per.setFrozenRows(1);
  }
}

// ----------------------------------------------------------------
// getData — devuelve todo el estado de la caja
// ----------------------------------------------------------------
function getData(ss) {
  // Config
  var cfgSheet = ss.getSheetByName('Config');
  var cfgValues = cfgSheet.getDataRange().getValues();
  var config = {};
  cfgValues.forEach(function(row) { config[String(row[0])] = row[1]; });

  // Movimientos
  var movSheet = ss.getSheetByName('Movimientos');
  var movValues = movSheet.getDataRange().getValues();
  var movimientos = [];

  if (movValues.length > 1) {
    var headers = movValues[0];
    for (var i = 1; i < movValues.length; i++) {
      var row = movValues[i];
      if (!row[0]) continue; // fila vacía
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = row[idx]; });

      var denoms = {};
      DENOMINACIONES.forEach(function(v) {
        denoms[v] = Number(obj['d' + v]) || 0;
      });

      movimientos.push({
        id:           String(obj.id),
        fecha:        String(obj.fecha),
        tipo:         String(obj.tipo),
        responsable:  String(obj.responsable),
        entregadoA:   String(obj.entregadoA || ''),
        descripcion:  String(obj.descripcion),
        monto:        Number(obj.monto) || 0,
        denominaciones: denoms
      });
    }
  }

  // Personas
  var perSheet = ss.getSheetByName('Personas');
  var perValues = perSheet.getDataRange().getValues();
  var personas = [];
  for (var j = 1; j < perValues.length; j++) {
    var nombre = String(perValues[j][0]).trim();
    if (nombre) personas.push(nombre);
  }

  return {
    ok: true,
    config: {
      saldoInicial:      Number(config.saldoInicial) || 0,
      efectivoPendiente: Number(config.efectivoPendiente) || 0,
      saldoExcel:        (config.saldoExcel === '' || config.saldoExcel === null) ? null : Number(config.saldoExcel),
      saldoSistema:      (config.saldoSistema === '' || config.saldoSistema === null) ? null : Number(config.saldoSistema)
    },
    movimientos: movimientos,
    personas:    personas
  };
}

// ----------------------------------------------------------------
// addMovimiento
// ----------------------------------------------------------------
function addMovimiento(ss, data) {
  var sheet = ss.getSheetByName('Movimientos');
  var row = [
    String(data.id),
    String(data.fecha),
    String(data.tipo),
    String(data.responsable),
    String(data.entregadoA || ''),
    String(data.descripcion),
    Number(data.monto) || 0
  ];
  DENOMINACIONES.forEach(function(v) {
    row.push(Number(data.denominaciones[v]) || 0);
  });
  sheet.appendRow(row);
  return { ok: true };
}

// ----------------------------------------------------------------
// deleteMovimiento
// ----------------------------------------------------------------
function deleteMovimiento(ss, id) {
  var sheet = ss.getSheetByName('Movimientos');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Movimiento no encontrado: ' + id };
}

// ----------------------------------------------------------------
// updateConfig — acepta uno o varios campos a la vez
// ----------------------------------------------------------------
function updateConfig(ss, data) {
  var sheet = ss.getSheetByName('Config');
  var values = sheet.getDataRange().getValues();
  var camposPermitidos = ['saldoInicial','efectivoPendiente','saldoExcel','saldoSistema'];

  values.forEach(function(row, i) {
    var key = String(row[0]);
    if (camposPermitidos.indexOf(key) !== -1 && data.hasOwnProperty(key)) {
      var val = data[key];
      // null/undefined → cadena vacía en la hoja
      sheet.getRange(i + 1, 2).setValue(val === null || val === undefined ? '' : val);
    }
  });
  return { ok: true };
}

// ----------------------------------------------------------------
// addPersona
// ----------------------------------------------------------------
function addPersona(ss, nombre) {
  var limpio = String(nombre).trim();
  if (!limpio) return { ok: false, error: 'Nombre vacío' };

  var sheet = ss.getSheetByName('Personas');
  var values = sheet.getDataRange().getValues();
  var existe = values.some(function(row) {
    return String(row[0]).trim().toLowerCase() === limpio.toLowerCase();
  });
  if (!existe) {
    sheet.appendRow([limpio]);
  }
  return { ok: true };
}
