// ================================================================
// FLUJO DE CAJA — Google Apps Script Backend
// Vincular a una Google Sheet antes de desplegar.
// Instrucciones completas en README.md
// ================================================================

var DENOMINACIONES = [500, 200, 100, 50, 20, 10, 5, 2, 1];

// ----------------------------------------------------------------
// Punto de entrada GET (lecturas y peticiones pequeñas, ?payload=JSON)
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

// ----------------------------------------------------------------
// Punto de entrada POST (peticiones grandes: arqueos con firmas base64)
// El cuerpo se envía como texto plano para evitar el preflight CORS.
// ----------------------------------------------------------------
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{"action":"getData"}';
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
    case 'updateMovimiento':  return updateMovimiento(ss, data);
    case 'addArqueo':        return addArqueo(ss, data);
    case 'updateArqueoFirmas': return updateArqueoFirmas(ss, data);
    case 'deleteArqueo':     return deleteArqueo(ss, data.id);
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
                   'd500','d200','d100','d50','d20','d10','d5','d2','d1',
                   'firmaEntrega','firmaRecibe','esPendiente','origenPendiente'];
    mov.getRange(1, 1, 1, headers.length).setValues([headers]);
    mov.setFrozenRows(1);
    // Formato moneda en columna monto (G)
    mov.getRange('G2:G').setNumberFormat('#,##0.00');
    // Forzar texto plano en columna fecha (B) para que Sheets no la reinterprete
    mov.getRange('B2:B').setNumberFormat('@');
  } else {
    // Migración: agregar columnas faltantes a hojas ya existentes (sin borrar datos)
    asegurarColumnas(ss.getSheetByName('Movimientos'),
      ['firmaEntrega','firmaRecibe','esPendiente','origenPendiente']);
  }

  // Hoja Personas
  if (!ss.getSheetByName('Personas')) {
    var per = ss.insertSheet('Personas');
    per.getRange('A1').setValue('nombre');
    per.setFrozenRows(1);
  }

  // Hoja Arqueos
  if (!ss.getSheetByName('Arqueos')) {
    var arq = ss.insertSheet('Arqueos');
    var h = ['id','fecha','hora','responsable','verificador','totalContado',
             'efectivoPendiente','teorico','sistema','diferencia','notas',
             'denominaciones','firmaResponsable','firmaVerificador','creadoEn'];
    arq.getRange(1, 1, 1, h.length).setValues([h]);
    arq.setFrozenRows(1);
    // Forzar texto plano en columnas que Sheets intentaría convertir a fecha/hora
    // B=fecha, C=hora, O=creadoEn  (evita corrupción al leer desde otro dispositivo)
    arq.getRange('B2:C').setNumberFormat('@');
    arq.getRange('O2:O').setNumberFormat('@');
  }
}

// ----------------------------------------------------------------
// asegurarColumnas — agrega al final las columnas que falten en el
// header de una hoja existente, sin borrar datos. Devuelve true si
// hubo cambios.
// ----------------------------------------------------------------
function asegurarColumnas(sheet, columnasRequeridas) {
  if (!sheet) return false;
  var lastCol = sheet.getLastColumn();
  var headerRange = sheet.getRange(1, 1, 1, Math.max(lastCol, 1));
  var headers = headerRange.getValues()[0].map(function(x){ return String(x); });
  var faltantes = [];
  columnasRequeridas.forEach(function(col) {
    if (headers.indexOf(col) === -1) faltantes.push(col);
  });
  if (faltantes.length === 0) return false;
  // Escribir las columnas faltantes a continuación del header actual
  sheet.getRange(1, lastCol + 1, 1, faltantes.length).setValues([faltantes]);
  return true;
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
        denominaciones: denoms,
        firmaEntrega: String(obj.firmaEntrega || ''),
        firmaRecibe:  String(obj.firmaRecibe || ''),
        esPendiente:  String(obj.esPendiente || '') === 'true' || obj.esPendiente === true,
        origenPendiente: String(obj.origenPendiente || '')
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

  // Arqueos
  var arqSheet = ss.getSheetByName('Arqueos');
  var arqValues = arqSheet.getDataRange().getValues();
  var arqueos = [];
  if (arqValues.length > 1) {
    var aHeaders = arqValues[0];
    for (var k = 1; k < arqValues.length; k++) {
      var arow = arqValues[k];
      if (!arow[0]) continue;
      var aobj = {};
      aHeaders.forEach(function(h, idx){ aobj[h] = arow[idx]; });
      var denoms;
      try { denoms = JSON.parse(aobj.denominaciones || '[]'); } catch(e) { denoms = []; }
      arqueos.push({
        id:                String(aobj.id),
        fecha:             normalizarFecha(aobj.fecha),
        hora:              normalizarHora(aobj.hora),
        responsable:       String(aobj.responsable || ''),
        verificador:       String(aobj.verificador || ''),
        totalContado:      Number(aobj.totalContado) || 0,
        efectivoPendiente: Number(aobj.efectivoPendiente) || 0,
        teorico:           Number(aobj.teorico) || 0,
        sistema:           Number(aobj.sistema) || 0,
        diferencia:        Number(aobj.diferencia) || 0,
        notas:             String(aobj.notas || ''),
        denominaciones:    denoms,
        firmaResponsable:  String(aobj.firmaResponsable || ''),
        firmaVerificador:  String(aobj.firmaVerificador || ''),
        creadoEn:          String(aobj.creadoEn || '')
      });
    }
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
    personas:    personas,
    arqueos:     arqueos
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
  // Guardar firmas de entrega/recibo si vienen (columnas extra R,S)
  row.push(String(data.firmaEntrega || ''));
  row.push(String(data.firmaRecibe || ''));
  row.push(data.esPendiente ? 'true' : 'false');
  row.push(String(data.origenPendiente || ''));
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 2, 1, 1).setNumberFormat('@'); // B = fecha como texto
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
// updateMovimiento — reemplaza la fila existente en Sheets
// ----------------------------------------------------------------
function updateMovimiento(ss, data) {
  var sheet = ss.getSheetByName('Movimientos');
  var values = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.id)) {
      rowIndex = i + 1; // 1-based para Sheets
      break;
    }
  }
  if (rowIndex === -1) return { ok: false, error: 'Movimiento no encontrado: ' + data.id };

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
  row.push(String(data.firmaEntrega || ''));
  row.push(String(data.firmaRecibe || ''));
  row.push(data.esPendiente ? 'true' : 'false');
  row.push(String(data.origenPendiente || ''));
  sheet.getRange(rowIndex, 2, 1, 1).setNumberFormat('@'); // fecha como texto
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  return { ok: true };
}

// ----------------------------------------------------------------
// addArqueo
// ----------------------------------------------------------------
function addArqueo(ss, data) {
  var sheet = ss.getSheetByName('Arqueos');
  var row = [
    String(data.id),
    String(data.fecha),
    String(data.hora || ''),
    String(data.responsable || ''),
    String(data.verificador || ''),
    Number(data.totalContado) || 0,
    Number(data.efectivoPendiente) || 0,
    Number(data.teorico) || 0,
    Number(data.sistema) || 0,
    Number(data.diferencia) || 0,
    String(data.notas || ''),
    JSON.stringify(data.denominaciones || []),
    String(data.firmaResponsable || ''),
    String(data.firmaVerificador || ''),
    String(data.creadoEn || new Date().toISOString())
  ];
  // Asegurar texto plano en fecha/hora/creadoEn antes de escribir
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 2, 1, 2).setNumberFormat('@'); // B,C = fecha,hora
  sheet.getRange(newRow, 15, 1, 1).setNumberFormat('@'); // O = creadoEn
  sheet.appendRow(row);
  return { ok: true };
}

// ----------------------------------------------------------------
// updateArqueoFirmas — actualiza solo las firmas de un arqueo existente
// ----------------------------------------------------------------
function updateArqueoFirmas(ss, data) {
  var sheet = ss.getSheetByName('Arqueos');
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var colFR = headers.indexOf('firmaResponsable');
  var colFV = headers.indexOf('firmaVerificador');
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.id)) {
      if (colFR !== -1) sheet.getRange(i + 1, colFR + 1).setValue(String(data.firmaResponsable || ''));
      if (colFV !== -1) sheet.getRange(i + 1, colFV + 1).setValue(String(data.firmaVerificador || ''));
      return { ok: true };
    }
  }
  return { ok: false, error: 'Arqueo no encontrado: ' + data.id };
}

// ----------------------------------------------------------------
// deleteArqueo
// ----------------------------------------------------------------
function deleteArqueo(ss, id) {
  var sheet = ss.getSheetByName('Arqueos');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Arqueo no encontrado: ' + id };
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

// ----------------------------------------------------------------
// Normalización de fecha/hora — repara valores que Sheets convirtió
// a objetos Date. Devuelve siempre texto en formato esperado por el
// frontend: fecha "YYYY-MM-DD", hora "HH:MM".
// ----------------------------------------------------------------
function pad2(n){ return ('0' + n).slice(-2); }

function normalizarFecha(val) {
  if (val === null || val === undefined || val === '') return '';
  // Si ya es texto YYYY-MM-DD, devolver tal cual
  if (typeof val === 'string') {
    var m = val.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    // Otro texto: intentar parsear como fecha
    var d2 = new Date(val);
    if (!isNaN(d2.getTime()) && d2.getFullYear() > 1900) {
      return d2.getFullYear() + '-' + pad2(d2.getMonth() + 1) + '-' + pad2(d2.getDate());
    }
    return String(val);
  }
  // Si es un objeto Date (lo que causaba el bug)
  if (Object.prototype.toString.call(val) === '[object Date]') {
    if (val.getFullYear() > 1900) {
      return val.getFullYear() + '-' + pad2(val.getMonth() + 1) + '-' + pad2(val.getDate());
    }
    return '';
  }
  return String(val);
}

function normalizarHora(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'string') {
    var m = val.match(/(\d{1,2}):(\d{2})/);
    if (m) return pad2(parseInt(m[1], 10)) + ':' + m[2];
    return String(val);
  }
  // Si es un objeto Date (hora con fecha base 1899)
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return pad2(val.getHours()) + ':' + pad2(val.getMinutes());
  }
  return String(val);
}
