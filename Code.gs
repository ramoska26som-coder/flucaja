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

  // Acciones de solo lectura: no necesitan bloqueo.
  if (action === 'getData') {
    return getData(ss);
  }
  if (action === 'getFirmas') {
    return getFirmas(ss, data);
  }

  // Escrituras: serializar con LockService para evitar duplicados por
  // ejecuciones concurrentes (reintentos que llegan casi al mismo tiempo).
  var lock = LockService.getScriptLock();
  var obtenido = false;
  try {
    obtenido = lock.tryLock(25000); // esperar hasta 25s por el turno
  } catch (e) {
    obtenido = false;
  }
  if (!obtenido) {
    return { ok: false, error: 'Servidor ocupado, reintenta en un momento.', _busy: true };
  }

  try {
    return despachar(ss, action, data);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function despachar(ss, action, data) {
  switch (action) {
    case 'addMovimiento':    return addMovimiento(ss, data);
    case 'deleteMovimiento': return deleteMovimiento(ss, data.id);
    case 'updateConfig':     return updateConfig(ss, data);
    case 'addPersona':       return addPersona(ss, data.nombre);
    case 'updateMovimiento':  return updateMovimiento(ss, data);
    case 'addArqueo':        return addArqueo(ss, data);
    case 'updateArqueoFirmas': return updateArqueoFirmas(ss, data);
    case 'deleteArqueo':     return deleteArqueo(ss, data.id);
    case 'addEtiqueta':      return addEtiqueta(ss, data.texto);
    case 'deleteEtiqueta':   return deleteEtiqueta(ss, data.texto);
    case 'addOrden':         return addOrden(ss, data);
    case 'updateOrden':      return updateOrden(ss, data);
    case 'deleteOrden':      return deleteOrden(ss, data.id);
    default:                 return { ok: false, error: 'Acción desconocida: ' + action };
  }
}

// Busca el número de fila (1-based) de un registro por su id en la columna A.
// Devuelve -1 si no existe. Reutilizable para idempotencia.
function buscarFilaPorId(sheet, id) {
  if (!id) return -1;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

// ----------------------------------------------------------------
// Inicializar hojas si no existen
// ----------------------------------------------------------------
function inicializarHojas(ss) {
  // Hoja Config
  if (!ss.getSheetByName('Config')) {
    var cfg = ss.insertSheet('Config');
    cfg.getRange('A1:B5').setValues([
      ['saldoInicial',      0],
      ['efectivoPendiente', 0],
      ['saldoExcel',        ''],
      ['saldoSistema',      ''],
      ['saldoMinimo',       '']
    ]);
    cfg.hideSheet();
  } else {
    // Migración: agregar claves de config que falten sin tocar las existentes
    asegurarClavesConfig(ss.getSheetByName('Config'), ['saldoMinimo']);
  }

  // Hoja Movimientos
  if (!ss.getSheetByName('Movimientos')) {
    var mov = ss.insertSheet('Movimientos');
    var headers = ['id','fecha','tipo','responsable','entregadoA','descripcion','monto',
                   'd500','d200','d100','d50','d20','d10','d5','d2','d1',
                   'firmaEntrega','firmaRecibe','esPendiente','origenPendiente','esExterno','cambioDe'];
    mov.getRange(1, 1, 1, headers.length).setValues([headers]);
    mov.setFrozenRows(1);
    // Formato moneda en columna monto (G)
    mov.getRange('G2:G').setNumberFormat('#,##0.00');
    // Forzar texto plano en columna fecha (B) para que Sheets no la reinterprete
    mov.getRange('B2:B').setNumberFormat('@');
  } else {
    // Migración: agregar columnas faltantes a hojas ya existentes (sin borrar datos)
    asegurarColumnas(ss.getSheetByName('Movimientos'),
      ['firmaEntrega','firmaRecibe','esPendiente','origenPendiente','esExterno','cambioDe']);
  }
  // Migrar Arqueos si ya existe
  if (ss.getSheetByName('Arqueos')) {
    asegurarColumnas(ss.getSheetByName('Arqueos'), ['pendientesLista']);
  }

  // Hoja Personas
  if (!ss.getSheetByName('Personas')) {
    var per = ss.insertSheet('Personas');
    per.getRange('A1').setValue('nombre');
    per.setFrozenRows(1);
  }

  // Hoja Etiquetas (descripciones rápidas)
  if (!ss.getSheetByName('Etiquetas')) {
    var et = ss.insertSheet('Etiquetas');
    et.getRange('A1').setValue('texto');
    et.setFrozenRows(1);
  }

  // Hoja OrdenesCompra
  if (!ss.getSheetByName('OrdenesCompra')) {
    var oc = ss.insertSheet('OrdenesCompra');
    var ocH = ['id','fecha','descripcion','monto','recibe','autorizador',
               'firmaAutoriza','estado','movimientoId','creadoEn'];
    oc.getRange(1, 1, 1, ocH.length).setValues([ocH]);
    oc.setFrozenRows(1);
    // Texto plano: fecha, estado y creadoEn (evita reinterpretación de Sheets)
    oc.getRange('B2:B').setNumberFormat('@'); // fecha
    oc.getRange('H2:H').setNumberFormat('@'); // estado
    oc.getRange('J2:J').setNumberFormat('@'); // creadoEn
    oc.getRange('D2:D').setNumberFormat('#,##0.00'); // monto
  } else {
    asegurarColumnas(ss.getSheetByName('OrdenesCompra'),
      ['id','fecha','descripcion','monto','recibe','autorizador',
       'firmaAutoriza','estado','movimientoId','creadoEn']);
  }

  // Hoja Arqueos
  if (!ss.getSheetByName('Arqueos')) {
    var arq = ss.insertSheet('Arqueos');
    var h = ['id','fecha','hora','responsable','verificador','totalContado',
             'efectivoPendiente','teorico','sistema','diferencia','notas',
             'denominaciones','firmaResponsable','firmaVerificador','creadoEn','pendientesLista'];
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

// Asegura que existan filas clave/valor en la hoja Config (migración no destructiva)
function asegurarClavesConfig(sheet, clavesRequeridas) {
  if (!sheet) return false;
  var values = sheet.getDataRange().getValues();
  var existentes = values.map(function(r){ return String(r[0]); });
  var faltantes = [];
  clavesRequeridas.forEach(function(k) {
    if (existentes.indexOf(k) === -1) faltantes.push([k, '']);
  });
  if (faltantes.length === 0) return false;
  sheet.getRange(values.length + 1, 1, faltantes.length, 2).setValues(faltantes);
  return true;
}

// ================================================================
// REPARACIÓN (correr UNA VEZ desde el editor de Apps Script)
// ----------------------------------------------------------------
// Arregla columnas duplicadas en la hoja Movimientos causadas por la
// tabla de Sheets que renombró encabezados a "Columna 3", "Columna 4",
// etc. Copia los datos de la columna vieja a la correcta (solo donde la
// correcta esté vacía) y elimina la columna vieja duplicada.
//
// CÓMO USARLA:
//   1) En el editor de Apps Script, selecciona "repararColumnasDuplicadas"
//      en el menú de funciones (arriba, junto a Ejecutar/Depurar).
//   2) Pulsa Ejecutar. Autoriza si lo pide.
//   3) Revisa el registro (Ver → Registros). Te dice qué hizo.
//   Es segura para correr más de una vez: si no hay duplicados, no hace nada.
// ----------------------------------------------------------------
function repararColumnasDuplicadas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Movimientos');
  if (!sheet) { Logger.log('No existe la hoja Movimientos.'); return; }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) {
    Logger.log('No se pudo obtener el lock. Reintenta en un momento.'); return;
  }

  try {
    // Estos son los encabezados "reales" y sus datos que un encabezado
    // genérico ("Columna N") pudo haber reemplazado. Reconocemos las
    // columnas viejas por su CONTENIDO característico.
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log('No hay datos que reparar.'); return; }

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    var datos = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Índice (0-based) de cada encabezado correcto
    function idxDe(nombre){ return headers.indexOf(nombre); }

    // Detecta una columna "huérfana" (encabezado genérico o vacío) cuyo
    // contenido corresponda a `tipoContenido`. Devuelve su índice o -1.
    //  tipoContenido: 'boolTF' (TRUE/FALSE) u 'origen' (banco/caja/vacío)
    function detectarHuerfana(tipoContenido, indicesExcluidos) {
      for (var c = 0; c < headers.length; c++) {
        if (indicesExcluidos.indexOf(c) !== -1) continue;
        var h = String(headers[c]).trim();
        var esGenerica = (h === '' || /^Columna\s*\d+$/i.test(h));
        if (!esGenerica) continue;
        // Revisar el contenido de la columna
        var tieneContenido = false, coincide = true;
        for (var r = 0; r < datos.length; r++) {
          var val = String(datos[r][c]).trim().toLowerCase();
          if (val === '') continue;
          tieneContenido = true;
          if (tipoContenido === 'boolTF') {
            if (val !== 'true' && val !== 'false') { coincide = false; break; }
          } else if (tipoContenido === 'origen') {
            if (val !== 'banco' && val !== 'caja') { coincide = false; break; }
          }
        }
        if (tieneContenido && coincide) return c;
      }
      return -1;
    }

    var reporte = [];
    var columnasABorrar = []; // índices 1-based, se borran al final en orden desc

    // --- Reparar esPendiente (TRUE/FALSE) ---
    var idxPend = idxDe('esPendiente');
    if (idxPend !== -1) {
      var huerfPend = detectarHuerfana('boolTF', [idxPend]);
      if (huerfPend !== -1) {
        var copiados = 0;
        for (var r = 0; r < datos.length; r++) {
          var destino = String(datos[r][idxPend]).trim();
          var origen = String(datos[r][huerfPend]).trim();
          // Copiar solo si el destino está vacío y el origen tiene algo
          if (destino === '' && origen !== '') {
            sheet.getRange(r + 2, idxPend + 1).setNumberFormat('@').setValue(origen.toLowerCase());
            copiados++;
          }
        }
        columnasABorrar.push(huerfPend + 1);
        reporte.push('esPendiente: copiados ' + copiados + ' valores desde la columna "' +
                     headers[huerfPend] + '" (col ' + (huerfPend + 1) + '), que se eliminará.');
      }
    }

    // --- Reparar origenPendiente (banco/caja) ---
    var idxOrig = idxDe('origenPendiente');
    if (idxOrig !== -1) {
      // Excluir la columna que ya marcamos para borrar, para no reusarla
      var excl = [idxOrig].concat(columnasABorrar.map(function(x){ return x - 1; }));
      var huerfOrig = detectarHuerfana('origen', excl);
      if (huerfOrig !== -1) {
        var copiadosO = 0;
        for (var r2 = 0; r2 < datos.length; r2++) {
          var destinoO = String(datos[r2][idxOrig]).trim();
          var origenO = String(datos[r2][huerfOrig]).trim();
          if (destinoO === '' && origenO !== '') {
            sheet.getRange(r2 + 2, idxOrig + 1).setNumberFormat('@').setValue(origenO.toLowerCase());
            copiadosO++;
          }
        }
        columnasABorrar.push(huerfOrig + 1);
        reporte.push('origenPendiente: copiados ' + copiadosO + ' valores desde la columna "' +
                     headers[huerfOrig] + '" (col ' + (huerfOrig + 1) + '), que se eliminará.');
      }
    }

    if (columnasABorrar.length === 0) {
      Logger.log('✓ No se detectaron columnas duplicadas. No se hizo ningún cambio.');
      return;
    }

    // Borrar de mayor a menor índice para no descuadrar posiciones
    columnasABorrar.sort(function(a, b){ return b - a; });
    columnasABorrar.forEach(function(colNum){
      sheet.deleteColumn(colNum);
    });

    Logger.log('REPARACIÓN COMPLETADA:\n' + reporte.join('\n') +
               '\n\nColumnas duplicadas eliminadas: ' + columnasABorrar.length +
               '.\nRecarga la app para confirmar que todo sigue correcto.');
  } finally {
    try { lock.releaseLock(); } catch (e) {}
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
        denominaciones: denoms,
        // Firmas NO se incluyen en getData (son base64 pesadas). Solo el indicador.
        tieneFirmaEntrega: String(obj.firmaEntrega || '') !== '',
        tieneFirmaRecibe:  String(obj.firmaRecibe || '') !== '',
        esPendiente:  (obj.esPendiente === true || String(obj.esPendiente).toLowerCase() === 'true'),
        origenPendiente: String(obj.origenPendiente || ''),
        esExterno:    (obj.esExterno === true || String(obj.esExterno).toLowerCase() === 'true'),
        cambioDe:     (function(){
                        var raw = String(obj.cambioDe || '').trim();
                        if (!raw) return [];
                        try { var arr = JSON.parse(raw); return Array.isArray(arr) ? arr.map(String) : []; }
                        catch(e){ return raw.split(',').map(function(s){ return s.trim(); }).filter(Boolean); }
                      })()
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
        tieneFirmaResponsable: String(aobj.firmaResponsable || '') !== '',
        tieneFirmaVerificador: String(aobj.firmaVerificador || '') !== '',
        creadoEn:          String(aobj.creadoEn || ''),
        pendientesLista:   (function(){ try{ return JSON.parse(aobj.pendientesLista || '[]'); }catch(e){ return []; } })()
      });
    }
  }

  // Etiquetas (descripciones rápidas)
  var etSheet = ss.getSheetByName('Etiquetas');
  var etValues = etSheet.getDataRange().getValues();
  var etiquetas = [];
  for (var e = 1; e < etValues.length; e++) {
    var txt = String(etValues[e][0]).trim();
    if (txt) etiquetas.push(txt);
  }

  // Órdenes de compra
  var ocSheet = ss.getSheetByName('OrdenesCompra');
  var ocValues = ocSheet.getDataRange().getValues();
  var ordenes = [];
  if (ocValues.length > 1) {
    var ocHeaders = ocValues[0];
    for (var o = 1; o < ocValues.length; o++) {
      var orow = ocValues[o];
      if (!orow[0]) continue;
      var oobj = {};
      ocHeaders.forEach(function(h, idx){ oobj[h] = orow[idx]; });
      ordenes.push({
        id:            String(oobj.id),
        fecha:         String(oobj.fecha || ''),
        descripcion:   String(oobj.descripcion || ''),
        monto:         Number(oobj.monto) || 0,
        recibe:        String(oobj.recibe || ''),
        autorizador:   String(oobj.autorizador || ''),
        tieneFirmaAutoriza: String(oobj.firmaAutoriza || '') !== '',
        estado:        String(oobj.estado || 'pendiente'),
        movimientoId:  String(oobj.movimientoId || ''),
        creadoEn:      String(oobj.creadoEn || '')
      });
    }
  }

  return {
    ok: true,
    config: {
      saldoInicial:      Number(config.saldoInicial) || 0,
      efectivoPendiente: Number(config.efectivoPendiente) || 0,
      saldoExcel:        (config.saldoExcel === '' || config.saldoExcel === null) ? null : Number(config.saldoExcel),
      saldoSistema:      (config.saldoSistema === '' || config.saldoSistema === null) ? null : Number(config.saldoSistema),
      saldoMinimo:       (config.saldoMinimo === '' || config.saldoMinimo === null || config.saldoMinimo === undefined) ? null : Number(config.saldoMinimo)
    },
    movimientos: movimientos,
    personas:    personas,
    arqueos:     arqueos,
    etiquetas:   etiquetas,
    ordenes:     ordenes
  };
}

// ----------------------------------------------------------------
// getFirmas — devuelve solo las firmas de UN registro (bajo demanda).
//  data = { tipo:'movimiento'|'arqueo'|'orden', id:'...' }
// Mantiene liviano el getData inicial.
// ----------------------------------------------------------------
function getFirmas(ss, data) {
  var tipo = data && data.tipo;
  var id = data && data.id;
  if (!tipo || !id) return { ok: false, error: 'Faltan tipo o id' };

  var mapa = {
    movimiento: { hoja: 'Movimientos',    campos: ['firmaEntrega','firmaRecibe'] },
    arqueo:     { hoja: 'Arqueos',        campos: ['firmaResponsable','firmaVerificador'] },
    orden:      { hoja: 'OrdenesCompra',  campos: ['firmaAutoriza'] }
  };
  var cfg = mapa[tipo];
  if (!cfg) return { ok: false, error: 'Tipo desconocido: ' + tipo };

  var sheet = ss.getSheetByName(cfg.hoja);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada' };
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(String);

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      var firmas = {};
      cfg.campos.forEach(function(c) {
        var idx = headers.indexOf(c);
        firmas[c] = (idx !== -1) ? String(values[i][idx] || '') : '';
      });
      return { ok: true, firmas: firmas };
    }
  }
  return { ok: false, error: 'Registro no encontrado: ' + id };
}
function addMovimiento(ss, data) {
  var sheet = ss.getSheetByName('Movimientos');

  // Idempotencia: si ya existe un movimiento con este id, no duplicar.
  // (Un reintento del cliente devuelve OK sin crear otra fila.)
  if (buscarFilaPorId(sheet, data.id) !== -1) {
    return { ok: true, _duplicado: true };
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var newRow = sheet.getLastRow() + 1;

  // Construir la fila alineada al header real (por nombre de columna)
  var valores = {};
  valores['id'] = String(data.id);
  valores['fecha'] = String(data.fecha);
  valores['tipo'] = String(data.tipo);
  valores['responsable'] = String(data.responsable);
  valores['entregadoA'] = String(data.entregadoA || '');
  valores['descripcion'] = String(data.descripcion);
  valores['monto'] = Number(data.monto) || 0;
  DENOMINACIONES.forEach(function(v){ valores['d'+v] = Number(data.denominaciones[v]) || 0; });
  valores['firmaEntrega'] = String(data.firmaEntrega || '');
  valores['firmaRecibe'] = String(data.firmaRecibe || '');
  valores['esPendiente'] = data.esPendiente ? 'true' : 'false';
  valores['origenPendiente'] = String(data.origenPendiente || '');
  valores['esExterno'] = data.esExterno ? 'true' : 'false';
  valores['cambioDe'] = (data.cambioDe && data.cambioDe.length) ? JSON.stringify(data.cambioDe) : '';

  var fila = headers.map(function(h){ return valores.hasOwnProperty(h) ? valores[h] : ''; });

  // Forzar texto en fecha, esPendiente, origenPendiente, esExterno y cambioDe
  ['fecha','esPendiente','origenPendiente','esExterno','cambioDe'].forEach(function(col){
    var idx = headers.indexOf(col);
    if(idx !== -1) sheet.getRange(newRow, idx+1, 1, 1).setNumberFormat('@');
  });

  sheet.getRange(newRow, 1, 1, fila.length).setValues([fila]);
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
  var camposPermitidos = ['saldoInicial','efectivoPendiente','saldoExcel','saldoSistema','saldoMinimo'];
  var clavesEnHoja = values.map(function(r){ return String(r[0]); });

  values.forEach(function(row, i) {
    var key = String(row[0]);
    if (camposPermitidos.indexOf(key) !== -1 && data.hasOwnProperty(key)) {
      var val = data[key];
      // null/undefined → cadena vacía en la hoja
      sheet.getRange(i + 1, 2).setValue(val === null || val === undefined ? '' : val);
    }
  });

  // Crear filas para claves permitidas que aún no existan en la hoja
  var nuevas = [];
  camposPermitidos.forEach(function(key) {
    if (data.hasOwnProperty(key) && clavesEnHoja.indexOf(key) === -1) {
      var v = data[key];
      nuevas.push([key, (v === null || v === undefined) ? '' : v]);
    }
  });
  if (nuevas.length) {
    sheet.getRange(values.length + 1, 1, nuevas.length, 2).setValues(nuevas);
  }
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

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);

  // Preservar las firmas existentes si el cliente manda vacío (evita borrarlas
  // por accidente, ya que las firmas no viajan en getData).
  var filaActual = values[rowIndex - 1];
  function firmaExistente(col){
    var idx = headers.indexOf(col);
    return idx !== -1 ? String(filaActual[idx] || '') : '';
  }
  var fEIn = String(data.firmaEntrega || '');
  var fRIn = String(data.firmaRecibe || '');
  if (fEIn === '') fEIn = firmaExistente('firmaEntrega');
  if (fRIn === '') fRIn = firmaExistente('firmaRecibe');

  var valores = {};
  valores['id'] = String(data.id);
  valores['fecha'] = String(data.fecha);
  valores['tipo'] = String(data.tipo);
  valores['responsable'] = String(data.responsable);
  valores['entregadoA'] = String(data.entregadoA || '');
  valores['descripcion'] = String(data.descripcion);
  valores['monto'] = Number(data.monto) || 0;
  DENOMINACIONES.forEach(function(v){ valores['d'+v] = Number(data.denominaciones[v]) || 0; });
  valores['firmaEntrega'] = fEIn;
  valores['firmaRecibe'] = fRIn;
  valores['esPendiente'] = data.esPendiente ? 'true' : 'false';
  valores['origenPendiente'] = String(data.origenPendiente || '');
  valores['esExterno'] = data.esExterno ? 'true' : 'false';
  valores['cambioDe'] = (data.cambioDe && data.cambioDe.length) ? JSON.stringify(data.cambioDe) : '';

  var fila = headers.map(function(h){ return valores.hasOwnProperty(h) ? valores[h] : ''; });

  ['fecha','esPendiente','origenPendiente','esExterno','cambioDe'].forEach(function(col){
    var idx = headers.indexOf(col);
    if(idx !== -1) sheet.getRange(rowIndex, idx+1, 1, 1).setNumberFormat('@');
  });

  sheet.getRange(rowIndex, 1, 1, fila.length).setValues([fila]);
  return { ok: true };
}

// ----------------------------------------------------------------
// addArqueo
// ----------------------------------------------------------------
function addArqueo(ss, data) {
  var sheet = ss.getSheetByName('Arqueos');
  if (buscarFilaPorId(sheet, data.id) !== -1) {
    return { ok: true, _duplicado: true };
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var newRow = sheet.getLastRow() + 1;

  var v = {};
  v['id'] = String(data.id);
  v['fecha'] = String(data.fecha);
  v['hora'] = String(data.hora || '');
  v['responsable'] = String(data.responsable || '');
  v['verificador'] = String(data.verificador || '');
  v['totalContado'] = Number(data.totalContado) || 0;
  v['efectivoPendiente'] = Number(data.efectivoPendiente) || 0;
  v['teorico'] = Number(data.teorico) || 0;
  v['sistema'] = Number(data.sistema) || 0;
  v['diferencia'] = Number(data.diferencia) || 0;
  v['notas'] = String(data.notas || '');
  v['denominaciones'] = JSON.stringify(data.denominaciones || []);
  v['firmaResponsable'] = String(data.firmaResponsable || '');
  v['firmaVerificador'] = String(data.firmaVerificador || '');
  v['creadoEn'] = String(data.creadoEn || new Date().toISOString());
  v['pendientesLista'] = JSON.stringify(data.pendientesLista || []);

  var fila = headers.map(function(h){ return v.hasOwnProperty(h) ? v[h] : ''; });
  // Texto plano en fecha, hora y creadoEn
  ['fecha','hora','creadoEn'].forEach(function(col){
    var idx = headers.indexOf(col);
    if(idx !== -1) sheet.getRange(newRow, idx+1, 1, 1).setNumberFormat('@');
  });
  sheet.getRange(newRow, 1, 1, fila.length).setValues([fila]);
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
// addEtiqueta / deleteEtiqueta — descripciones rápidas
// ----------------------------------------------------------------
function addEtiqueta(ss, texto) {
  var limpio = String(texto).trim();
  if (!limpio) return { ok: false, error: 'Etiqueta vacía' };
  var sheet = ss.getSheetByName('Etiquetas');
  var values = sheet.getDataRange().getValues();
  var existe = values.some(function(row) {
    return String(row[0]).trim().toLowerCase() === limpio.toLowerCase();
  });
  if (!existe) sheet.appendRow([limpio]);
  return { ok: true };
}

function deleteEtiqueta(ss, texto) {
  var limpio = String(texto).trim();
  var sheet = ss.getSheetByName('Etiquetas');
  var values = sheet.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]).trim().toLowerCase() === limpio.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

// ----------------------------------------------------------------
// Órdenes de compra — construir fila alineada por nombre de columna
// ----------------------------------------------------------------
function ordenValores(data) {
  var v = {};
  v['id'] = String(data.id);
  v['fecha'] = String(data.fecha || '');
  v['descripcion'] = String(data.descripcion || '');
  v['monto'] = Number(data.monto) || 0;
  v['recibe'] = String(data.recibe || '');
  v['autorizador'] = String(data.autorizador || '');
  v['firmaAutoriza'] = String(data.firmaAutoriza || '');
  v['estado'] = String(data.estado || 'pendiente');
  v['movimientoId'] = String(data.movimientoId || '');
  v['creadoEn'] = String(data.creadoEn || ahoraLocalISO_gs());
  return v;
}

// Timestamp local con offset (equivalente al del frontend). Evita
// depender de toISOString() que devuelve UTC.
function ahoraLocalISO_gs() {
  var d = new Date();
  var off = -d.getTimezoneOffset();
  var signo = off >= 0 ? '+' : '-';
  var abs = Math.abs(off);
  return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) +
         'T' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds()) +
         signo + pad2(Math.floor(abs / 60)) + ':' + pad2(abs % 60);
}

function addOrden(ss, data) {
  var sheet = ss.getSheetByName('OrdenesCompra');
  if (buscarFilaPorId(sheet, data.id) !== -1) {
    return { ok: true, _duplicado: true };
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var newRow = sheet.getLastRow() + 1;
  var v = ordenValores(data);
  var fila = headers.map(function(h){ return v.hasOwnProperty(h) ? v[h] : ''; });
  ['fecha','estado','creadoEn','movimientoId'].forEach(function(col){
    var idx = headers.indexOf(col);
    if (idx !== -1) sheet.getRange(newRow, idx + 1, 1, 1).setNumberFormat('@');
  });
  sheet.getRange(newRow, 1, 1, fila.length).setValues([fila]);
  return { ok: true };
}

function updateOrden(ss, data) {
  var sheet = ss.getSheetByName('OrdenesCompra');
  var values = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.id)) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) return { ok: false, error: 'Orden no encontrada: ' + data.id };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var v = ordenValores(data);
  // Conservar creadoEn original si no viene en data
  if (!data.creadoEn) {
    var idxC = headers.indexOf('creadoEn');
    if (idxC !== -1) v['creadoEn'] = String(values[rowIndex - 1][idxC] || v['creadoEn']);
  }
  // Conservar la firma existente si el cliente manda vacío (no viaja en getData)
  if (!v['firmaAutoriza']) {
    var idxF = headers.indexOf('firmaAutoriza');
    if (idxF !== -1) v['firmaAutoriza'] = String(values[rowIndex - 1][idxF] || '');
  }
  var fila = headers.map(function(h){ return v.hasOwnProperty(h) ? v[h] : ''; });
  ['fecha','estado','creadoEn','movimientoId'].forEach(function(col){
    var idx = headers.indexOf(col);
    if (idx !== -1) sheet.getRange(rowIndex, idx + 1, 1, 1).setNumberFormat('@');
  });
  sheet.getRange(rowIndex, 1, 1, fila.length).setValues([fila]);
  return { ok: true };
}

function deleteOrden(ss, id) {
  var sheet = ss.getSheetByName('OrdenesCompra');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Orden no encontrada: ' + id };
}
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
