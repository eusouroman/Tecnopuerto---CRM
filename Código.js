// TECNOPUERTO CRM — Código.gs
// Reemplazá el SHEET_ID con el ID de tu Google Sheet

const SHEET_ID = '1X7d94VCxflut0joYaSie4t1C1TxkW8vbUmRBrhI6wHY';

const HOJAS = { stock: 'Stock', ventas: 'Ventas', reparaciones: 'Reparaciones', gastos: 'Gastos', ventasBorradas: 'VentasBorradas', cierres: 'CierresCaja' };
const HEADERS = {
  stock:        ['id', 'nombre', 'cantidad', 'costo', 'precioVenta', 'stockMinimo'],
  ventas:       ['id', 'fecha', 'hora', 'nombre', 'cantidad', 'precioVenta', 'costoUnit', 'medio', 'medio1', 'montoMedio1', 'medio2', 'montoMedio2',
                 'cliente', 'clienteWA', 'saldoPendiente', 'fechaRecordatorio', 'cobradoSaldo', 'fechaCobroSaldo', 'medioSaldo'],
  reparaciones: ['id', 'fecha', 'cliente', 'whatsapp', 'equipo', 'problema',
                 'precioReparacion', 'costoRepuesto', 'tecnico', 'estado', 'notas', 'medio', 'medio1', 'montoMedio1', 'medio2', 'montoMedio2', 'avisado', 'fechaEntrega', 'fechaListo'],
  gastos:       ['id', 'fecha', 'categoria', 'detalle', 'monto', 'medio'],
  ventasBorradas: ['id', 'fecha', 'hora', 'nombre', 'cantidad', 'precioVenta', 'costoUnit', 'medio', 'medio1', 'montoMedio1', 'medio2', 'montoMedio2', 'motivo', 'fechaBorrado'],
  cierres:      ['id', 'fecha', 'efectivo', 'transferencia', 'debito', 'credito', 'pagoQR', 'otros', 'total', 'cerradoEn']
};

function getSheet(nombre) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    const key = Object.keys(HOJAS).find(function(k) { return HOJAS[k] === nombre; });
    if (key) hoja.appendRow(HEADERS[key]);
  }
  return hoja;
}

function sheetToArray(hoja) {
  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) {
      var val = row[i];
      if (val instanceof Date) {
        // La columna 'hora' la guarda Sheets como hora del dia (fecha base 1899-12-30);
        // hay que formatearla como HH:mm, no como fecha, si no se pierde la hora y se rompe el orden.
        val = (h === 'hora')
          ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm')
          : Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[h] = val;
    });
    return obj;
  });
}

function jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function doGet(e) {
  const accion = e.parameter.accion;
  if (accion === 'leer') {
    const hoja = e.parameter.hoja;
    const sheet = getSheet(HOJAS[hoja]);
    return jsonResponse({ ok: true, data: sheetToArray(sheet) });
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Tecnopuerto CRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const accion = body.accion;

  if (accion === 'agregar') {
    const sheet = getSheet(HOJAS[body.hoja]);
    // Alinear con el encabezado REAL de la hoja y crear columnas faltantes (ej: 'hora')
    var lastCol = sheet.getLastColumn();
    var heads = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h); }) : [];
    HEADERS[body.hoja].forEach(function(h) {
      if (heads.indexOf(h) === -1) { heads.push(h); sheet.getRange(1, heads.length).setValue(h); }
    });
    var fila = heads.map(function(h) { return body.item[h] !== undefined ? body.item[h] : ''; });
    sheet.appendRow(fila);
    return jsonResponse({ ok: true });
  }

  if (accion === 'actualizar') {
    const sheet = getSheet(HOJAS[body.hoja]);
    const data = sheet.getDataRange().getValues();
    const heads = data[0];
    const idIdx = heads.indexOf('id');
    let camIdx = heads.indexOf(body.campo);
    // Si la columna no existe, crearla automáticamente
    if (camIdx === -1) {
      camIdx = heads.length;
      sheet.getRange(1, camIdx + 1).setValue(body.campo);
    }
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(body.id)) {
        sheet.getRange(i + 1, camIdx + 1).setValue(body.valor);
        break;
      }
    }
    return jsonResponse({ ok: true });
  }

  if (accion === 'importarStock') {
    const sheet = getSheet(HOJAS.stock);
    const data = sheet.getDataRange().getValues();
    const heads = data[0];
    const iNombre = heads.indexOf('nombre');
    const iCant = heads.indexOf('cantidad');
    const iCosto = heads.indexOf('costo');
    const reemplazar = body.reemplazar === true;

    // Índice de existentes para búsqueda rápida
    var existentes = {};
    for (var i = 1; i < data.length; i++) {
      existentes[String(data[i][iNombre]).toUpperCase()] = i;
    }

    var nuevasFilas = [];
    var huboActualizaciones = false;

    body.items.forEach(function(item) {
      var key = String(item.nombre).toUpperCase();
      if (existentes[key] !== undefined) {
        var fila = existentes[key];
        var nuevaCant = reemplazar ? item.cantidad : (Number(data[fila][iCant]) || 0) + item.cantidad;
        data[fila][iCant] = nuevaCant;   // modificar en memoria
        data[fila][iCosto] = item.costo; // modificar en memoria
        huboActualizaciones = true;
      } else {
        nuevasFilas.push(HEADERS.stock.map(function(h) { return item[h] !== undefined ? item[h] : ''; }));
      }
    });

    // Una sola escritura para todas las actualizaciones
    if (huboActualizaciones && data.length > 1) {
      sheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
    }

    // Insertar todos los nuevos de una sola vez
    if (nuevasFilas.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
    }

    return jsonResponse({ ok: true, importados: body.items.length });
  }
  if (accion === 'importarPrecios') {
    // Actualiza precioVenta por nombre en la hoja Stock (matchea contra lo que haya EN VIVO en la planilla)
    const sheet = getSheet(HOJAS.stock);
    const data = sheet.getDataRange().getValues();
    const heads = data[0];
    const iNombre = heads.indexOf('nombre');
    let iPV = heads.indexOf('precioVenta');
    if (iPV === -1) { iPV = heads.length; sheet.getRange(1, iPV + 1).setValue('precioVenta'); }

    var indice = {};
    for (var i = 1; i < data.length; i++) indice[String(data[i][iNombre]).toUpperCase().trim()] = i;

    var actualizados = 0;
    body.items.forEach(function(item) {
      var fila = indice[String(item.nombre).toUpperCase().trim()];
      if (fila !== undefined) {
        while (data[fila].length <= iPV) data[fila].push('');
        data[fila][iPV] = item.precioVenta;
        actualizados++;
      }
    });

    if (actualizados > 0) {
      var ancho = iPV + 1;
      sheet.getRange(2, 1, data.length - 1, ancho).setValues(data.slice(1).map(function(r) {
        while (r.length < ancho) r.push('');
        return r;
      }));
    }

    return jsonResponse({ ok: true, actualizados: actualizados, total: body.items.length });
  }

  if (accion === 'borrar') {
    const sheet = getSheet(HOJAS[body.hoja]);
    const data = sheet.getDataRange().getValues();
    const idIdx = data[0].indexOf('id');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(body.id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return jsonResponse({ ok: true });
  }

  // Borrado LOGICO de venta: la mueve a la hoja VentasBorradas (queda el historial)
  if (accion === 'borrarVenta') {
    const sheet = getSheet(HOJAS.ventas);
    const data = sheet.getDataRange().getValues();
    const heads = data[0];
    const idIdx = heads.indexOf('id');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(body.id)) {
        var obj = {};
        heads.forEach(function(h, j) { obj[h] = data[i][j]; });
        obj.motivo = body.motivo || '';
        obj.fechaBorrado = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        var bsheet = getSheet(HOJAS.ventasBorradas);
        bsheet.appendRow(HEADERS.ventasBorradas.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; }));
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return jsonResponse({ ok: true });
  }

  if (accion === 'actualizarStockPorNombre') {
    const sheet = getSheet(HOJAS.stock);
    const data = sheet.getDataRange().getValues();
    const heads = data[0];
    const iNombre = heads.indexOf('nombre');
    const iCant = heads.indexOf('cantidad');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iNombre]).toUpperCase().trim() === String(body.nombre).toUpperCase().trim()) {
        sheet.getRange(i + 1, iCant + 1).setValue(body.valor);
        break;
      }
    }
    return jsonResponse({ ok: true });
  }
  return jsonResponse({ ok: false, error: 'Accion no reconocida' });
}

// ── BACKUP SEMANAL AUTOMÁTICO ──────────────────────────────────
// Configurá un trigger semanal apuntando a esta función (ver instrucciones)
function backupSemanal() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var nombreCopia = 'Backup_Tecnopuerto_' + fecha;

  // Buscar o crear carpeta "Backups TecnoPuerto" en Drive
  var carpeta;
  var carpetas = DriveApp.getFoldersByName('Backups TecnoPuerto');
  carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder('Backups TecnoPuerto');

  // Copiar el spreadsheet completo y moverlo a la carpeta
  var copia = ss.copy(nombreCopia);
  DriveApp.getFileById(copia.getId()).moveTo(carpeta);

  // Limpieza: conservar solo los últimos 14 backups
  limpiarBackupsViejos(carpeta, 14);

  Logger.log('Backup creado: ' + nombreCopia);
}

// Borra los backups más viejos, deja solo los N más recientes
function limpiarBackupsViejos(carpeta, conservar) {
  var archivos = [];
  var it = carpeta.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf('Backup_Tecnopuerto_') === 0) archivos.push(f);
  }
  archivos.sort(function(a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (var i = conservar; i < archivos.length; i++) {
    archivos[i].setTrashed(true);
  }
}

// ── REPORTE MENSUAL AUTOMÁTICO POR EMAIL ────────────────────────
// El día 1 de cada mes manda un resumen del mes recién cerrado (ventas +
// reparaciones + gastos + ganancia neta) al mail de la cuenta dueña del script.
function calcularReporte(desdeStr, hastaStr) {
  var ventas = sheetToArray(getSheet(HOJAS.ventas));
  var reparaciones = sheetToArray(getSheet(HOJAS.reparaciones));
  var gastos = sheetToArray(getSheet(HOJAS.gastos));

  function enRango(f) { return f && f >= desdeStr && f <= hastaStr; }

  var vF = ventas.filter(function(v) { return enRango(v.fecha); });
  var tV = vF.reduce(function(s, v) { return s + (parseFloat(v.precioVenta) || 0) * (parseInt(v.cantidad) || 0); }, 0);
  var cV = vF.reduce(function(s, v) { return s + (parseFloat(v.costoUnit) || 0) * (parseInt(v.cantidad) || 0); }, 0);
  var gV = tV - cV;

  var rEnt = reparaciones.filter(function(r) { return r.estado === 'Entregado' && enRango(r.fechaEntrega || r.fecha); });
  var tR = rEnt.reduce(function(s, r) { return s + (parseFloat(r.precioReparacion) || 0); }, 0);
  var cR = rEnt.reduce(function(s, r) { return s + (parseFloat(r.costoRepuesto) || 0); }, 0);
  var gR = tR - cR;

  var gastosF = gastos.filter(function(g) { return enRango(g.fecha) && g.categoria !== 'Mercadería/Stock'; });
  var totGastos = gastosF.reduce(function(s, g) { return s + (parseFloat(g.monto) || 0); }, 0);

  var facturado = tV + tR;
  var brutaTotal = gV + gR; // ventas + reparaciones, igual que en Estadísticas (ver crm-ganancia-neta-fix)
  var neta = brutaTotal - totGastos;
  var ops = vF.length + rEnt.length;
  var ticket = ops > 0 ? facturado / ops : 0;

  return {
    facturado: facturado, brutaTotal: brutaTotal, gV: gV, gR: gR,
    reparacionesCant: rEnt.length, totGastos: totGastos, neta: neta,
    ticket: ticket, ops: ops
  };
}

function fmARS(n) {
  return '$' + Math.round(n || 0).toLocaleString('es-AR');
}

function enviarReporteMensual() {
  var hoy = new Date();
  var inicioMesAnt = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  var finMesAnt = new Date(hoy.getFullYear(), hoy.getMonth(), 0);
  var desdeStr = Utilities.formatDate(inicioMesAnt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var hastaStr = Utilities.formatDate(finMesAnt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var nombreMes = Utilities.formatDate(inicioMesAnt, Session.getScriptTimeZone(), 'MMMM yyyy');

  var r = calcularReporte(desdeStr, hastaStr);

  var html =
    '<h2>Tecnopuerto — Reporte de ' + nombreMes + '</h2>' +
    '<table cellpadding="6" style="border-collapse:collapse;font-family:sans-serif">' +
    '<tr><td>Total facturado</td><td><b>' + fmARS(r.facturado) + '</b></td></tr>' +
    '<tr><td>Ganancia bruta (ventas + reparaciones)</td><td><b>' + fmARS(r.brutaTotal) + '</b></td></tr>' +
    '<tr><td>Reparaciones entregadas</td><td>' + r.reparacionesCant + ' — ganancia ' + fmARS(r.gR) + '</td></tr>' +
    '<tr><td>Gastos operativos</td><td>' + fmARS(r.totGastos) + '</td></tr>' +
    '<tr><td>Ganancia neta</td><td><b>' + fmARS(r.neta) + '</b></td></tr>' +
    '<tr><td>Ticket promedio</td><td>' + fmARS(r.ticket) + ' (' + r.ops + ' operaciones)</td></tr>' +
    '</table>' +
    '<p style="color:#888;font-size:12px">Generado automáticamente el ' + Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '.</p>';

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'Tecnopuerto — Reporte de ' + nombreMes,
    htmlBody: html
  });
}

// ── EJECUTAR UNA SOLA VEZ ───────────────────────────────────────
// Corré esta función UNA vez desde el editor de Apps Script.
// Instala el envío automático del reporte mensual: el día 1 de cada mes a
// las 8 AM, manda por mail el resumen del mes recién cerrado.
function instalarReporteMensualAutomatico() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'enviarReporteMensual') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('enviarReporteMensual')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();
  Logger.log('✓ Reporte mensual automático instalado. Se manda el día 1 de cada mes a las 8 AM.');
}

// ── EJECUTAR UNA SOLA VEZ ───────────────────────────────────────
// Corré esta función UNA vez desde el editor de Apps Script.
// Instala el backup automático diario (entre 2 y 3 AM) y ya está:
// después corre solo todos los días, sin tocar nada.
function instalarBackupAutomatico() {
  // Evita duplicar el disparador si ya existe
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'backupSemanal') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('backupSemanal')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  Logger.log('✓ Backup automático diario instalado. Corre todas las noches ~2 AM.');
}