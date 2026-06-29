/***** =========================================================
 * CONTABILIDAD.GS — Sistema de Contabilidad Vicelly Sánchez Studio
 * ---------------------------------------------------------------
 * Backend único para el nuevo frontend (agenda-contabilidad).
 * Proyecto de Apps Script SEPARADO del de la agenda de citas.
 *
 * Responsabilidades:
 *   - Registrar ventas (reemplaza el Google Form)
 *   - Calcular comisión por tipo de servicio
 *   - Registrar adelantos
 *   - Generar resumen de dashboard (con filtros de fecha/colaborador)
 *   - Generar y enviar comprobante de pago en PDF (sin archivar en Drive)
 *   - Cerrar quincena (archiva en hoja oculta + limpia Respuestas/Adelantos)
 *
 * Version: 2026.06 — reconstrucción limpia
 * ========================================================= *****/

const SHEET_ID = '1q3cX8fW7rCjaDPKDLoq6tf82_ReLh7r5GWBqPg01KTU';
const LOGO_FILE_ID = '1h32RqyFZfQj0_hYaLFWdFO6KHAIkGUU5';

const HOJA_RESPUESTAS = 'Respuestas';
const HOJA_ADELANTOS = 'Adelantos';
const HOJA_PROPINAS = 'Propinas';
const HOJA_CATALOGO = 'CatalogoServicios';
const HOJA_CONFIG = 'Config';
const HOJA_COLABORADORES = 'Colaboradores';

/***** ============== ROUTER PRINCIPAL ============== *****/
function doGet(e) {
  var action = e.parameter.action || '';
  if (action === 'getCatalogoServicios') return getCatalogoServicios();
  if (action === 'getColaboradores') return getColaboradores();
  if (action === 'getDashboardResumen') return getDashboardResumen(e);
  if (action === 'getDetalleServicios') return getDetalleServicios(e);
  if (action === 'getComisionesPorColaborador') return getComisionesPorColaborador();
  if (action === 'getPropinasPendientes') return getPropinasPendientes(e);
  return jsonOutput_({ ok: false, error: 'Accion no reconocida' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'registrarVenta') return registrarVenta(body);
    if (body.action === 'registrarAdelanto') return registrarAdelanto(body);
    if (body.action === 'registrarPropina') return registrarPropina(body);
    if (body.action === 'marcarPropinaPagada') return marcarPropinaPagada(body);
    if (body.action === 'generarComprobante') return generarComprobante(body);
    if (body.action === 'cerrarQuincena') return cerrarQuincena(body);
    return jsonOutput_({ ok: false, error: 'Accion no reconocida' });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

/***** ============== UTILIDADES BASE ============== *****/
function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(nombre) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(nombre);
  if (!sh) throw new Error('No existe la hoja "' + nombre + '".');
  return sh;
}

function normalizar_(texto) {
  return String(texto || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().trim();
}

function formatoFechaLocal_(fecha) {
  var d = fecha instanceof Date ? fecha : new Date(fecha);
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var dia = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + dia;
}

function buildPeriodo_(fecha) {
  var d = fecha instanceof Date ? fecha : new Date(fecha);
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var q = d.getDate() <= 15 ? 'Q1' : 'Q2';
  return y + '-' + m + ' ' + q;
}

function getConfigMap_() {
  var sh = getSheet_(HOJA_CONFIG);
  var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  var map = {};
  vals.forEach(function (row) {
    if (row[0]) map[String(row[0]).trim()] = String(row[1]).trim();
  });
  return map;
}

function getColaboradoresMap_() {
  var sh = getSheet_(HOJA_COLABORADORES);
  var map = {};
  var last = sh.getLastRow();
  if (last < 2) return map;
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  vals.forEach(function (row) {
    if (row[0] && row[1]) map[String(row[0]).trim()] = String(row[1]).trim();
  });
  return map;
}

function lookupPorcentajeServicio_(servicio) {
  if (!servicio) return 0;
  var sh = getSheet_(HOJA_CATALOGO);
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var data = sh.getRange(2, 1, last - 1, 2).getValues();
  var target = normalizar_(servicio);
  for (var i = 0; i < data.length; i++) {
    var srv = data[i][0], pct = data[i][1];
    if (!srv) continue;
    if (normalizar_(String(srv)) === target) {
      var p = parseFloat(pct);
      if (isNaN(p)) return 0;
      return p >= 1 ? p / 100 : p;
    }
  }
  return 0;
}

function formatCurrency_(n) {
  var v = isNaN(n) ? 0 : Number(n);
  return v.toLocaleString('es-PA', { style: 'currency', currency: 'PAB', minimumFractionDigits: 2 });
}

function sanitize_(texto) {
  if (texto === null || texto === undefined) return '';
  return String(texto).replace(/[&<>"']/g, function (c) {
    return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c];
  });
}

/***** =====================================================
 * REGISTRAR VENTA — reemplaza el Google Form
 * Body esperado: { cliente, colaborador, servicio, costo, metodoPago, notas }
 * ===================================================== *****/
function registrarVenta(body) {
  try {
    var sh = getSheet_(HOJA_RESPUESTAS);

    var fecha = new Date();
    var cliente = body.cliente || '';
    var colaborador = body.colaborador || '';
    var servicio = body.servicio || '';
    var categoria = (body.categoria || '').toLowerCase().trim();
    var costo = parseFloat(body.costo) || 0;
    var metodoPago = body.metodoPago || '';
    var notas = body.notas || '';

    var pct = lookupPorcentajeServicio_(servicio);
    var comisionSocio = costo * pct;
    var comisionStudio = costo - comisionSocio;
    var periodo = buildPeriodo_(fecha);

    var notasFinal = notas;
    if (pct === 0 && servicio) {
      notasFinal = notas ? notas + ' [Servicio no encontrado en catalogo: ' + servicio + ']' : '[Servicio no encontrado en catalogo: ' + servicio + ']';
    }

    var servicioEstilista = categoria === 'unas' ? '' : servicio;
    var servicioUnas = categoria === 'unas' ? servicio : '';

    sh.appendRow([
      fecha, cliente, colaborador, servicioEstilista, servicioUnas, costo,
      notasFinal, metodoPago, pct, comisionSocio, comisionStudio, periodo
    ]);

    var venta = {
      fecha: fecha, socio: colaborador, cliente: cliente, servicio: servicio,
      precio: costo, pct: pct, comisionSocio: comisionSocio,
      comisionEstudio: comisionStudio, notas: notasFinal, periodo: periodo
    };

    enviarCorreoVenta_(venta);

    return jsonOutput_({ ok: true, venta: venta });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

/***** ============== CORREO INMEDIATO DE VENTA ============== *****/
function enviarCorreoVenta_(data) {
  try {
    var cfg = getConfigMap_();
    var estudio = cfg['NombreEstudio'] || 'Estudio';
    var correoVentas = (cfg['CorreoVentas'] || '').trim();
    var mapColab = getColaboradoresMap_();
    var correoColab = mapColab[data.socio] || '';

    var fechaStr = Utilities.formatDate(data.fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var pctStr = Math.round((data.pct || 0) * 100) + '%';
    var precioStr = formatCurrency_(data.precio);
    var comSocStr = formatCurrency_(data.comisionSocio);
    var comEstStr = formatCurrency_(data.comisionEstudio);

    if (correoVentas) {
      var asuntoVentas = 'Venta registrada - ' + (data.socio || 'N/D') + ' - ' + (data.servicio || 'Servicio') + ' ' + precioStr;
      var htmlVentas =
        '<table width="100%" cellpadding="0" cellspacing="0" style="font-family:Arial,Helvetica,sans-serif">' +
        '<tr><td style="padding:14px;background:#111827;color:#fff;">' + sanitize_(estudio) + ' - Venta registrada</td></tr>' +
        '<tr><td style="padding:14px;border:1px solid #e5e7eb;">' +
        '<div><b>Fecha:</b> ' + fechaStr + ' &middot; <b>Periodo:</b> ' + sanitize_(data.periodo) + '</div>' +
        '<div><b>Socio:</b> ' + sanitize_(data.socio || '-') + '</div>' +
        '<div><b>Cliente:</b> ' + sanitize_(data.cliente || '-') + '</div>' +
        '<div><b>Servicio:</b> ' + sanitize_(data.servicio || '-') + '</div>' +
        '<div><b>Precio:</b> ' + precioStr + '</div>' +
        '<div><b>% aplicado:</b> ' + pctStr + '</div>' +
        '<div><b>Comision socio:</b> ' + comSocStr + '</div>' +
        '<div><b>Comision estudio:</b> ' + comEstStr + '</div>' +
        (data.notas ? '<div><b>Notas:</b> ' + sanitize_(data.notas) + '</div>' : '') +
        '</td></tr></table>';
      MailApp.sendEmail({ to: correoVentas, subject: asuntoVentas, htmlBody: htmlVentas });
    }

    if (correoColab) {
      var asuntoColab = 'Tu venta registrada - ' + (data.servicio || 'Servicio') + ' ' + precioStr;
      var nombreTC = String(data.socio || '').split(' ').filter(Boolean)
        .map(function (s) { return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); }).join(' ');
      var htmlColab =
        '<div style="font-family:Arial,Helvetica,sans-serif;color:#111827;">' +
        '<p>Hola <b>' + sanitize_(nombreTC) + '</b>,</p>' +
        '<p>Se registro tu venta del periodo <b>' + sanitize_(data.periodo) + '</b>:</p>' +
        '<ul style="margin:0 0 12px 18px;padding:0;">' +
        '<li><b>Cliente:</b> ' + sanitize_(data.cliente || '-') + '</li>' +
        '<li><b>Servicio:</b> ' + sanitize_(data.servicio || '-') + '</li>' +
        '<li><b>Precio:</b> ' + precioStr + '</li>' +
        '<li><b>% aplicado:</b> ' + pctStr + '</li>' +
        '<li><b>Total a recibir:</b> <b>' + comSocStr + '</b></li>' +
        '</ul>' +
        (data.notas ? '<p><b>Notas:</b> ' + sanitize_(data.notas) + '</p>' : '') +
        '<p style="margin-top:10px;">Gracias por tu trabajo.</p></div>';
      MailApp.sendEmail({ to: correoColab, subject: asuntoColab, htmlBody: htmlColab });
    }
  } catch (err) {
    Logger.log('Error en enviarCorreoVenta_: ' + err.message);
  }
}

/***** =====================================================
 * REGISTRAR ADELANTO
 * Body esperado: { colaborador, monto, motivo }
 * ===================================================== *****/
function registrarAdelanto(body) {
  try {
    var sh = getSheet_(HOJA_ADELANTOS);
    sh.appendRow([new Date(), body.colaborador || '', parseFloat(body.monto) || 0, body.motivo || '']);
    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

/***** =====================================================
 * REGISTRAR PROPINA
 * Body esperado: { colaborador, monto, cliente, estado, metodoPago }
 * estado: "Contado" o "Por pagar"
 * Estructura de la hoja Propinas: Timestamp | Colaborador | Monto | Cliente | Estado | MetodoPago
 * La propina es 100% para el socio, no se descuenta comision.
 * Solo las propinas en estado "Por pagar" se suman al Neto a pagar de la quincena.
 * ===================================================== *****/
function registrarPropina(body) {
  try {
    var sh = getSheet_(HOJA_PROPINAS);
    var estado = body.estado === 'Contado' ? 'Contado' : 'Por pagar';
    sh.appendRow([new Date(), body.colaborador || '', parseFloat(body.monto) || 0, body.cliente || '', estado, body.metodoPago || '']);
    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

/***** =====================================================
 * GET PROPINAS PENDIENTES (estado "Por pagar") de un socio
 * Params: colaborador
 * Devuelve cada fila con su numero de fila real en el Sheet,
 * para poder marcarla como pagada despues sin ambiguedad.
 * ===================================================== *****/
function getPropinasPendientes(e) {
  try {
    var colaborador = e.parameter.colaborador || '';
    var sh = getSheet_(HOJA_PROPINAS);
    var last = sh.getLastRow();
    var pendientes = [];
    if (last > 1 && colaborador) {
      var data = sh.getRange(2, 1, last - 1, 6).getValues();
      for (var i = 0; i < data.length; i++) {
        var fila = data[i];
        var fechaObj = fila[0];
        var col = String(fila[1] || '').trim();
        var monto = Number(fila[2]) || 0;
        var cliente = fila[3] || '';
        var estado = String(fila[4] || '').trim();
        if (col !== colaborador) continue;
        if (estado !== 'Por pagar') continue;
        pendientes.push({
          fila: i + 2,
          fecha: fechaObj instanceof Date ? formatoFechaLocal_(fechaObj) : '',
          monto: monto,
          cliente: cliente
        });
      }
    }
    return jsonOutput_({ ok: true, pendientes: pendientes });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message, pendientes: [] });
  }
}

/***** =====================================================
 * MARCAR PROPINA COMO PAGADA
 * Body esperado: { fila }
 * Actualiza el Estado de esa fila exacta en la hoja Propinas
 * de "Por pagar" a "Contado", sin crear un registro nuevo.
 * ===================================================== *****/
function marcarPropinaPagada(body) {
  try {
    var fila = parseInt(body.fila);
    if (!fila || fila < 2) return jsonOutput_({ ok: false, error: 'Numero de fila invalido.' });
    var sh = getSheet_(HOJA_PROPINAS);
    var estadoActual = String(sh.getRange(fila, 5).getValue() || '').trim();
    if (estadoActual !== 'Por pagar') {
      return jsonOutput_({ ok: false, error: 'Esa propina ya no esta en estado Por pagar.' });
    }
    sh.getRange(fila, 5).setValue('Contado');
    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

/***** =====================================================
 * GET CATALOGO DE SERVICIOS
 * ===================================================== *****/
function getCatalogoServicios() {
  try {
    var sh = getSheet_(HOJA_CATALOGO);
    var last = sh.getLastRow();
    if (last < 2) return jsonOutput_({ ok: true, servicios: [] });
    var data = sh.getRange(2, 1, last - 1, 2).getValues();
    var servicios = data.filter(function (r) { return r[0]; }).map(function (r) {
      return { servicio: String(r[0]).trim(), porcentaje: parseFloat(r[1]) || 0 };
    });
    return jsonOutput_({ ok: true, servicios: servicios });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message, servicios: [] });
  }
}

/***** =====================================================
 * GET COLABORADORES
 * ===================================================== *****/
function getColaboradores() {
  try {
    var sh = getSheet_(HOJA_COLABORADORES);
    var last = sh.getLastRow();
    if (last < 2) return jsonOutput_({ ok: true, colaboradores: [] });
    var data = sh.getRange(2, 1, last - 1, 2).getValues();
    var colaboradores = data.filter(function (r) { return r[0]; }).map(function (r) {
      return { nombre: String(r[0]).trim(), correo: String(r[1] || '').trim() };
    });
    return jsonOutput_({ ok: true, colaboradores: colaboradores });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message, colaboradores: [] });
  }
}

/***** =====================================================
 * GET COMISIONES POR COLABORADOR (historico, sin filtro)
 * ===================================================== *****/
function getComisionesPorColaborador() {
  try {
    var sh = getSheet_(HOJA_RESPUESTAS);
    var last = sh.getLastRow();
    if (last < 2) return jsonOutput_({ ok: true, comisiones: {} });
    var data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    var resultado = {};
    data.forEach(function (row) {
      var colaborador = row[2];
      var comision = Number(row[9]) || 0;
      if (!colaborador) return;
      resultado[colaborador] = (resultado[colaborador] || 0) + comision;
    });
    return jsonOutput_({ ok: true, comisiones: resultado });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message, comisiones: {} });
  }
}

/***** =====================================================
 * GET DASHBOARD RESUMEN
 * Params: inicio, fin, colaborador
 * ===================================================== *****/
function getDashboardResumen(e) {
  try {
    var p = e.parameter;
    var sh = getSheet_(HOJA_RESPUESTAS);
    var last = sh.getLastRow();

    var resultado = {
      ventas: 0, comision: 0, studio: 0, servicios: 0,
      adelantos: 0, propinas: 0, propinasPagadas: 0, neto: 0, comisionEstilismo: 0, comisionUnas: 0,
      pagos: { Efectivo: 0, Yappy: 0, Transferencia: 0, Tarjeta: 0, Otro: 0 }
    };

    if (last < 2) return jsonOutput_({ ok: true, resumen: resultado });

    var data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    var inicio = p.inicio || null;
    var fin = p.fin || null;
    var colaboradorFiltro = p.colaborador || null;

    data.forEach(function (fila) {
      var fechaObj = fila[0];
      if (!(fechaObj instanceof Date)) return;
      var colaborador = fila[2];
      var servicioEstilista = fila[3];
      var servicioUnas = fila[4];
      var costo = Number(fila[5]) || 0;
      var metodo = fila[7];
      var comisionSocio = Number(fila[9]) || 0;
      var comisionStudio = Number(fila[10]) || 0;

      var fecha = formatoFechaLocal_(fechaObj);
      if (inicio && fecha < inicio) return;
      if (fin && fecha > fin) return;
      if (colaboradorFiltro && colaborador !== colaboradorFiltro) return;

      resultado.ventas += costo;
      resultado.comision += comisionSocio;
      resultado.studio += comisionStudio;
      resultado.servicios++;

      if (metodo && resultado.pagos.hasOwnProperty(metodo)) {
        resultado.pagos[metodo] += costo;
      } else if (metodo) {
        resultado.pagos.Otro += costo;
      }

      if (servicioEstilista) resultado.comisionEstilismo += comisionSocio;
      if (servicioUnas) resultado.comisionUnas += comisionSocio;
    });

    if (colaboradorFiltro) {
      var shAdel = getSheet_(HOJA_ADELANTOS);
      var lastAdel = shAdel.getLastRow();
      if (lastAdel > 1) {
        var datosAdel = shAdel.getRange(2, 1, lastAdel - 1, 4).getValues();
        datosAdel.forEach(function (fila) {
          var empleado = fila[1];
          var monto = Number(fila[2]) || 0;
          if (empleado === colaboradorFiltro) resultado.adelantos += monto;
        });
      }

      var shProp = getSheet_(HOJA_PROPINAS);
      var lastProp = shProp.getLastRow();
      if (lastProp > 1) {
        var datosProp = shProp.getRange(2, 1, lastProp - 1, 6).getValues();
        datosProp.forEach(function (fila) {
          var empleado = fila[1];
          var monto = Number(fila[2]) || 0;
          var estado = String(fila[4] || '').trim();
          if (empleado !== colaboradorFiltro) return;
          if (estado === 'Por pagar') resultado.propinas += monto;
          else if (estado === 'Contado') resultado.propinasPagadas += monto;
        });
      }
    }

    resultado.neto = resultado.comision - resultado.adelantos + resultado.propinas;

    return jsonOutput_({ ok: true, resumen: resultado });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

/***** =====================================================
 * GET DETALLE DE SERVICIOS (tabla detallada + adelantos)
 * Params: inicio, fin, colaborador
 * ===================================================== *****/
function getDetalleServicios(e) {
  try {
    var p = e.parameter;
    var inicio = p.inicio || null;
    var fin = p.fin || null;
    var colaboradorFiltro = p.colaborador || null;

    var shResp = getSheet_(HOJA_RESPUESTAS);
    var lastResp = shResp.getLastRow();
    var resultado = [];

    if (lastResp > 1) {
      var dataResp = shResp.getRange(2, 1, lastResp - 1, shResp.getLastColumn()).getValues();
      dataResp.forEach(function (fila) {
        var fechaObj = fila[0];
        if (!(fechaObj instanceof Date)) return;
        var empleado = fila[2];
        var servicio = fila[3] || fila[4];
        var costo = Number(fila[5]) || 0;
        var metodo = fila[7];
        var cliente = fila[1];

        var fecha = formatoFechaLocal_(fechaObj);
        if (inicio && fecha < inicio) return;
        if (fin && fecha > fin) return;
        if (colaboradorFiltro && empleado !== colaboradorFiltro) return;

        resultado.push({ empleado: empleado, fecha: fecha, servicio: servicio, metodo: metodo, cliente: cliente, monto: costo });
      });
    }

    var shAdel = getSheet_(HOJA_ADELANTOS);
    var lastAdel = shAdel.getLastRow();
    if (lastAdel > 1 && colaboradorFiltro) {
      var dataAdel = shAdel.getRange(2, 1, lastAdel - 1, 4).getValues();
      dataAdel.forEach(function (fila) {
        var fechaObj = fila[0];
        if (!(fechaObj instanceof Date)) return;
        var empleado = fila[1];
        var monto = Number(fila[2]) || 0;
        var motivo = fila[3] || '-';

        var fecha = formatoFechaLocal_(fechaObj);
        if (inicio && fecha < inicio) return;
        if (fin && fecha > fin) return;
        if (empleado !== colaboradorFiltro) return;

        resultado.push({ empleado: empleado, fecha: fecha, servicio: 'Adelanto', metodo: '-', cliente: motivo, monto: -Math.abs(monto) });
      });
    }

    var shProp = getSheet_(HOJA_PROPINAS);
    var lastProp = shProp.getLastRow();
    if (lastProp > 1 && colaboradorFiltro) {
      var dataProp = shProp.getRange(2, 1, lastProp - 1, 6).getValues();
      dataProp.forEach(function (fila) {
        var fechaObj = fila[0];
        if (!(fechaObj instanceof Date)) return;
        var empleado = fila[1];
        var monto = Number(fila[2]) || 0;
        var cliente = fila[3] || '-';
        var estado = String(fila[4] || 'Por pagar').trim();
        var metodoPago = fila[5] || '-';

        var fecha = formatoFechaLocal_(fechaObj);
        if (inicio && fecha < inicio) return;
        if (fin && fecha > fin) return;
        if (empleado !== colaboradorFiltro) return;

        resultado.push({ empleado: empleado, fecha: fecha, servicio: 'Propina (' + estado + ')', metodo: metodoPago, cliente: cliente, monto: Math.abs(monto) });
      });
    }

    resultado.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });

    return jsonOutput_({ ok: true, detalle: resultado });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message, detalle: [] });
  }
}

/***** =====================================================
 * GENERAR COMPROBANTE PDF (sin archivar en Drive)
 * Body esperado: { colaborador, inicio, fin }
 * Devuelve el PDF en base64 para descargar/enviar desde el frontend,
 * y ademas lo envia por correo al colaborador.
 * ===================================================== *****/
function generarComprobante(body) {
  try {
    var colaborador = body.colaborador;
    var inicio = body.inicio;
    var fin = body.fin;
    if (!colaborador || !inicio || !fin) {
      return jsonOutput_({ ok: false, error: 'Faltan datos: colaborador, inicio o fin.' });
    }

    var fechaFin = new Date(fin);
    var periodo = buildPeriodo_(fechaFin);

    var sh = getSheet_(HOJA_RESPUESTAS);
    var last = sh.getLastRow();
    var filasHTML = '';
    var totalComision = 0;
    var subtotal = 0;

    if (last > 1) {
      var data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
      data.forEach(function (row) {
        var fechaObj = row[0];
        if (!(fechaObj instanceof Date)) return;
        var col = row[2];
        var cliente = row[1] || '';
        var servicio = row[3] || row[4];
        var costoServicio = Number(row[5]) || 0;
        var comision = Number(row[9]) || 0;
        var fecha = formatoFechaLocal_(fechaObj);

        if (col !== colaborador) return;
        if (fecha < inicio || fecha > fin) return;

        totalComision += comision;
        subtotal += costoServicio;
        filasHTML += '<tr><td>' + fecha + '</td><td>' + sanitize_(servicio) + '</td><td>' + sanitize_(cliente) + '</td><td style="text-align:right;">USD ' + comision.toFixed(2) + '</td></tr>';
      });
    }

    var shAdel = getSheet_(HOJA_ADELANTOS);
    var lastAdel = shAdel.getLastRow();
    var totalAdelantos = 0;
    if (lastAdel > 1) {
      var dataAdel = shAdel.getRange(2, 1, lastAdel - 1, 4).getValues();
      dataAdel.forEach(function (row) {
        var fechaObj = row[0];
        if (!(fechaObj instanceof Date)) return;
        var col = row[1];
        var monto = Number(row[2]) || 0;
        var fecha = formatoFechaLocal_(fechaObj);

        if (col !== colaborador) return;
        if (fecha < inicio || fecha > fin) return;

        totalAdelantos += monto;
      });
    }

    var shProp = getSheet_(HOJA_PROPINAS);
    var lastProp = shProp.getLastRow();
    var totalPropinas = 0;
    var totalPropinasPagadas = 0;
    var filasPropinasPorPagarHTML = '';
    if (lastProp > 1) {
      var dataProp = shProp.getRange(2, 1, lastProp - 1, 6).getValues();
      dataProp.forEach(function (row) {
        var fechaObj = row[0];
        if (!(fechaObj instanceof Date)) return;
        var col = row[1];
        var monto = Number(row[2]) || 0;
        var clientePropina = row[3] || 'Sin nombre';
        var estado = String(row[4] || '').trim();
        var fecha = formatoFechaLocal_(fechaObj);

        if (col !== colaborador) return;
        if (fecha < inicio || fecha > fin) return;

        if (estado === 'Por pagar') {
          totalPropinas += monto;
          filasPropinasPorPagarHTML += '<tr><td>' + fecha + '</td><td>' + sanitize_(clientePropina) + '</td><td style="text-align:right;">USD ' + monto.toFixed(2) + '</td></tr>';
        } else if (estado === 'Contado') {
          totalPropinasPagadas += monto;
        }
      });
    }

    var netoFinal = totalComision - totalAdelantos + totalPropinas;

    var logoFile = DriveApp.getFileById(LOGO_FILE_ID);
    var logoBlob = logoFile.getBlob();
    var logoBase64 = Utilities.base64Encode(logoBlob.getBytes());
    var logoType = logoBlob.getContentType();

    var html =
      '<html><head><style>' +
      'body{font-family:Arial,sans-serif;padding:60px 70px;color:#222;}' +
      '.logo{text-align:center;margin-bottom:10px;}.logo img{height:110px;}' +
      '.brand{text-align:center;font-weight:bold;font-size:20px;margin-top:10px;}' +
      '.ruc{text-align:center;font-size:12px;color:#666;margin-bottom:35px;}' +
      '.info{margin-bottom:30px;font-size:14px;}' +
      'h3{font-size:13px;color:#555;margin:30px 0 6px;}' +
      'table{width:100%;border-collapse:collapse;margin-top:10px;}' +
      'th{border-bottom:2px solid #000;text-align:left;padding-bottom:10px;font-size:13px;}' +
      'td{padding:10px 0;border-bottom:1px solid #eee;font-size:13px;}' +
      '.resumen{margin-top:30px;text-align:right;font-size:14px;}' +
      '.resumen-linea{margin:6px 0;}' +
      '.nota-aparte{margin-top:20px;padding:12px 16px;background:#f9f9f9;border-left:3px solid #ccc;font-size:12px;color:#666;}' +
      '.bloque-total{margin-top:25px;padding:20px;background:#f5f5f5;border-radius:8px;text-align:right;}' +
      '.total-final-label{font-size:14px;color:#555;}' +
      '.total-final-valor{font-size:30px;font-weight:bold;margin-top:5px;}' +
      '.footer{margin-top:50px;font-size:11px;text-align:center;color:#888;}' +
      '</style></head><body>' +
      '<div class="logo"><img src="data:' + logoType + ';base64,' + logoBase64 + '"></div>' +
      '<div class="brand">VICELLY SANCHEZ STUDIO</div>' +
      '<div class="ruc">RUC: 8-NT-2-721034 DV41</div>' +
      '<div class="info"><strong>Colaborador:</strong> ' + sanitize_(colaborador) + '<br><strong>Periodo:</strong> ' + periodo + '</div>' +
      '<h3>Servicios realizados</h3>' +
      '<table><tr><th>Fecha</th><th>Servicio</th><th>Cliente</th><th style="text-align:right;">Comision</th></tr>' + filasHTML + '</table>' +
      (filasPropinasPorPagarHTML ?
        '<h3>Propinas por pagar</h3>' +
        '<table><tr><th>Fecha</th><th>Cliente</th><th style="text-align:right;">Monto</th></tr>' + filasPropinasPorPagarHTML + '</table>'
        : '') +
      '<div class="resumen">' +
      '<div class="resumen-linea">Comision total: USD ' + totalComision.toFixed(2) + '</div>' +
      '<div class="resumen-linea">Adelantos: USD -' + totalAdelantos.toFixed(2) + '</div>' +
      '<div class="resumen-linea">Propinas por pagar: USD +' + totalPropinas.toFixed(2) + '</div>' +
      '</div>' +
      '<div class="bloque-total"><div class="total-final-label">TOTAL A PAGAR</div>' +
      '<div class="total-final-valor">USD ' + netoFinal.toFixed(2) + '</div></div>' +
      (totalPropinasPagadas > 0 ?
        '<div class="nota-aparte">Nota: ya recibiste USD ' + totalPropinasPagadas.toFixed(2) + ' en propinas directamente durante este periodo. Ese monto no esta incluido en el total a pagar porque ya esta en tu poder.</div>'
        : '') +
      '<div class="footer">Documento generado automaticamente el ' + new Date().toLocaleString() + '</div>' +
      '</body></html>';

    var blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
    var nombreArchivo = 'Comprobante_' + colaborador + '_' + periodo.replace(' ', '_') + '.pdf';

    var pdfBase64 = Utilities.base64Encode(blob.getBytes());

    return jsonOutput_({
      ok: true,
      nombreArchivo: nombreArchivo,
      pdfBase64: pdfBase64,
      totalComision: totalComision,
      totalAdelantos: totalAdelantos,
      totalPropinas: totalPropinas,
      totalPropinasPagadas: totalPropinasPagadas,
      netoFinal: netoFinal
    });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

/***** =====================================================
 * CERRAR QUINCENA
 * Body esperado: { periodo }  Ej: "2026-06 Q1"
 * - Genera y envia comprobante a CADA colaboradora con movimientos en el periodo
 * - Archiva las filas del periodo en una hoja oculta "Cierre_YYYY-MM_QX"
 * - Limpia Adelantos del periodo
 * - Envia resumen al correo del estudio
 * ===================================================== *****/
function cerrarQuincena(body) {
  try {
    var periodo = body.periodo;
    var colaboradorFiltro = body.colaborador || null;
    if (!periodo) return jsonOutput_({ ok: false, error: 'Debe indicar el periodo a cerrar.' });

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var shResp = getSheet_(HOJA_RESPUESTAS);
    var last = shResp.getLastRow();
    if (last < 2) return jsonOutput_({ ok: false, error: 'No hay datos para cerrar.' });

    var headers = shResp.getRange(1, 1, 1, shResp.getLastColumn()).getValues()[0];
    var data = shResp.getRange(2, 1, last - 1, shResp.getLastColumn()).getValues();
    var idxPeriodo = headers.indexOf('Periodo');
    var idxColab = headers.indexOf('Socio de Belleza');

    var filasPeriodo = data.filter(function (r) {
      var coincideP = String(r[idxPeriodo] || '').trim() === periodo;
      var coincideC = colaboradorFiltro ? String(r[idxColab] || '').trim() === colaboradorFiltro : true;
      return coincideP && coincideC;
    });
    if (filasPeriodo.length === 0) {
      return jsonOutput_({ ok: false, error: 'No hay filas para cerrar con esos filtros.' });
    }

    var colaboradoresPeriodo = {};
    filasPeriodo.forEach(function (r) {
      var col = String(r[idxColab] || '').trim();
      if (col) colaboradoresPeriodo[col] = true;
    });

    var partes = periodo.split(' ');
    var anioMes = partes[0];
    var quincena = partes[1];
    var dia = quincena === 'Q1' ? 1 : 16;
    var fechaRef = new Date(anioMes + '-' + String(dia).padStart(2, '0'));
    var inicioStr = anioMes + '-' + (quincena === 'Q1' ? '01' : '16');
    var ultimoDiaMes = new Date(fechaRef.getFullYear(), fechaRef.getMonth() + 1, 0).getDate();
    var finStr = anioMes + '-' + (quincena === 'Q1' ? '15' : String(ultimoDiaMes));

    var enviados = [];
    var fallidos = [];
    var totalComisionGeneral = 0;
    var totalAdelantosGeneral = 0;
    var totalPropinasGeneral = 0;
    var totalNetoGeneral = 0;
    var detalleLineas = [];

    Object.keys(colaboradoresPeriodo).forEach(function (colaborador) {
      try {
        var resultado = generarYEnviarComprobanteInterno_(colaborador, inicioStr, finStr, periodo);
        totalComisionGeneral += resultado.totalComision;
        totalAdelantosGeneral += resultado.totalAdelantos;
        totalPropinasGeneral += resultado.totalPropinas;
        totalNetoGeneral += resultado.netoFinal;
        detalleLineas.push({
          colaborador: colaborador,
          comision: resultado.totalComision,
          adelantos: resultado.totalAdelantos,
          propinas: resultado.totalPropinas,
          neto: resultado.netoFinal,
          enviado: resultado.enviado
        });
        if (resultado.enviado) enviados.push(colaborador);
        else fallidos.push(colaborador + ' (sin correo configurado)');
      } catch (errColab) {
        fallidos.push(colaborador + ' (' + errColab.message + ')');
      }
    });

    var sufijoHoja = colaboradorFiltro ? '_' + colaboradorFiltro.replace(/\s+/g, '') : '';
    var periodoHoja = 'Cierre_' + periodo.replace(' ', '_') + sufijoHoja;
    var old = ss.getSheetByName(periodoHoja);
    if (old) ss.deleteSheet(old);

    var shCierre = ss.insertSheet(periodoHoja);
    shCierre.getRange(1, 1, 1, headers.length).setValues([headers]);
    shCierre.getRange(2, 1, filasPeriodo.length, headers.length).setValues(filasPeriodo);
    var prot = shCierre.protect().setDescription('Cierre quincenal');
    prot.removeEditors(prot.getEditors());
    shCierre.hideSheet();

    var filasEliminar = [];
    for (var i = 2; i <= shResp.getLastRow(); i++) {
      var valP = shResp.getRange(i, idxPeriodo + 1).getValue();
      var valC = shResp.getRange(i, idxColab + 1).getValue();
      var coincideP = String(valP).trim() === periodo;
      var coincideC = colaboradorFiltro ? String(valC).trim() === colaboradorFiltro : true;
      if (coincideP && coincideC) filasEliminar.push(i);
    }
    for (var j = filasEliminar.length - 1; j >= 0; j--) {
      shResp.deleteRow(filasEliminar[j]);
    }

    var shAdel = getSheet_(HOJA_ADELANTOS);
    if (shAdel.getLastRow() > 1) {
      var iAdel = shAdel.getLastRow();
      while (iAdel >= 2) {
        var fechaAdel = shAdel.getRange(iAdel, 1).getValue();
        var colAdel = String(shAdel.getRange(iAdel, 2).getValue() || '').trim();
        var coincidePeriodoAdel = fechaAdel instanceof Date && buildPeriodo_(fechaAdel) === periodo;
        var coincideColabAdel = colaboradorFiltro ? colAdel === colaboradorFiltro : true;
        if (coincidePeriodoAdel && coincideColabAdel) {
          shAdel.deleteRow(iAdel);
        }
        iAdel--;
      }
    }

    var shProp = getSheet_(HOJA_PROPINAS);
    if (shProp.getLastRow() > 1) {
      var iProp = shProp.getLastRow();
      while (iProp >= 2) {
        var fechaProp = shProp.getRange(iProp, 1).getValue();
        var colProp = String(shProp.getRange(iProp, 2).getValue() || '').trim();
        var coincidePeriodoProp = fechaProp instanceof Date && buildPeriodo_(fechaProp) === periodo;
        var coincideColabProp = colaboradorFiltro ? colProp === colaboradorFiltro : true;
        if (coincidePeriodoProp && coincideColabProp) {
          shProp.deleteRow(iProp);
        }
        iProp--;
      }
    }

    var cfg = getConfigMap_();
    var emailEstudio = cfg['CorreoEstudio'] || '';
    if (emailEstudio) {
      var filasTabla = detalleLineas.map(function (d) {
        return '<tr><td style="padding:6px 10px;border-bottom:1px solid #eee;">' + sanitize_(d.colaborador) + '</td>' +
          '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;">USD ' + d.comision.toFixed(2) + '</td>' +
          '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;color:#b91c1c;">USD ' + d.adelantos.toFixed(2) + '</td>' +
          '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;color:#15803d;">USD ' + d.propinas.toFixed(2) + '</td>' +
          '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;font-weight:bold;">USD ' + d.neto.toFixed(2) + '</td></tr>';
      }).join('');

      var mensaje =
        '<div style="font-family:Arial,sans-serif;max-width:560px;">' +
        '<p>Se ha completado el cierre' + (colaboradorFiltro ? ' de <b>' + sanitize_(colaboradorFiltro) + '</b>' : '') + ' del periodo <b>' + periodo + '</b>.</p>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:14px;">' +
        '<tr style="background:#f5f5f5;"><th style="padding:6px 10px;text-align:left;">Colaboradora</th><th style="padding:6px 10px;text-align:right;">Comision</th><th style="padding:6px 10px;text-align:right;">Adelantos</th><th style="padding:6px 10px;text-align:right;">Propinas</th><th style="padding:6px 10px;text-align:right;">Neto pagado</th></tr>' +
        filasTabla +
        '<tr style="background:#fafafa;font-weight:bold;"><td style="padding:8px 10px;">TOTAL</td><td style="padding:8px 10px;text-align:right;">USD ' + totalComisionGeneral.toFixed(2) + '</td><td style="padding:8px 10px;text-align:right;">USD ' + totalAdelantosGeneral.toFixed(2) + '</td><td style="padding:8px 10px;text-align:right;">USD ' + totalPropinasGeneral.toFixed(2) + '</td><td style="padding:8px 10px;text-align:right;">USD ' + totalNetoGeneral.toFixed(2) + '</td></tr>' +
        '</table>' +
        '<p style="margin-top:16px;">Comprobantes enviados: <b>' + enviados.length + '</b></p>' +
        (fallidos.length ? '<p style="color:#b91c1c;"><b>No se pudo enviar a:</b> ' + fallidos.join(', ') + '</p>' : '') +
        '</div>';
      MailApp.sendEmail({ to: emailEstudio, subject: 'Cierre Quincenal - ' + periodo + (colaboradorFiltro ? ' - ' + colaboradorFiltro : ''), htmlBody: mensaje });
    }

    return jsonOutput_({ ok: true, enviados: enviados.length, fallidos: fallidos, totalComision: totalComisionGeneral, totalAdelantos: totalAdelantosGeneral, totalPropinas: totalPropinasGeneral, totalNeto: totalNetoGeneral });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

/***** Funcion interna compartida entre generarComprobante() y cerrarQuincena() *****/
function generarYEnviarComprobanteInterno_(colaborador, inicio, fin, periodo) {
  var sh = getSheet_(HOJA_RESPUESTAS);
  var last = sh.getLastRow();
  var filasHTML = '';
  var totalComision = 0;
  var subtotal = 0;

  if (last > 1) {
    var data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    data.forEach(function (row) {
      var fechaObj = row[0];
      if (!(fechaObj instanceof Date)) return;
      var col = row[2];
      var cliente = row[1] || '';
      var servicio = row[3] || row[4];
      var costoServicio = Number(row[5]) || 0;
      var comision = Number(row[9]) || 0;
      var fecha = formatoFechaLocal_(fechaObj);

      if (col !== colaborador) return;
      if (fecha < inicio || fecha > fin) return;

      totalComision += comision;
      subtotal += costoServicio;
      filasHTML += '<tr><td>' + fecha + '</td><td>' + sanitize_(servicio) + '</td><td>' + sanitize_(cliente) + '</td><td style="text-align:right;">USD ' + comision.toFixed(2) + '</td></tr>';
    });
  }

  var shAdel = getSheet_(HOJA_ADELANTOS);
  var lastAdel = shAdel.getLastRow();
  var totalAdelantos = 0;
  if (lastAdel > 1) {
    var dataAdel = shAdel.getRange(2, 1, lastAdel - 1, 4).getValues();
    dataAdel.forEach(function (row) {
      var fechaObj = row[0];
      if (!(fechaObj instanceof Date)) return;
      var col = row[1];
      var monto = Number(row[2]) || 0;
      var fecha = formatoFechaLocal_(fechaObj);

      if (col !== colaborador) return;
      if (fecha < inicio || fecha > fin) return;

      totalAdelantos += monto;
    });
  }

  var shProp = getSheet_(HOJA_PROPINAS);
  var lastProp = shProp.getLastRow();
  var totalPropinas = 0;
  var totalPropinasPagadas = 0;
  var filasPropinasPorPagarHTML = '';
  if (lastProp > 1) {
    var dataProp = shProp.getRange(2, 1, lastProp - 1, 6).getValues();
    dataProp.forEach(function (row) {
      var fechaObj = row[0];
      if (!(fechaObj instanceof Date)) return;
      var col = row[1];
      var monto = Number(row[2]) || 0;
      var clientePropina = row[3] || 'Sin nombre';
      var estado = String(row[4] || '').trim();
      var fecha = formatoFechaLocal_(fechaObj);

      if (col !== colaborador) return;
      if (fecha < inicio || fecha > fin) return;

      if (estado === 'Por pagar') {
        totalPropinas += monto;
        filasPropinasPorPagarHTML += '<tr><td>' + fecha + '</td><td>' + sanitize_(clientePropina) + '</td><td style="text-align:right;">USD ' + monto.toFixed(2) + '</td></tr>';
      } else if (estado === 'Contado') {
        totalPropinasPagadas += monto;
      }
    });
  }

  var netoFinal = totalComision - totalAdelantos + totalPropinas;

  var logoFile = DriveApp.getFileById(LOGO_FILE_ID);
  var logoBlob = logoFile.getBlob();
  var logoBase64 = Utilities.base64Encode(logoBlob.getBytes());
  var logoType = logoBlob.getContentType();

  var html =
    '<html><head><style>' +
    'body{font-family:Arial,sans-serif;padding:60px 70px;color:#222;}' +
    '.logo{text-align:center;margin-bottom:10px;}.logo img{height:110px;}' +
    '.brand{text-align:center;font-weight:bold;font-size:20px;margin-top:10px;}' +
    '.ruc{text-align:center;font-size:12px;color:#666;margin-bottom:35px;}' +
    '.info{margin-bottom:30px;font-size:14px;}' +
    'h3{font-size:13px;color:#555;margin:30px 0 6px;}' +
    'table{width:100%;border-collapse:collapse;margin-top:10px;}' +
    'th{border-bottom:2px solid #000;text-align:left;padding-bottom:10px;font-size:13px;}' +
    'td{padding:10px 0;border-bottom:1px solid #eee;font-size:13px;}' +
    '.resumen{margin-top:30px;text-align:right;font-size:14px;}' +
    '.resumen-linea{margin:6px 0;}' +
    '.nota-aparte{margin-top:20px;padding:12px 16px;background:#f9f9f9;border-left:3px solid #ccc;font-size:12px;color:#666;}' +
    '.bloque-total{margin-top:25px;padding:20px;background:#f5f5f5;border-radius:8px;text-align:right;}' +
    '.total-final-label{font-size:14px;color:#555;}' +
    '.total-final-valor{font-size:30px;font-weight:bold;margin-top:5px;}' +
    '.footer{margin-top:50px;font-size:11px;text-align:center;color:#888;}' +
    '</style></head><body>' +
    '<div class="logo"><img src="data:' + logoType + ';base64,' + logoBase64 + '"></div>' +
    '<div class="brand">VICELLY SANCHEZ STUDIO</div>' +
    '<div class="ruc">RUC: 8-NT-2-721034 DV41</div>' +
    '<div class="info"><strong>Colaborador:</strong> ' + sanitize_(colaborador) + '<br><strong>Periodo:</strong> ' + periodo + '</div>' +
    '<h3>Servicios realizados</h3>' +
    '<table><tr><th>Fecha</th><th>Servicio</th><th>Cliente</th><th style="text-align:right;">Comision</th></tr>' + filasHTML + '</table>' +
    (filasPropinasPorPagarHTML ?
      '<h3>Propinas por pagar</h3>' +
      '<table><tr><th>Fecha</th><th>Cliente</th><th style="text-align:right;">Monto</th></tr>' + filasPropinasPorPagarHTML + '</table>'
      : '') +
    '<div class="resumen">' +
    '<div class="resumen-linea">Comision total: USD ' + totalComision.toFixed(2) + '</div>' +
    '<div class="resumen-linea">Adelantos: USD -' + totalAdelantos.toFixed(2) + '</div>' +
    '<div class="resumen-linea">Propinas por pagar: USD +' + totalPropinas.toFixed(2) + '</div>' +
    '</div>' +
    '<div class="bloque-total"><div class="total-final-label">TOTAL A PAGAR</div>' +
    '<div class="total-final-valor">USD ' + netoFinal.toFixed(2) + '</div></div>' +
    (totalPropinasPagadas > 0 ?
      '<div class="nota-aparte">Nota: ya recibiste USD ' + totalPropinasPagadas.toFixed(2) + ' en propinas directamente durante este periodo. Ese monto no esta incluido en el total a pagar porque ya esta en tu poder.</div>'
      : '') +
    '<div class="footer">Documento generado automaticamente el ' + new Date().toLocaleString() + '</div>' +
    '</body></html>';

  var blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  var nombreArchivo = 'Comprobante_' + colaborador + '_' + periodo.replace(' ', '_') + '.pdf';

  var mapColab = getColaboradoresMap_();
  var correoColab = mapColab[colaborador];
  var enviado = false;

  if (correoColab) {
    MailApp.sendEmail({
      to: correoColab,
      subject: 'Comprobante de Pago - ' + colaborador + ' (' + periodo + ')',
      htmlBody: '<p>Hola <b>' + sanitize_(colaborador) + '</b>,</p><p>Adjunto tu comprobante de pago correspondiente al periodo <b>' + periodo + '</b>.</p>',
      attachments: [blob.setName(nombreArchivo)]
    });
    enviado = true;
  }

  return { enviado: enviado, totalComision: totalComision, totalAdelantos: totalAdelantos, totalPropinas: totalPropinas, totalPropinasPagadas: totalPropinasPagadas, netoFinal: netoFinal };
}
