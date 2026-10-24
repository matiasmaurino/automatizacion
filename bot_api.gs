/***********************************************************
 * API PARA EL BOT DE WHATSAPP
 *
 * Agregar este archivo al proyecto de Apps Script (junto a modulos.gs
 * y email.gs) y volver a implementar la web app (Implementar >
 * Administrar implementaciones > Editar > Nueva versión).
 *
 * Antes: crear la propiedad del script BOT_SECRET
 * (Configuración del proyecto > Propiedades del script) con un valor
 * largo y aleatorio. El bot manda ese mismo valor en cada pedido.
 *
 * Las claves nunca salen de acá: el bot solo recibe un "ok", el email
 * enmascarado y los datos necesarios para armar el pedido (paciente,
 * prestaciones, valores, CUIT receptores). Las facturas se cargan en la
 * hoja WEBAPP con las mismas funciones que usa la web (guardarFactura y
 * guardarFacturaNoIoma).
 ***********************************************************/

const BOT_COL_NOMBRE = 0;     // A
const BOT_COL_CUIT = 1;       // B
const BOT_COL_CLAVE_ARCA = 2; // C
const BOT_COL_CLAVE_ARBA = 3; // D
const BOT_COL_ALAS = 4;       // E
const BOT_COL_EMAIL = 6;      // G
const BOT_COL_HONORARIOS = 12; // M

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return _botJson({ ok: false, code: 'ERROR' });
  }

  const secret = PropertiesService.getScriptProperties().getProperty('BOT_SECRET');
  if (!secret || req.secret !== secret) {
    return _botJson({ ok: false, code: 'ERROR' });
  }

  try {
    return _botJson(_botDispatch(req));
  } catch (err) {
    Logger.log('bot_api error: ' + err);
    return _botJson(_botMapError(err));
  }
}

function _botDispatch(req) {
  // Consultas que no dependen del cliente
  if (req.action === 'valor_hora') {
    const v = obtenerValorHora(String(req.servicio || ''), Number(req.mes), Number(req.anio));
    return { ok: true, valorK: v.valorK, valorM: v.valorM };
  }
  if (req.action === 'meses_retro') {
    return { ok: true, meses: obtenerMesesRetroactivos(String(req.servicio || '')) };
  }

  const cliente = _botBuscarCliente(req.dni);
  if (cliente.count === 0) return { ok: false, code: 'NOT_FOUND' };
  if (cliente.count > 1) return { ok: false, code: 'MULTIPLE' };

  const c = cliente.fila;
  const base = { ok: true, emailMasked: _botMaskEmail(c.email), sinEmail: !c.email };

  switch (req.action) {
    case 'identificar':
      return base;

    case 'opcion':
      guardarOpcion({ cuit: c.cuit, claveAfip: c.claveArca, clienteNombre: c.nombre });
      return base;

    case 'claves':
      _botRegistrarConsulta(c);
      return base;

    case 'deuda':
      _botRegistrarDeuda(c);
      return base;

    case 'honorarios':
      return { ok: true, valor: c.honorarios };

    // ---- Facturación IOMA (reemplaza index.html / guardarFactura) ----
    case 'ioma_paciente': {
      const p = buscarPacientePorDni(c.fila, req.dniPaciente);
      if (!p) return { ok: false, code: 'PACIENTE_NOT_FOUND' };
      return {
        ok: true,
        paciente: { nombre: p.nombre, dni: p.dni, tramites: p.tramites || [] },
        servicios: listarServiciosCliente(c.fila)
      };
    }

    case 'ioma_guardar': {
      const p = buscarPacientePorDni(c.fila, req.dniPaciente);
      if (!p) return { ok: false, code: 'PACIENTE_NOT_FOUND' };
      const mes = Number(req.mes), anio = Number(req.anio), horas = Number(req.horas);
      const retro = req.retroactivo === true;
      if (!_botMesOk(mes, anio) || !(horas > 0)) return { ok: false, code: 'ERROR' };
      if (retro && (!req.puntoVenta || !req.nroComprobante)) return { ok: false, code: 'ERROR' };
      const res = guardarFactura({
        cuit: c.cuit, claveAfip: c.claveArca, clienteNombre: c.nombre, email: c.email,
        pacienteNombre: p.nombre, numeroAfiliado: p.numeroAfiliado, estado: p.estado, dniPaciente: p.dni,
        numeroTramite: String(req.numeroTramite || ''), servicio: String(req.servicio || ''),
        mesNumero: mes, anio: anio, horas: horas, cuitReceptor: CUIT_IOMA,
        retroactivo: retro, subimoAIoma: req.subimoAIoma === true,
        puntoVenta: req.puntoVenta, nroComprobante: req.nroComprobante
      });
      return Object.assign({}, base, {
        resolucion: res.resolucion,
        valorHora: retro ? (res.valorHoraM || res.valorHora) : res.valorHora
      });
    }

    // ---- Facturación sin CAE (reemplaza index no ioma.html) ----
    case 'noioma_opciones':
      return {
        ok: true,
        servicios: listarServiciosClienteNoIoma(c.fila),
        pacientes: listarPacientesClienteNoIoma(c.fila),
        cuits: _botCuitsNoIoma(c.fila)
      };

    case 'noioma_guardar': {
      const receptor = _soloDigitos(req.cuitReceptor);
      if (_botCuitsNoIoma(c.fila).indexOf(receptor) === -1) return { ok: false, code: 'ERROR' };
      const mes = Number(req.mes), anio = Number(req.anio);
      if (!_botMesOk(mes, anio)) return { ok: false, code: 'ERROR' };
      const modo = req.modo === 'directo' ? 'directo' : 'horas';
      let horas = 0, valorHora = 0, importe = 0;
      if (modo === 'horas') {
        horas = Number(req.horas);
        valorHora = Number(req.valorHora);
        importe = horas * valorHora;
        if (!(horas > 0 && valorHora > 0)) return { ok: false, code: 'ERROR' };
      } else {
        importe = Number(req.importeTotal);
        if (!(importe > 0)) return { ok: false, code: 'ERROR' };
      }
      guardarFacturaNoIoma({
        cuit: c.cuit, claveAfip: c.claveArca, clienteNombre: c.nombre, email: c.email,
        pacienteNombre: '', numeroAfiliado: '', estado: '', dniPaciente: '', numeroTramite: '',
        servicio: String(req.servicio || ''), paciente: String(req.paciente || ''),
        mesNumero: mes, anio: anio, horas: horas, valorHoraManual: valorHora, importeTotal: importe,
        cuitReceptor: receptor, modoFacturacion: modo, retroactivo: false
      });
      return Object.assign({}, base, { importeTotal: importe });
    }

    default:
      return { ok: false, code: 'ERROR' };
  }
}

function _botMesOk(mes, anio) {
  return mes >= 1 && mes <= 12 && anio >= 2000 && anio <= 2100;
}

function _botCuitsNoIoma(fila) {
  return listarCuitsReceptorClienteNoIoma(fila).filter(function (x) { return x !== CUIT_IOMA; });
}

/***********************************************************
 * BÚSQUEDA POR DNI
 * El cliente escribe su DNI y se compara con los dígitos 3 a 10
 * del CUIT de la columna B (11 dígitos: 2 de prefijo + 8 de DNI
 * + 1 verificador). Un DNI de 7 dígitos se completa con un cero.
 * Coincidencia exacta, no "contiene".
 ***********************************************************/
function _botDniDeCuit(cuit) {
  const c = _soloDigitos(cuit);
  return c.length === 11 ? c.substring(2, 10) : '';
}

function _botBuscarCliente(dniInput) {
  const dni = _soloDigitos(dniInput);
  if (dni.length < 6 || dni.length > 8) return { count: 0 };
  const buscado = ('00000000' + dni).slice(-8);

  const sheet = SpreadsheetApp.openById(EXENTOS_SS_ID).getSheetByName(NOMBRE_HOJA_EXENTOS);
  const valores = sheet.getDataRange().getValues();
  const visibles = sheet.getDataRange().getDisplayValues();

  const encontrados = [];
  for (let r = 1; r < valores.length; r++) {
    if (_botDniDeCuit(valores[r][BOT_COL_CUIT]) === buscado) {
      encontrados.push({
        fila: r + 1,
        nombre: String(valores[r][BOT_COL_NOMBRE] || '').trim(),
        cuit: String(valores[r][BOT_COL_CUIT] || '').trim(),
        claveArca: String(valores[r][BOT_COL_CLAVE_ARCA] || '').trim(),
        claveArba: String(valores[r][BOT_COL_CLAVE_ARBA] || '').trim(),
        alas: valores[r][BOT_COL_ALAS],
        email: String(valores[r][BOT_COL_EMAIL] || '').trim(),
        honorarios: String(visibles[r][BOT_COL_HONORARIOS] || '').trim()
      });
    }
  }
  return { count: encontrados.length, fila: encontrados[0] };
}

/***********************************************************
 * REGISTRO DE PEDIDOS (mismas hojas y columnas que la web)
 * Se escriben desde la fila ya resuelta en EXENTOS, así no
 * dependen de que otra hoja ("Exentos") tenga el mismo orden.
 ***********************************************************/
function _botRegistrarConsulta(c) {
  const ssAuto = SpreadsheetApp.openById(AUTOMATIZACION_SS_ID);
  let sheetDP = ssAuto.getSheetByName('DATOS PERSONALES');
  if (!sheetDP) {
    sheetDP = ssAuto.insertSheet('DATOS PERSONALES');
    sheetDP.getRange(1, 1, 1, 7).setValues([[
      'CUIT', 'CLIENTE', 'CLAVE ARCA', 'CLAVE ARBA', 'VENCIMIENTO ALAS', 'EMAIL', 'FECHA SOLICITUD'
    ]]);
  }
  sheetDP.appendRow([c.cuit, c.nombre, c.claveArca, c.claveArba, c.alas, c.email, new Date()]);
}

function _botRegistrarDeuda(c) {
  const ssAuto = SpreadsheetApp.openById(AUTOMATIZACION_SS_ID);
  let hojaCcma = ssAuto.getSheetByName('CCMA');
  if (!hojaCcma) {
    hojaCcma = ssAuto.insertSheet('CCMA');
    hojaCcma.appendRow(['CUIT', 'CLIENTE', 'CLAVE AFIP', 'FECHA PEDIDO', 'ESTADO', 'EMAIL ENVIADO', 'EMAIL']);
  }
  const siguienteFila = obtenerUltimaFilaReal(hojaCcma) + 1;
  hojaCcma.getRange(siguienteFila, 1, 1, 7).setValues([[
    c.cuit, c.nombre, c.claveArca, new Date(), 'PROCESADO', '', c.email
  ]]);
}

/***********************************************************
 * UTILIDADES
 ***********************************************************/
function _botMaskEmail(email) {
  const m = String(email || '').trim().match(/^(.)([^@]*)@(.+)$/);
  if (!m) return '';
  return m[1] + '***@' + m[3];
}

function _botMapError(err) {
  const msg = String((err && err.message) || err);
  if (msg.indexOf('Ya existe un pedido') !== -1) return { ok: false, code: 'PEDIDO_EN_CURSO' };
  if (msg.indexOf('clave fiscal') !== -1) return { ok: false, code: 'SIN_CLAVE' };
  if (msg.indexOf('No tenés pacientes') !== -1) return { ok: false, code: 'SIN_PACIENTES' };
  if (msg.indexOf('captura de pantalla') !== -1) return { ok: false, code: 'SIN_DATOS' };
  return { ok: false, code: 'ERROR' };
}

function _botJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
