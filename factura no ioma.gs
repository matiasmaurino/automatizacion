// 1. Desglose de servicios de la columna H del renglón actual
function listarServiciosClienteNoIoma(filaCliente) {
  try {
    const sheet = _sheetExentos();
    const celdaH = sheet.getRange(filaCliente, 8).getValue();

    if (!celdaH) {
      return SERVICIOS_DISPONIBLES; 
    }

    const servicios = String(celdaH)
      .split(/[\n,;\/]+/)
      .map(s => s.trim())
      .filter(s => s !== '');

    return servicios.length === 0 ? SERVICIOS_DISPONIBLES : servicios;
  } catch(e) {
    return SERVICIOS_DISPONIBLES;
  }
}

// 2. Desglose de CUITs Receptores de la columna F del renglón actual
function listarCuitsReceptorClienteNoIoma(filaCliente) {
  try {
    const sheet = _sheetExentos();
    const celdaF = sheet.getRange(filaCliente, 6).getValue();
    const arr = [];

    if (celdaF) {
      String(celdaF).split(/[\n,;\/]+/).forEach(function(linea) {
        const v = String(linea || '').replace(/\D/g, '');
        if (v.length >= 10 && arr.indexOf(v) === -1) {
          arr.push(v);
        }
      });
    }

    if (arr.indexOf(CUIT_IOMA) === -1 && arr.length === 0) {
      arr.push(CUIT_IOMA);
    }
    
    return arr;
  } catch(e) {
    return [CUIT_IOMA];
  }
}

// 3. Función exclusiva de guardado
function guardarFacturaNoIoma(payload) {
  const rango = _mesARango(Number(payload.mesNumero), Number(payload.anio));
  
  let mesNombre = '';
  if (!isNaN(payload.mesNumero) && Number(payload.mesNumero) >= 1 && Number(payload.mesNumero) <= 12) {
    mesNombre = MESES[Number(payload.mesNumero) - 1];
  } else {
    mesNombre = String(payload.mesNumero).toUpperCase().trim();
  }

  const pacienteTexto = payload.paciente ? (' ' + payload.paciente) : '';

  let descripcion = '';
  let cantidadCeldas = 1;
  let precioUnitarioCeldas = Number(payload.importeTotal);

  if (payload.modoFacturacion === 'horas') {
    descripcion = payload.servicio + pacienteTexto + ' — del mes de ' + mesNombre + ' ' + payload.anio + ' por ' + payload.horas + ' horas a un valor de $' + payload.valorHoraManual;
    cantidadCeldas = Number(payload.horas);
    precioUnitarioCeldas = Number(payload.valorHoraManual);
  } else {
    descripcion = payload.servicio + pacienteTexto + ' — Prestación correspondiente al mes de ' + mesNombre + ' ' + payload.anio;
  }

  const sheet = _sheetFacturar();
  sheet.appendRow([
    payload.cuit,                              // A
    payload.claveAfip,                         // B
    payload.clienteNombre,                     // C
    rango.fechaFactura,                        // D FECHA
    'Factura C',                               // E
    rango.desde,                               // F DESDE
    rango.hasta,                               // G HASTA
    rango.fechaEmision,                        // H VENCIMIENTO
    payload.cuitReceptor,                      // I
    obtenerCondicionIva(payload.cuitReceptor), // J — condición IVA del receptor
    descripcion,                               // K
    cantidadCeldas,                            // L Cant
    'otras unidades',                          // M
    precioUnitarioCeldas,                      // N Prec
    '',                                        // O FACTURA
    '',                                        // P OPCION Y CREDENCIAL
    '',                                        // Q CAE
    payload.email || ''                        // R EMAIL
  ]);

  // Agregar cuitReceptor a CONDICION IVA si no existe todavía
  const ssAuto = SpreadsheetApp.openById(AUTOMATIZACION_SS_ID);
  const sheetCondIva = ssAuto.getSheetByName('CONDICION IVA');
  if (sheetCondIva) {
    const cuitReceptorLimpio = String(payload.cuitReceptor || '').replace(/\D/g, '');
    const dataCond = sheetCondIva.getDataRange().getValues();
    const yaExiste = dataCond.some(function(fila) {
      return String(fila[0] || '').replace(/\D/g, '') === cuitReceptorLimpio;
    });
    if (!yaExiste && cuitReceptorLimpio.length >= 10) {
      sheetCondIva.appendRow([cuitReceptorLimpio, '']);
    }
  }

  return {
    ok: true,
    descripcion: descripcion
  };
}

// 4. Desglose de Pacientes de la columna L del renglón actual
function listarPacientesClienteNoIoma(filaCliente) {
  try {
    const sheet = _sheetExentos();
    const celdaL = sheet.getRange(filaCliente, 12).getValue();

    if (!celdaL) {
      return [];
    }

    const pacientes = String(celdaL)
      .split(/[\n,;]+/)
      .map(s => s.trim())
      .filter(s => s !== '');

    return pacientes;
  } catch(e) {
    return [];
  }
}

// 5. Obtener condición IVA desde hoja CONDICION IVA
function obtenerCondicionIva(cuit) {
  const ss = SpreadsheetApp.openById(AUTOMATIZACION_SS_ID);
  const sheet = ss.getSheetByName('CONDICION IVA');
  if (!sheet) return 'Exento';

  const data = sheet.getDataRange().getValues();
  const cuitBuscado = String(cuit || '').replace(/\D/g, '');

  for (let r = 0; r < data.length; r++) {
    const cuitFila = String(data[r][0] || '').replace(/\D/g, '');
    if (cuitFila.length >= 10 && cuitFila === cuitBuscado) {
      return String(data[r][1] || '').trim() || 'Exento';
    }
  }
  return 'Exento';
}