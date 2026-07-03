// 1. Desglose de servicios de la columna H del renglón actual
function listarServiciosClienteNoIoma(filaCliente) {
  try {
    const sheet = _sheetExentos();
    const celdaH = sheet.getRange(filaCliente, 8).getValue(); // Columna H

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
    const celdaF = sheet.getRange(filaCliente, 6).getValue(); // Columna F
    const arr = [];

    if (celdaF) {
      String(celdaF).split(/[\n,;\/]+/).forEach(function(linea) {
        const v = String(linea || '').replace(/\D/g, ''); // Solo dígitos
        if (v.length >= 10 && arr.indexOf(v) === -1) {
          arr.push(v);
        }
      });
    }

    // Por defecto, si está vacío o querés asegurar que esté IOMA en la lista
    if (arr.indexOf(CUIT_IOMA) === -1 && arr.length === 0) {
      arr.push(CUIT_IOMA);
    }
    
    return arr;
  } catch(e) {
    return [CUIT_IOMA];
  }
}

// 3. Función exclusiva de guardado simplificado
function guardarFacturaNoIoma(payload) {
  const rango = _mesARango(Number(payload.mesNumero), Number(payload.anio));
  
  let mesNombre = '';
  if (!isNaN(payload.mesNumero) && Number(payload.mesNumero) >= 1 && Number(payload.mesNumero) <= 12) {
    mesNombre = MESES[Number(payload.mesNumero) - 1];
  } else {
    mesNombre = String(payload.mesNumero).toUpperCase().trim();
  }

  // Armamos la descripción y celdas correspondientes según el modo de entrada elegido
  let descripcion = '';
  let cantidadCeldas = 1;
  let precioUnitarioCeldas = Number(payload.importeTotal);

  if (payload.modoFacturacion === 'horas') {
    descripcion = payload.servicio + ' — del mes de ' + mesNombre + ' ' + payload.anio + ' por ' + payload.horas + ' horas a un valor de $' + payload.valorHoraManual;
    cantidadCeldas = Number(payload.horas);
    precioUnitarioCeldas = Number(payload.valorHoraManual);
  } else {
    // Si es Importe Directo, no discriminamos horas en el texto
    descripcion = payload.servicio + ' — Prestación correspondiente al mes de ' + mesNombre + ' ' + payload.anio;
  }

  const sheet = _sheetFacturar();
  sheet.appendRow([
    payload.cuit,            // A
    payload.claveAfip,       // B
    payload.clienteNombre,   // C
    rango.fechaFactura,      // D FECHA
    'Factura C',             // E
    rango.desde,             // F DESDE
    rango.hasta,             // G HASTA
    rango.fechaEmision,      // H VENCIMIENTO
    payload.cuitReceptor,    // I
    'Exento',                // J
    descripcion,             // K
    cantidadCeldas,          // L Cant
    'otras unidades',        // M
    precioUnitarioCeldas     // N Prec
  ]);

  return {
    ok: true,
    descripcion: descripcion
  };
}