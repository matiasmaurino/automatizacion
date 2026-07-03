/***********************************************************
 * CONFIGURACIÓN
 ***********************************************************/
const EXENTOS_SS_ID = '1JjdVfUdiIhSMO1McU4FK2o0U2k-Qyd7clPDfQ0fzKjU';
const AUTOMATIZACION_SS_ID = '17xJc3GF9M3XkYJMmJn-LIzJ42ui1PyAmRwyU3BsvEHs';

const NOMBRE_HOJA_EXENTOS = 'EXENTOS';
const NOMBRE_HOJA_TABLA_AUX = 'TABLAS AUX';
const NOMBRE_HOJA_FACTURAR = 'WEBAPP';

const CUIT_IOMA = '30628249527';

const SERVICIOS_DISPONIBLES = [
  'Cuidado domiciliario',
  'Acompañamiento terapeutico',
  'Enfermeria domiciliaria'
];

const MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
  'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

/***********************************************************
 * ROUTER
 ***********************************************************/
function doGet(e) {
  const modulo = e && e.parameter ? e.parameter.modulo : null;

  if (modulo === 'facturacion') {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Facturación ARCA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (modulo === 'opcion') {
    return HtmlService.createHtmlOutputFromFile('opcion')
      .setTitle('Opción Monotributo - ARCA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (modulo === 'consulta') {
    return HtmlService.createHtmlOutputFromFile('consulta')
      .setTitle('Consultar Clave Fiscal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('AUTOMATIZACION ARCA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***********************************************************
 * HELPERS DE HOJAS
 ***********************************************************/
function _sheetExentos() {
  return SpreadsheetApp.openById(EXENTOS_SS_ID).getSheetByName(NOMBRE_HOJA_EXENTOS);
}
function _sheetTablaAux() {
  return SpreadsheetApp.openById(EXENTOS_SS_ID).getSheetByName(NOMBRE_HOJA_TABLA_AUX);
}
function _sheetFacturar() {
  return SpreadsheetApp.openById(AUTOMATIZACION_SS_ID).getSheetByName(NOMBRE_HOJA_FACTURAR);
}
function _soloDigitos(v) {
  return String(v || '').replace(/\D/g, '');
}

function _getPlanillaExentosExterna() {
  const ID_PLANILLA_EXTERNA = "1JjdVfUdiIhSMO1McU4FK2o0U2k-Qyd7clPDfQ0fzKjU";
  const ss = SpreadsheetApp.openById(ID_PLANILLA_EXTERNA);
  return ss.getSheetByName("Exentos");
}

/***********************************************************
 * BÚSQUEDA DE CLIENTE POR DNI (CON ADICIÓN DE COLUMNA K)
 ***********************************************************/
function buscarClientePorDni(dniInput) {
  const dni = _soloDigitos(dniInput);
  if (dni.length < 6) {
    throw new Error('Ingresá un DNI válido.');
  }
  
  const sheet = _getPlanillaExentosExterna();
  const data = sheet.getDataRange().getValues();
  const resultados = [];

  for (let r = 1; r < data.length; r++) {
    const cuit = _soloDigitos(data[r][1]); // Columna B (CUIT)
    if (cuit.length < 10) continue;

    const dniConCero = dni.length === 7 ? '0' + dni : dni;
    const dniSinCero = dni.replace(/^0+/, '');

    if (cuit.indexOf(dni) !== -1 ||
        cuit.indexOf(dniConCero) !== -1 ||
        (dniSinCero && cuit.indexOf(dniSinCero) !== -1)) {
      
      // Formateo de fecha de la Columna K (Índice 10)
      let fechaFormateada = 'No registra';
      if (data[r][10]) {
        try {
          fechaFormateada = Utilities.formatDate(new Date(data[r][10]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        } catch(e) {
          fechaFormateada = String(data[r][10]);
        }
      }

      resultados.push({
        fila: r + 1,
        cliente: data[r][0], // Columna A
        cuit: data[r][1],    // Columna B
        claveAfip: data[r][2],
        cuitReceptorPlanilla: _soloDigitos(data[r][5]),
        vencimientoTramiteK: fechaFormateada
      });
    }
  }
  return resultados;
}

/***********************************************************
 * MÓDULO FACTURACIÓN — parseo de pacientes
 ***********************************************************/
function _parsearLineaPaciente(linea, idx) {
  const numAfiliadoMatch = linea.match(/(\d{6,10})\s*\/\s*(\d{2})/);
  const dniExplicitoMatch = linea.match(/DNI\s*([\d.]+)/i);

  const estadoMatch = linea.match(
    /\b(OBLIGATORIO|OBLIGATRIO|OBLIGAGORIO|OBLIGAIGORIO|VOLUNTARIO\s+INDIVIDUAL|VOLUNTARIO|COLECTIVO)\b/i
  );

  let dni = '';
  if (dniExplicitoMatch) {
    dni = _soloDigitos(dniExplicitoMatch[1]);
  } else if (numAfiliadoMatch) {
    dni = _soloDigitos(numAfiliadoMatch[1]).replace(/^0+/, '');
  }

  let nombre = linea;
  const cortes = linea.match(/^[^\d]+/);
  if (cortes) nombre = cortes[0].trim();

  return {
    lineIndex: idx,
    textoOriginal: linea.trim(),
    nombre: nombre,
    numeroAfiliado: numAfiliadoMatch ? numAfiliadoMatch[1] : '',
    dni: dni,
    estado: estadoMatch ? estadoMatch[0] : ''
  };
}

function buscarPacientePorDni(filaCliente, dniPacienteInput) {
  const dni = _soloDigitos(dniPacienteInput);
  if (dni.length < 6) {
    throw new Error('Ingresá un DNI de paciente válido.');
  }

  const sheet = _sheetExentos();
  const celdaI = sheet.getRange(filaCliente, 9).getValue();
  const celdaJ = sheet.getRange(filaCliente, 10).getValue();

  if (!celdaI) {
    throw new Error('Esa fila no tiene pacientes cargados en la columna PACIENTE.');
  }

  const lineasI = String(celdaI).split('\n').filter(l => l.trim() !== '');
  const lineasJ = celdaJ ? String(celdaJ).split('\n').filter(l => l.trim() !== '') : [];

  const dniSinCero = dni.replace(/^0+/, '');

  for (let i = 0; i < lineasI.length; i++) {
    const p = _parsearLineaPaciente(lineasI[i], i);
    if (!p.dni) continue;

    if (p.dni === dni || p.dni === dniSinCero || ('0' + p.dni) === dni) {
      const lineaJ = lineasJ[i] !== undefined ? lineasJ[i] : (lineasJ.length === 1 ? lineasJ[0] : '');
      const tramites = String(lineaJ)
        .split(/[,;\s\/]+/)
        .map(t => t.trim())
        .filter(t => /^\d{4,}$/.test(t));

      p.tramites = tramites;
      return p;
    }
  }

  return null;
}

/***********************************************************
 * FILTRADO DINÁMICO DE PRESTACIONES ÚNICAS SIN REPETIR (COLUMNA H)
 ***********************************************************/
function listarServiciosUnicosCliente(filaCliente) {
  const sheet = _getPlanillaExentosExterna();
  const celdaH = sheet.getRange(filaCliente, 8).getValue(); // Columna H (Prestación)

  if (!celdaH) {
    return ["Cuidado domiciliario"]; // Valor por defecto si estuviese vacío
  }

  // Separamos por comas, saltos de línea o barras
  const serviciosSucios = String(celdaH).split(/[\n,;\/]+/);
  const serviciosLimpios = [];

  serviciosSucios.forEach(function(s) {
    const sTrimmed = s.trim();
    // CONDICIONAL CLAVE: Validamos que no esté vacío y que NO esté repetido en la lista
    if (sTrimmed !== '' && serviciosLimpios.indexOf(sTrimmed) === -1) {
      serviciosLimpios.push(sTrimmed);
    }
  });

  return serviciosLimpios;
}

/***********************************************************
 * CUITS RECEPTORES
 ***********************************************************/
function listarCuitsReceptor(filaCliente) {
  const sheet = _sheetExentos();
  const arr = [];

  if (filaCliente) {
    const celdaF = String(sheet.getRange(filaCliente, 6).getValue() || '');
    celdaF.split('\n').forEach(function(linea) {
      const v = _soloDigitos(linea.trim());
      if (v.length >= 10 && arr.indexOf(v) === -1) arr.push(v);
    });
  }

  if (arr.indexOf(CUIT_IOMA) === -1) arr.push(CUIT_IOMA);
  arr.sort(function(a, b) { return a === CUIT_IOMA ? -1 : b === CUIT_IOMA ? 1 : 0; });
  return arr;
}

function listarTodosLosCuitsReceptor() {
  const sheet = _sheetExentos();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
  const set = {};
  data.forEach(row => {
    const v = _soloDigitos(row[0]);
    if (v.length >= 10) set[v] = true;
  });
  set[CUIT_IOMA] = true;
  const arr = Object.keys(set);
  arr.sort((a, b) => (a === CUIT_IOMA ? -1 : b === CUIT_IOMA ? 1 : 0));
  return arr;
}

function listarServicios() {
  return SERVICIOS_DISPONIBLES;
}

function listarServiciosCliente(filaCliente) {
  const sheet = _sheetExentos();
  const celdaH = sheet.getRange(filaCliente, 8).getValue();

  if (!celdaH) {
    throw new Error('Esa fila no tiene ninguna prestación cargada en la columna H.');
  }

  const servicios = String(celdaH)
    .split(/[\n,;\/]+/)
    .map(s => s.trim())
    .filter(s => s !== '');

  if (servicios.length === 0) {
    throw new Error('Esa fila no tiene ninguna prestación cargada en la columna H.');
  }

  return servicios;
}

/***********************************************************
 * TABLAS AUX — valor hora y resolución
 * G=clave, H=PERIODO, I=RES IOMA, J=SERVICIO,
 * K=VALOR HOI, L=MES PARA ORDEN, M=Aumento (retroactivo)
 ***********************************************************/
function _getValorYResolucion(servicio, mesNombre, anio) {
  const sheet = _sheetTablaAux(); 
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  
  // Leemos desde la fila 2, columna 1 (A) hasta la última fila y columna M (13 columnas)
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  
  // Normalizamos lo que el usuario ingresó para buscar (en mayúsculas y sin espacios extras)
  const periodoBuscado = (mesNombre + ' ' + anio).toUpperCase().trim();
  const servicioBuscado = String(servicio).toUpperCase().trim();
  
  for (let i = 0; i < data.length; i++) {
    // Columna H es el índice 7 (A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7)
    const periodoPlanilla = String(data[i][7]).toUpperCase().trim();
    
    // Columna J es el índice 9 (I=8, J=9)
    const servicioPlanilla = String(data[i][9]).toUpperCase().trim();
    
    // Si coinciden de forma independiente el período y el servicio:
    if (periodoPlanilla === periodoBuscado && servicioPlanilla === servicioBuscado) {
      return {
        resolucion: data[i][8],  // Columna I (índice 8)
        valorHora:  data[i][10], // Columna K (índice 10)
        valorHoraM: data[i][12]  // Columna M (índice 12)
      };
    }
  }
  return null;
}
/***********************************************************
 * FECHAS
 ***********************************************************/
function _mesARango(mesNumero, anio) {
  const tz = Session.getScriptTimeZone();
  const desde        = new Date(anio, mesNumero - 1, 1);
  const hasta        = new Date(anio, mesNumero, 0);
  const fechaEmision = new Date(anio, mesNumero, 1); // 1ro mes siguiente (vencimiento)

  const hoy = new Date();
  const esMesEnCurso = (mesNumero === (hoy.getMonth() + 1)) && (Number(anio) === hoy.getFullYear());
  const fechaFactura = esMesEnCurso ? fechaEmision : hoy;

  return {
    desde:        Utilities.formatDate(desde,        tz, 'dd/MM/yyyy'),
    hasta:        Utilities.formatDate(hasta,        tz, 'dd/MM/yyyy'),
    fechaEmision: Utilities.formatDate(fechaEmision, tz, 'dd/MM/yyyy'),
    fechaFactura: Utilities.formatDate(fechaFactura, tz, 'dd/MM/yyyy')
  };
}

/***********************************************************
 * GUARDAR FACTURA EN HOJA WEBAPP
 ***********************************************************/
function guardarFactura(payload) {
  const rango = _mesARango(Number(payload.mesNumero), Number(payload.anio)); //
  const mesNombre = MESES[Number(payload.mesNumero) - 1]; //

  const vr = _getValorYResolucion(payload.servicio, mesNombre, payload.anio); //
  if (!vr) {
    throw new Error('No se encontró el valor de hora / resolución para ese servicio.'); //
  }

  const cuitReceptorFinal = payload.retroactivo ? CUIT_IOMA : payload.cuitReceptor; //
  
  // 🟢 ACTUALIZACIÓN: Removimos payload.estado de la descripción y añadimos el Vencimiento del Trámite (Columna K)
  let descripcion =
    payload.servicio + ' ' +
    payload.pacienteNombre + ' ' +
    payload.numeroAfiliado + '/00 ' +
    'DNI ' + payload.dniPaciente + ' ' +
    'tramite ' + payload.numeroTramite + ' ' +
    '(Vence: ' + payload.vencimientoTramiteK + ') ' +
    'segun resolucion ' + vr.resolucion + ' ' +
    'del mes de ' + mesNombre + ' ' + payload.anio + ' ' +
    'por ' + payload.horas + ' horas a un valor de $' + vr.valorHora; //

  if (payload.retroactivo) {
    descripcion = 'RETROACTIVO de Factura Pto.Vta ' + payload.puntoVenta +
      ' Nro ' + payload.nroComprobante + ' — ' + descripcion; //
  }

  const sheet = _sheetFacturar(); //
  sheet.appendRow([
    payload.cuit,            // A
    payload.claveAfip,       // B
    payload.clienteNombre,   // C
    rango.fechaFactura,      // D
    'Factura C',             // E
    rango.desde,             // F
    rango.hasta,             // G
    rango.fechaEmision,      // H
    cuitReceptorFinal,       // I
    'Exento',                // J
    descripcion,             // K
    Number(payload.horas),   // L
    'otras unidades',        // M
    vr.valorHora             // N
  ]);
  return { ok: true }; //
}

/***********************************************************
 * MÓDULO OPCIÓN MONOTRIBUTO
 ***********************************************************/
function guardarOpcion(payload) {
  const cuit  = _soloDigitos(payload.cuit);
  const clave = String(payload.claveAfip || '').trim();
  const nombre = String(payload.clienteNombre || '').trim();

  if (cuit.length < 10) throw new Error('CUIT inválido.');
  if (!clave) throw new Error('No se encontró la clave fiscal para ese CUIT.');

  const sheet = _sheetFacturar();

  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const cuitFila = _soloDigitos(data[r][0]);
    const estado   = String(data[r][15] || '').trim().toLowerCase();
    if (cuitFila === cuit && estado !== 'descargada') {
      throw new Error('Ese CUIT ya está cargado en WEBAPP y todavía no fue descargado.');
    }
  }

  sheet.appendRow([
    cuit, clave, nombre,
    '', '', '', '', '', '', '', '', '', '', ''
  ]);

  return { ok: true, cuit: cuit, nombre: nombre };
}

/***********************************************************
 * MÓDULO CONSULTAR CLAVE FISCAL ARCA / ARBA
 ***********************************************************/
function generarOpcionesVerificacion(filaCliente) {
  const sheet = _sheetExentos();
  const lastRow = sheet.getLastRow();

  const nombreCorrecto = String(sheet.getRange(filaCliente, 1).getValue() || '').trim();
  if (!nombreCorrecto) throw new Error('No se pudo obtener el nombre del cliente.');

  const todosNombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => String(r[0] || '').trim())
    .filter(n => n && n !== nombreCorrecto);

  const unicos = Array.from(new Set(todosNombres));
  for (let i = unicos.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [unicos[i], unicos[j]] = [unicos[j], unicos[i]];
  }
  const distractores = unicos.slice(0, 3);

  const opciones = distractores.concat([nombreCorrecto]);
  for (let i = opciones.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [opciones[i], opciones[j]] = [opciones[j], opciones[i]];
  }

  return opciones;
}

function verificarYObtenerDatosClave(filaCliente, nombreElegido) {
  const sheet = _sheetExentos();
  const nombreReal = String(sheet.getRange(filaCliente, 1).getValue() || '').trim();

  if (!nombreElegido || nombreElegido.trim() !== nombreReal) {
    throw new Error('La selección no coincide. Volvé a intentarlo desde el DNI.');
  }

  const fila = sheet.getRange(filaCliente, 1, 1, 5).getValues()[0];

  return {
    cliente:          fila[0],
    cuit:             fila[1],
    claveArca:        fila[2],
    claveArba:        fila[3],
    vencimientoAlas:  fila[4]
  };
}

/***********************************************************
 * VERIFICAR SI HAY RETROACTIVO PARA UN MES/SERVICIO
 * Devuelve el valor de col M (0 si no hay retroactivo)
 ***********************************************************/
function obtenerTarifaAuxiliar(servicio, mesNumero, anio) {
  const mesNombre = MESES[Number(mesNumero) - 1];
  const vr = _getValorYResolucion(servicio, mesNombre, anio);
  return vr ? Number(vr.valorHoraM) : 0;
}

/***********************************************************
 * GENERAR OPCIONES DE VERIFICACIÓN PARA MÚLTIPLES FILAS
 * (consulta.html: cuando el mismo DNI matchea varios clientes)
 ***********************************************************/
function generarOpcionesMultiFila(filas) {
  const sheet = _sheetExentos();
  const lastRow = sheet.getLastRow();

  // Nombres correctos (todos los clientes que matchearon)
  const nombresCorrectos = filas.map(function(fila) {
    return String(sheet.getRange(fila, 1).getValue() || '').trim();
  }).filter(function(n) { return n; });

  // Distractores: nombres de la hoja que NO son correctos
  const todosNombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(function(r) { return String(r[0] || '').trim(); })
    .filter(function(n) { return n && nombresCorrectos.indexOf(n) === -1; });

  // Mezclar distractores
  for (let i = todosNombres.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = todosNombres[i]; todosNombres[i] = todosNombres[j]; todosNombres[j] = tmp;
  }

  // Tomar suficientes distractores para llegar a 4 opciones en total
  const cantDistractores = Math.max(0, 4 - nombresCorrectos.length);
  const distractores = todosNombres.slice(0, cantDistractores);

  // Mezclar opciones finales
  const opciones = nombresCorrectos.concat(distractores);
  for (let i = opciones.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = opciones[i]; opciones[i] = opciones[j]; opciones[j] = tmp;
  }

  return { opciones: opciones, nombresCorrectos: nombresCorrectos };
}

/***********************************************************
 * MÓDULO CONSULTA — nuevo flujo por email
 * Col G de EXENTOS = EMAIL del cliente
 * Escribe en hoja "DATOS PERSONALES" y envía email
 ***********************************************************/
function enviarDatosPersonales(dniInput) {
  const dni = _soloDigitos(dniInput);
  if (dni.length < 6) throw new Error('Ingresá un DNI válido.');

  const sheet = _sheetExentos();
  const data  = sheet.getDataRange().getValues();
  let cliente = null;

  for (let r = 1; r < data.length; r++) {
    const cuit = _soloDigitos(data[r][1]);
    if (cuit.length < 10) continue;
    const dniConCero = dni.length === 7 ? '0' + dni : dni;
    const dniSinCero = dni.replace(/^0+/, '');
    if (cuit.indexOf(dni) !== -1 || cuit.indexOf(dniConCero) !== -1 ||
        (dniSinCero && cuit.indexOf(dniSinCero) !== -1)) {
      cliente = {
        nombre:          String(data[r][0] || '').trim(),
        cuit:            String(data[r][1] || '').trim(),
        claveArca:       String(data[r][2] || '').trim(),
        claveArba:       String(data[r][3] || '').trim(),
        vencimientoAlas: data[r][4],
        email:           String(data[r][6] || '').trim()
      };
      break;
    }
  }

  if (!cliente) throw new Error('No encontramos ese DNI en nuestra base de datos.');

  // Guardar pedido en hoja DATOS PERSONALES
  const ssAuto = SpreadsheetApp.openById(AUTOMATIZACION_SS_ID);
  let sheetDP  = ssAuto.getSheetByName('DATOS PERSONALES');
  if (!sheetDP) {
    sheetDP = ssAuto.insertSheet('DATOS PERSONALES');
    sheetDP.getRange(1, 1, 1, 7).setValues([[
      'CUIT','CLIENTE','CLAVE ARCA','CLAVE ARBA','VENCIMIENTO ALAS','EMAIL','FECHA SOLICITUD'
    ]]);
  }
  sheetDP.appendRow([
    cliente.cuit,
    cliente.nombre,
    cliente.claveArca,
    cliente.claveArba,
    cliente.vencimientoAlas,
    cliente.email,
    new Date()
  ]);

return 'ok';
}
function registrarPedidoConsulta(fila) {
  const sheet = _sheetExentos();
  const row = sheet.getRange(fila, 1, 1, 7).getValues()[0];
  const ssAuto = SpreadsheetApp.openById(AUTOMATIZACION_SS_ID);
  let sheetDP = ssAuto.getSheetByName('DATOS PERSONALES');
  if (!sheetDP) {
    sheetDP = ssAuto.insertSheet('DATOS PERSONALES');
    sheetDP.getRange(1,1,1,7).setValues([['CUIT','CLIENTE','CLAVE ARCA','CLAVE ARBA','VENCIMIENTO ALAS','EMAIL','FECHA SOLICITUD']]);
  }
  sheetDP.appendRow([row[1], row[0], row[2], row[3], row[4], row[6], new Date()]);
  return 'ok';
}