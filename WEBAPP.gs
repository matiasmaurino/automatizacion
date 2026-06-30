/***********************************************************
 * CONFIGURACIÓN — AJUSTÁ ACÁ SI CAMBIAN IDs O NOMBRES DE HOJA
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
 * REEMPLAZAR el doGet() actual de Code.gs por este,
 * que dirige a cada módulo según el parámetro ?modulo=
 ***********************************************************/
function doGet(e) {
  const modulo = e && e.parameter ? e.parameter.modulo : null;

  if (modulo === 'facturacion') {
    return HtmlService.createHtmlOutputFromFile('index') // tu index.html actual de facturación
      .setTitle('Facturación ARCA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (modulo === 'opcion') {
    return HtmlService.createHtmlOutputFromFile('opcion')
      .setTitle('Opción Monotributo - ARCA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Sin parámetro -> pantalla inicial con selector de módulo
  return HtmlService.createHtmlOutputFromFile('Inicio')
    .setTitle('AUTOMATIZACION ARCA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


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

function buscarClientePorDni(dniInput) {
  const dni = _soloDigitos(dniInput);
  if (dni.length < 6) {
    throw new Error('Ingresá un DNI válido.');
  }
  const sheet = _sheetExentos();
  const data = sheet.getDataRange().getValues();
  const resultados = [];

  for (let r = 1; r < data.length; r++) {
    const cuit = _soloDigitos(data[r][1]); // columna B
    if (cuit.length < 10) continue;

    const dniConCero = dni.length === 7 ? '0' + dni : dni;
    const dniSinCero = dni.replace(/^0+/, '');

    if (cuit.indexOf(dni) !== -1 ||
        cuit.indexOf(dniConCero) !== -1 ||
        (dniSinCero && cuit.indexOf(dniSinCero) !== -1)) {
      resultados.push({
        fila: r + 1,
        cliente: data[r][0],
        cuit: data[r][1],
        claveAfip: data[r][2],
        cuitReceptorPlanilla: _soloDigitos(data[r][5])
      });
    }
  }
  return resultados;
}

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

function listarCuitsReceptor(filaCliente) {
  const sheet = _sheetExentos();
  const arr = [];

  if (filaCliente) {
    const cuitFila = _soloDigitos(sheet.getRange(filaCliente, 6).getValue());
    if (cuitFila.length >= 10) arr.push(cuitFila);
  }

  if (arr.indexOf(CUIT_IOMA) === -1) arr.push(CUIT_IOMA);
  arr.sort((a, b) => (a === CUIT_IOMA ? -1 : b === CUIT_IOMA ? 1 : 0));
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

function _getValorYResolucion(servicio, mesNombre, anio) {
  const sheet = _sheetTablaAux();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 7, lastRow - 1, 6).getValues();

  const claveBuscada = (mesNombre + ' ' + anio + servicio).toUpperCase().replace(/\s+/g, ' ').trim();

  for (let i = 0; i < data.length; i++) {
    const clave = String(data[i][0]).toUpperCase().replace(/\s+/g, ' ').trim();
    if (clave === claveBuscada) {
      return {
        resolucion: data[i][2],
        valorHora: data[i][4]
      };
    }
  }
  return null;
}

function _mesARango(mesNumero, anio) {
  const tz = Session.getScriptTimeZone();
  const desde = new Date(anio, mesNumero - 1, 1);
  const hasta = new Date(anio, mesNumero, 0);
  const fechaEmision = new Date(anio, mesNumero, 1); // vencimiento, siempre 1ro del mes siguiente

  const hoy = new Date();
  const esMesEnCurso = (mesNumero === (hoy.getMonth() + 1)) && (Number(anio) === hoy.getFullYear());
  // Columna D (FECHA): si es el mes en curso, 1ro del mes siguiente; si es mes anterior, hoy.
  const fechaFactura = esMesEnCurso ? fechaEmision : hoy;

  return {
    desde: Utilities.formatDate(desde, tz, 'dd/MM/yyyy'),
    hasta: Utilities.formatDate(hasta, tz, 'dd/MM/yyyy'),
    fechaEmision: Utilities.formatDate(fechaEmision, tz, 'dd/MM/yyyy'),
    fechaFactura: Utilities.formatDate(fechaFactura, tz, 'dd/MM/yyyy')
  };
}

function guardarFactura(payload) {
  const rango = _mesARango(Number(payload.mesNumero), Number(payload.anio));
  const mesNombre = MESES[Number(payload.mesNumero) - 1];

  const vr = _getValorYResolucion(payload.servicio, mesNombre, payload.anio);
  if (!vr) {
    throw new Error('No se encontró el valor de hora / resolución para ese servicio y período en TABLAS AUX. Revisá que el período exista en esa hoja.');
  }

  const descripcion =
    payload.servicio + ' ' +
    payload.pacienteNombre + ' ' +
    payload.numeroAfiliado + '/00 ' +
    payload.estado + ' ' +
    'DNI ' + payload.dniPaciente + ' ' +
    'tramite ' + payload.numeroTramite + ' ' +
    'segun resolucion ' + vr.resolucion + ' ' +
    'del mes de ' + mesNombre + ' ' + payload.anio + ' ' +
    'por ' + payload.horas + ' horas a un valor de $' + vr.valorHora;

  const sheet = _sheetFacturar();
  sheet.appendRow([
    payload.cuit,
    payload.claveAfip,
    payload.clienteNombre,
    rango.fechaFactura,
    'Factura C',
    rango.desde,
    rango.hasta,
    rango.fechaEmision,
    payload.cuitReceptor,
    'Exento',
    descripcion,
    Number(payload.horas),
    'otras unidades',
    vr.valorHora
  ]);

  return { ok: true, valorHora: vr.valorHora, resolucion: vr.resolucion, descripcion: descripcion };
}

/***********************************************************
 * REEMPLAZAR el doGet() actual de Code.gs por este,
 * que dirige a cada módulo según el parámetro ?modulo=
 ***********************************************************/
function doGet(e) {
  const modulo = e && e.parameter ? e.parameter.modulo : null;

  if (modulo === 'facturacion') {
    return HtmlService.createHtmlOutputFromFile('index') // tu index.html actual de facturación
      .setTitle('Facturación ARCA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (modulo === 'opcion') {
    return HtmlService.createHtmlOutputFromFile('opcion')
      .setTitle('Opción Monotributo - ARCA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Sin parámetro -> pantalla inicial con selector de módulo
  return HtmlService.createHtmlOutputFromFile('Inicio')
    .setTitle('AUTOMATIZACION ARCA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
