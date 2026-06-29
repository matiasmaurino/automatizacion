// =========================================================================
// CONFIGURACIÓN PRINCIPAL (Seguridad Blindada: la API Key se lee de forma oculta)
// =========================================================================
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const CARPETA_ORIGEN_ID = '1cLnlPOvel1V7q-Syegm0KGWNd0F4a_Ws';

// Crear el menú unificado en tu Google Sheets
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧾 Actualizacion') // Menú principal solicitado
    .addItem('CAE', 'procesarFacturas') // Subtarea 1
    .addItem('Actualizar CUIT y CLAVES', 'actualizarcuityclaves') // Subtarea 2
    .addToUi();
}

// =========================================================================
// TAREA 1: PROCESAR FACTURAS (EXTRACCIÓN CON IA)
// =========================================================================
function procesarFacturas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CAE');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Error: No se encontró la pestaña llamada "CAE". Asegurate de que el nombre sea exacto.');
    return;
  }

  // Validación de seguridad por si no se guardó la propiedad oculta aún
  if (!GEMINI_API_KEY) {
    SpreadsheetApp.getUi().alert('❌ Error: No se configuró la GEMINI_API_KEY en las Propiedades del Script (Configuración ⚙️).');
    return;
  }
  
  let carpetaOrigen;
  try {
    carpetaOrigen = DriveApp.getFolderById(CARPETA_ORIGEN_ID);
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: El ID de la carpeta origen es incorrecto.');
    return; 
  }
  
  const archivos = carpetaOrigen.getFilesByType(MimeType.PDF);
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange('A1').setValue('Nombre del Archivo');
    sheet.getRange('B1').setValue('URL QR (Clic para CAE)'); 
    sheet.getRange('A1:B1').setFontWeight('bold');
  }
  
  const ultimaFila = sheet.getLastRow(); 
  let archivosYaRegistrados = [];
  if (ultimaFila > 1) {
    archivosYaRegistrados = sheet.getRange(2, 1, ultimaFila - 1, 1).getValues().map(row => row[0]);
  } 
  
  let contador = 0;
  while (archivos.hasNext()) {
    const archivo = archivos.next(); 
    const nombreArchivo = archivo.getName();
    
    if (archivosYaRegistrados.includes(nombreArchivo)) {
      continue;
    }
    if (nombreArchivo.toLowerCase().includes('credencial') || nombreArchivo.toLowerCase().includes('opcion')) {
      continue;
    }
    
    try {
      const datosFactura = extraerDatosConGemini(archivo);
      if (datosFactura) { 
        const urlCompletaARCA = generarUrlOficialQR(datosFactura);
        const proximaFila = sheet.getLastRow() + 1;
        sheet.getRange(proximaFila, 1).setValue(nombreArchivo); 
        sheet.getRange(proximaFila, 2).setValue(urlCompletaARCA);
        
        contador++;
      }
    } catch (e) {
      Logger.log('Error en archivo ' + nombreArchivo + ': ' + e.toString());
    }
  }
  
  SpreadsheetApp.getUi().alert('Proceso terminado. Se procesaron ' + contador + ' archivos.');
}

function extraerDatosConGemini(archivo) {
  const blob = archivo.getBlob();
  const base64Pdf = Utilities.base64Encode(blob.getBytes()); 
  
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + GEMINI_API_KEY;
  const prompt = "Analiza el texto de esta factura de AFIP/ARCA de Argentina y extrae los siguientes datos numéricos exactos para armar el JSON del QR oficial. " + 
                 "Devuelve ÚNICAMENTE un objeto JSON válido con estas claves exactas (sin formato markdown, sin texto extra, solo el objeto puro): " + 
                 "{\n" +
                 "  \"fecha\": \"fecha de emision en formato AAAA-MM-DD\",\n" + 
                 "  \"cuitEmisor\": número de CUIT del emisor sin guiones,\n" + 
                 "  \"ptoVta\": número del punto de venta,\n" + 
                 "  \"tipoCmp\": 11,\n" + 
                 "  \"nroCmp\": número de comprobante,\n" + 
                 "  \"importe\": importe total como número (usa punto para decimales),\n" + 
                 "  \"cuitReceptor\": número de CUIT del receptor/cliente sin guiones,\n" + 
                 "  \"cae\": número de CAE de 14 dígitos\n" + 
                 "}";

  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inlineData": { "mimeType": "application/pdf", "data": base64Pdf } } 
      ]
    }]
  };
  const opciones = {
    "method": "post", 
    "contentType": "application/json", 
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  };
  const respuesta = UrlFetchApp.fetch(url, opciones); 
  const jsonRespuesta = JSON.parse(respuesta.getContentText());
  
  if (jsonRespuesta.candidates && jsonRespuesta.candidates[0].content.parts[0].text) {
    let textoJson = jsonRespuesta.candidates[0].content.parts[0].text.trim();
    textoJson = textoJson.replace(/```json/g, '').replace(/```/g, '').trim();
    return JSON.parse(textoJson); 
  }
  return null;
}

function generarUrlOficialQR(datos) {
  const objetoQrAfip = {
    "ver": 1,
    "fecha": datos.fecha,
    "cuit": parseInt(datos.cuitEmisor, 10),
    "ptoVta": parseInt(datos.ptoVta, 10),
    "tipoCmp": parseInt(datos.tipoCmp, 10),
    "nroCmp": parseInt(datos.nroCmp, 10),
    "importe": parseFloat(datos.importe),
    "moneda": "PES",
    "ctz": 1,
    "tipoDocRec": 80, 
    "nroDocRec": parseInt(datos.cuitReceptor, 10),
    "tipoCodAut": "E", 
    "codAut": parseInt(datos.cae, 10) 
  };
  
  const cadenaJson = JSON.stringify(objetoQrAfip);
  const base64Token = Utilities.base64Encode(cadenaJson, Utilities.Charset.UTF_8);
  return 'https://servicioscf.afip.gob.ar/publico/comprobantes/cae.aspx?p=' + base64Token; 
}

// =========================================================================
// TAREA 2: ACTUALIZAR CUIT Y CLAVES (EMULACIÓN DE ARRAYFORMULA + BUSCARX)
// =========================================================================
function actualizarcuityclaves() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtener la hoja principal (asumo que es la primera pestaña)
  const hojaPrincipal = ss.getSheets()[0]; 
  
  // Obtener la hoja de referencia
  const hojaOrigen = ss.getSheetByName('CUIT y CLAVES');
  if (!hojaOrigen) {
    Logger.log("No se encontró la pestaña 'CUIT y CLAVES'");
    return;
  }

  const ultimaFilaOrigen = hojaOrigen.getLastRow();
  if (ultimaFilaOrigen < 1) return;
  
  // Leemos desde la columna I (9) hasta la K (11), es decir, 3 columnas (I, J, K)
  const datosOrigen = hojaOrigen.getRange(1, 9, ultimaFilaOrigen, 3).getValues();
  const mapaBuscarX = new Map();
  
  datosOrigen.forEach(fila => {
    const clave = fila[0].toString().trim(); // Columna I
    const valorJ = fila[1];                  // Columna J
    const valorK = fila[2];                  // Columna K
    
    if (clave !== "") {
      mapaBuscarX.set(clave, { j: valorJ, k: valorK });
    }
  });

  const ultimaFilaPrincipal = hojaPrincipal.getLastRow();
  if (ultimaFilaPrincipal < 2) return; 
  
  const valoresC = hojaPrincipal.getRange(2, 3, ultimaFilaPrincipal - 1, 1).getValues();
  const nuevosValoresResultado = []; 

  valoresC.forEach(fila => {
    const valorC = fila[0].toString().trim();
    
    if (valorC !== "" && mapaBuscarX.has(valorC)) {
      const coincidencia = mapaBuscarX.get(valorC);
      nuevosValoresResultado.push([coincidencia.j, coincidencia.k]);
    } else {
      nuevosValoresResultado.push(["", ""]);
    }
  });

  // Escribe los resultados simultáneamente en las columnas D y E de la hoja principal
  hojaPrincipal.getRange(2, 4, nuevosValoresResultado.length, nuevosValoresResultado[0].length).setValues(nuevosValoresResultado);
}