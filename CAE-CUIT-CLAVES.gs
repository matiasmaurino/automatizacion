// =========================================================================
// CONFIGURACIÓN PRINCIPAL (Seguridad Blindada: la API Key se lee de forma oculta)
// =========================================================================
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const CARPETA_ORIGEN_ID = '1cLnlPOvel1V7q-Syegm0KGWNd0F4a_Ws';

// ID de la hoja de cálculo origen (extraído de tu fórmula IMPORTRANGE)
const SPREADSHEET_ORIGEN_ID = '1JjdVfUdiIhSMO1McU4FK2o0U2k-Qyd7clPDfQ0fzKjU'; 

// Crear el menú unificado en tu Google Sheets
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧾 Actualizacion') // Menú principal solicitado
    .addItem('CAE', 'procesarFacturas') // Subtarea 1
    .addItem('Actualizar CUIT y CLAVES', 'actualizarCuitYClaves') // Subtarea 2 (Nueva)
    .addToUi();
}

// =========================================================================
// TAREA 2: ACTUALIZAR CUIT Y CLAVES (REEMPLAZO ESTABLE DE IMPORTRANGE)
// =========================================================================
function actualizarCuitYClaves() {
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDestino = ssDestino.getSheetByName('CUIT y CLAVES');
  
  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert('❌ Error: No se encontró la pestaña llamada "CUIT y CLAVES".');
    return;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Conectando con la base de datos origen...', 'Actualización', 3);
  
  try {
    const ssOrigen = SpreadsheetApp.openById(SPREADSHEET_ORIGEN_ID);
    const hojaOrigen = ssOrigen.getSheetByName('EXENTOS');
    
    if (!hojaOrigen) {
      SpreadsheetApp.getUi().alert('❌ Error: No se encontró la pestaña "EXENTOS" en el archivo origen.');
      return;
    }
    
    const ultimaFilaOrigen = hojaOrigen.getLastRow();
    if (ultimaFilaOrigen === 0) {
      SpreadsheetApp.getUi().alert('⚠️ La hoja de origen está vacía.');
      return;
    }
    
    const valoresOrigen = hojaOrigen.getRange(1, 1, ultimaFilaOrigen, 3).getValues();
    
    const ultimaFilaDestino = hojaDestino.getLastRow();
    if (ultimaFilaDestino > 0) {
      hojaDestino.getRange(1, 1, ultimaFilaDestino, 3).clearContent();
    }
    
    hojaDestino.getRange(1, 1, valoresOrigen.length, 3).setValues(valoresOrigen);
    
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert('✅ Hoja "CUIT y CLAVES" actualizada con éxito de forma estática.');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error al intentar acceder al archivo origen:\n' + error.toString() + 
    '\n\nVerifica que tengas permisos de edición o lectura en el archivo compartido.');
  }
}

// =========================================================================
// TAREA 1: PROCESAR FACTURAS (EXTRACCIÓN CON IA) - DETECTA RESPUESTAS INVÁLIDAS
// =========================================================================
function procesarFacturas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CAE');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Error: No se encontró la pestaña llamada "CAE". Asegurate de que el nombre sea exacto.');
    return;
  }

  if (!GEMINI_API_KEY) {
    SpreadsheetApp.getUi().alert('❌ Error: No se configuró la GEMINI_API_KEY en las Propiedades del Script.');
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

  // Borra las filas cuyo resultado quedó en "-" (no se pudo generar el CAE),
  // así esos archivos se vuelven a intentar procesar en esta misma corrida.
  limpiarFilasConGuion(sheet);
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange('A1').setValue('Nombre del Archivo');
    sheet.getRange('B1').setValue('URL QR (Clic para CAE)'); 
    sheet.getRange('C1').setValue('Estado / Motivo de Error');
    sheet.getRange('A1:C1').setFontWeight('bold');
  } else {
    sheet.getRange('C1').setValue('Estado / Motivo de Error').setFontWeight('bold');
  }
  
  const ultimaFila = sheet.getLastRow(); 
  let archivosYaRegistrados = [];
  if (ultimaFila > 1) {
    archivosYaRegistrados = sheet.getRange(2, 1, ultimaFila - 1, 1).getValues().map(row => row[0]);
  } 
  
  let contadorExitosos = 0;
  let contadorFallidos = 0;
  
  while (archivos.hasNext()) {
    const archivo = archivos.next(); 
    const nombreArchivo = archivo.getName();
    
    if (archivosYaRegistrados.includes(nombreArchivo)) {
      continue;
    }
    if (nombreArchivo.toLowerCase().includes('credencial') || nombreArchivo.toLowerCase().includes('opcion')) {
      continue;
    }
    
    const proximaFila = sheet.getLastRow() + 1;
    
    try {
      const resultadoGemini = extraerDatosConGemini(archivo);
      
      if (resultadoGemini && resultadoGemini.datos) {
        const d = resultadoGemini.datos;

        // Validación de campos obligatorios antes de generar la URL.
        // Si Gemini extrajo el JSON pero le faltó algún dato clave,
        // esto evita que se genere una URL "OK" pero rota (con NaN adentro).
        const camposObligatorios = ['fecha', 'cuitEmisor', 'ptoVta', 'nroCmp', 'importe', 'cuitReceptor', 'cae'];
        const faltantes = camposObligatorios.filter(c => {
          const v = d[c];
          return v === null || v === undefined || v === '' || (typeof v === 'string' && v.trim() === '');
        });

        if (faltantes.length > 0) {
          sheet.getRange(proximaFila, 1).setValue(nombreArchivo);
          sheet.getRange(proximaFila, 2).setValue('-');
          sheet.getRange(proximaFila, 3).setValue('Faltan campos: ' + faltantes.join(', ') + ' | JSON: ' + JSON.stringify(d).substring(0, 200)).setFontColor('orange');
          contadorFallidos++;
        } else {
          const urlCompletaARCA = generarUrlOficialQR(d);
          sheet.getRange(proximaFila, 1).setValue(nombreArchivo); 
          sheet.getRange(proximaFila, 2).setValue(urlCompletaARCA);
          sheet.getRange(proximaFila, 3).setValue('OK').setFontColor('green');
          contadorExitosos++;
        }
      } else {
        const respuestaTexto = resultadoGemini ? resultadoGemini.respuestaCruda : 'Sin respuesta (revisar cuota/API o bloqueo de contenido)';
        sheet.getRange(proximaFila, 1).setValue(nombreArchivo);
        sheet.getRange(proximaFila, 2).setValue('-');
        sheet.getRange(proximaFila, 3).setValue('Error Estructura. Gemini respondió: ' + respuestaTexto.substring(0, 150)).setFontColor('orange');
        contadorFallidos++;
      }
    } catch (e) {
      sheet.getRange(proximaFila, 1).setValue(nombreArchivo);
      sheet.getRange(proximaFila, 2).setValue('-');
      sheet.getRange(proximaFila, 3).setValue('Error Crítico: ' + e.toString()).setFontColor('red');
      contadorFallidos++;
    }
    
    SpreadsheetApp.flush();
  }
  
  SpreadsheetApp.getUi().alert('Proceso terminado. Exitosos: ' + contadorExitosos + ' | Fallidos: ' + contadorFallidos);
}

// Borra las filas de la hoja "CAE" cuya columna B (URL) quedó marcada con "-",
// es decir, las que fallaron en una corrida anterior. Recorre de abajo hacia
// arriba para que borrar una fila no desplace los índices de las que faltan revisar.
function limpiarFilasConGuion(sheet) {
  const ultimaFila = sheet.getLastRow();
  if (ultimaFila < 2) return;

  const valoresB = sheet.getRange(2, 2, ultimaFila - 1, 1).getValues();
  let filasBorradas = 0;

  for (let i = valoresB.length - 1; i >= 0; i--) {
    const valor = valoresB[i][0];
    if (valor === '-' || valor === '' ) {
      const numeroFila = i + 2; // +2 porque el rango arranca en la fila 2
      sheet.deleteRow(numeroFila);
      filasBorradas++;
    }
  }

  if (filasBorradas > 0) {
    SpreadsheetApp.flush();
    Logger.log('Filas borradas por estar marcadas con "-": ' + filasBorradas);
  }
}

function extraerDatosConGemini(archivo) {
  const blob = archivo.getBlob();
  const base64Pdf = Utilities.base64Encode(blob.getBytes()); 
  
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + GEMINI_API_KEY;
  const prompt = "Analiza el texto de esta factura de AFIP/ARCA de Argentina y extrae los siguientes datos numéricos exactos para armar el JSON del QR oficial. " + 
                  "Si alguno de estos datos no aparece en el documento, devolvé el valor null para ese campo en vez de inventarlo. " +
                  "Devuelve ÚNICAMENTE un objeto JSON válido con estas claves exactas (sin formato markdown, sin texto extra, solo el objeto puro):\n" + 
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
    }],
    "generationConfig": {
      "responseMimeType": "application/json"
    }
  };
  
  const opciones = {
    "method": "post", 
    "contentType": "application/json", 
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  };
  
  const respuesta = UrlFetchApp.fetch(url, opciones);
  const codigoRespuesta = respuesta.getResponseCode();
  const jsonRespuesta = JSON.parse(respuesta.getContentText());

  // Si la API devolvió un error HTTP (cuota, key inválida, etc.), lo dejamos explícito
  if (codigoRespuesta !== 200) {
    const mensajeError = jsonRespuesta.error ? jsonRespuesta.error.message : 'Error HTTP ' + codigoRespuesta;
    return { datos: null, respuestaCruda: 'Error API (' + codigoRespuesta + '): ' + mensajeError };
  }

  // Si Gemini bloqueó la respuesta por safety filters u otro motivo, candidates puede venir vacío
  if (!jsonRespuesta.candidates || jsonRespuesta.candidates.length === 0) {
    const motivoBloqueo = jsonRespuesta.promptFeedback ? JSON.stringify(jsonRespuesta.promptFeedback) : 'Sin candidates en la respuesta';
    return { datos: null, respuestaCruda: motivoBloqueo };
  }
  
  if (jsonRespuesta.candidates[0].content && jsonRespuesta.candidates[0].content.parts[0].text) {
    let textoJson = jsonRespuesta.candidates[0].content.parts[0].text.trim();
    textoJson = textoJson.replace(/```json/g, '').replace(/```/g, '').trim();
    
    try {
      const datosObjeto = JSON.parse(textoJson);
      return { datos: datosObjeto, respuestaCruda: textoJson };
    } catch(e) {
      return { datos: null, respuestaCruda: textoJson };
    }
  }

  // Caso borde: candidate sin texto (por ejemplo finishReason MAX_TOKENS o SAFETY)
  const finishReason = jsonRespuesta.candidates[0].finishReason || 'desconocido';
  return { datos: null, respuestaCruda: 'Respuesta sin texto. finishReason: ' + finishReason };
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
