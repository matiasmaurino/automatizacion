const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const CARPETA_ORIGEN_ID = '1cLnlPOvel1V7q-Syegm0KGWNd0F4a_Ws';
const SPREADSHEET_ORIGEN_ID = '1JjdVfUdiIhSMO1McU4FK2o0U2k-Qyd7clPDfQ0fzKjU';

// =========================================================================
// TAREA 2: ACTUALIZAR CUIT Y CLAVES
// =========================================================================
function actualizarCuitYClaves() {
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDestino = ssDestino.getSheetByName('CUIT y CLAVES');
  
  if (!hojaDestino) {
    ssDestino.toast('❌ Error: No se encontró la pestaña llamada "CUIT y CLAVES".', 'Actualización', 5);
    Logger.log('Error: No se encontró la pestaña "CUIT y CLAVES".');
    return;
  }
  
  ssDestino.toast('Conectando con la base de datos origen...', 'Actualización', 3);
  
  try {
    const ssOrigen = SpreadsheetApp.openById(SPREADSHEET_ORIGEN_ID);
    const hojaOrigen = ssOrigen.getSheetByName('EXENTOS');
    
    if (!hojaOrigen) {
      ssDestino.toast('❌ Error: No se encontró la pestaña "EXENTOS" en el archivo origen.', 'Actualización', 5);
      Logger.log('Error: No se encontró la pestaña "EXENTOS" en el archivo origen.');
      return;
    }
    
    const ultimaFilaOrigen = hojaOrigen.getLastRow();
    if (ultimaFilaOrigen === 0) {
      ssDestino.toast('⚠️ La hoja de origen está vacía.', 'Actualización', 5);
      Logger.log('Advertencia: la hoja de origen está vacía.');
      return;
    }
    
    const valoresOrigen = hojaOrigen.getRange(1, 1, ultimaFilaOrigen, 7).getValues();
    
    const datosProcesados = valoresOrigen.map(fila => [
      fila[0], // Columna A
      fila[1], // Columna B
      fila[2], // Columna C
      fila[6]  // Columna G (Email)
    ]);
    
    const ultimaFilaDestino = hojaDestino.getLastRow();
    if (ultimaFilaDestino > 0) {
      hojaDestino.getRange(1, 1, ultimaFilaDestino, 4).clearContent();
    }
    
    hojaDestino.getRange(1, 1, datosProcesados.length, 4).setValues(datosProcesados);
    SpreadsheetApp.flush();
    
    ssDestino.toast('✅ Hoja "CUIT, CLAVES y Email" actualizada correctamente.', 'Actualización', 5);
    Logger.log('Actualización completada correctamente.');
    
  } catch (error) {
    ssDestino.toast('❌ Error: ' + error.toString(), 'Actualización', 8);
    Logger.log('Error: ' + error.toString());
  }
}