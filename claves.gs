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
    
    // 1. Modificado: Traemos hasta la columna 7 (Columna G)
    const valoresOrigen = hojaOrigen.getRange(1, 1, ultimaFilaOrigen, 7).getValues();
    
    // 2. Procesamos los datos para quedarnos solo con A, B, C y G (columna indexada 0, 1, 2 y 6)
    const datosProcesados = valoresOrigen.map(fila => [
      fila[0], // Columna A
      fila[1], // Columna B
      fila[2], // Columna C
      fila[6]  // Columna G (Email)
    ]);
    
    const ultimaFilaDestino = hojaDestino.getLastRow();
    // 3. Modificado: Limpiamos 4 columnas en lugar de 3 en el destino (A, B, C y D)
    if (ultimaFilaDestino > 0) {
      hojaDestino.getRange(1, 1, ultimaFilaDestino, 4).clearContent();
    }
    
    // 4. Modificado: Pegamos la nueva matriz que incluye el email en la columna D del destino
    hojaDestino.getRange(1, 1, datosProcesados.length, 4).setValues(datosProcesados);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert('✅ Hoja "CUIT, CLAVES y Email".');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.toString());
  }
}