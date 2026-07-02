const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const CARPETA_ORIGEN_ID = '1cLnlPOvel1V7q-Syegm0KGWNd0F4a_Ws';
const SPREADSHEET_ORIGEN_ID = '1JjdVfUdiIhSMO1McU4FK2o0U2k-Qyd7clPDfQ0fzKjU';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Actualizacion')
    .addItem('Actualizar CUIT y CLAVES', 'actualizarCuitYClaves')
    .addToUi();
}

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
    
    const valoresOrigen = hojaOrigen.getRange(1, 1, ultimaFilaOrigen, 3).getValues();
    const ultimaFilaDestino = hojaDestino.getLastRow();
    if (ultimaFilaDestino > 0) {
      hojaDestino.getRange(1, 1, ultimaFilaDestino, 3).clearContent();
    }
    
    hojaDestino.getRange(1, 1, valoresOrigen.length, 3).setValues(valoresOrigen);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert('✅ Hoja "CUIT y CLAVES" actualizada con éxito.');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.toString());
  }
}
