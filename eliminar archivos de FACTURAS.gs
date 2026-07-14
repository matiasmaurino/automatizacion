function eliminarFacturasEnviadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Facturar') || ss.getSheetByName('WEBAPP') || ss.getActiveSheet();
  
  Logger.log(`--- INICIANDO BÚSQUEDA POR CUIT ---`);
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) return;
  
  // Leemos hasta la columna S (19)
  const datos = hoja.getRange(2, 1, ultimaFila - 1, 19).getValues();
  
  const ID_CARPETA_DRIVE = '1cLnlPOvel1V7q-Syegm0KGWNd0F4a_Ws';
  let carpeta;
  try {
    carpeta = DriveApp.getFolderById(ID_CARPETA_DRIVE);
  } catch(e) {
    Logger.log(`ERROR CRÍTICO: No se pudo acceder a la carpeta. Detalle: ${e.toString()}`);
    return; 
  }

  let borrados = 0;

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];
    const numeroFilaExcel = i + 2;
    
    const cuit = String(fila[0]).trim();     // Columna A (Índice 0)
    const estado = String(fila[18]).trim();  // Columna S (Índice 18)

    // Condición: Estado debe ser "Email enviado" y el CUIT no debe estar vacío
    if (estado === "Email enviado" && cuit !== "") {
      Logger.log(`[Fila ${numeroFilaExcel}] Procesando CUIT: "${cuit}"...`);
      
      // Traemos TODOS los archivos de la carpeta para evaluar sus nombres
      const archivos = carpeta.getFiles();
      let archivosEncontradosParaCuit = 0;
      
      while (archivos.hasNext()) {
        const archivo = archivos.next();
        const nombreArchivo = archivo.getName();
        
        // Verificamos si el nombre del archivo empieza con el CUIT del renglón
        if (nombreArchivo.startsWith(cuit)) {
          archivosEncontradosParaCuit++;
          Logger.log(`   🔍 Encontrado archivo que coincide: "${nombreArchivo}"`);
          
          try {
            // Intento 1: Mandar a la papelera (si eres el dueño)
            archivo.setTrashed(true);
            borrados++;
            Logger.log(`   ✅ Enviado a la papelera.`);
          } catch (error) {
            // Intento 2: Remover de la carpeta (si no eres el dueño)
            try {
              carpeta.removeFile(archivo);
              borrados++;
              Logger.log(`   ⚠️ Removido de la carpeta (No eras el propietario).`);
            } catch (errorRemover) {
              Logger.log(`   ❌ No se pudo procesar. Error: ${errorRemover.toString()}`);
            }
          }
        }
      }
      
      if (archivosEncontradosParaCuit === 0) {
        Logger.log(`   ❌ No se encontró ningún archivo que empiece con el CUIT ${cuit} en la carpeta.`);
      }
    }
  }

  Logger.log(`--- PROCESO FINALIZADO ---`);
  Logger.log(`Total de archivos procesados/quitados: ${borrados}`);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Limpieza finalizada. Se procesaron ${borrados} archivos vinculados a los CUITs enviados.`, 
    "Limpieza de Drive", 
    5
  );
}