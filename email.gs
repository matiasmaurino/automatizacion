function enviarCorreosClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Cambia 'DATOS PERSONALES' por el nombre exacto de tu pestaña si es otra (ej. 'Hoja 1')
  const hoja = ss.getSheetByName('DATOS PERSONALES') || ss.getActiveSheet(); 
  
  // Obtener todos los datos de la hoja
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) {
    SpreadsheetApp.getUi().alert('La hoja está vacía o solo contiene los encabezados.');
    return;
  }
  
  // Leemos desde la fila 2 (asumiendo que la 1 tiene títulos) hasta la última fila, y las columnas A a la F (6 columnas)
  const rango = hoja.getRange(2, 1, ultimaFila - 1, 6);
  const datos = rango.getValues();
  
  // Recorrer cada fila de datos
  datos.forEach(function(fila, indice) {
    // Mapeo de columnas (A=0, B=1, C=2, D=3, E=4, F=5)
    const columnaA = fila[0];
    const columnaB = fila[1];
    const columnaC = fila[2];
    const columnaD = fila[3];
    const columnaE = fila[4];
    const emailDestino = String(fila[5]).trim();
    
    // Validar que la columna F contenga un correo electrónico válido
    if (emailDestino && emailDestino.indexOf('@') !== -1) {
      
      // Asunto del correo
      const asunto = "Envío de datos - Estudio Contable CBMM";
      
      // Cuerpo del correo (puedes personalizar este texto a tu gusto)
      const cuerpo = `Estimado/a cliente,
      
Le enviamos los datos registrados en nuestro sistema:

• CUIT / Dato A: ${columnaA}
• Cliente / Dato B: ${columnaB}
• Clave ARCA / Dato C: ${columnaC}
• Clave ARBA / Dato D: ${columnaD}
• Vencimiento / Dato E: ${columnaE}

Atentamente,
Estudio Contable CB & MM`;

      try {
        // Enviar el correo electrónico
        MailApp.sendEmail(emailDestino, asunto, cuerpo);
        Logger.log(`Correo enviado correctamente a: ${emailDestino}`);
      } catch (error) {
        Logger.log(`Error al enviar correo a ${emailDestino}: ${error.toString()}`);
      }
      
    } else {
      Logger.log(`Fila ${indice + 2}: No se envió correo porque la columna F está vacía o no es válida.`);
    }
  });
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Proceso de envío de correos finalizado.', 'Éxito', 5);
}