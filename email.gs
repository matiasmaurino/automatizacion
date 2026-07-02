function enviarCorreosClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('DATOS PERSONALES') || ss.getActiveSheet(); 
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) {
    SpreadsheetApp.getUi().alert('La hoja está vacía o solo contiene los encabezados.');
    return;
  }
  
  const rango = hoja.getRange(2, 1, ultimaFila - 1, 6);
  const datos = rango.getValues();
  
  datos.forEach(function(fila, indice) {
    const columnaA = fila[0];
    const columnaB = fila[1];
    const columnaC = fila[2];
    const columnaD = fila[3];
    const columnaE = fila[4];
    const emailDestino = String(fila[5]).trim();
    
    if (emailDestino && emailDestino.indexOf('@') !== -1) {
      const asunto = "Envío de datos - Estudio Contable CBMM";
      const cuerpo = `Estimado/a cliente,
      
Te enviamos los datos registrados en nuestro base de datos:

• CUIT / Dato A: ${columnaA}
• ${columnaB}
• Clave ARCA (ex AFIP) / Dato C: ${columnaC}
• Clave ARBA / Dato D: ${columnaD}
• Vencimiento de tu exencion en ingresos brutos (ALAS) / Dato E: ${columnaE}

Muchas gracias
Saludos

Estudio Contable CB & MM
Contadores Publicos
Celular/Whatsapp 221.544.0900
estudiocontablecbmm@gmail.com
estudiocontable-cb-mm.web.app/
`;

      try {
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

function enviarFacturasAgrupadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Facturar') || ss.getSheetByName('WEBAPP') || ss.getActiveSheet();
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) {
    SpreadsheetApp.getUi().alert('La hoja está vacía o solo contiene los encabezados.');
    return;
  }
  
  // Revisamos las últimas 40 filas para mantener la máxima velocidad
  const CANTIDAD_FILAS_A_REVISAR = 40;
  let filaInicio = ultimaFila - CANTIDAD_FILAS_A_REVISAR + 1;
  if (filaInicio < 2) filaInicio = 2; 
  const filasALeer = ultimaFila - filaInicio + 1;

  const rango = hoja.getRange(filaInicio, 1, filasALeer, 19);
  const datos = rango.getValues();
  
  let gruposPorEmail = {};

  // 1. PASO: Agrupamos las filas por dirección de correo electrónico
  for (let i = datos.length - 1; i >= 0; i--) {
    const fila = datos[i];
    const numeroFilaReal = filaInicio + i; 
    
    const cuitCliente = String(fila[0]).trim();  // Columna A (CUIT)
    const emailDestino = String(fila[17]).trim(); // Columna R (EMAIL)
    const estadoEnvio = String(fila[18]).trim();   // Columna S (EMAIL ENVIADO)
    
    if (emailDestino && emailDestino.indexOf('@') !== -1 && estadoEnvio !== "Email enviado") {
      if (!gruposPorEmail[emailDestino]) {
        gruposPorEmail[emailDestino] = {
          clienteNombre: fila[2] || 'Cliente', // Columna C (Cliente)
          cuit: cuitCliente,
          renglones: []
        };
      }
      
      gruposPorEmail[emailDestino].renglones.push({
        filaHoja: numeroFilaReal,
        facturaTexto: fila[4] 
      });
    }
  }

  const listaEmails = Object.keys(gruposPorEmail);
  if (listaEmails.length === 0) {
    SpreadsheetApp.getUi().alert('No se encontraron facturas pendientes de envío en las últimas filas.');
    return;
  }

  // ID de tu carpeta de Google Drive provista
  const ID_CARPETA_DRIVE = '1cLnlPOvel1V7q-Syegm0KGWNd0F4a_Ws';
  let carpeta;
  try {
    carpeta = DriveApp.getFolderById(ID_CARPETA_DRIVE);
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: No se pudo acceder a la carpeta de Google Drive. Verificá el ID.');
    return;
  }

  let correosEnviadosContador = 0;

  // 2. PASO: Procesar cada cliente, buscar sus archivos en Drive y enviarlos
  for (let email in gruposPorEmail) {
    const infoCliente = gruposPorEmail[email];
    const cuitBuscar = infoCliente.cuit;
    
    if (!cuitBuscar) continue;

    let adjuntos = [];

    // Buscamos TODOS los archivos en tu carpeta de Drive que comiencen con el CUIT del cliente
    const archivos = carpeta.getFiles();
    while (archivos.hasNext()) {
      const archivo = archivos.next();
      const nombreArchivo = archivo.getName();
      
      if (nombreArchivo.indexOf(cuitBuscar) === 0) {
        adjuntos.push(archivo.getAs(MimeType.PDF));
      }
    }

    // Si no hay archivos, registramos la leyenda en la columna S y saltamos al siguiente
    if (adjuntos.length === 0) {
      Logger.log("Fila omitida para " + infoCliente.clienteNombre + ": Sin archivos.");
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 19).setValue("Sin archivos en Facturas"); 
      });
      continue;
    }

    // Armamos el cuerpo del correo de manera prolija usando concatenación simple para evitar fallas
    const asunto = "Envío FACTURA, CAE y OPCION - Estudio Contable CB & MM";
    let cuerpo = "Estimado/a " + infoCliente.clienteNombre + ",\n\n" +
                 "Te enviamos la/s factura/s, Opción Monotributo y/o Credencial de Pago.\n\n" +
                 "Estudio Contable CB & MM\n" +
                 "Contadores Publicos\n" +
                 "Celular/Whatsapp 221.544.0900\n" +
                 "estudiocontablecbmm@gmail.com";

    try {
      // Enviamos el mail incluyendo los archivos adjuntos de Google Drive
      MailApp.sendEmail({
        to: email,
        subject: asunto,
        body: cuerpo,
        attachments: adjuntos
      });
      
      correosEnviadosContador++;
      
      // 3. PASO: Marcar "Email enviado" en la columna S (columna 19)
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 19).setValue("Email enviado"); 
      });
      
      Utilities.sleep(500); 

    } catch (error) {
      Logger.log("Error al enviar correo a " + email + ": " + error.toString());
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "¡Listo! Se procesaron y enviaron correos con sus respectivos PDF adjuntos a " + correosEnviadosContador + " clientes.", 
    "Envío finalizado", 
    5
  );
}