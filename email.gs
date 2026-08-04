// =========================================================================
// FUNCIÓN AUXILIAR: Registrar envíos en la hoja "ENVIO DE EMAIL"
// =========================================================================
function registrarEnvioEmail(nombreCliente, nombresArchivos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaLog = ss.getSheetByName('ENVIO DE EMAIL');

  // Si la pestaña no existe, se crea automáticamente con sus encabezados
  if (!hojaLog) {
    hojaLog = ss.insertSheet('ENVIO DE EMAIL');
    hojaLog.appendRow(['Cliente', 'Fecha de Envío', 'Adjunto 1', 'Adjunto 2', 'Adjunto 3']);
  }

  const fechaEnvio = new Date(); // Fecha y hora actual
  let filaLog = [nombreCliente, fechaEnvio];

  // Si hay archivos adjuntos, se agregan a partir de la columna C en adelante
  if (Array.isArray(nombresArchivos) && nombresArchivos.length > 0) {
    filaLog = filaLog.concat(nombresArchivos);
  }

  hojaLog.appendRow(filaLog);
}

// =========================================================================
// 1. ENVÍO DE CORREOS CLIENTES (DATOS PERSONALES)
// =========================================================================
function enviarCorreosClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('DATOS PERSONALES') || ss.getActiveSheet(); 
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) {
    SpreadsheetApp.getUi().alert('La hoja está vacía o solo contiene los encabezados.');
    return;
  }
  
  const rango = hoja.getRange(2, 1, ultimaFila - 1, 8);
  const datos = rango.getValues();
  const hoy = new Date(); // Fecha actual para comparar
  
  datos.forEach(function(fila, indice) {
    const columnaA = fila[0];
    const columnaB = fila[1];
    const columnaC = fila[2];
    const columnaD = fila[3];
    const columnaE = fila[4]; // Fecha de vencimiento ALAS
    const emailDestino = String(fila[5]).trim();
    const estadoEnvio = String(fila[7]).trim(); 
    
    if (emailDestino && emailDestino.indexOf('@') !== -1 && estadoEnvio !== "Enviado") {
      const asunto = "Envío de claves personales y vencimiento de ALAS - Estudio Contable CB&MM";
      
      // 1. Formatear la fecha a DD/MM/AAAA de forma limpia
      let fechaFormateada = "";
      let advertenciaVencido = "";
      
      if (columnaE instanceof Date) {
        fechaFormateada = Utilities.formatDate(columnaE, Session.getScriptTimeZone(), "d/M/yyyy");
        
        // 2. Validar si la fecha está vencida (menor a hoy)
        if (columnaE < hoy) {
          advertenciaVencido = ' <br><span style="color: red; font-weight: bold;">⚠️ Tu exención en ingresos brutos está vencida</span>';
        }
      } else if (columnaE) {
        // Por si acaso viene como texto y no como objeto Date
        fechaFormateada = columnaE;
      }

      // 3. Armamos el cuerpo en HTML reemplazando los saltos de línea por <br>
      const cuerpoHtml = `
        <p>Estimado/a cliente,</p>
        <p>Te enviamos los datos registrados en nuestra base de datos:</p>
        <ul>
          <li><strong>CUIT:</strong> ${columnaA}</li>
          <li><strong>${columnaB}</strong></li>
          <li><strong>Clave ARCA (ex AFIP):</strong> ${columnaC}</li>
          <li><strong>Clave ARBA:</strong> ${columnaD}</li>
          <li><strong>Vencimiento de tu exención en ingresos brutos (ALAS):</strong> ${fechaFormateada}${advertenciaVencido}</li>
        </ul>
        <p>Muchas gracias<br>Saludos</p>
        <p><strong>Estudio Contable CB & MM</strong><br>
        Contadores Públicos<br>
        Celular/Whatsapp 221.544.0900<br>
        <a href="mailto:estudiocontablecbmm@gmail.com">estudiocontablecbmm@gmail.com</a><br>
        <a href="https://estudiocontable-cb-mm.web.app/">estudiocontable-cb-mm.web.app/</a></p>
      `;

      try {
        // Usamos htmlBody para que reconozca los estilos y el color rojo
        MailApp.sendEmail({
          to: emailDestino,
          subject: asunto,
          htmlBody: cuerpoHtml
        });
        
        Logger.log(`Correo enviado correctamente a: ${emailDestino}`);
        
        const filaReal = indice + 2; 
        hoja.getRange(filaReal, 8).setValue('Enviado');
        
        // --- REGISTRO EN HOJA "ENVIO DE EMAIL" ---
        const clienteNombre = columnaB || columnaA;
        registrarEnvioEmail(clienteNombre, []);
        
      } catch (error) {
        Logger.log(`Error al enviar correo a ${emailDestino}: ${error.toString()}`);
      }
    } else if (estadoEnvio === "Enviado") {
      Logger.log(`Fila ${indice + 2}: Ya fue enviado anteriormente.`);
    } else {
      Logger.log(`Fila ${indice + 2}: No se envió correo porque la columna F está vacía o no es válida.`);
    }
  });
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Proceso de envío de correos finalizado.', 'Éxito', 5);

  // --- EJECUCIÓN DE ELIMINACIÓN DE ARCHIVOS ---
  try {
    eliminarFacturasEnviadas();
  } catch (error) {
    Logger.log("Error al intentar eliminar archivos: " + error.toString());
  }
}

// =========================================================================
// 2. ENVÍO DE FACTURAS (PESTAÑA FACTURAR)
// =========================================================================
function enviarFacturasFacturar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Facturar');
  
  if (!hoja) {
    SpreadsheetApp.getUi().alert('❌ Error: No se encontró la pestaña llamada "Facturar".');
    return;
  }
  
  const ultimaFila = obtenerUltimaFilaReal(hoja);
  if (ultimaFila <= 1) {
    SpreadsheetApp.getUi().alert('La hoja "Facturar" está vacía o solo contiene los encabezados.');
    return;
  }
  
  // Revisamos TODAS las filas con datos (desde la fila 2 hasta la última),
  // para no dejar afuera pendientes más viejos que las últimas N filas.
  const filaInicio = 2;
  const filasALeer = ultimaFila - filaInicio + 1;

  const rango = hoja.getRange(filaInicio, 1, filasALeer, 21);
  const datos = rango.getValues();
  
  let gruposPorEmail = {};

  // 1. PASO: Agrupamos las filas por dirección de correo electrónico
  for (let i = datos.length - 1; i >= 0; i--) {
    const fila = datos[i];
    const numeroFilaReal = filaInicio + i; 
    
    const cuitCliente = String(fila[0]).trim();    // Columna A (CUIT)
    const emailDestino = String(fila[17]).trim();  // Columna R (EMAIL) [Índice 17]
    const estadoEnvio = String(fila[18]).trim();    // Columna S (ESTADO ENVÍO) [Índice 18]
    
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
    SpreadsheetApp.getUi().alert('No se encontraron facturas pendientes de envío en la pestaña "Facturar".');
    return;
  }

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
    let nombresArchivos = []; // Guardará los nombres de los adjuntos para el registro

    const archivos = carpeta.getFiles();
    while (archivos.hasNext()) {
      const archivo = archivos.next();
      const nombreArchivo = archivo.getName();
      
      if (nombreArchivo.indexOf(cuitBuscar) === 0) {
        adjuntos.push(archivo.getAs(MimeType.PDF));
        nombresArchivos.push(nombreArchivo);
      }
    }

    if (adjuntos.length === 0) {
      Logger.log("Fila omitida para " + infoCliente.clienteNombre + " en Facturar: Sin archivos.");
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 19).setValue("Sin archivos en Facturas"); 
      });
      continue;
    }

    const asunto = "Envío FACTURA - Estudio Contable CB & MM";
    let cuerpo = "Estimado/a " + infoCliente.clienteNombre + ",\n\n" +
                 "Te enviamos la/s factura/s, Opción Monotributo y/o Credencial de Pago.\n\n" +
                 "Estudio Contable CB & MM\n" +
                 "Contadores Publicos\n" +
                 "Celular/Whatsapp 221.544.0900\n" +
                 "estudiocontablecbmm@gmail.com";

    try {
      MailApp.sendEmail({
        to: email,
        subject: asunto,
        body: cuerpo,
        attachments: adjuntos
      });
      
      correosEnviadosContador++;
      
      // Marcar "Email enviado" en la columna S (columna 19)
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 19).setValue("Email enviado"); 
      });

      // --- REGISTRO EN HOJA "ENVIO DE EMAIL" ---
      registrarEnvioEmail(infoCliente.clienteNombre, nombresArchivos);
      
      Utilities.sleep(500); 

    } catch (error) {
      Logger.log("Error al enviar correo a " + email + ": " + error.toString());
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "¡Listo! Se procesaron y enviaron correos con sus respectivos PDF adjuntos a " + correosEnviadosContador + " clientes de Facturar.", 
    "Envío Facturar Finalizado", 
    5
  );
}

// =========================================================================
// 3. ENVÍO DE FACTURAS (PESTAÑA WEBAPP)
// =========================================================================
function enviarFacturasWEBAPP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('WEBAPP');
  
  if (!hoja) {
    SpreadsheetApp.getUi().alert('❌ Error: No se encontró la pestaña llamada "WEBAPP".');
    return;
  }
  
  const ultimaFila = obtenerUltimaFilaReal(hoja);
  if (ultimaFila <= 1) {
    SpreadsheetApp.getUi().alert('La hoja "WEBAPP" está vacía o solo contiene los encabezados.');
    return;
  }
  
  // Revisamos TODAS las filas con datos (desde la fila 2 hasta la última),
  // para no dejar afuera pendientes más viejos que las últimas N filas.
  const filaInicio = 2;
  const filasALeer = ultimaFila - filaInicio + 1;

  const rango = hoja.getRange(filaInicio, 1, filasALeer, 21);
  const datos = rango.getValues();
  
  let gruposPorEmail = {};

  // 1. PASO: Agrupamos las filas por dirección de correo electrónico
  for (let i = datos.length - 1; i >= 0; i--) {
    const fila = datos[i];
    const numeroFilaReal = filaInicio + i; 
    
    const cuitCliente = String(fila[0]).trim();    // Columna A (CUIT)
    const emailDestino = String(fila[17]).trim();  // Columna R (EMAIL) [Índice 17]
    const estadoEnvio = String(fila[18]).trim();    // Columna S (ESTADO ENVÍO) [Índice 18]
    
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
    SpreadsheetApp.getUi().alert('No se encontraron facturas pendientes de envío en la pestaña "WEBAPP".');
    return;
  }

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
    let nombresArchivos = []; // Guardará los nombres de los adjuntos para el registro

    const archivos = carpeta.getFiles();
    while (archivos.hasNext()) {
      const archivo = archivos.next();
      const nombreArchivo = archivo.getName();
      
      if (nombreArchivo.indexOf(cuitBuscar) === 0) {
        adjuntos.push(archivo.getAs(MimeType.PDF));
        nombresArchivos.push(nombreArchivo);
      }
    }

    if (adjuntos.length === 0) {
      Logger.log("Fila omitida para " + infoCliente.clienteNombre + " en WEBAPP: Sin archivos.");
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 19).setValue("Sin archivos en Facturas"); 
      });
      continue;
    }

    const asunto = "Envío FACTURA, CAE y OPCION - Estudio Contable CB & MM";
    let cuerpo = "Estimado/a " + infoCliente.clienteNombre + ",\n\n" +
                 "Te enviamos la/s factura/s, Opción Monotributo y/o Credencial de Pago.\n\n" +
                 "Estudio Contable CB & MM\n" +
                 "Contadores Publicos\n" +
                 "Celular/Whatsapp 221.544.0900\n" +
                 "estudiocontablecbmm@gmail.com";

    try {
      MailApp.sendEmail({
        to: email,
        subject: asunto,
        body: cuerpo,
        attachments: adjuntos
      });
      
      correosEnviadosContador++;
      
      // Marcar "Email enviado" en la columna S (columna 19)
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 19).setValue("Email enviado"); 
      });

      // --- REGISTRO EN HOJA "ENVIO DE EMAIL" ---
      registrarEnvioEmail(infoCliente.clienteNombre, nombresArchivos);
      
      Utilities.sleep(500); 

    } catch (error) {
      Logger.log("Error al enviar correo a " + email + ": " + error.toString());
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "¡Listo! Se procesaron y enviaron correos con sus respectivos PDF adjuntos a " + correosEnviadosContador + " clientes de WEBAPP.", 
    "Envío WEBAPP Finalizado", 
    5
  );
}

// =========================================================================
// 4. FUNCIONES AUXILIARES Y GUARDADO
// =========================================================================
function obtenerUltimaFilaReal(hoja) {
  const valores = hoja.getRange("A:A").getValues();
  for (let i = valores.length - 1; i >= 0; i--) {
    if (String(valores[i][0]).trim() !== "") {
      return i + 1;
    }
  }
  return 1;
}