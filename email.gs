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
}

function enviarFacturasFacturar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Facturar');
  
  if (!hoja) {
    SpreadsheetApp.getUi().alert('❌ Error: No se encontró la pestaña llamada "Facturar".');
    return;
  }
  
  // Obtenemos la última fila real con datos en la columna A (CUIT) para evitar fallos por filas vacías
  const ultimaFila = obtenerUltimaFilaReal(hoja);
  if (ultimaFila <= 1) {
    SpreadsheetApp.getUi().alert('La hoja "Facturar" está vacía o solo contiene los encabezados.');
    return;
  }
  
  // Revisamos las últimas 40 filas con datos reales para mantener la velocidad
  const CANTIDAD_FILAS_A_REVISAR = 40;
  let filaInicio = ultimaFila - CANTIDAD_FILAS_A_REVISAR + 1;
  if (filaInicio < 2) filaInicio = 2; 
  const filasALeer = ultimaFila - filaInicio + 1;

  // Leemos hasta la columna 21 (Columna U) para incluir el Email y el Estado de envío
  const rango = hoja.getRange(filaInicio, 1, filasALeer, 21);
  const datos = rango.getValues();
  
  let gruposPorEmail = {};

  // 1. PASO: Agrupamos las filas por dirección de correo electrónico
  for (let i = datos.length - 1; i >= 0; i--) {
    const fila = datos[i];
    const numeroFilaReal = filaInicio + i; 
    
    const cuitCliente = String(fila[0]).trim();    // Columna A (CUIT)
    const emailDestino = String(fila[19]).trim();  // Columna T (EMAIL) [Índice 19]
    const estadoEnvio = String(fila[20]).trim();    // Columna U (ESTADO ENVÍO) [Índice 20]
    
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
    SpreadsheetApp.getUi().alert('No se encontraron facturas pendientes de envío en las últimas filas de la pestaña "Facturar".');
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

    // Si no hay archivos, registramos la leyenda en la columna U y saltamos al siguiente
    if (adjuntos.length === 0) {
      Logger.log("Fila omitida para " + infoCliente.clienteNombre + " en Facturar: Sin archivos.");
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 21).setValue("Sin archivos en Facturas"); 
      });
      continue;
    }

    // Armamos el cuerpo del correo
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
      
      // 3. PASO: Marcar "Email enviado" en la columna U (columna 21)
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 21).setValue("Email enviado"); 
      });
      
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

function enviarFacturasWEBAPP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('WEBAPP');
  
  if (!hoja) {
    SpreadsheetApp.getUi().alert('❌ Error: No se encontró la pestaña llamada "WEBAPP".');
    return;
  }
  
  // Obtenemos la última fila real con datos en la columna A (CUIT) para evitar fallos por filas vacías
  const ultimaFila = obtenerUltimaFilaReal(hoja);
  if (ultimaFila <= 1) {
    SpreadsheetApp.getUi().alert('La hoja "WEBAPP" está vacía o solo contiene los encabezados.');
    return;
  }
  
  // Revisamos las últimas 40 filas con datos reales para mantener la velocidad
  const CANTIDAD_FILAS_A_REVISAR = 40;
  let filaInicio = ultimaFila - CANTIDAD_FILAS_A_REVISAR + 1;
  if (filaInicio < 2) filaInicio = 2; 
  const filasALeer = ultimaFila - filaInicio + 1;

  // Leemos hasta la columna 21 (Columna U) para incluir el Email y el Estado de envío
  const rango = hoja.getRange(filaInicio, 1, filasALeer, 21);
  const datos = rango.getValues();
  
  let gruposPorEmail = {};

  // 1. PASO: Agrupamos las filas por dirección de correo electrónico
  for (let i = datos.length - 1; i >= 0; i--) {
    const fila = datos[i];
    const numeroFilaReal = filaInicio + i; 
    
    const cuitCliente = String(fila[0]).trim();    // Columna A (CUIT)
    const emailDestino = String(fila[19]).trim();  // Columna T (EMAIL) [Índice 19]
    const estadoEnvio = String(fila[20]).trim();    // Columna U (ESTADO ENVÍO) [Índice 20]
    
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
    SpreadsheetApp.getUi().alert('No se encontraron facturas pendientes de envío en las últimas filas de la pestaña "WEBAPP".');
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

    // Si no hay archivos, registramos la leyenda en la columna U y saltamos al siguiente
    if (adjuntos.length === 0) {
      Logger.log("Fila omitida para " + infoCliente.clienteNombre + " en WEBAPP: Sin archivos.");
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 21).setValue("Sin archivos en Facturas"); 
      });
      continue;
    }

    // Armamos el cuerpo del correo
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
      
      // 3. PASO: Marcar "Email enviado" en la columna U (columna 21)
      infoCliente.renglones.forEach(function(renglon) {
        hoja.getRange(renglon.filaHoja, 21).setValue("Email enviado"); 
      });
      
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
// 2. FUNCIÓN AUXILIAR: Detectar la última fila REAL ignorando celdas vacías
// =========================================================================
function obtenerUltimaFilaReal(hoja) {
  const valores = hoja.getRange("A:A").getValues(); // Revisamos la columna CUIT
  for (let i = valores.length - 1; i >= 0; i--) {
    if (String(valores[i][0]).trim() !== "") {
      return i + 1;
    }
  }
  return 1;
}

// =========================================================================
// 3. GUARDAR FACTURA CORREGIDA (Auto-completa el Email en la Columna T)
// =========================================================================
function guardarFactura(payload) {
  const rango = _mesARango(Number(payload.mesNumero), Number(payload.anio)); 
  
  let mesNombre = '';
  if (!isNaN(payload.mesNumero) && Number(payload.mesNumero) >= 1 && Number(payload.mesNumero) <= 12) {
    mesNombre = MESES[Number(payload.mesNumero) - 1];
  } else {
    mesNombre = String(payload.mesNumero).toUpperCase().trim();
  }

  const vr = _getValorYResolucion(payload.servicio, mesNombre, payload.anio); 
  if (!vr) {
    throw new Error('No se encontró el valor de hora / resolución para ese servicio y período en TABLAS AUX. Revisá que el período exista en esa hoja.'); 
  }

  const cuitReceptorFinal = payload.retroactivo ? CUIT_IOMA : payload.cuitReceptor; 

  let descripcion =
    payload.servicio + ' ' +
    payload.pacienteNombre + ' ' +
    payload.numeroAfiliado + '/00 ' +
    payload.estado + ' ' +
    'DNI ' + payload.dniPaciente + ' ' +
    'tramite ' + payload.numeroTramite + ' ' +
    'segun resolucion ' + vr.resolucion + ' ' +
    'del mes de ' + mesNombre + ' ' + payload.anio + ' ' +
    'por ' + payload.horas + ' horas a un valor de $' + vr.valorHora; 

  if (payload.retroactivo) {
    descripcion = 'RETROACTIVO de Factura Pto.Vta ' + payload.puntoVenta +
      ' Nro ' + payload.nroComprobante + ' — ' + descripcion; 
  }

  // --- NUEVA LÓGICA DE AUTOCOMPLETADO DE EMAIL ---
  let emailCliente = payload.email || '';
  if (!emailCliente && payload.cuit) {
    try {
      const sheetExentos = _getPlanillaExentosExterna();
      const dataExentos = sheetExentos.getDataRange().getValues();
      const cuitLimpioPayload = String(payload.cuit).replace(/\D/g, '');
      
      for (let r = 1; r < dataExentos.length; r++) {
        const cuitExento = String(dataExentos[r][1]).replace(/\D/g, '');
        if (cuitExento === cuitLimpioPayload) {
          emailCliente = String(dataExentos[r][6] || '').trim(); // Columna G (EMAIL) en Exentos
          break;
        }
      }
    } catch (e) {
      Logger.log("No se pudo autocompletar el email: " + e.toString());
    }
  }
  // ----------------------------------------------

  const sheet = _sheetFacturar(); 
  sheet.appendRow([
    payload.cuit,            // A (1)
    payload.claveAfip,       // B (2)
    payload.clienteNombre,   // C (3)
    rango.fechaFactura,      // D (4) FECHA
    'Factura C',             // E (5)
    rango.desde,             // F (6) DESDE
    rango.hasta,             // G (7) HASTA
    rango.fechaEmision,      // H (8) VENCIMIENTO
    cuitReceptorFinal,       // I (9)
    'Exento',                // J (10)
    descripcion,             // K (11)
    Number(payload.horas),   // L (12) Cant
    'otras unidades',        // M (13)
    vr.valorHora,            // N (14) Prec
    '',                      // O (15) FACTURA (vacío, lo completa el script)
    '',                      // P (16) OPCION Y CREDENCIAL (vacío)
    '',                      // Q (17) CAE (vacío, lo completa el script)
    payload.retroactivo ? 'SI' : '',   // R (18) RETROACTIVO
    payload.subimoAIoma ? 'SI' : '',   // S (19) SUBIMOS A IOMA
    emailCliente,            // T (20) EMAIL (Nuevo campo guardado automáticamente)
    ''                       // U (21) ESTADO ENVÍO EMAIL (Nuevo campo de control)
  ]);

  return {
    ok: true, 
    valorHora:  vr.valorHora, 
    valorHoraM: vr.valorHoraM, 
    resolucion: vr.resolucion, 
    descripcion: descripcion 
  };
}