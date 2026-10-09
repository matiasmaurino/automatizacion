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
}

// =========================================================================
// 2. ENVÍO DE FACTURAS POR CUIT (basado en archivos de Drive + hoja "CUIT y CLAVES")
//
// Reemplaza a las antiguas enviarFacturasFacturar() / enviarFacturasWEBAPP():
// en vez de depender de las columnas EMAIL/ESTADO ENVÍO de las hojas "Facturar"
// o "WEBAPP" (fuente de varios bugs de desalineación de columnas), trabaja
// directo sobre los archivos de la carpeta de Drive y los cruza por CUIT
// contra la hoja "CUIT y CLAVES". Sirve para CUALQUIER archivo que esté en
// FACTURAS, sin importar de qué proceso haya salido.
//
// Lógica:
// 1. Lee "CUIT y CLAVES" y arma un mapa CUIT(col B) -> {nombre(col A), email(col D)}
// 2. Recorre los archivos de la carpeta de Drive (FACTURAS)
// 3. Para cada archivo, busca qué CUIT conocido aparece al inicio del nombre
// 4. Agrupa los archivos por EMAIL (un email puede recibir archivos de varios CUIT)
// 5. Envía un mail por cada email con todos sus archivos adjuntos
// 6. Si el envío fue exitoso, mueve esos archivos a la carpeta "ENVIADO"
//    (subcarpeta de FACTURAS) para no reenviarlos en la próxima corrida
// 7. Los archivos con CUIT no registrado, o con CUIT registrado pero sin
//    email cargado, se listan en un alert al final en vez de saltearse
//    en silencio
// =========================================================================
function enviarFacturasPorCuit() {
  const ID_CARPETA_DRIVE = '1cLnlPOvel1V7q-Syegm0KGWNd0F4a_Ws';
  const ID_CARPETA_ENVIADOS = '1gLvuWNZfeZa-QtrPrRGgONQcbe2kx4Fl'; // "ENVIADO", dentro de FACTURAS
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCuitClaves = ss.getSheetByName('CUIT y CLAVES');

  if (!hojaCuitClaves) {
    SpreadsheetApp.getUi().alert('❌ Error: No se encontró la pestaña "CUIT y CLAVES".');
    return;
  }

  // 1. PASO: Mapa CUIT -> {nombre, email} a partir de las columnas A, B y D.
  //    Guardamos TODOS los CUIT válidos, tengan o no email cargado, para poder
  //    distinguir "CUIT no registrado" de "CUIT registrado pero sin email".
  const datosCuitClaves = hojaCuitClaves.getDataRange().getValues();
  const mapaCuitInfo = {};

  for (let r = 0; r < datosCuitClaves.length; r++) {
    const nombre = String(datosCuitClaves[r][0] || '').trim();              // Columna A
    const cuit   = String(datosCuitClaves[r][1] || '').replace(/\D/g, '');  // Columna B
    const email  = String(datosCuitClaves[r][3] || '').trim();              // Columna D

    if (cuit.length >= 10) {
      mapaCuitInfo[cuit] = {
        nombre: nombre,
        email: (email && email.indexOf('@') !== -1) ? email : ''
      };
    }
  }

  const cuitsConocidos = Object.keys(mapaCuitInfo);
  if (cuitsConocidos.length === 0) {
    SpreadsheetApp.getUi().alert('No se encontraron CUIT cargados en "CUIT y CLAVES".');
    return;
  }

  let carpeta;
  try {
    carpeta = DriveApp.getFolderById(ID_CARPETA_DRIVE);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: No se pudo acceder a la carpeta de Google Drive. Verificá el ID.');
    return;
  }

  let carpetaEnviados;
  try {
    carpetaEnviados = DriveApp.getFolderById(ID_CARPETA_ENVIADOS);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: No se pudo acceder a la carpeta "ENVIADO". Verificá el ID.');
    return;
  }

  // 2. PASO: Recorremos los archivos de la carpeta UNA sola vez y los agrupamos por email.
  //    Los que no se puedan agrupar (CUIT desconocido o CUIT sin email) se listan aparte
  //    para avisar al final, en vez de saltearlos en silencio.
  let gruposPorEmail = {};
  let archivosSinCuitRegistrado = [];
  let archivosSinEmail = []; // { archivo: nombreArchivo, cuit, cliente }

  const archivos = carpeta.getFiles();
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName();

    // Buscamos cuál CUIT conocido aparece al inicio del nombre del archivo
    const cuitEncontrado = cuitsConocidos.find(function (cuit) {
      return nombreArchivo.indexOf(cuit) === 0;
    });

    if (!cuitEncontrado) {
      archivosSinCuitRegistrado.push(nombreArchivo);
      continue;
    }

    const info = mapaCuitInfo[cuitEncontrado];

    if (!info.email) {
      archivosSinEmail.push({
        archivo: nombreArchivo,
        cuit: cuitEncontrado,
        cliente: info.nombre || '(sin nombre)'
      });
      continue;
    }

    const emailKey = info.email.toLowerCase(); // normalizado para no separar por mayúsculas

    if (!gruposPorEmail[emailKey]) {
      gruposPorEmail[emailKey] = {
        clienteNombre: info.nombre || 'Cliente',
        archivosDrive: [],   // objetos File de Drive (para poder moverlos después)
        adjuntos: [],        // Blobs PDF para el mail
        nombresArchivos: []  // nombres para el registro en ENVIO DE EMAIL
      };
    }

    gruposPorEmail[emailKey].archivosDrive.push(archivo);
    gruposPorEmail[emailKey].adjuntos.push(archivo.getAs(MimeType.PDF));
    gruposPorEmail[emailKey].nombresArchivos.push(nombreArchivo);
  }

  const listaEmails = Object.keys(gruposPorEmail);
  if (listaEmails.length === 0 && archivosSinCuitRegistrado.length === 0 && archivosSinEmail.length === 0) {
    SpreadsheetApp.getUi().alert('No se encontraron archivos en la carpeta.');
    return;
  }
  if (listaEmails.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No se pudo enviar ningún archivo.\n\n' +
      (archivosSinCuitRegistrado.length > 0
        ? archivosSinCuitRegistrado.length + ' archivo(s) con CUIT no registrado en "CUIT y CLAVES":\n' + archivosSinCuitRegistrado.join('\n') + '\n\n'
        : '') +
      (archivosSinEmail.length > 0
        ? archivosSinEmail.length + ' archivo(s) con CUIT registrado pero sin email cargado:\n' +
          archivosSinEmail.map(function (a) { return a.archivo + ' (CUIT ' + a.cuit + ' — ' + a.cliente + ')'; }).join('\n')
        : '')
    );
    return;
  }

  // 3. PASO: Enviar un mail por cada email, con todos sus archivos adjuntos
  let correosEnviadosContador = 0;

  for (let email in gruposPorEmail) {
    const info = gruposPorEmail[email];

    const asunto = "Envío FACTURA, CAE y OPCION - Estudio Contable CB & MM";
    const cuerpo = "Estimado/a " + info.clienteNombre + ",\n\n" +
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
        attachments: info.adjuntos
      });

      correosEnviadosContador++;

      // --- REGISTRO EN HOJA "ENVIO DE EMAIL" ---
      registrarEnvioEmail(info.clienteNombre, info.nombresArchivos);

      // --- MOVEMOS LOS ARCHIVOS YA ENVIADOS A "ENVIADO" PARA NO REPETIRLOS LA PRÓXIMA VEZ ---
      info.archivosDrive.forEach(function (archivoDrive) {
        try {
          archivoDrive.moveTo(carpetaEnviados);
        } catch (errorMover) {
          // Fallback por si moveTo no está disponible o no somos dueños del archivo:
          // lo agregamos a ENVIADO y lo sacamos de la carpeta de origen a mano.
          try {
            carpetaEnviados.addFile(archivoDrive);
            carpeta.removeFile(archivoDrive);
          } catch (errorFallback) {
            Logger.log('No se pudo mover el archivo ' + archivoDrive.getName() + ': ' + errorFallback.toString());
          }
        }
      });

      Utilities.sleep(500);

    } catch (error) {
      Logger.log("Error al enviar correo a " + email + ": " + error.toString());
      // Si falló el envío, NO movemos los archivos: quedan en FACTURAS para reintentar en la próxima corrida
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "¡Listo! Se enviaron correos a " + correosEnviadosContador + " clientes según los archivos de Drive.",
    "Envío por CUIT Finalizado",
    5
  );

  // Si quedaron archivos sin poder procesarse, lo mostramos en un alert aparte
  // (el toast desaparece solo y se puede pasar por alto; esto requiere un click).
  if (archivosSinCuitRegistrado.length > 0 || archivosSinEmail.length > 0) {
    let mensajePendientes = 'Se enviaron ' + correosEnviadosContador + ' correos, pero quedaron archivos sin procesar en FACTURAS:\n\n';

    if (archivosSinCuitRegistrado.length > 0) {
      mensajePendientes += archivosSinCuitRegistrado.length + ' archivo(s) con CUIT no registrado en "CUIT y CLAVES":\n' +
        archivosSinCuitRegistrado.join('\n') + '\n\n';
    }

    if (archivosSinEmail.length > 0) {
      mensajePendientes += archivosSinEmail.length + ' archivo(s) con CUIT registrado pero sin email cargado:\n' +
        archivosSinEmail.map(function (a) { return a.archivo + ' (CUIT ' + a.cuit + ' — ' + a.cliente + ')'; }).join('\n');
    }

    SpreadsheetApp.getUi().alert(mensajePendientes);
  }
}

// =========================================================================
// 3. FUNCIONES AUXILIARES Y GUARDADO
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
    throw new Error('Error, hace una captura de pantalla y escribinos por <a href="https://wa.me/542215440900" target="_blank" style="color: #25D366; font-weight: bold; text-decoration: none;">WhatsApp</a> para que podamos darte una respuesta.'); 
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

  let emailCliente = payload.email || '';
  if (!emailCliente && payload.cuit) {
    try {
      const sheetExentos = _getPlanillaExentosExterna();
      const dataExentos = sheetExentos.getDataRange().getValues();
      const cuitLimpioPayload = String(payload.cuit).replace(/\D/g, '');
      
      for (let r = 1; r < dataExentos.length; r++) {
        const cuitExento = String(dataExentos[r][1]).replace(/\D/g, '');
        if (cuitExento === cuitLimpioPayload) {
          emailCliente = String(dataExentos[r][6] || '').trim();
          break;
        }
      }
    } catch (e) {
      Logger.log("No se pudo autocompletar el email: " + e.toString());
    }
  }

  const sheet = _sheetFacturar(); 
  sheet.appendRow([
    payload.cuit,
    payload.claveAfip,
    payload.clienteNombre,
    rango.fechaFactura,
    'Factura C',
    rango.desde,
    rango.hasta,
    null,
    cuitReceptorFinal,
    'Exento',
    descripcion,
    Number(payload.horas),
    '',
    vr.valorHora,
    '',
    '',
    '',
    emailCliente,
    '',
    payload.retroactivo ? 'SI' : '',
    payload.subimoAIoma ? 'SI' : ''
  ]);

  return {
    ok: true, 
    valorHora:  vr.valorHora, 
    valorHoraM: vr.valorHoraM, 
    resolucion: vr.resolucion, 
    descripcion: descripcion 
  };
}