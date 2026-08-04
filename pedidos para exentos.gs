function copiarWebappATareas() {
  const ss = SpreadsheetApp.openById('17xJc3GF9M3XkYJMmJn-LIzJ42ui1PyAmRwyU3BsvEHs');
  const sheetWebapp = ss.getSheetByName('WEBAPP');
  const sheetTareas = ss.getSheetByName('TAREAS');

  if (!sheetWebapp || !sheetTareas) {
    SpreadsheetApp.getUi().alert('No se encontró WEBAPP o TAREAS.');
    return;
  }

  // ── Limpiar TAREAS antes de empezar (se deja la fila 1 de encabezados) ──
  const ultimaFilaTareas = sheetTareas.getLastRow();
  if (ultimaFilaTareas > 1) {
    sheetTareas.getRange(2, 1, ultimaFilaTareas - 1, sheetTareas.getLastColumn())
      .clearContent();
  }

  const dataWebapp = sheetWebapp.getDataRange().getValues();
  let filasNuevas = 0;

  for (let r = 1; r < dataWebapp.length; r++) {
    const fila = dataWebapp[r];

    const cliente    = fila[2];   // C = CLIENTE
    const desde      = fila[5];   // F = DESDE
    const cuitRecep  = String(fila[8]  || '').trim(); // I = CUIT RECEPTOR
    const desc       = fila[10];  // K = DESCRIPCIÓN
    const cant       = fila[11];  // L = CANTIDAD
    const precio     = fila[13];  // N = PRECIO UNITARIO
    const retroact   = String(fila[19] || '').trim().toUpperCase(); // T
    const subimoIoma = String(fila[20] || '').trim().toUpperCase(); // U
    const exportado  = String(fila[21] || '').trim(); // V = marca de procesado

    if (!cliente || exportado === 'EXPORTADO') continue;

    // ── Fila 1: AFIP FC ... ───────────────────────────────────────
    let tipoTarea;
    if (cuitRecep === '30628249527') {
      tipoTarea = retroact === 'SI' ? 'AFIP FC IOMA RETROACTIVO' : 'AFIP FC IOMA';
    } else {
      tipoTarea = 'AFIP FC';
    }

    const sig1 = sheetTareas.getLastRow() + 1;
    sheetTareas.getRange(sig1, 6).setValue(cliente);   // F
    sheetTareas.getRange(sig1, 10).setValue(tipoTarea); // J
    sheetTareas.getRange(sig1, 12).setValue(desde);    // L
    sheetTareas.getRange(sig1, 14).setValue(cant);     // N
    sheetTareas.getRange(sig1, 15).setValue(precio);   // O
    sheetTareas.getRange(sig1, 16).setValue(desc);     // P
    filasNuevas++;

    // ── Fila 2 (solo si col U = SI): IOMA FC DIGITAL PRESENTAR ───
    if (subimoIoma === 'SI') {
      const sig2 = sheetTareas.getLastRow() + 1;
      sheetTareas.getRange(sig2, 6).setValue(cliente);
      sheetTareas.getRange(sig2, 10).setValue('IOMA FC DIGITAL PRESENTAR');
      sheetTareas.getRange(sig2, 12).setValue(desde);
      sheetTareas.getRange(sig2, 14).setValue(cant);
      sheetTareas.getRange(sig2, 15).setValue(precio);
      sheetTareas.getRange(sig2, 16).setValue(desc);
      filasNuevas++;
    }

    // ── Marcar fila como exportada en col V de WEBAPP ─────────────
    sheetWebapp.getRange(r + 1, 22).setValue('EXPORTADO');
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    filasNuevas + ' filas agregadas en TAREAS.',
    'Listo', 4
  );
}