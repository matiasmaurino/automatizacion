function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Estudio CB & MM')
  .addItem('Actualizar CUIT, CLAVES y EMAIL', 'actualizarCuitYClaves')
  .addItem('Copiar WEBAPP → PEDIDOS PARA EXENTOS', 'copiarWebappATareas')
  .addItem('Registrar los CAE ya generados', 'completarNombresArchivos')
    .addSeparator()
    .addItem('📧 Enviar Datos Personales', 'enviarCorreosClientes')
    .addItem('📄 Enviar Facturas (Webapp)', 'enviarFacturasWEBAPP')
    .addItem('Eliminar archivos enviados de Facturas', 'eliminarFacturasEnviadas')
    .addItem('Enviar Deuda CCMA', 'enviarFacturasCCMA')
    
    .addToUi();
}