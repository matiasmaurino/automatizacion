function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Estudio CB & MM')
  .addItem('Actualizar CUIT, CLAVES y EMAIL', 'actualizarCuitYClaves')
    .addItem('🚀 ENVIAR TODO (1-Clic)', 'enviarTodo')
    .addSeparator()
    .addItem('📧 Enviar Clientes (Datos Personales)', 'enviarCorreosClientes')
    .addItem('📄 Enviar Facturas (Pestaña Facturar)', 'enviarFacturasFacturar')
    .addItem('🌐 Enviar Facturas (Pestaña WEBAPP)', 'enviarFacturasWEBAPP')
    .addItem('Enviar Deuda CCMA', 'enviarFacturasCCMA')
    .addToUi();
}