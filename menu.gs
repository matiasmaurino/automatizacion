function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Estudio CB & MM')
  .addItem('Actualizar CUIT, CLAVES y EMAIL', 'actualizarCuitYClaves')
  .addItem('Copiar WEBAPP → PEDIDOS PARA EXENTOS', 'copiarWebappATareas')
  .addItem('Registrar los CAE ya generados', 'completarNombresArchivos')
    .addSeparator()
    .addItem('📧 Enviar Datos Personales', 'enviarCorreosClientes')
    .addItem('📄 Enviar Archivos en la carpeta Facturas (por CUIT)', 'enviarFacturasPorCuit')
    .addItem('Enviar Deuda CCMA', 'enviarFacturasCCMA')
    
    .addToUi();
}