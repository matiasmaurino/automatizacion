function completarNombresArchivos() {
  var folderId = "1cLnlPOvel1V7q-Syegm0KGWNd0F4a_Ws";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Facturar");
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  var fileMap = {};
  var cuitRegex = /\b\d{2}[-.\s]?\d{8}[-.\s]?\d\b|\d{11}/;
  
  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    
    // Evalúa PRIMERO si el archivo contiene "constataci"
    if (name.toLowerCase().includes("constataci")) {
      var match = name.match(cuitRegex);
      if (match) {
        var cuitLimpio = match[0].replace(/\D/g, '');
        fileMap[cuitLimpio] = name;
      }
    }
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  var cuits = sheet.getRange("A2:A" + lastRow).getDisplayValues();
  var output = [];
  
  for (var i = 0; i < cuits.length; i++) {
    var cuitSheet = cuits[i][0].toString().replace(/\D/g, '');
    // Si no encuentra un archivo de constatación, guarda un texto vacío ""
    var nombreArchivo = fileMap[cuitSheet] || "";
    output.push([nombreArchivo]);
  }
  
  // Escribe los resultados en la columna S (columna 19) a partir de la fila 2
  sheet.getRange(2, 19, output.length, 1).setValues(output);
}