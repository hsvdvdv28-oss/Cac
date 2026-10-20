function copiarBasadoEnSeleccion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = ss.getSheetByName("PORTES");
  if (!hojaOrigen) {
    SpreadsheetApp.getUi().alert("La hoja 'PORTES' no existe.");
    return;
  }
  var hojaDestino = ss.getSheetByName("PROCESO DE INFORMACION");
  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert("La hoja 'PROCESO DE INFORMACION' no existe.");
    return;
  }
  var seleccion = hojaDestino.getRange("F3").getValue().trim();
  var rangoACopiar = obtenerRangoPorSeleccion(seleccion, hojaOrigen);
  if (!rangoACopiar) {
    SpreadsheetApp.getUi().alert("No se encontró la opción seleccionada en los mapeos.");
    return;
  }
  rangoACopiar.copyTo(hojaDestino.getRange("G6"), {contentsOnly: true});
}

function obtenerRangoPorSeleccion(seleccion, hojaOrigen) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mappingsSheet = ss.getSheetByName("CONTROL DE DATOS");
  if (!mappingsSheet) {
    SpreadsheetApp.getUi().alert("La hoja 'CONTROL DE DATOS' no existe.");
    return null;
  }
  var data = mappingsSheet.getDataRange().getValues();
  var rangoMap = {};
  for (var i = 0; i < data.length; i++) {
    var key = data[i][0].toString().trim();
    var value = data[i][1].toString().trim();
    rangoMap[key] = value;
  }
  var rangoACopiarStr = rangoMap[seleccion];
  if (rangoACopiarStr) {
    return hojaOrigen.getRange(rangoACopiarStr);
  } else {
    return null;
  }
}
