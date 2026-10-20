function STOP() {
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
  var seleccion = hojaDestino.getRange("AP3").getValue().trim();
  var rangoACopiar = obtenerRangoPorSeleccion(seleccion, hojaOrigen);
  if (!rangoACopiar) {
    SpreadsheetApp.getUi().alert("No se encontró la opción seleccionada en los mapeos.");
    return;
  }
  rangoACopiar.copyTo(hojaDestino.getRange("AQ6"), {contentsOnly: true});
}
