function generateSequence() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONTROL DE DATOS");
  var totalFilas = 2000;
  var valores = [];
  for (var i = 0; i < totalFilas; i++) {
    var filaInicio = 3 + 11 * Math.floor(i / 7);
    var columnaInicio = 2 + 11 * (i % 7);
    valores.push([letraColumna(columnaInicio) + filaInicio + ':' + letraColumna(columnaInicio + 9) + (filaInicio + 8)]);
  }
  hoja.getRange(1, 2, totalFilas, 1).setValues(valores);
}

function letraColumna(numero) {
  var letra = '';
  while (numero > 0) {
    var resto = (numero - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    numero = Math.floor((numero - 1) / 26);
  }
  return letra;
}
