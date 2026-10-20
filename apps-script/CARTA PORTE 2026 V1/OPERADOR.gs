var OPERADOR = {
  HOJA: 'PROCESO DE INFORMACION',
  HOJA_DATOS: 'OP EXTERNOS',
  FILA_OPERADOR: 37,
  COLUMNA_OPERADOR: 13,
  FILA_UNIDAD: 38,
  COLUMNA_UNIDAD: 14
};

function onEdit(e) {
  if (!e || !e.range) return;
  var hoja = e.range.getSheet();
  if (hoja.getName() !== OPERADOR.HOJA) return;
  var fila = e.range.getRow(), columna = e.range.getColumn();
  if (fila === OPERADOR.FILA_OPERADOR && columna === OPERADOR.COLUMNA_OPERADOR) llenarOperador_(hoja, e.range.getValue());
  if (fila === OPERADOR.FILA_UNIDAD && columna === OPERADOR.COLUMNA_UNIDAD) llenarUnidad_(hoja, e.range.getValue());
}

function llenarOperador_(hoja, seleccion) {
  var clave = textoOperador_(seleccion);
  var nombre = '', licencia = '';
  if (clave) {
    var encontrada = filasOperadores_(hoja, 6, 3).filter(function (f) { return textoOperador_(f[2]) === clave; })[0];
    if (encontrada) {
      nombre = textoOperador_(encontrada[0]);
      licencia = textoOperador_(encontrada[1]);
    } else {
      var partes = clave.split(' I ');
      nombre = partes.shift();
      licencia = partes.join(' I ');
    }
  }
  hoja.getRange('P31').setValue(nombre);
  hoja.getRange('P32').setValue(licencia);
}

function llenarUnidad_(hoja, seleccion) {
  var clave = textoOperador_(seleccion);
  if (!clave) {
    hoja.getRangeList(['P33', 'T33', 'T34']).clearContent();
    return;
  }
  var encontrada = filasOperadores_(hoja, 1, 3).filter(function (f) { return textoOperador_(f[2]) === clave; })[0];
  var placas = separarPlacas_(encontrada ? encontrada[2] : clave);
  var unidad = '';
  if (encontrada) unidad = typeof encontrada[1] === 'number' ? encontrada[1] : textoOperador_(encontrada[1]);
  hoja.getRange('P33').setValue(unidad);
  hoja.getRange('T33').setValue(placas.tx);
  hoja.getRange('T34').setValue(placas.mx);
}

function filasOperadores_(hoja, columna, ancho) {
  var datos = hoja.getParent().getSheetByName(OPERADOR.HOJA_DATOS);
  if (!datos || datos.getLastRow() < 2) return [];
  return datos.getRange(2, columna, datos.getLastRow() - 1, ancho).getValues();
}

function separarPlacas_(texto) {
  var tx = [], mx = [];
  textoOperador_(texto).split('/').forEach(function (parte) {
    parte = parte.trim();
    if (!parte) return;
    if (/^R\d+$/.test(parte)) tx.push(parte); else mx.push(parte);
  });
  return { tx: tx.join(' / '), mx: mx.join(' / ') };
}

function textoOperador_(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[ㅤ​-‍⁠﻿­]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
}
