var CPV = {
  HOJA_PLAN: 'PLAN MEXICO',
  HOJA_REGISTRO: 'CARTAS PORTE VINCULADAS',
  HOJA_COPIA: 'MASTER SAFETY',
  COLUMNAS: 9,
  TITULOS: 1,
  MAX_ACTIVAS: 10,
  REGISTRO: ['ID', 'NOMBRE', 'ACTIVA', 'AGREGADA', 'ÚLTIMA COPIA', 'ÚLTIMO AVISO'],
  CLAVES: { 0: 'FECHA', 1: 'NO. OP', 2: 'NOMBRE', 5: 'FOLIO' },
  DISPARADOR: 'cpv_alEditar',
  SIMPLES: ['onOpen', 'onEdit']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('CARTAS PORTE VINCULADAS')
    .addItem('Agregar copia', 'cpv_agregar')
    .addItem('Desactivar copia', 'cpv_desactivar')
    .addItem('Revisar copias', 'cpv_revisar')
    .addItem('Realinear copia con PLAN MEXICO', 'cpv_realinear')
    .addSeparator()
    .addItem('Instalar disparador', 'cpv_instalar')
    .addToUi();
}

function cpv_instalar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  cpvHojaRegistro_(ss);
  const r = cpvSincronizarDisparadores_(ss);
  cpvAvisar_('CARTAS PORTE VINCULADAS', 'DISPARADOR LISTO: CADA EDICIÓN EN ' + CPV.HOJA_PLAN + ' (A:I) SE COPIA A LAS COPIAS ACTIVAS, Y CADA EDICIÓN EN ' +
    CPV.HOJA_COPIA + ' DE UNA COPIA ACTIVA SE COPIA A ' + CPV.HOJA_PLAN + ' Y A LAS DEMÁS COPIAS ACTIVAS.' +
    (r.viejos ? '\nSE BORRARON ' + r.viejos + ' DISPARADORES VIEJOS O REPETIDOS.' : '') + '\n' + cpvTextoDisparadores_(r) +
    '\nCOPIAS EN EL REGISTRO: ' + cpvLeerRegistro_(ss).length + '.');
}

function cpv_alEditar(e) {
  if (!e || !e.range) return;
  const hoja = e.range.getSheet();
  const libro = hoja.getParent();
  const idPlan = PropertiesService.getScriptProperties().getProperty('ID_PLAN');
  if (!idPlan || libro.getId() === idPlan) {
    if (hoja.getName() === CPV.HOJA_PLAN) cpvDesdePlan_(e, hoja);
    return;
  }
  if (hoja.getName() === CPV.HOJA_COPIA) cpvDesdeCopia_(e, hoja, SpreadsheetApp.openById(idPlan));
}

function cpvDesdePlan_(e, hoja) {
  const c1 = e.range.getColumn(), c2 = Math.min(e.range.getLastColumn(), CPV.COLUMNAS);
  if (c1 > c2) return;
  const ss = hoja.getParent();
  const f1 = e.range.getRow();
  const fuente = { valores: hoja.getRange(f1, 1, e.range.getNumRows(), CPV.COLUMNAS).getValues(), tz: ss.getSpreadsheetTimeZone() };
  const reg = ss.getSheetByName(CPV.HOJA_REGISTRO);
  const copiadas = [];
  cpvActivas_(ss).forEach(function (x) {
    let aviso = '';
    try {
      const destino = SpreadsheetApp.openById(x.id).getSheetByName(CPV.HOJA_COPIA);
      if (!destino) throw new Error('NO TIENE LA HOJA ' + CPV.HOJA_COPIA);
      const omitidas = cpvEspejarFilas_(fuente, destino, f1, c1, c2);
      copiadas.push(x.fila);
      if (omitidas.length) aviso = cpvTextoOmitidas_(omitidas);
    } catch (err) {
      aviso = 'ERROR: ' + cpvMensaje_(err);
    }
    if (aviso) cpvAnotar_(ss, reg, x.fila, 'DESDE ' + CPV.HOJA_PLAN + ': ' + aviso);
  });
  if (reg && copiadas.length) reg.getRangeList(copiadas.map(function (f) { return 'E' + f; })).setValue(new Date());
}

function cpvDesdeCopia_(e, hoja, plan) {
  const c1 = e.range.getColumn(), c2 = Math.min(e.range.getLastColumn(), CPV.COLUMNAS);
  if (c1 > c2) return;
  const libro = hoja.getParent();
  const activas = cpvActivas_(plan);
  const propia = activas.filter(function (x) { return x.id === libro.getId(); })[0];
  if (!propia) return;
  const reg = plan.getSheetByName(CPV.HOJA_REGISTRO);
  const f1 = e.range.getRow();
  const fuente = { valores: hoja.getRange(f1, 1, e.range.getNumRows(), CPV.COLUMNAS).getValues(), tz: libro.getSpreadsheetTimeZone() };
  let aviso = '';
  try {
    const destino = plan.getSheetByName(CPV.HOJA_PLAN);
    if (!destino) throw new Error('NO EXISTE LA HOJA ' + CPV.HOJA_PLAN);
    const omitidas = cpvEspejarFilas_(fuente, destino, f1, c1, c2);
    if (omitidas.length) aviso = cpvTextoOmitidas_(omitidas);
  } catch (err) {
    aviso = 'ERROR: ' + cpvMensaje_(err);
  }
  if (aviso) {
    cpvAnotar_(plan, reg, propia.fila, 'HACIA ' + CPV.HOJA_PLAN + ': ' + aviso + ' NO SE ENVIÓ A LAS DEMÁS COPIAS.');
    return;
  }
  const copiadas = [propia.fila];
  activas.forEach(function (x) {
    if (x.id === propia.id) return;
    let nota = '';
    try {
      const otra = SpreadsheetApp.openById(x.id).getSheetByName(CPV.HOJA_COPIA);
      if (!otra) throw new Error('NO TIENE LA HOJA ' + CPV.HOJA_COPIA);
      const omitidas = cpvEspejarFilas_(fuente, otra, f1, c1, c2);
      copiadas.push(x.fila);
      if (omitidas.length) nota = cpvTextoOmitidas_(omitidas);
    } catch (err) {
      nota = 'ERROR: ' + cpvMensaje_(err);
    }
    if (nota) cpvAnotar_(plan, reg, x.fila, 'DESDE ' + (propia.nombre || propia.id) + ': ' + nota);
  });
  if (reg) reg.getRangeList(copiadas.map(function (f) { return 'E' + f; })).setValue(new Date());
}

function cpv_agregar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = cpvPedirId_('AGREGAR COPIA', 'PEGA EL ENLACE O EL ID DE LA HOJA DE CÁLCULO (CARTA PORTE 2027, UNA COPIA O LA 2026):');
  if (!id) return;
  if (id === ss.getId()) {
    cpvAvisar_('AGREGAR COPIA', 'ESE ES EL MISMO ARCHIVO DE ' + CPV.HOJA_PLAN + '.');
    return;
  }
  const activas = cpvActivas_(ss).filter(function (x) { return x.id !== id; });
  if (activas.length >= CPV.MAX_ACTIVAS) {
    cpvAvisar_('AGREGAR COPIA', 'YA HAY ' + activas.length + ' COPIAS ACTIVAS (MÁXIMO ' + CPV.MAX_ACTIVAS + '). DESACTIVA UNA ANTES DE AGREGAR OTRA:\n\n' +
      activas.map(function (x) { return x.fila + ' · ' + (x.nombre || x.id); }).join('\n'));
    return;
  }
  let libro;
  try {
    libro = SpreadsheetApp.openById(id);
  } catch (err) {
    cpvAvisar_('AGREGAR COPIA', 'NO SE PUDO ABRIR ESE ARCHIVO CON ESTA CUENTA: ' + cpvMensaje_(err));
    return;
  }
  const copia = libro.getSheetByName(CPV.HOJA_COPIA);
  if (!copia) {
    cpvAvisar_('AGREGAR COPIA', '"' + libro.getName() + '" NO TIENE LA HOJA ' + CPV.HOJA_COPIA + '.');
    return;
  }
  const hoja = cpvHojaRegistro_(ss);
  const actual = cpvLeerRegistro_(ss).filter(function (x) { return x.id === id; })[0];
  let fila;
  if (actual) {
    fila = actual.fila;
    hoja.getRange(fila, 2, 1, 2).setValues([[libro.getName(), 'SI']]);
  } else {
    fila = hoja.getLastRow() + 1;
    hoja.getRange(fila, 1, 1, 3).setNumberFormat('@').setValues([[id, libro.getName(), 'SI']]);
    hoja.getRange(fila, 4).setValue(new Date());
  }
  cpvAvisar_('AGREGAR COPIA', '"' + libro.getName() + '" QUEDÓ ACTIVA EN LA FILA ' + fila + ' DEL REGISTRO (' + (activas.length + 1) + ' DE ' + CPV.MAX_ACTIVAS + ').\n' +
    cpvComparar_(ss, copia) + '\n' + cpvActualizarSiInstalador_(ss));
}

function cpv_desactivar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lista = cpvLeerRegistro_(ss);
  if (!lista.length) {
    cpvAvisar_('DESACTIVAR COPIA', 'EL REGISTRO ESTÁ VACÍO.');
    return;
  }
  const texto = lista.map(function (x) { return x.fila + ' · ' + (x.nombre || x.id) + ' · ' + (x.activa ? 'ACTIVA' : 'INACTIVA'); }).join('\n');
  const respuesta = cpvPreguntar_('DESACTIVAR COPIA', 'ESCRIBE EL NÚMERO DE FILA, EL ENLACE O EL ID:\n\n' + texto);
  if (respuesta === null) return;
  const id = cpvIdDe_(respuesta);
  const elegida = lista.filter(function (x) { return String(x.fila) === respuesta.trim() || x.id === id; })[0];
  if (!elegida) {
    cpvAvisar_('DESACTIVAR COPIA', 'NO ESTÁ EN EL REGISTRO.');
    return;
  }
  ss.getSheetByName(CPV.HOJA_REGISTRO).getRange(elegida.fila, 3).setValue('NO');
  cpvAvisar_('DESACTIVAR COPIA', '"' + (elegida.nombre || elegida.id) + '" YA NO RECIBE NI ENVÍA COPIAS DE ' + CPV.HOJA_COPIA + '. PARA VOLVER A ACTIVARLA USA "AGREGAR COPIA".\n' +
    cpvActualizarSiInstalador_(ss));
}

function cpv_revisar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CPV.HOJA_REGISTRO);
  const lista = cpvLeerRegistro_(ss);
  const lineas = lista.map(function (x) {
    let texto;
    try {
      const libro = SpreadsheetApp.openById(x.id);
      hoja.getRange(x.fila, 2).setValue(libro.getName());
      const copia = libro.getSheetByName(CPV.HOJA_COPIA);
      texto = copia ? cpvComparar_(ss, copia) : 'NO TIENE LA HOJA ' + CPV.HOJA_COPIA + '.';
    } catch (err) {
      texto = 'NO SE PUDO ABRIR: ' + cpvMensaje_(err);
    }
    return 'FILA ' + x.fila + ' · ' + (x.nombre || x.id) + (x.activa ? '' : ' (INACTIVA)') + '\n   ' + texto;
  });
  const activas = lista.filter(function (x) { return x.activa; }).length;
  cpvAvisar_('REVISAR COPIAS', lineas.length ? 'ACTIVAS: ' + activas + ' DE ' + CPV.MAX_ACTIVAS + '.\n' + lineas.join('\n') : 'EL REGISTRO ESTÁ VACÍO.');
}

function cpv_realinear() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = cpvPedirId_('REALINEAR COPIA', 'PEGA EL ENLACE O EL ID DE LA COPIA. SU HOJA ' + CPV.HOJA_COPIA + ' (COLUMNAS A:I) QUEDARÁ IGUAL A ' + CPV.HOJA_PLAN + ':');
  if (!id) return;
  let copia;
  try {
    copia = SpreadsheetApp.openById(id).getSheetByName(CPV.HOJA_COPIA);
  } catch (err) {
    cpvAvisar_('REALINEAR COPIA', 'NO SE PUDO ABRIR ESE ARCHIVO: ' + cpvMensaje_(err));
    return;
  }
  if (!copia) {
    cpvAvisar_('REALINEAR COPIA', 'ESE ARCHIVO NO TIENE LA HOJA ' + CPV.HOJA_COPIA + '.');
    return;
  }
  const plan = ss.getSheetByName(CPV.HOJA_PLAN);
  const n = cpvUltimaConDato_(plan);
  const ui = SpreadsheetApp.getUi();
  const si = ui.alert('REALINEAR COPIA', 'SE REESCRIBEN LAS COLUMNAS A:I DE ' + CPV.HOJA_COPIA + ' EN "' + copia.getParent().getName() + '" CON LAS ' + n +
    ' FILAS DE ' + CPV.HOJA_PLAN + ' (SIN LA FILA DE TÍTULOS) Y SE BORRA LO QUE ESTÉ MÁS ABAJO EN ESAS COLUMNAS. ¿CONTINUAR?', ui.ButtonSet.YES_NO);
  if (si !== ui.Button.YES) return;
  const tzO = ss.getSpreadsheetTimeZone(), tzD = copia.getParent().getSpreadsheetTimeZone();
  const desde = CPV.TITULOS + 1, filas = n - CPV.TITULOS;
  cpvAsegurarTamano_(copia, n, CPV.COLUMNAS);
  if (filas > 0) {
    const formatos = copia.getRange(desde, 1, filas, CPV.COLUMNAS).getNumberFormats();
    const valores = plan.getRange(desde, 1, filas, CPV.COLUMNAS).getValues().map(function (fila, i) {
      return fila.map(function (v, j) { return cpvValorParaCopiar_(v, formatos[i][j], tzO, tzD); });
    });
    copia.getRange(desde, 1, filas, CPV.COLUMNAS).setValues(valores);
  }
  if (copia.getMaxRows() > n) copia.getRange(n + 1, 1, copia.getMaxRows() - n, CPV.COLUMNAS).clearContent();
  const x = cpvLeerRegistro_(ss).filter(function (r) { return r.id === id; })[0];
  if (x) {
    const reg = ss.getSheetByName(CPV.HOJA_REGISTRO);
    reg.getRange(x.fila, 5).setValue(new Date());
    cpvAnotar_(ss, reg, x.fila, 'REALINEADA CON ' + CPV.HOJA_PLAN + ' (' + n + ' FILAS)');
  }
  cpvAvisar_('REALINEAR COPIA', 'LISTO: ' + n + ' FILAS COPIADAS.\n' + cpvComparar_(ss, copia));
}

function cpvSincronizarDisparadores_(ss) {
  const global = Function('return this')();
  let viejos = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const h = t.getHandlerFunction();
    if (h === CPV.DISPARADOR || CPV.SIMPLES.indexOf(h) >= 0 || typeof global[h] !== 'function') {
      if (h !== CPV.DISPARADOR) viejos++;
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger(CPV.DISPARADOR).forSpreadsheet(ss).onEdit().create();
  PropertiesService.getScriptProperties().setProperties({ ID_PLAN: ss.getId(), INSTALADOR: Session.getEffectiveUser().getEmail() });
  const todas = cpvLeerRegistro_(ss).filter(function (x) { return x.activa; });
  const conectadas = [], fallidas = [];
  todas.slice(0, CPV.MAX_ACTIVAS).forEach(function (x) {
    try {
      ScriptApp.newTrigger(CPV.DISPARADOR).forSpreadsheet(x.id).onEdit().create();
      conectadas.push(x.nombre || x.id);
    } catch (err) {
      fallidas.push((x.nombre || x.id) + ': ' + cpvMensaje_(err));
    }
  });
  return { viejos: viejos, conectadas: conectadas, fallidas: fallidas, sobrantes: Math.max(0, todas.length - CPV.MAX_ACTIVAS) };
}

function cpvTextoDisparadores_(r) {
  return 'COPIAS QUE ENVÍAN A ' + CPV.HOJA_PLAN + ': ' + r.conectadas.length + ' DE ' + CPV.MAX_ACTIVAS + (r.conectadas.length ? ' (' + r.conectadas.join(', ') + ')' : '') + '.' +
    (r.fallidas.length ? '\nNO SE PUDIERON CONECTAR: ' + r.fallidas.join('; ') + '.' : '') +
    (r.sobrantes ? '\nHAY ' + r.sobrantes + ' COPIAS ACTIVAS DE MÁS: SOLO SE CONECTARON LAS PRIMERAS ' + CPV.MAX_ACTIVAS + '. DESACTIVA LAS QUE NO USES.' : '');
}

function cpvActualizarSiInstalador_(ss) {
  const instalador = PropertiesService.getScriptProperties().getProperty('INSTALADOR');
  if (instalador && instalador === Session.getEffectiveUser().getEmail()) return cpvTextoDisparadores_(cpvSincronizarDisparadores_(ss));
  return 'PARA QUE EL CAMBIO APLIQUE AL ENVÍO DESDE LAS COPIAS, LA CUENTA QUE INSTALÓ EL DISPARADOR' + (instalador ? ' (' + instalador + ')' : '') + ' DEBE USAR "INSTALAR DISPARADOR".';
}

function cpvActivas_(ss) {
  return cpvLeerRegistro_(ss).filter(function (x) { return x.activa; }).slice(0, CPV.MAX_ACTIVAS);
}

function cpvComparar_(ss, copia) {
  const plan = ss.getSheetByName(CPV.HOJA_PLAN);
  const tzA = copia.getParent().getSpreadsheetTimeZone(), tzB = ss.getSpreadsheetTimeZone();
  const a = cpvUltimaConDato_(copia), b = cpvUltimaConDato_(plan);
  const fa = copia.getRange(a, 6).getValue(), fb = plan.getRange(b, 6).getValue();
  if (a === b && cpvMismoDato_(fa, tzA, fb, tzB)) return 'COINCIDE CON ' + CPV.HOJA_PLAN + ': TERMINA EN LA FILA ' + a + ' CON EL FOLIO ' + (cpvTextoSync_(fa, tzA) || '—') + '.';
  return 'NO COINCIDE: TERMINA EN LA FILA ' + a + ' (FOLIO ' + (cpvTextoSync_(fa, tzA) || '—') + ') Y ' + CPV.HOJA_PLAN + ' EN LA FILA ' + b +
    ' (FOLIO ' + (cpvTextoSync_(fb, tzB) || '—') + '). USA "REALINEAR COPIA CON PLAN MEXICO".';
}

function cpvAnotar_(ss, reg, fila, texto) {
  console.error('CARTAS PORTE VINCULADAS, FILA ' + fila + ': ' + texto);
  if (!reg) return;
  const linea = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm') + ' · ' + texto;
  reg.getRange(fila, CPV.REGISTRO.length).setValue("'" + linea.slice(0, 2000));
}

function cpvHojaRegistro_(ss) {
  const hoja = ss.getSheetByName(CPV.HOJA_REGISTRO) || ss.insertSheet(CPV.HOJA_REGISTRO);
  const n = CPV.REGISTRO.length;
  if (cpvTextoLimpio_(hoja.getRange(1, 1).getValue()) !== CPV.REGISTRO[0]) {
    cpvAsegurarTamano_(hoja, 2, n);
    hoja.getRange(1, 1, 1, n).setValues([CPV.REGISTRO]).setFontWeight('bold').setBackground('#263238').setFontColor('#FFFFFF');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function cpvLeerRegistro_(ss) {
  const hoja = ss.getSheetByName(CPV.HOJA_REGISTRO);
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getDisplayValues().map(function (f, i) {
    return { id: cpvTextoLimpio_(f[0]), nombre: cpvTextoLimpio_(f[1]), activa: cpvTextoLimpio_(f[2]).toUpperCase() === 'SI', fila: i + 2 };
  }).filter(function (x) { return x.id; });
}

function cpvPedirId_(titulo, texto) {
  const r = cpvPreguntar_(titulo, texto);
  if (r === null) return '';
  const id = cpvIdDe_(r);
  if (!id) cpvAvisar_(titulo, 'NO SE RECONOCIÓ UN ENLACE O ID DE HOJA DE CÁLCULO.');
  return id;
}

function cpvPreguntar_(titulo, texto) {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(titulo, texto, ui.ButtonSet.OK_CANCEL);
  return r.getSelectedButton() === ui.Button.OK ? r.getResponseText() : null;
}

function cpvIdDe_(texto) {
  const s = String(texto || '').trim();
  const m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(s);
  if (m) return m[1];
  return /^[A-Za-z0-9_-]{20,}$/.test(s) ? s : '';
}

function cpvEspejarFilas_(fuente, destino, f1, c1, c2) {
  const N = CPV.COLUMNAS, w = c2 - c1 + 1, n = fuente.valores.length;
  const tzO = fuente.tz, tzD = destino.getParent().getSpreadsheetTimeZone();
  cpvAsegurarTamano_(destino, f1 + n - 1, N);
  const actual = destino.getRange(f1, 1, n, N).getValues();
  const formatos = destino.getRange(f1, c1, n, w).getNumberFormats();
  const omitidas = [];
  const ok = fuente.valores.map(function (fila, i) {
    const m = cpvFilaDistinta_(fila, tzO, actual[i], tzD, c1, c2);
    if (m) omitidas.push({ fila: f1 + i, motivo: m });
    return !m;
  });
  let i = 0;
  while (i < n) {
    if (!ok[i]) { i++; continue; }
    let k = i;
    while (k < n && ok[k]) k++;
    const vals = [];
    for (let r = i; r < k; r++) {
      vals.push(fuente.valores[r].slice(c1 - 1, c2).map(function (v, j) { return cpvValorParaCopiar_(v, formatos[r][j], tzO, tzD); }));
    }
    destino.getRange(f1 + i, c1, k - i, w).setValues(vals);
    i = k;
  }
  return omitidas;
}

function cpvFilaDistinta_(o, tzO, d, tzD, c1, c2) {
  if (d.every(function (v) { return v === '' || v === null; })) return '';
  const claves = CPV.CLAVES;
  const cols = Object.keys(claves).map(Number);
  const distinto = function (j) {
    return claves[j] + ' ' + (cpvTextoSync_(o[j], tzO) || '(VACÍO)') + ' ≠ ' + (cpvTextoSync_(d[j], tzD) || '(VACÍO)');
  };
  const fijas = cols.filter(function (j) { return j + 1 < c1 || j + 1 > c2; });
  for (let x = 0; x < fijas.length; x++) {
    if (!cpvMismoDato_(o[fijas[x]], tzO, d[fijas[x]], tzD)) return distinto(fijas[x]);
  }
  if (fijas.length) return '';
  const hay = function (v, tz) { return cpvTextoSync_(v, tz) !== ''; };
  if (!cols.some(function (j) { return hay(o[j], tzO); })) return 'LA FILA SE VACIÓ; SI ES EL MISMO REGISTRO, BÓRRALA A MANO EN EL OTRO ARCHIVO';
  if (hay(o[5], tzO) && hay(d[5], tzD)) return cpvMismoDato_(o[5], tzO, d[5], tzD) ? '' : 'EL FOLIO NO COINCIDE: ' + distinto(5);
  const mismos = [0, 2].every(function (j) { return hay(o[j], tzO) && cpvMismoDato_(o[j], tzO, d[j], tzD); });
  return mismos ? '' : 'NINGÚN DATO CLAVE COINCIDE: ' + cols.map(distinto).join(', ');
}

function cpvMismoDato_(a, tzA, b, tzB) {
  const fa = cpvFechaSync_(a, tzA), fb = cpvFechaSync_(b, tzB);
  if (fa && fb) return fa.dia === fb.dia || (fa.t !== null && fb.t !== null && Math.abs(fa.t - fb.t) < 12 * 3600000);
  return cpvClaveSync_(a, tzA) === cpvClaveSync_(b, tzB);
}

function cpvFechaSync_(v, tz) {
  if (v && typeof v === 'object' && typeof v.getTime === 'function') return { dia: Utilities.formatDate(v, tz, 'yyyy-MM-dd'), t: v.getTime() };
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(cpvTextoLimpio_(v));
  if (!m) return null;
  const dos = function (s) { return (s.length < 2 ? '0' : '') + s; };
  return { dia: m[3] + '-' + dos(m[2]) + '-' + dos(m[1]), t: null };
}

function cpvClaveSync_(v, tz) {
  const s = cpvTextoSync_(v, tz).toUpperCase();
  return /^-?\d+(?:\.\d+)?$/.test(s) ? String(Number(s)) : s;
}

function cpvTextoSync_(v, tz) {
  if (v && typeof v === 'object' && typeof v.getTime === 'function') return Utilities.formatDate(v, tz, 'dd/MM/yyyy');
  return cpvTextoLimpio_(v);
}

function cpvTextoOmitidas_(omitidas) {
  const lista = omitidas.slice(0, 5).map(function (o) { return o.fila + ' (' + o.motivo + ')'; }).join('; ');
  return 'NO SE COPIÓ ' + (omitidas.length === 1 ? 'LA FILA ' : omitidas.length + ' FILAS: ') + lista + (omitidas.length > 5 ? '…' : '') +
    '. ESA FILA ES OTRO REGISTRO EN EL DESTINO: REVISA QUE LAS FILAS DE LOS ARCHIVOS COINCIDAN (INSERTAR, BORRAR U ORDENAR FILAS NO SE COPIA).';
}

function cpvValorParaCopiar_(v, formato, tzO, tzD) {
  if (v && typeof v === 'object' && typeof v.getTime === 'function') return cpvMismaHora_(v, tzO, tzD);
  if (typeof v === 'string' && v !== '' && !/^@/.test(String(formato || '').trim()) &&
    (/^\s*[=+\-@'.,$%(\d]/.test(v) || /^\s*(?:TRUE|FALSE|VERDADERO|FALSO)\s*$/i.test(v))) return "'" + v;
  return v;
}

function cpvMismaHora_(fecha, tzO, tzD) {
  if (!tzO || !tzD || tzO === tzD) return fecha;
  const patron = 'yyyy-MM-dd HH:mm:ss';
  const reloj = cpvMsReloj_(Utilities.formatDate(fecha, tzO, patron));
  let t = fecha.getTime() + reloj - cpvMsReloj_(Utilities.formatDate(fecha, tzD, patron));
  t += reloj - cpvMsReloj_(Utilities.formatDate(new Date(t), tzD, patron));
  return new Date(t);
}

function cpvMsReloj_(texto) {
  const m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/.exec(texto);
  return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
}

function cpvUltimaConDato_(hoja) {
  const max = hoja.getMaxRows();
  let ultima = 1;
  [1, 6].forEach(function (col) {
    if (col > hoja.getMaxColumns()) return;
    const celda = hoja.getRange(max, col);
    const f = celda.getValue() !== '' ? max : celda.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    if (f > ultima) ultima = f;
  });
  return ultima;
}

function cpvTextoLimpio_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return typeof v.getTime === 'function' ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
  return String(v).replace(/[ㅤᅟᅠﾠ​-‍⁠﻿­]/g, '').replace(/[\s   ]+/g, ' ').trim();
}

function cpvAsegurarTamano_(hoja, filas, cols) {
  const mf = hoja.getMaxRows(), mc = hoja.getMaxColumns();
  if (mf < filas) hoja.insertRowsAfter(mf, filas - mf);
  if (mc < cols) hoja.insertColumnsAfter(mc, cols - mc);
}

function cpvMensaje_(err) {
  return err && err.message ? err.message : String(err);
}

function cpvAvisar_(titulo, mensaje) {
  console.log(titulo + '\n' + mensaje);
  try {
    SpreadsheetApp.getUi().alert(titulo, mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    console.log('SIN INTERFAZ PARA EL AVISO.');
  }
}
