/**
 * API — Control de Gastos (backend para la app alojada en GitHub Pages)
 * ------------------------------------------------------------
 * Este archivo YA NO sirve el HTML directamente (eso ahora vive en
 * GitHub Pages). Aquí solo se responde a peticiones con datos, en
 * formato JSON, protegidas con una llave secreta.
 *
 * INSTALACIÓN:
 * 1. Pega este archivo completo en "Code.gs" (reemplaza todo).
 * 2. Cambia el valor de API_KEY de abajo por tu llave (te la di aparte,
 *    o genera una tú: cualquier texto largo y random sirve).
 * 3. Guarda.
 * 4. Implementar > Administrar implementaciones > ✏️ > cambia:
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario (Anyone)
 *    Versión: Nueva versión > Implementar.
 *    (Sí, "Cualquier usuario" sí suena raro para algo privado — pero
 *    sin la llave correcta, nadie puede hacer nada: cada petición se
 *    rechaza si no trae el API_KEY exacto.)
 * 5. Copia la URL que termina en /exec, la vas a necesitar en el
 *    frontend (archivo config.js de GitHub Pages).
 */

var API_KEY = 'q0RUm-tKprAogLl1QRoqOQr-QfX_eSLhW24NOKoV2Bo'; // <-- tu llave secreta

var SHEET_RESUMEN = 'Resumen';
var SHEET_REGISTRO = 'Registro';
var SHEET_CATEGORIAS = 'Categorías';
var SHEET_AHORRO = 'Ahorro';
var SHEET_ESTADO = 'Estado Pagos';
var MESES_ABR = ['Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

// GET solo para confirmar que la API está viva (no sirve HTML).
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, mensaje: 'API de Control de Gastos activa.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Todas las acciones reales entran por aquí, como POST con el body:
// { "key": "...", "action": "nombreFuncion", "params": {...} }
function doPost(e) {
  var respuesta;
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.key !== API_KEY) {
      respuesta = { ok: false, mensaje: 'Llave inválida.' };
    } else {
      respuesta = enrutarAccion_(body.action, body.params || {});
    }
  } catch (err) {
    respuesta = { ok: false, mensaje: 'Error: ' + err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(respuesta))
    .setMimeType(ContentService.MimeType.JSON);
}

function enrutarAccion_(action, p) {
  switch (action) {
    case 'getMesesDisponibles': return getMesesDisponibles();
    case 'getResumenMes': return getResumenMes(p.mesAbr);
    case 'getUltimosGastos': return getUltimosGastos(p.n);
    case 'getCategorias': return getCategorias();
    case 'registrarGasto': return registrarGasto(p.datos);
    case 'getFijosEstado': return getFijosEstado(p.mesAbr);
    case 'marcarFijoPagado': return marcarFijoPagado(p.fila, p.monto, p.fechaStr);
    case 'desmarcarFijoPagado': return desmarcarFijoPagado(p.fila);
    case 'actualizarMontoFijo': return actualizarMontoFijo(p.fila, p.nuevoMonto, p.aplicarFuturo);
    case 'actualizarDiaPago': return actualizarDiaPago(p.categoria, p.dia);
    case 'agregarFijoEstado': return agregarFijoEstado(p.categoria, p.monto, p.aplicarATodos, p.mesUnico);
    case 'eliminarFijoEstado': return eliminarFijoEstado(p.fila, p.aplicarATodos);
    case 'getAhorroResumen': return getAhorroResumen();
    case 'actualizarAhorroMes': return actualizarAhorroMes(p.fila, p.aportado, p.retiro, p.nota);
    case 'crearCategoria': return crearCategoria(p.nombre, p.tipo, p.montoInicial);
    case 'getSpreadsheetUrl': return getSpreadsheetUrl();
    default: return { ok: false, mensaje: 'Acción no reconocida: ' + action };
  }
}

// ============================================================
// UTILIDADES
// ============================================================
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function buscarFilaPorEtiqueta_(hoja, etiqueta, maxFila) {
  maxFila = maxFila || 80;
  var col = hoja.getRange(1, 1, maxFila, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (col[i][0] && col[i][0].toString().trim() === etiqueta) return i + 1;
  }
  return -1;
}

function buscarColumnaEnFila_(hoja, fila, etiqueta, maxCol) {
  maxCol = maxCol || 30;
  var row = hoja.getRange(fila, 1, 1, maxCol).getValues()[0];
  for (var i = 0; i < row.length; i++) {
    if (row[i] && row[i].toString().trim() === etiqueta) return i + 1;
  }
  return -1;
}

function mesActualIndex_() {
  var m = new Date().getMonth() + 1; // 1-12
  var idx = m - 7; // Jul = 0
  if (idx < 0 || idx > 5) idx = 0;
  return idx;
}

function encontrarFilaLibre_(hoja) {
  var col = hoja.getRange('A4:A' + hoja.getMaxRows()).getValues();
  for (var i = 0; i < col.length; i++) {
    if (col[i][0] === '' || col[i][0] === null) return i + 4;
  }
  return -1;
}

function formatoFecha_(d) {
  if (!d) return '';
  var dd = ('0' + d.getDate()).slice(-2);
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  return dd + '/' + mm + '/' + d.getFullYear();
}

// ============================================================
// INICIO: meses disponibles
// ============================================================
function getMesesDisponibles() {
  return { meses: MESES_ABR, actual: mesActualIndex_() };
}

// ============================================================
// INICIO: resumen del mes (disponible, gasto real, chart de categorías)
// ============================================================
function getResumenMes(mesAbr) {
  var hoja = ss_().getSheetByName(SHEET_RESUMEN);
  if (!hoja) return { error: 'No encontré la hoja "Resumen".' };

  var rowIngreso = buscarFilaPorEtiqueta_(hoja, 'Ingreso del mes');
  var rowHeadMes = rowIngreso - 1;
  var colMes = buscarColumnaEnFila_(hoja, rowHeadMes, mesAbr);
  var rowAhorro = buscarFilaPorEtiqueta_(hoja, 'Ahorro aportado (hoja Ahorro)');
  var rowRetiro = buscarFilaPorEtiqueta_(hoja, 'Retiro del ahorro (imprevisto)');
  var rowPlaneado = buscarFilaPorEtiqueta_(hoja, 'Disponible planeado');
  var rowGastoReal = buscarFilaPorEtiqueta_(hoja, 'Gasto real del mes');
  var rowRestante = buscarFilaPorEtiqueta_(hoja, 'Disponible restante');

  if (colMes === -1 || rowIngreso === -1) {
    return { error: 'No encontré los datos de ese mes en Resumen.' };
  }

  var resumen = {
    ingreso: hoja.getRange(rowIngreso, colMes).getValue(),
    ahorro: hoja.getRange(rowAhorro, colMes).getValue(),
    retiro: hoja.getRange(rowRetiro, colMes).getValue(),
    planeado: hoja.getRange(rowPlaneado, colMes).getValue(),
    gastoReal: hoja.getRange(rowGastoReal, colMes).getValue(),
    restante: hoja.getRange(rowRestante, colMes).getValue()
  };

  // Desglose por categoría para la gráfica
  var rowTituloCat = buscarFilaPorEtiqueta_(hoja, 'Gasto por Categoría y Mes');
  var categorias = [];
  if (rowTituloCat > -1) {
    var rowHeadCat = rowTituloCat + 1;
    var colMesCat = buscarColumnaEnFila_(hoja, rowHeadCat, mesAbr);
    for (var r = rowHeadCat + 1; r < rowHeadCat + 100; r++) {
      var nombre = hoja.getRange(r, 1).getValue();
      if (nombre === 'TOTAL') break;
      if (nombre && nombre.toString().trim() !== '') {
        var valor = hoja.getRange(r, colMesCat).getValue();
        if (valor && valor > 0) categorias.push({ nombre: nombre, valor: valor });
      }
    }
  }
  categorias.sort(function (a, b) { return b.valor - a.valor; });
  var top = categorias.slice(0, 5);
  var restoValor = categorias.slice(5).reduce(function (s, c) { return s + c.valor; }, 0);
  if (restoValor > 0) top.push({ nombre: 'Otros', valor: restoValor });
  resumen.categorias = top;

  return resumen;
}

// ============================================================
// ÚLTIMOS GASTOS (para Inicio)
// ============================================================
function getUltimosGastos(n) {
  var hoja = ss_().getSheetByName(SHEET_REGISTRO);
  if (!hoja) return [];
  var datos = hoja.getRange('A4:D' + hoja.getMaxRows()).getValues();
  var filas = [];
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] !== '' && datos[i][0] !== null) {
      filas.push({
        fecha: formatoFecha_(datos[i][0]),
        categoria: datos[i][1],
        descripcion: datos[i][2],
        monto: datos[i][3]
      });
    }
  }
  return filas.slice(-n).reverse();
}

// ============================================================
// CATEGORÍAS (para el selector de Agregar)
// ============================================================
function getCategorias() {
  var hoja = ss_().getSheetByName(SHEET_CATEGORIAS);
  if (!hoja) return [];
  var valores = hoja.getRange('A2:B' + hoja.getLastRow()).getValues();
  return valores
    .filter(function (fila) { return fila[0] && fila[0].toString().trim() !== ''; })
    .map(function (fila) { return { nombre: fila[0], tipo: fila[1] }; });
}

// ============================================================
// CREAR CATEGORÍA NUEVA (desde el formulario de Agregar, o desde migraciones)
// ============================================================
function crearCategoria(nombre, tipo, montoInicial) {
  try {
    var hoja = ss_().getSheetByName(SHEET_CATEGORIAS);
    if (!hoja) return { ok: false, mensaje: 'No encontré la hoja "Categorías".' };
    nombre = (nombre || '').toString().trim();
    if (!nombre) return { ok: false, mensaje: 'Escribe un nombre para la categoría.' };

    var existentes = getCategorias().map(function (c) { return c.nombre.toString().toLowerCase(); });
    if (existentes.indexOf(nombre.toLowerCase()) > -1) {
      return { ok: false, mensaje: 'Ya existe una categoría con ese nombre.' };
    }

    var col = hoja.getRange('A2:A' + hoja.getLastRow()).getValues();
    var filaLibre = -1;
    for (var i = 0; i < col.length; i++) {
      if (!col[i][0] || col[i][0].toString().trim() === '') { filaLibre = i + 2; break; }
    }
    if (filaLibre === -1) {
      return { ok: false, mensaje: 'No quedan filas en blanco en "Categorías". Agrega una fila más ahí manualmente.' };
    }

    hoja.getRange(filaLibre, 1).setValue(nombre);
    hoja.getRange(filaLibre, 2).setValue(tipo || 'Variable');
    hoja.getRange(filaLibre, 3).setValue('Mensual');
    hoja.getRange(filaLibre, 4).setValue(parseFloat(montoInicial) || 0);

    return { ok: true, mensaje: 'Categoría "' + nombre + '" creada ✅', nombre: nombre };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// ============================================================
// AGREGAR GASTO (o REEMBOLSO: mismo formulario, resta del total)
// ============================================================
function registrarGasto(datos) {
  try {
    var hojaReg = ss_().getSheetByName(SHEET_REGISTRO);
    if (!hojaReg) return { ok: false, mensaje: 'No encontré la hoja "Registro".' };

    var categorias = getCategorias().map(function (c) { return c.nombre; });
    if (categorias.indexOf(datos.categoria) === -1) {
      return { ok: false, mensaje: 'Esa categoría ya no existe. Recarga la página.' };
    }
    var monto = parseFloat(datos.monto);
    if (isNaN(monto) || monto <= 0) return { ok: false, mensaje: 'El monto no es válido.' };
    if (datos.tipo === 'reembolso') monto = -monto; // se resta del total de la categoría

    var fecha = datos.fecha ? new Date(datos.fecha + 'T00:00:00') : new Date();
    var filaLibre = encontrarFilaLibre_(hojaReg);
    if (filaLibre === -1) return { ok: false, mensaje: 'No queda espacio en "Registro".' };

    var descripcion = datos.descripcion || '';
    if (datos.tipo === 'reembolso') descripcion = '(Reembolso) ' + descripcion;

    hojaReg.getRange(filaLibre, 1, 1, 4).setValues([[fecha, datos.categoria, descripcion, monto]]);
    return { ok: true, mensaje: (datos.tipo === 'reembolso' ? 'Reembolso registrado ✅' : 'Gasto guardado ✅') };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// ============================================================
// FIJOS: checklist con partes/cuotas
// ============================================================
function getFijosEstado(mesAbr) {
  var hoja = ss_().getSheetByName(SHEET_ESTADO);
  if (!hoja) return { error: 'No encontré la hoja "Estado Pagos".', items: [] };

  var dias = getDiasPago_();
  var datos = hoja.getRange(4, 1, hoja.getLastRow() - 3, 8).getValues();
  var items = [];
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] === mesAbr) {
      items.push({
        fila: i + 4,
        categoria: datos[i][1],
        parte: datos[i][2],
        montoPlaneado: datos[i][3],
        pagado: datos[i][4] === true,
        montoPagado: datos[i][5] || null,
        fechaPago: datos[i][6] ? formatoFecha_(new Date(datos[i][6])) : null,
        diaPago: dias[datos[i][1]] || null
      });
    }
  }
  var pagados = items.filter(function (it) { return it.pagado; }).length;
  return { items: items, pagados: pagados, total: items.length };
}

function marcarFijoPagado(fila, monto, fechaStr) {
  try {
    var hoja = ss_().getSheetByName(SHEET_ESTADO);
    var hojaReg = ss_().getSheetByName(SHEET_REGISTRO);
    if (!hoja || !hojaReg) return { ok: false, mensaje: 'Faltan hojas del archivo.' };

    var categoria = hoja.getRange(fila, 2).getValue();
    var parte = hoja.getRange(fila, 3).getValue();
    var montoFinal = parseFloat(monto);
    if (isNaN(montoFinal) || montoFinal <= 0) return { ok: false, mensaje: 'Monto no válido.' };
    var fecha = fechaStr ? new Date(fechaStr + 'T00:00:00') : new Date();

    var filaLibre = encontrarFilaLibre_(hojaReg);
    if (filaLibre === -1) return { ok: false, mensaje: 'No hay espacio en "Registro".' };

    var totalPartes = (parte || '').toString();
    var descripcion = totalPartes.indexOf('/1') > -1 ? categoria : categoria + ' (parte ' + parte + ')';
    hojaReg.getRange(filaLibre, 1, 1, 4).setValues([[fecha, categoria, descripcion, montoFinal]]);

    hoja.getRange(fila, 5).setValue(true);
    hoja.getRange(fila, 6).setValue(montoFinal);
    hoja.getRange(fila, 7).setValue(fecha);
    hoja.getRange(fila, 8).setValue(filaLibre);

    return { ok: true, mensaje: 'Marcado como pagado ✅' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

function desmarcarFijoPagado(fila) {
  try {
    var hoja = ss_().getSheetByName(SHEET_ESTADO);
    var hojaReg = ss_().getSheetByName(SHEET_REGISTRO);
    var filaRegistro = hoja.getRange(fila, 8).getValue();
    if (filaRegistro && hojaReg) {
      hojaReg.getRange(filaRegistro, 1, 1, 4).clearContent();
    }
    hoja.getRange(fila, 5).setValue(false);
    hoja.getRange(fila, 6).clearContent();
    hoja.getRange(fila, 7).clearContent();
    hoja.getRange(fila, 8).clearContent();
    return { ok: true, mensaje: 'Se deshizo el pago.' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// Edita el monto planeado de un fijo. Si aplicarFuturo=true, también
// actualiza la misma categoría/parte en los meses siguientes.
function actualizarMontoFijo(fila, nuevoMonto, aplicarFuturo) {
  try {
    var hoja = ss_().getSheetByName(SHEET_ESTADO);
    var monto = parseFloat(nuevoMonto);
    if (isNaN(monto) || monto <= 0) return { ok: false, mensaje: 'Monto no válido.' };

    var mesFila = hoja.getRange(fila, 1).getValue();
    var categoria = hoja.getRange(fila, 2).getValue();
    var parte = hoja.getRange(fila, 3).getValue();
    hoja.getRange(fila, 4).setValue(monto);

    var actualizados = 1;
    if (aplicarFuturo) {
      var idxMes = MESES_ABR.indexOf(mesFila);
      var datos = hoja.getRange(4, 1, hoja.getLastRow() - 3, 8).getValues();
      for (var i = 0; i < datos.length; i++) {
        var filaActual = i + 4;
        if (filaActual === fila) continue;
        var idxActual = MESES_ABR.indexOf(datos[i][0]);
        if (datos[i][1] === categoria && datos[i][2] === parte && idxActual > idxMes && datos[i][4] !== true) {
          hoja.getRange(filaActual, 4).setValue(monto);
          actualizados++;
        }
      }
    }
    return { ok: true, mensaje: 'Monto actualizado en ' + actualizados + ' mes(es) ✅' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// ============================================================
// RECORDATORIOS: día de pago aproximado por categoría fija
// ============================================================
var COL_DIA_PAGO = 7; // columna G en "Categorías" (después de Nota en F)

// Ejecuta esto UNA sola vez (▶ en el editor) para preparar la columna.
// Es seguro volver a ejecutarlo: no borra datos existentes.
function configurarDiaPago() {
  var hoja = ss_().getSheetByName(SHEET_CATEGORIAS);
  if (!hoja) { Logger.log('No existe la hoja Categorías'); return; }
  var encabezado = hoja.getRange(1, COL_DIA_PAGO).getValue();
  if (encabezado !== 'Día de Pago (1-31)') {
    hoja.getRange(1, COL_DIA_PAGO).setValue('Día de Pago (1-31)');
  }
  SpreadsheetApp.getUi().alert('Listo: la columna "Día de Pago" ya está disponible en Categorías (columna G). Ahora puedes definir el día desde la pestaña Fijos de la app.');
}

function getCategoriaFila_(nombre) {
  var hoja = ss_().getSheetByName(SHEET_CATEGORIAS);
  var nombres = hoja.getRange('A2:A' + hoja.getLastRow()).getValues();
  for (var i = 0; i < nombres.length; i++) {
    if (nombres[i][0] === nombre) return i + 2;
  }
  return -1;
}

function getDiasPago_() {
  var hoja = ss_().getSheetByName(SHEET_CATEGORIAS);
  var mapa = {};
  if (!hoja) return mapa;
  var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, COL_DIA_PAGO).getValues();
  datos.forEach(function (fila) {
    if (fila[0]) mapa[fila[0]] = fila[COL_DIA_PAGO - 1] || null;
  });
  return mapa;
}

function actualizarDiaPago(categoria, dia) {
  try {
    var fila = getCategoriaFila_(categoria);
    if (fila === -1) return { ok: false, mensaje: 'No encontré esa categoría.' };
    var diaNum = parseInt(dia, 10);
    if (isNaN(diaNum) || diaNum < 1 || diaNum > 31) return { ok: false, mensaje: 'Escribe un día entre 1 y 31.' };
    ss_().getSheetByName(SHEET_CATEGORIAS).getRange(fila, COL_DIA_PAGO).setValue(diaNum);
    return { ok: true, mensaje: 'Día de pago guardado ✅' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// Revisa los fijos del mes actual sin pagar y cuyo día ya pasó (o es hoy),
// y te manda un correo recordatorio. Para automatizarlo, ver instrucciones
// al final de este archivo ("ACTIVAR RECORDATORIOS DIARIOS").
function revisarRecordatorios() {
  var hoy = new Date();
  var idxMes = hoy.getMonth() + 1 - 7;
  if (idxMes < 0 || idxMes > 5) return;
  var mesAbr = MESES_ABR[idxMes];
  var diaHoy = hoy.getDate();

  var dias = getDiasPago_();
  var estado = getFijosEstado(mesAbr);
  if (estado.error) return;

  var pendientes = estado.items.filter(function (it) {
    if (it.pagado) return false;
    var dia = dias[it.categoria];
    return dia && diaHoy >= dia;
  });

  if (pendientes.length === 0) return;

  var cuerpo = 'Tienes ' + pendientes.length + ' pago(s) fijo(s) pendiente(s) este mes:\n\n';
  pendientes.forEach(function (it) {
    cuerpo += '• ' + it.categoria + (it.parte.indexOf('/1') === -1 ? ' (' + it.parte + ')' : '') +
      ' — ' + it.montoPlaneado.toLocaleString('es-CO') + ' COP\n';
  });
  cuerpo += '\nMárcalos como pagados en la app cuando los hagas.';

  MailApp.sendEmail(Session.getActiveUser().getEmail(), '💰 Pagos fijos pendientes (' + mesAbr + ')', cuerpo);
}

/**
 * ACTIVAR RECORDATORIOS DIARIOS (una sola vez):
 * 1. En el editor de Apps Script, ícono del reloj (⏰) en el menú izquierdo
 *    ("Activadores" / "Triggers").
 * 2. "Añadir activador" (esquina inferior derecha).
 * 3. Función a ejecutar: revisarRecordatorios
 * 4. Origen del evento: Basado en tiempo
 * 5. Tipo de activador basado en tiempo: Temporizador diario
 * 6. Elige un horario, ej. "entre las 8:00 y las 9:00".
 * 7. Guardar (te pedirá autorizar el envío de correos, es normal).
 *
 * A partir de ahí, Google ejecuta la revisión sola todos los días,
 * sin que abras nada, y te llega un correo solo si hay algo pendiente.
 */

// Agrega un nuevo pago fijo al checklist. Si aplicarATodos=true, lo agrega
// en los 6 meses (Jul-Dic); si no, solo en mesUnico.
function agregarFijoEstado(categoria, monto, aplicarATodos, mesUnico) {
  try {
    var hoja = ss_().getSheetByName(SHEET_ESTADO);
    var montoNum = parseFloat(monto);
    if (isNaN(montoNum) || montoNum <= 0) return { ok: false, mensaje: 'Monto no válido.' };

    var meses = aplicarATodos ? MESES_ABR : [mesUnico];
    var filaInicio = hoja.getLastRow() + 1;
    var filas = meses.map(function (m) { return [m, categoria, '1/1', montoNum, false, null, null, null]; });
    hoja.getRange(filaInicio, 1, filas.length, 8).setValues(filas);

    return { ok: true, mensaje: 'Fijo agregado a ' + filas.length + ' mes(es) ✅' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// Elimina un fijo del checklist. Si aplicarATodos=true, elimina esa misma
// categoría/parte en TODOS los meses (pasados y futuros); si no, solo esa fila.
function eliminarFijoEstado(fila, aplicarATodos) {
  try {
    var hoja = ss_().getSheetByName(SHEET_ESTADO);
    if (!aplicarATodos) {
      hoja.deleteRow(fila);
      return { ok: true, mensaje: 'Fijo eliminado de ese mes ✅' };
    }
    var categoria = hoja.getRange(fila, 2).getValue();
    var parte = hoja.getRange(fila, 3).getValue();
    var datos = hoja.getRange(4, 1, hoja.getLastRow() - 3, 8).getValues();
    var filasABorrar = [];
    for (var i = 0; i < datos.length; i++) {
      if (datos[i][1] === categoria && datos[i][2] === parte) filasABorrar.push(i + 4);
    }
    filasABorrar.sort(function (a, b) { return b - a; }); // de abajo hacia arriba
    filasABorrar.forEach(function (f) { hoja.deleteRow(f); });
    return { ok: true, mensaje: 'Fijo eliminado de ' + filasABorrar.length + ' mes(es) ✅' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// ============================================================
// MIGRACIÓN (ejecutar UNA sola vez, manualmente ▶ en el editor):
// separa "Plan de datos y parqueadero" en dos categorías: "Plan de
// datos" y "Parqueadero". Divide el monto 50/50; ajusta después desde
// la app con el lápiz ✎ si no era mitad y mitad.
// ============================================================
function dividirParqueaderoYDatos() {
  var ui = SpreadsheetApp.getUi();
  var NOMBRE_VIEJO = 'Plan de datos y parqueadero';
  var hojaCat = ss_().getSheetByName(SHEET_CATEGORIAS);
  var filaCombo = getCategoriaFila_(NOMBRE_VIEJO);
  if (filaCombo === -1) {
    ui.alert('No encontré "' + NOMBRE_VIEJO + '" en Categorías. Puede que ya lo hayas dividido antes.');
    return;
  }

  var montoTotal = hojaCat.getRange(filaCombo, 4).getValue();
  var mitad = montoTotal / 2;

  hojaCat.getRange(filaCombo, 1).setValue('Plan de datos');
  hojaCat.getRange(filaCombo, 4).setValue(mitad);

  var resultado = crearCategoria('Parqueadero', 'Fijo', mitad);
  if (!resultado.ok) {
    ui.alert('Separé "Plan de datos" en Categorías, pero no pude crear "Parqueadero": ' + resultado.mensaje);
    return;
  }

  var hojaEstado = ss_().getSheetByName(SHEET_ESTADO);
  var datos = hojaEstado.getRange(4, 1, hojaEstado.getLastRow() - 3, 8).getValues();
  var filasNuevas = [];
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][1] === NOMBRE_VIEJO) {
      var filaReal = i + 4;
      hojaEstado.getRange(filaReal, 2).setValue('Plan de datos');
      hojaEstado.getRange(filaReal, 4).setValue(mitad);
      filasNuevas.push([datos[i][0], 'Parqueadero', '1/1', mitad, false, null, null, null]);
    }
  }
  if (filasNuevas.length > 0) {
    var filaDestino = hojaEstado.getLastRow() + 1;
    hojaEstado.getRange(filaDestino, 1, filasNuevas.length, 8).setValues(filasNuevas);
  }

  ui.alert('Listo: "Plan de datos" y "Parqueadero" ya están separados, cada uno con $' +
    Math.round(mitad).toLocaleString('es-CO') + '. Si no era mitad y mitad, ajusta el monto de cada uno desde la app (lápiz ✎ en la pestaña Fijos).');
}

// ============================================================
// AJUSTES: link directo a la hoja de cálculo (para el botón de Configuración)
// ============================================================
function getSpreadsheetUrl() {
  return ss_().getUrl();
}
// ============================================================
// AHORRO
// ============================================================
function getAhorroResumen() {
  var hoja = ss_().getSheetByName(SHEET_AHORRO);
  if (!hoja) return { error: 'No encontré la hoja "Ahorro".' };

  var rowHead = buscarFilaPorEtiqueta_(hoja, 'Mes');
  var meses = [];
  for (var r = rowHead + 1; r < rowHead + 30; r++) {
    var mes = hoja.getRange(r, 1).getValue();
    if (!mes || mes === 'TOTAL') break;
    meses.push({
      fila: r,
      mes: mes,
      aportado: hoja.getRange(r, 2).getValue(),
      retiro: hoja.getRange(r, 3).getValue(),
      nota: hoja.getRange(r, 4).getValue() || '',
      saldo: hoja.getRange(r, 5).getValue()
    });
  }
  var saldoActual = meses.length ? meses[meses.length - 1].saldo : 0;
  return { meses: meses, saldoActual: saldoActual };
}

function actualizarAhorroMes(fila, aportado, retiro, nota) {
  try {
    var hoja = ss_().getSheetByName(SHEET_AHORRO);
    hoja.getRange(fila, 2).setValue(parseFloat(aportado) || 0);
    hoja.getRange(fila, 3).setValue(parseFloat(retiro) || 0);
    hoja.getRange(fila, 4).setValue(nota || '');
    return { ok: true, mensaje: 'Ahorro actualizado ✅' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}
