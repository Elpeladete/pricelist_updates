// Última actualización: 2026-04-29 — fuente unificada (nueva planilla)
// =============================================================================
// CONFIGURACIÓN — Modificar solo estos valores si cambian los IDs o nombres
// =============================================================================

// Planilla ORIGEN: col A = "[SKU / INTERNO] descripción", col B = "USD X,XX"
var ORIGEN_ID   = "1xVIFLsVGg4-GF65BkeM2xkGDlGAp6pH_91dgX31WyT0";
var ORIGEN_HOJA = "";  // Nombre de la hoja. Vacío = primera hoja del archivo.

// Planilla DESTINO: col A = CODIGO DYE, col B = CODIGO DJI, col X = precio actualizado
var DESTINO_ID = "1x5nI3DsJOhcYnO71rQww_vXYteTGvXcbGLhCPUCEw4w";

// Hojas del destino a procesar
var HOJAS_DESTINO = [
  { nombre: "PEDIDOS ACTUALIZADOS", colDYE: 1, colDJI: 2, colPrecio: 24, filaInicio: 2 },
  { nombre: "DRON T100",            colDYE: 1, colDJI: 2, colPrecio: 24, filaInicio: 2 },
];

var TEXTO_NO_ENCONTRADO = "NO ENCONTRADO";

// =============================================================================
// NORMALIZACIÓN DE CÓDIGOS
// =============================================================================

/**
 * Elimina el último segmento separado por punto.
 * "CP.AG.575.01" → "CP.AG.575"
 */
function normalizarCodigo(codigo) {
  var ultimoPunto = codigo.lastIndexOf('.');
  if (ultimoPunto === -1) return codigo;
  return codigo.substring(0, ultimoPunto);
}

/**
 * Elimina todos los guiones.
 * "DJI-R891" → "DJIR891"
 */
function sinGuiones(codigo) {
  return codigo.replace(/-/g, '');
}

// =============================================================================
// FUNCIÓN PRINCIPAL
// =============================================================================

/**
 * Lee los precios de la planilla origen y los escribe en cada hoja destino.
 * Se ejecuta automáticamente cada 1 hora (ver instalarTrigger).
 */
function actualizarPrecios() {
  try {
    var ssOrigen  = SpreadsheetApp.openById(ORIGEN_ID);
    var mapas     = construirMapasLookup(ssOrigen);
    var ssDestino = SpreadsheetApp.openById(DESTINO_ID);

    for (var h = 0; h < HOJAS_DESTINO.length; h++) {
      procesarHojaDestino(ssDestino, HOJAS_DESTINO[h], mapas);
    }
  } catch (e) {
    Logger.log("ERROR en actualizarPrecios: " + e.message);
    throw e;
  }
}

// =============================================================================
// PROCESAR HOJA DESTINO
// =============================================================================

function procesarHojaDestino(ssDestino, def, mapas) {
  var hoja = ssDestino.getSheetByName(def.nombre);
  if (!hoja) {
    Logger.log('ADVERTENCIA: Hoja destino "' + def.nombre + '" no encontrada. Se omite.');
    return;
  }

  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < def.filaInicio) {
    Logger.log('Hoja "' + def.nombre + '": sin datos.');
    return;
  }

  var filasData = ultimaFila - def.filaInicio + 1;
  var codigos   = hoja.getRange(def.filaInicio, def.colDYE, filasData, 2).getValues();
  var precios   = [];

  for (var i = 0; i < codigos.length; i++) {
    var codigoDYE = String(codigos[i][0]).trim();  // Col A — CODIGO DYE
    var codigoDJI = String(codigos[i][1]).trim();  // Col B — CODIGO DJI

    if ((codigoDJI === "" || codigoDJI === "0") && (codigoDYE === "" || codigoDYE === "0")) {
      precios.push([""]);
      continue;
    }

    var precio = buscarPrecio(codigoDJI, codigoDYE, mapas);
    precios.push([precio !== null ? precio : TEXTO_NO_ENCONTRADO]);
  }

  hoja.getRange(def.filaInicio, def.colPrecio, filasData, 1).setValues(precios);
  Logger.log('Hoja "' + def.nombre + '": ' + filasData + ' filas procesadas.');
}

// =============================================================================
// BÚSQUEDA DE PRECIO
// =============================================================================

/**
 * Intenta primero por CODIGO DJI (→ SKU del origen),
 * luego por CODIGO DYE (→ código interno del origen).
 * Cada búsqueda prueba: exacto, sin guiones, sin último sufijo, combinación.
 */
function buscarPrecio(codigoDJI, codigoDYE, mapas) {
  if (codigoDJI !== "" && codigoDJI !== "0") {
    var p = buscarEnMapa(codigoDJI, mapas.sku);
    if (p !== null) return p;
  }
  if (codigoDYE !== "" && codigoDYE !== "0") {
    var p2 = buscarEnMapa(codigoDYE, mapas.interno);
    if (p2 !== null) return p2;
  }
  return null;
}

/**
 * Busca un código en un mapa probando 4 variantes de normalización.
 */
function buscarEnMapa(codigo, mapa) {
  if (!mapa) return null;
  // 1. Exacto
  if (mapa.hasOwnProperty(codigo)) return mapa[codigo];
  // 2. Sin guiones
  var norm = sinGuiones(codigo);
  if (norm !== codigo && mapa.hasOwnProperty(norm)) return mapa[norm];
  // 3. Sin último sufijo
  var base = normalizarCodigo(codigo);
  if (base !== codigo && mapa.hasOwnProperty(base)) return mapa[base];
  // 4. Sin sufijo + sin guiones
  var baseNorm = sinGuiones(base);
  if (baseNorm !== base && mapa.hasOwnProperty(baseNorm)) return mapa[baseNorm];
  return null;
}

// =============================================================================
// CONSTRUIR MAPAS DE LOOKUP DESDE LA PLANILLA ORIGEN
// =============================================================================

/**
 * Lee la planilla origen y construye dos mapas de búsqueda:
 *   mapas.sku     → { SKU: precio }      (búsqueda por CODIGO DJI)
 *   mapas.interno → { INTERNO: precio }  (búsqueda por CODIGO DYE)
 *
 * Formato esperado en col A: "[SKU / INTERNO] descripción del producto"
 * Formato esperado en col B: "USD 1,50"
 *
 * Cada código se indexa también en sus variantes (sin guiones, sin sufijo).
 */
function construirMapasLookup(ssOrigen) {
  var hoja = ORIGEN_HOJA ? ssOrigen.getSheetByName(ORIGEN_HOJA) : ssOrigen.getSheets()[0];
  if (!hoja) {
    Logger.log('ERROR: No se encontró la hoja origen.');
    return { sku: {}, interno: {} };
  }

  var datos       = hoja.getDataRange().getValues();
  var mapaSKU     = {};
  var mapaInterno = {};
  var patron      = /^\[([^\]\/]+?)\s*\/\s*([^\]]+?)\]/;  // Extrae [SKU / INTERNO]

  for (var r = 0; r < datos.length; r++) {
    var celda  = String(datos[r][0]).trim();
    var precio = parsearPrecio(String(datos[r][1]).trim());
    if (precio === null) continue;

    var match = patron.exec(celda);
    if (!match) continue;

    indexarCodigo(match[1].trim(), precio, mapaSKU);
    indexarCodigo(match[2].trim(), precio, mapaInterno);
  }

  Logger.log('Origen cargado: ' + Object.keys(mapaSKU).length + ' SKUs | ' + Object.keys(mapaInterno).length + ' códigos internos.');
  return { sku: mapaSKU, interno: mapaInterno };
}

/**
 * Agrega un código y sus variantes normalizadas al mapa.
 */
function indexarCodigo(codigo, precio, mapa) {
  if (!codigo) return;
  mapa[codigo] = precio;
  var norm = sinGuiones(codigo);
  if (norm !== codigo) mapa[norm] = precio;
  var base = normalizarCodigo(codigo);
  if (base !== codigo && !mapa.hasOwnProperty(base)) mapa[base] = precio;
  var baseNorm = sinGuiones(base);
  if (baseNorm !== base && !mapa.hasOwnProperty(baseNorm)) mapa[baseNorm] = precio;
}

/**
 * Parsea "USD 1,50" o "USD 12.00" → 1.5 / 12.0
 * Devuelve null si no puede parsear.
 */
function parsearPrecio(texto) {
  var match = /USD\s*([\d]+[,\.][\d]+)/i.exec(texto);
  if (!match) return null;
  return parseFloat(match[1].replace(',', '.'));
}

// =============================================================================
// DIAGNÓSTICO
// =============================================================================

function diagnostico() {
  try { Logger.log("Cuenta activa: " + Session.getActiveUser().getEmail()); }
  catch (e) { Logger.log("Cuenta: (no disponible) — " + e.message); }

  try {
    var o = SpreadsheetApp.openById(ORIGEN_ID);
    var h = ORIGEN_HOJA ? o.getSheetByName(ORIGEN_HOJA) : o.getSheets()[0];
    Logger.log("ORIGEN OK: " + o.getName() + " | Hoja: " + (h ? h.getName() : 'NO ENCONTRADA') + " | Filas: " + (h ? h.getLastRow() : 0));
  } catch (e) { Logger.log("ORIGEN ERROR: " + e.message); }

  try {
    var d = SpreadsheetApp.openById(DESTINO_ID);
    Logger.log("DESTINO OK: " + d.getName());
  } catch (e) { Logger.log("DESTINO ERROR: " + e.message); }
}

// =============================================================================
// GESTIÓN DE TRIGGERS
// =============================================================================

/**
 * Instala un trigger que ejecuta actualizarPrecios() cada 1 hora.
 * Ejecutar solo una vez manualmente.
 */
function instalarTrigger() {
  eliminarTriggers();
  ScriptApp.newTrigger("actualizarPrecios").timeBased().everyHours(1).create();
  Logger.log("Trigger instalado: actualizarPrecios cada 1 hora.");
}

function eliminarTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "actualizarPrecios") {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log("Trigger eliminado: " + triggers[i].getUniqueId());
    }
  }
}
