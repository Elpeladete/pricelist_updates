// =============================================================================
// CONFIGURACIÓN — Modificar solo estos valores si cambian los IDs o nombres
// =============================================================================

var ORIGEN_ID  = "1SeUISHBX7Zep0U6yCClz1z6pU2aWrAklqI2OBIcOa2c";
var DESTINO_ID = "1x5nI3DsJOhcYnO71rQww_vXYteTGvXcbGLhCPUCEw4w";

var DESTINO_HOJA           = "PEDIDOS ACTUALIZADOS";
var DESTINO_COL_CODIGODYE  = 1;   // Columna A  → "CODIGO DYE"
var DESTINO_COL_CODIGO     = 2;   // Columna B  → "CODIGO DJI"
var DESTINO_COL_PRECIO     = 24;  // Columna X  → "Precio actualizado"
var DESTINO_FILA_INICIO    = 2;   // Primera fila de datos (fila 1 = encabezados)

// Lista de hojas del archivo DESTINO a procesar.
// Todas comparten la misma estructura: col A = CODIGO DYE, col B = CODIGO DJI, col X = precio.
var HOJAS_DESTINO = [
  { nombre: "PEDIDOS ACTUALIZADOS", colDYE: 1, colDJI: 2, colPrecio: 24, filaInicio: 2 },
  { nombre: "DRON T100",            colDYE: 1, colDJI: 2, colPrecio: 24, filaInicio: 2 },
];

// Definición de las hojas del archivo origen.
// colPrecio   → índice base 1 de la columna con el precio
// colInterno  → índice base 1 de la columna "Código interno" (búsqueda secundaria por CODIGO DYE)
var HOJAS_ORIGEN = [
  { nombre: "T100",   colPrecio: 7, colInterno: 2 },  // Precio: Col G | Código interno: Col B
  { nombre: "T70p",   colPrecio: 7, colInterno: 2 },  // Precio: Col G | Código interno: Col B
  { nombre: "D14000", colPrecio: 6, colInterno: 1 },  // Precio: Col F | Código interno: Col A
];

var TEXTO_NO_ENCONTRADO  = "NO ENCONTRADO";
var ENCABEZADO_SKU       = "Código SKU";      // Encabezado en origen para búsqueda por CODIGO DJI
var ENCABEZADO_INTERNO   = "Código interno";  // Encabezado en origen para búsqueda por CODIGO DYE

// Fuente terciaria: PDF en Google Drive
// Formato en PDF: [CODIGO_DJI / CODIGO_DYE] Descripción ... USD X,XX
var PDF_ID        = "1358QcGtL3-Cto59ho9OCYTk1jZcUn7pO";
var PDF_CACHE_KEY = "mapaPDF_v1";  // Cambiar la versión fuerza re-parseo

// =============================================================================
// NORMALIZACIÓN DE CÓDIGOS
// =============================================================================

/**
 * Elimina el último segmento separado por punto de un código.
 * Ejemplo: "CP.AG.00000575.01" → "CP.AG.00000575"
 * Si el código no tiene punto, devuelve el mismo código sin cambios.
 */
function normalizarCodigo(codigo) {
  var ultimoPunto = codigo.lastIndexOf('.');
  if (ultimoPunto === -1) return codigo;
  return codigo.substring(0, ultimoPunto);
}

/**
 * Elimina todos los guiones de un código.
 * Ejemplo: "DJI-R891" → "DJIR891"
 */
function sinGuiones(codigo) {
  return codigo.replace(/-/g, '');
}

// =============================================================================
// FUNCIÓN PRINCIPAL
// =============================================================================

/**
 * Lee los códigos DJI del archivo destino, los busca en el archivo origen
 * y escribe el precio actualizado en la columna X.
 * Se ejecuta automáticamente cada 6 horas (ver instalarTrigger).
 */
function actualizarPrecios() {
  try {
    // ------------------------------------------------------------------
    // 1. Construir mapas de lookup desde el archivo origen (una sola vez)
    // ------------------------------------------------------------------
    var ssOrigen = SpreadsheetApp.openById(ORIGEN_ID);
    var mapas    = construirMapasLookup(ssOrigen);

    // ------------------------------------------------------------------
    // 2. Cargar mapa del PDF (cacheado en PropertiesService)
    // ------------------------------------------------------------------
    var mapaPDF = obtenerMapaPDF();

    // ------------------------------------------------------------------
    // 3. Procesar cada hoja del archivo destino
    // ------------------------------------------------------------------
    var ssDestino = SpreadsheetApp.openById(DESTINO_ID);

    for (var h = 0; h < HOJAS_DESTINO.length; h++) {
      var def = HOJAS_DESTINO[h];
      procesarHojaDestino(ssDestino, def, mapas, mapaPDF);
    }

  } catch (e) {
    Logger.log("ERROR en actualizarPrecios: " + e.message);
    throw e;
  }
}

// =============================================================================
// FUNCIÓN AUXILIAR: procesar una hoja del destino
// =============================================================================

/**
 * Lee los códigos de una hoja del archivo destino, busca los precios en los
 * mapas del origen y escribe los resultados en la columna de precio indicada.
 *
 * @param {Spreadsheet} ssDestino - Spreadsheet del archivo destino.
 * @param {Object}      def       - Definición de la hoja: { nombre, colDYE, colDJI, colPrecio, filaInicio }.
 * @param {Object}      mapas     - Mapas de lookup construidos por construirMapasLookup().
 */
function procesarHojaDestino(ssDestino, def, mapas, mapaPDF) {
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

  var filasData    = ultimaFila - def.filaInicio + 1;
  var rangoCodigos = hoja.getRange(def.filaInicio, def.colDYE, filasData, 2);
  var codigos      = rangoCodigos.getValues();  // [[dyeCod, djiCod], ...]
  var precios      = [];

  for (var i = 0; i < codigos.length; i++) {
    var codigoDYE = String(codigos[i][0]).trim();  // Col A — CODIGO DYE
    var codigoDJI = String(codigos[i][1]).trim();  // Col B — CODIGO DJI

    if ((codigoDJI === "" || codigoDJI === "0") && (codigoDYE === "" || codigoDYE === "0")) {
      precios.push([""]);
      continue;
    }

    var precioEncontrado = null;

    // ---- BÚSQUEDA PRIMARIA: CODIGO DJI → "Código SKU" del origen ----
    if (codigoDJI !== "" && codigoDJI !== "0") {
      var codigoDJIBase = normalizarCodigo(codigoDJI);
      for (var j = 0; j < HOJAS_ORIGEN.length; j++) {
        var entrada = mapas[HOJAS_ORIGEN[j].nombre];
        if (!entrada) continue;
        if (entrada.exacto.hasOwnProperty(codigoDJI)) {
          precioEncontrado = entrada.exacto[codigoDJI]; break;
        }
        if (codigoDJIBase !== codigoDJI && entrada.exacto.hasOwnProperty(codigoDJIBase)) {
          precioEncontrado = entrada.exacto[codigoDJIBase]; break;
        }
        if (entrada.base.hasOwnProperty(codigoDJI)) {
          precioEncontrado = entrada.base[codigoDJI]; break;
        }
      }
    }

    // ---- BÚSQUEDA SECUNDARIA: CODIGO DYE → "Código interno" del origen ----
    if (precioEncontrado === null && codigoDYE !== "" && codigoDYE !== "0") {
      var codigoDYEBase = normalizarCodigo(codigoDYE);
      var codigoDYENorm = sinGuiones(codigoDYE);  // e.g. "DJIR891" para buscar "DJI-R891" en origen
      for (var k = 0; k < HOJAS_ORIGEN.length; k++) {
        var entradaDYE = mapas[HOJAS_ORIGEN[k].nombre];
        if (!entradaDYE) continue;
        // 1. Exacto
        if (entradaDYE.interno.hasOwnProperty(codigoDYE)) {
          precioEncontrado = entradaDYE.interno[codigoDYE]; break;
        }
        // 2. Sin último segmento
        if (codigoDYEBase !== codigoDYE && entradaDYE.interno.hasOwnProperty(codigoDYEBase)) {
          precioEncontrado = entradaDYE.interno[codigoDYEBase]; break;
        }
        // 3. Origen con sufijo, destino sin él
        if (entradaDYE.internoBase.hasOwnProperty(codigoDYE)) {
          precioEncontrado = entradaDYE.internoBase[codigoDYE]; break;
        }
        // 4. Sin guiones (DJI-R891 ↔ DJIR891)
        if (codigoDYENorm !== codigoDYE && entradaDYE.internoNorm.hasOwnProperty(codigoDYENorm)) {
          precioEncontrado = entradaDYE.internoNorm[codigoDYENorm]; break;
        }
        // 5. Sin guiones buscando el código del destino directamente en internoNorm
        if (entradaDYE.internoNorm.hasOwnProperty(codigoDYE)) {
          precioEncontrado = entradaDYE.internoNorm[codigoDYE]; break;
        }
      }
    }
      // ---- BÚSQUEDA TERCIARIA: PDF (fallback final) ----
      // Solo si las búsquedas primaria y secundaria no encontraron precio.
      // En el PDF: [CODIGO_DJI / CODIGO_DYE] → el primero es DJI, el segundo es DYE.
      if (precioEncontrado === null && mapaPDF) {
        // Intentar con CODIGO DJI (col B)
        if (codigoDJI !== "" && codigoDJI !== "0") {
          var djiNorm = sinGuiones(codigoDJI);
          var djiBase = normalizarCodigo(codigoDJI);
          if      (mapaPDF.hasOwnProperty(codigoDJI)) { precioEncontrado = mapaPDF[codigoDJI]; }
          else if (mapaPDF.hasOwnProperty(djiNorm))   { precioEncontrado = mapaPDF[djiNorm]; }
          else if (djiBase !== codigoDJI && mapaPDF.hasOwnProperty(djiBase)) { precioEncontrado = mapaPDF[djiBase]; }
        }
        // Intentar con CODIGO DYE (col A) si DJI no resolvió
        if (precioEncontrado === null && codigoDYE !== "" && codigoDYE !== "0") {
          var dyeNorm = sinGuiones(codigoDYE);
          var dyeBase = normalizarCodigo(codigoDYE);
          if      (mapaPDF.hasOwnProperty(codigoDYE)) { precioEncontrado = mapaPDF[codigoDYE]; }
          else if (mapaPDF.hasOwnProperty(dyeNorm))   { precioEncontrado = mapaPDF[dyeNorm]; }
          else if (dyeBase !== codigoDYE && mapaPDF.hasOwnProperty(dyeBase)) { precioEncontrado = mapaPDF[dyeBase]; }
        }
      }
    precios.push([precioEncontrado !== null ? precioEncontrado : TEXTO_NO_ENCONTRADO]);
  }

  hoja.getRange(def.filaInicio, def.colPrecio, filasData, 1).setValues(precios);
  Logger.log('Hoja "' + def.nombre + '": ' + filasData + ' filas procesadas.');
}

// =============================================================================
// FUENTE TERCIARIA: PDF
// =============================================================================

/**
 * Devuelve el mapa { codigo: precio } del PDF.
 * Usa PropertiesService como caché; el PDF se parsea solo la primera vez
 * (o cuando se ejecuta refrescarCachePDF() manualmente).
 */
function obtenerMapaPDF() {
  try {
    var cached = PropertiesService.getScriptProperties().getProperty(PDF_CACHE_KEY);
    if (cached) {
      var mapa = JSON.parse(cached);
      Logger.log('PDF: usando caché (' + Object.keys(mapa).length + ' códigos).');
      return mapa;
    }
    return parsearPDF();
  } catch (e) {
    Logger.log('ERROR en obtenerMapaPDF: ' + e.message);
    return {};
  }
}

/**
 * Convierte el PDF a texto usando OCR (copia temporal como Google Doc),
 * parsea los códigos y precios, cachea el resultado y elimina la copia.
 */
function parsearPDF() {
  var copiaTempId = null;
  try {
    // Crear copia como Google Doc para que Drive aplique OCR automáticamente
    var copia = Drive.Files.copy(
      { title: 'ocr_tmp_pricelist_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
      PDF_ID
    );
    copiaTempId = copia.id;
    Utilities.sleep(5000);  // Esperar a que el OCR termine

    var texto = DocumentApp.openById(copiaTempId).getBody().getText();
    var mapa  = parsearTextoPDF(texto);

    PropertiesService.getScriptProperties().setProperty(PDF_CACHE_KEY, JSON.stringify(mapa));
    Logger.log('PDF parseado y cacheado: ' + Object.keys(mapa).length + ' códigos indexados.');
    return mapa;

  } catch (e) {
    Logger.log('ERROR en parsearPDF: ' + e.message);
    return {};
  } finally {
    if (copiaTempId) {
      try { DriveApp.getFileById(copiaTempId).setTrashed(true); } catch (ex) {}
    }
  }
}

/**
 * Extrae { codigo: precio } del texto OCR del PDF.
 *
 * Formato esperado: [CODIGO_DJI / CODIGO_DYE] Descripción ... USD X,XX
 * (el usuario confirmó: primero=DJI, segundo=DYE)
 *
 * Indexa ambos códigos con y sin guiones al mismo precio.
 */
function parsearTextoPDF(texto) {
  var mapa   = {};
  // Captura: grupo1=CODIGO_DJI, grupo2=CODIGO_DYE, grupo3=precio numérico
  var patron = /\[([^\]\/]+?)\s*\/\s*([^\]]+?)\][^\[]*?USD\s*([\d]+[,\.][\d]+)/gi;
  var match;

  while ((match = patron.exec(texto)) !== null) {
    var codigoDJI = match[1].trim();
    var codigoDYE = match[2].trim();
    // Normalizar separador decimal: coma → punto
    var precio    = parseFloat(match[3].replace(',', '.'));
    if (isNaN(precio)) continue;

    [codigoDJI, codigoDYE].forEach(function(cod) {
      if (!cod) return;
      mapa[cod] = precio;
      var norm = sinGuiones(cod);
      if (norm !== cod) mapa[norm] = precio;
      var base = normalizarCodigo(cod);
      if (base !== cod && !mapa.hasOwnProperty(base)) mapa[base] = precio;
    });
  }
  return mapa;
}

/**
 * Borra la caché del PDF y lo reparsea inmediatamente.
 * Ejecutar manualmente desde el editor si el PDF fue reemplazado.
 */
function refrescarCachePDF() {
  PropertiesService.getScriptProperties().deleteProperty(PDF_CACHE_KEY);
  var mapa = parsearPDF();
  Logger.log('Caché del PDF refrescada. Códigos indexados: ' + Object.keys(mapa).length);
}

// =============================================================================
// FUNCIÓN AUXILIAR: construir mapas de lookup
// =============================================================================

/**
 * Por cada hoja definida en HOJAS_ORIGEN:
 *   - Encuentra la columna "Código SKU" buscando el encabezado en la fila 1.
 *   - Lee todos los datos y crea un objeto { codigoSKU: precio }.
 *
 * @param {Spreadsheet} ssOrigen - Objeto Spreadsheet del archivo origen.
 * @returns {Object} Diccionario con clave = nombre de hoja, valor = mapa { sku: precio }.
 */
function construirMapasLookup(ssOrigen) {
  var resultado = {};

  for (var i = 0; i < HOJAS_ORIGEN.length; i++) {
    var def  = HOJAS_ORIGEN[i];
    var hoja = ssOrigen.getSheetByName(def.nombre);

    if (!hoja) {
      Logger.log('ADVERTENCIA: Hoja "' + def.nombre + '" no encontrada en el origen. Se omite.');
      resultado[def.nombre] = {};
      continue;
    }

    var datos = hoja.getDataRange().getValues();  // [[fila1col1, fila1col2, ...], ...]
    if (datos.length < 2) {
      Logger.log('ADVERTENCIA: Hoja "' + def.nombre + '" tiene menos de 2 filas. Se omite.');
      resultado[def.nombre] = {};
      continue;
    }

    // Buscar las columnas por encabezado (búsqueda dinámica)
    var encabezados = datos[0];
    var colSKU      = -1;
    var colInterno  = def.colInterno - 1;  // Posición fija según configuración (base 0)

    for (var c = 0; c < encabezados.length; c++) {
      if (String(encabezados[c]).trim() === ENCABEZADO_SKU) {
        colSKU = c;
        break;
      }
    }

    if (colSKU === -1) {
      Logger.log('ADVERTENCIA: Columna "' + ENCABEZADO_SKU + '" no encontrada en hoja "' + def.nombre + '". Se omite.');
      resultado[def.nombre] = { exacto: {}, base: {}, interno: {}, internoBase: {} };
      continue;
    }

    // Columna de precio (índice base 0)
    var colPrecio = def.colPrecio - 1;

    // Construir los cuatro mapas de búsqueda
    var mapaExacto      = {};  // SKU exacto         → precio  (búsqueda primaria)
    var mapaBase        = {};  // SKU sin sufijo      → precio  (búsqueda primaria fallback)
    var mapaInterno     = {};  // Código interno exacto          → precio  (búsqueda secundaria)
    var mapaInternoBase = {};  // Código interno sin sufijo       → precio  (búsqueda secundaria fallback)
    var mapaInternoNorm = {};  // Código interno sin guiones      → precio  (búsqueda secundaria fallback)

    for (var r = 1; r < datos.length; r++) {
      var precio = datos[r][colPrecio];

      // Mapa SKU (primario)
      var sku = String(datos[r][colSKU]).trim();
      if (sku !== "" && sku !== "0") {
        mapaExacto[sku] = precio;
        var skuBase = normalizarCodigo(sku);
        if (skuBase !== sku && !mapaBase.hasOwnProperty(skuBase)) {
          mapaBase[skuBase] = precio;
        }
      }

      // Mapa Código interno (secundario)
      var interno = String(datos[r][colInterno]).trim();
      if (interno !== "" && interno !== "0") {
        mapaInterno[interno] = precio;
        var internoBase = normalizarCodigo(interno);
        if (internoBase !== interno && !mapaInternoBase.hasOwnProperty(internoBase)) {
          mapaInternoBase[internoBase] = precio;
        }
        // Índice sin guiones: DJI-R891 y DJIR891 resuelven al mismo precio
        var internoNorm = sinGuiones(interno);
        if (internoNorm !== interno && !mapaInternoNorm.hasOwnProperty(internoNorm)) {
          mapaInternoNorm[internoNorm] = precio;
        }
      }
    }

    resultado[def.nombre] = { exacto: mapaExacto, base: mapaBase, interno: mapaInterno, internoBase: mapaInternoBase, internoNorm: mapaInternoNorm };
    Logger.log('Hoja "' + def.nombre + '": ' + Object.keys(mapaExacto).length + ' SKUs | ' + Object.keys(mapaInterno).length + ' códigos internos cargados.');
  }

  return resultado;
}

// =============================================================================
// DIAGNÓSTICO — Ejecutar una vez para identificar la cuenta y verificar accesos
// =============================================================================

/**
 * Muestra en el log la cuenta que ejecuta el script y si tiene acceso
 * a los dos archivos de Sheets configurados.
 */
function diagnostico() {
  try { Logger.log("Cuenta activa: "    + Session.getActiveUser().getEmail()); }
  catch (e) { Logger.log("Cuenta activa: (no disponible) — " + e.message); }

  try { Logger.log("Cuenta efectiva: "  + Session.getEffectiveUser().getEmail()); }
  catch (e) { Logger.log("Cuenta efectiva: (no disponible) — " + e.message); }

  try {
    var o = SpreadsheetApp.openById(ORIGEN_ID);
    Logger.log("ORIGEN OK: " + o.getName());
  } catch (e) {
    Logger.log("ORIGEN ERROR: " + e.message);
  }

  try {
    var d = SpreadsheetApp.openById(DESTINO_ID);
    Logger.log("DESTINO OK: " + d.getName());
  } catch (e) {
    Logger.log("DESTINO ERROR: " + e.message);
  }
}

// =============================================================================
// GESTIÓN DE TRIGGERS
// =============================================================================

/**
 * Instala un trigger de tiempo que ejecuta actualizarPrecios() cada 1 hora.
 * EJECUTAR SOLO UNA VEZ manualmente desde el editor de Apps Script.
 * Si ya existe un trigger previo de esta función, lo elimina antes de crear uno nuevo.
 */
function instalarTrigger() {
  eliminarTriggers();  // Limpia duplicados

  ScriptApp.newTrigger("actualizarPrecios")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("Trigger instalado: actualizarPrecios se ejecutará cada 1 hora.");
}

/**
 * Elimina todos los triggers asociados a actualizarPrecios().
 * Útil para desinstalar o cambiar la frecuencia.
 */
function eliminarTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "actualizarPrecios") {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log("Trigger eliminado: " + triggers[i].getUniqueId());
    }
  }
}
