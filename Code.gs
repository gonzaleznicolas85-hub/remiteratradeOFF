/***** ==========================================================
 *  REMITERA – WebApp + PDFs desde Sheets (Google Apps Script)
 *  - Generación de PDFs "desde Sheets" con diseño y logo
 *  - Guarda PDF en Drive y linkea PDF_URL en REMITOS
 *  - API compatible con Netlify (soloPdf, pdfBase64, etc.)
 *  - Stock en tiempo real, alta de productos e histórico (GET)
 * =========================================================== */

/* =========================
 * CONFIG BÁSICA
 * ========================= */

// 🔴 Este script corre SIEMPRE bound a la planilla que lo contiene
// (la del clon OFF). Antes había acá un par USE_SHEET_ID/SHEET_ID que
// quedó colgado apuntando a la planilla del proyecto ORIGINAL: estaba
// inerte porque el flag venía en false, pero si alguien lo ponía en true
// el OFF empezaba a escribir sobre los datos del original. Se eliminó
// para que no exista esa posibilidad.

const HOJA_STOCK   = 'STOCK';
const HOJA_REMITOS = 'REMITOS';
const HOJA_DETALLE = 'DETALLEREMITOS';

const TOKEN_ESPERADO = 'REMITOSDADIGITAL-OFF';
const VALIDAR_SKU_EN_STOCK = true;

const FOLDER_ID_PDF = '1H7YJqkC_nds-U7l5T-DIPA4bLmpnDPEz';

// Carpeta de Drive donde se guardan las fotos de producto.
// Por defecto usa la misma carpeta que los PDFs; cambiala por el ID de una carpeta
// dedicada si preferís mantenerlas separadas.
const FOLDER_ID_IMAGENES = '1H7YJqkC_nds-U7l5T-DIPA4bLmpnDPEz';

// ID de la presentación de Slides "plantilla maestra" del remito (imagen +
// cuadros de texto con placeholders {{...}}). Se completa corriendo UNA VEZ
// crearPlantillaRemitoSlides_() desde el editor y pegando acá el ID que
// devuelve. Ver knowledge/ o AGENTS.md de este repo si hace falta rehacerla.
const TEMPLATE_SLIDE_ID = '1JtTPvAWSvgR61WMxcu__oIQdnkCjPhqc-JiBvVlH9ug';

// Pedidos de materiales: el Form de esta área guarda las respuestas en OTRA
// planilla (no en la de STOCK/REMITOS), hay que abrirla aparte por ID.
const PEDIDOS_SHEET_ID = '1W8xls7GeUj0SJSiTooigl6a3XzH9dzkIG5lh5hHuByI';
const PEDIDOS_TAB_NAME = 'Pedidos';

/* =========================
 * BRANDING / DISEÑO PDF
 * ========================= */
const BRAND = {
  TITLE: 'Remito · DEMAND ACCELERATION',
  LOGO_FILE_ID: '',       // opcional
  LOGO_WIDTH_PX: 0,
  LOGO_HEIGHT_PX: 0,
  LOGO_ALIGN: 'CENTER',
  COLOR_PRIMARY: '#000000',
  COLOR_ACCENT:  '#ef4444',
  TABLE_HEAD_BG: '#b5bdb6',
  TABLE_ZEBRA_BG:'#cfa500',
  FONT_FAMILY:   'Roboto'
};

/* ==========================================================
 * HELPERS GENERALES
 * ========================================================== */
function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Crea hojas SOLO si no existen. No pisa nada si ya están.
 */
function ensureSheets_(ss) {
  const defs = [
    { name: HOJA_STOCK,   headers: ['SKU','Descripcion','Tipo','Marca','StockInicial','Entregado','StockActual'] },
    { name: HOJA_REMITOS, headers: ['NroRemito','Fecha','PuntoVenta','Usuario','Obs','Timestamp','PDF_URL'] },
    { name: HOJA_DETALLE, headers: ['NroRemito','SKU','Descripción','Tipo','Marca','Cantidad'] },
  ];

  defs.forEach(def => {
    let sh = ss.getSheetByName(def.name);
    if (!sh) {
      sh = ss.insertSheet(def.name);
      sh.appendRow(def.headers);
    }
  });
}

function jsonOut(obj, status) {
  // NOTA: ContentService.TextOutput no tiene setHeader() ni setResponseCode()
  // (esos métodos no existen en la API de Apps Script) — llamarlos rompía
  // TODAS las respuestas con un TypeError. Un Web App público de Apps Script
  // ya responde con headers permisivos por defecto, así que no hacen falta.
  // El código de error/éxito va en el body (obj.ok), el status HTTP siempre es 200.
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Lee el body del POST (JSON o parámetros normales)
 */
function readPostBody_(e) {
  try {
    const ct = e.postData && e.postData.type ? String(e.postData.type).toLowerCase() : '';
    if (ct.indexOf('application/json') !== -1) {
      return JSON.parse(e.postData.contents || '{}');
    }

    // OJO: acá NO va una whitelist de campos. Antes se listaba parámetro por
    // parámetro y cada campo nuevo que se agregaba al front se perdía en
    // silencio si alguien se olvidaba de sumarlo a la lista (ya pasó con
    // requestId: llegaba del navegador, no se leía acá, y la deduplicación
    // de remitos quedaba muerta sin ningún error visible). Se copian TODOS
    // los parámetros y solo se les da tratamiento especial a los que son
    // JSON serializado o booleanos.
    const p = { header: {}, lines: [] };
    Object.keys(e.parameter || {}).forEach(function (k) { p[k] = e.parameter[k]; });

    ['header', 'lines', 'producto'].forEach(function (k) {
      if (!e.parameter[k]) return;
      try {
        p[k] = JSON.parse(e.parameter[k]);
      } catch (errJson) {
        // JSON roto: se vuelve al valor por defecto en vez de dejar un
        // string donde el resto del código espera un objeto/array.
        p[k] = (k === 'lines') ? [] : {};
        Logger.log('readPostBody_: no se pudo parsear ' + k + ': ' + errJson);
      }
    });
    p.token = e.parameter.token || '';
    p.soloPdf = (String(e.parameter.soloPdf || '') === 'true');
    if (!Array.isArray(p.lines)) p.lines = [];
    if (!p.header || typeof p.header !== 'object') p.header = {};
    return p;
  } catch (err) {
    throw new Error('No se pudo leer el cuerpo del POST: ' + err);
  }
}

/**
 * Convierte un valor de celda (Date real, número de serie de Sheets, o texto)
 * a un objeto Date de JS. Devuelve null si no se puede interpretar.
 */
function toJsDate_(val) {
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  if (typeof val === 'number' && isFinite(val)) {
    // Sheets/Excel: día 0 = 30/12/1899. 25569 = días entre esa fecha y 01/01/1970 (epoch Unix).
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }
  if (typeof val === 'string' && val.trim()) {
    const parsed = new Date(val);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function skuExisteEnStock_(sheetStock, sku) {
  const last = sheetStock.getLastRow();
  if (last < 2) return false;
  const vals = sheetStock.getRange(2,1,last-1,1).getValues();
  const skuBuscado = String(sku).trim();
  return vals.some(r => String(r[0]).trim() === skuBuscado);
}

/**
 * Próximo SKU disponible: toma el máximo código numérico ya usado en STOCK y le suma 1.
 * Ignora SKUs no numéricos (no los rompe, simplemente no los cuenta para el máximo).
 */
function nextStockSku_(shStock) {
  const last = shStock.getLastRow();
  if (last < 2) return '10000';
  const vals = shStock.getRange(2, 1, last - 1, 1).getValues();
  let max = 0;
  vals.forEach(r => {
    const n = parseInt(String(r[0]).trim(), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return String(max + 1);
}

function nextRemitoNumber_(sheetRemitos) {
  const last = sheetRemitos.getLastRow();
  if (last < 2) return 'R0001';
  const prev = String(sheetRemitos.getRange(last,1).getValue() || '').trim();
  const num  = Number((prev.match(/R(\d+)/) || [0,0])[1]) || 0;
  return 'R' + ('0000' + (num + 1)).slice(-4);
}

/* ==========================================================
 * STOCK helpers (columnas ubicadas por nombre, no por posición fija,
 * porque "StockInicial" trae fecha en el encabezado, ej: "StockInicial 29/9")
 * ========================================================== */
function stockColIndex_(headers, patterns) {
  // patterns: array de substrings (lowercase) a buscar en el header
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').toLowerCase();
    if (patterns.some(p => h.indexOf(p) !== -1)) return i; // 0-based
  }
  return -1;
}

function getStockMap_(sheetStock) {
  const lastRow = sheetStock.getLastRow();
  const lastCol = sheetStock.getLastColumn();
  const headers = sheetStock.getRange(1,1,1,lastCol).getValues()[0];

  return {
    headers: headers,
    idxSku:          stockColIndex_(headers, ['sku']),
    idxDescripcion:  stockColIndex_(headers, ['descripcion','descripción']),
    idxTipo:         stockColIndex_(headers, ['tipo']),
    idxMarca:        stockColIndex_(headers, ['marca']),
    idxStockInicial: stockColIndex_(headers, ['stockinicial']),
    idxEntregado:    stockColIndex_(headers, ['entregado']),
    idxStockActual:  stockColIndex_(headers, ['stockactual']),
    idxImagen:       stockColIndex_(headers, ['imagen']),
    idxMedidas:      stockColIndex_(headers, ['medida']),
    lastRow: lastRow,
    lastCol: lastCol
  };
}

/**
 * Asegura que exista la columna ImagenURL en STOCK y devuelve su índice (1-based).
 * No pisa ninguna columna existente: si no la encuentra, la agrega al final.
 */
function ensureImagenUrlHeader_(shStock) {
  const lastCol = Math.max(1, shStock.getLastColumn());
  const header  = shStock.getRange(1,1,1,lastCol).getValues()[0];

  let colIndex = stockColIndex_(header, ['imagen']); // 0-based
  if (colIndex === -1) {
    colIndex = header.length;
    shStock.getRange(1, colIndex + 1).setValue('ImagenURL');
  }
  return colIndex + 1; // 1-based
}

/**
 * Asegura que exista la columna Medidas en STOCK y devuelve su índice (1-based).
 * Mismo criterio que ensureImagenUrlHeader_: si no existe, la agrega al final.
 */
function ensureMedidasHeader_(shStock) {
  const lastCol = Math.max(1, shStock.getLastColumn());
  const header  = shStock.getRange(1,1,1,lastCol).getValues()[0];

  let colIndex = stockColIndex_(header, ['medida']); // 0-based
  if (colIndex === -1) {
    colIndex = header.length;
    shStock.getRange(1, colIndex + 1).setValue('Medidas');
  }
  return colIndex + 1; // 1-based
}

/**
 * Sube una foto (base64) a Drive y devuelve una URL apta para <img src>.
 */
function saveImagenToDrive_(base64Str, name) {
  const folder = DriveApp.getFolderById(FOLDER_ID_IMAGENES);
  const clean  = String(base64Str || '').split(',').pop();
  const blob   = Utilities.newBlob(
    Utilities.base64Decode(clean),
    'image/jpeg',
    name || ('producto_' + Date.now() + '.jpg')
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // "uc?export=view" es poco confiable para insertar directo en <img> (a veces Drive
  // devuelve una página intermedia en vez de la imagen). El endpoint de thumbnail sí
  // está pensado para esto: responde la imagen con CORS abierto.
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

/**
 * Busca la fila (1-based) de un SKU dentro de STOCK. Devuelve -1 si no existe.
 */
function findStockRowBySku_(shStock, sku) {
  const last = shStock.getLastRow();
  if (last < 2) return -1;
  const vals = shStock.getRange(2,1,last-1,1).getValues();
  const buscado = String(sku).trim();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === buscado) return i + 2;
  }
  return -1;
}

/**
 * Descuenta stock cuando se registra un remito: por cada línea válida
 * ([nro, sku, descripcion, tipo, marca, cantidad], el mismo formato que
 * arma doPost antes de escribir DETALLEREMITOS), suma la cantidad
 * entregada a la columna "Entregado" y la resta de "StockActual" del SKU
 * correspondiente en STOCK. Si el mismo SKU aparece en varias líneas del
 * remito, se agrupan para hacer una sola lectura/escritura por SKU.
 * No frena el registro del remito si un SKU puntual no se encuentra en
 * STOCK (ya se validó antes con VALIDAR_SKU_EN_STOCK, pero por las dudas
 * se ignora silenciosamente en vez de tirar error acá).
 */
function descontarStockPorRemito_(shStock, shRemitos, vrLineas, nroRemito) {
  const map = getStockMap_(shStock);
  if (map.idxEntregado < 0 && map.idxStockActual < 0) return; // hoja sin esas columnas

  const colAplicado = shRemitos ? ensureStockAplicadoHeader_(shRemitos) : -1;
  const rowRemito   = shRemitos ? findRemitoRowByNro_(shRemitos, nroRemito) : -1;

  /* ---------------------------------------------------------------
   * GUARD DE IDEMPOTENCIA (dos capas)
   *
   *  1) Marca DURABLE en REMITOS!StockAplicado. Es la que manda: vive en
   *     la planilla, no expira y se puede auditar a ojo. Si el remito ya
   *     tiene fecha ahí, su stock NO se vuelve a descontar nunca.
   *  2) Marca "en vuelo" en CacheService, para el caso en que dos
   *     ejecuciones arranquen casi simultáneas y ninguna haya llegado
   *     todavía a escribir la marca durable.
   *
   * El chequeo+marcado va bajo lock: sin eso, dos ejecuciones realmente
   * concurrentes podrían leer ambas marcas vacías ANTES de que ninguna
   * las haya escrito, y aplicar el descuento las dos (check-then-act).
   * --------------------------------------------------------------- */
  if (nroRemito) {
    const lock = LockService.getScriptLock();
    const gotLock = lock.tryLock(30000);
    let yaAplicado = false;
    try {
      if (rowRemito > 0 && colAplicado > 0) {
        const marca = shRemitos.getRange(rowRemito, colAplicado).getValue();
        if (String(marca || '').trim()) yaAplicado = true;
      }
      if (!yaAplicado) {
        const cache = CacheService.getScriptCache();
        const cacheKey = 'stockAplicado_' + nroRemito;
        if (cache.get(cacheKey)) yaAplicado = true;
        else cache.put(cacheKey, '1', 21600); // 6 h (máximo de CacheService)
      }
    } finally {
      if (gotLock) lock.releaseLock();
    }
    if (yaAplicado) {
      Logger.log('descontarStockPorRemito_: stock ya aplicado antes para ' + nroRemito + ', se omite.');
      return;
    }
  }

  const cantidadPorSku = {};
  for (let i = 0; i < vrLineas.length; i++) {
    const sku = String(vrLineas[i][1] || '').trim();
    const cant = Number(vrLineas[i][5] || 0);
    if (!sku || !(cant > 0)) continue;
    cantidadPorSku[sku] = (cantidadPorSku[sku] || 0) + cant;
  }

  const colInicial   = map.idxStockInicial >= 0 ? map.idxStockInicial + 1 : -1;
  const colEntregado = map.idxEntregado    >= 0 ? map.idxEntregado    + 1 : -1;
  const colActual    = map.idxStockActual  >= 0 ? map.idxStockActual  + 1 : -1;

  // Se LEE un rango que cubre StockInicial..Entregado..StockActual de una
  // sola vez, y se ESCRIBE solo el tramo Entregado..StockActual: así no se
  // pisa StockInicial (que es una carga manual) y, sobre todo, no se lee
  // ninguna celda después de haber escrito otra de la misma fila.
  const colsLeer = [colInicial, colEntregado, colActual].filter(function (c) { return c > 0; });
  const colsEscribir = [colEntregado, colActual].filter(function (c) { return c > 0; });
  const readStart  = Math.min.apply(null, colsLeer);
  const readEnd    = Math.max.apply(null, colsLeer);
  const writeStart = Math.min.apply(null, colsEscribir);
  const writeEnd   = Math.max.apply(null, colsEscribir);

  Object.keys(cantidadPorSku).forEach(function (sku) {
    const row = findStockRowBySku_(shStock, sku);
    if (row === -1) return; // SKU no encontrado, no hay nada que descontar

    const cant = cantidadPorSku[sku];
    const leidos = shStock.getRange(row, readStart, 1, readEnd - readStart + 1).getValues()[0];
    const val = function (col) { return Number(leidos[col - readStart] || 0); };

    let entregadoNuevo = null;
    let actualNuevo = null;

    if (colEntregado > 0) entregadoNuevo = val(colEntregado) + cant;

    if (colActual > 0) {
      if (colInicial > 0 && entregadoNuevo !== null) {
        // StockActual se recalcula SIEMPRE de forma ABSOLUTA a partir de
        // StockInicial - Entregado, nunca como "el StockActual que había
        // menos la cantidad". Esto lo vuelve idempotente por construcción:
        // aplicar el mismo remito dos veces da el mismo número, y además
        // es inmune a que la columna tenga una fórmula que se recalcule
        // sola al escribir Entregado (que es lo que hacía que el descuento
        // impactara dos veces).
        actualNuevo = val(colInicial) - entregadoNuevo;
      } else {
        actualNuevo = val(colActual) - cant;
      }
    }

    const escribir = [];
    for (let c = writeStart; c <= writeEnd; c++) {
      if (c === colEntregado)    escribir.push(entregadoNuevo);
      else if (c === colActual)  escribir.push(actualNuevo);
      else                       escribir.push(leidos[c - readStart]); // columna intermedia: se deja igual
    }
    shStock.getRange(row, writeStart, 1, escribir.length).setValues([escribir]);
  });

  // Recién acá se sella la marca durable: si algo explotó a mitad del
  // descuento, el remito queda SIN marcar y se puede reprocesar a mano
  // (la marca de cache lo frena por 6 h, no para siempre).
  if (rowRemito > 0 && colAplicado > 0) {
    shRemitos.getRange(rowRemito, colAplicado).setValue(new Date());
  }
}

/* ==========================================================
 * DRIVE / PDF helpers
 * ========================================================== */
function savePdfToDrive_(base64Str, name) {
  const folder = DriveApp.getFolderById(FOLDER_ID_PDF);
  const clean  = String(base64Str || '').split(',').pop();
  const blob   = Utilities.newBlob(
    Utilities.base64Decode(clean),
    'application/pdf',
    name || ('remito_' + Date.now() + '.pdf')
  );
  const file   = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

/**
 * Asegura que exista la cabecera RequestId y devuelve columna (1-based).
 * Se usa para deduplicar reintentos de un mismo remito (ver handleCrearRemito_).
 */
function ensureRequestIdHeader_(shRemitos) {
  const lastCol = Math.max(1, shRemitos.getLastColumn());
  const header  = shRemitos.getRange(1,1,1,lastCol).getValues()[0];

  let colIndex = header.indexOf('RequestId'); // 0-based
  if (colIndex === -1) {
    colIndex = header.length;
    shRemitos.getRange(1, colIndex + 1).setValue('RequestId');
  }
  return colIndex + 1; // 1-based
}

/**
 * Busca un remito ya registrado con este requestId (columna RequestId de
 * REMITOS). Devuelve el NroRemito si lo encuentra, o '' si no existe todavía.
 * Esto evita duplicar el remito cuando el cliente reintenta el envío
 * (por ejemplo: la red se cae justo cuando Apps Script ya había terminado
 * de escribir, y el navegador vuelve a mandar el mismo remito desde la cola
 * offline). Ver knowledge/ si hace falta documentar esto más a fondo.
 */
function buscarNroPorRequestId_(shRemitos, colRequestId, requestId) {
  const last = shRemitos.getLastRow();
  if (last < 2) return '';
  const vals = shRemitos.getRange(2, 1, last - 1, colRequestId).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][colRequestId - 1] || '').trim() === requestId) {
      return String(vals[i][0]).trim();
    }
  }
  return '';
}

/**
 * Asegura que exista la cabecera StockAplicado en REMITOS y devuelve su
 * columna (1-based). Es la marca DURABLE de idempotencia del descuento de
 * stock: una vez que un remito tiene fecha acá, su stock no se vuelve a
 * descontar nunca, aunque doPost se ejecute de nuevo para el mismo Nº.
 * A diferencia de CacheService (que expira a las 6 h y se puede desalojar),
 * esta marca vive en la planilla y es auditable a ojo.
 */
function ensureStockAplicadoHeader_(shRemitos) {
  const lastCol = Math.max(1, shRemitos.getLastColumn());
  const header  = shRemitos.getRange(1,1,1,lastCol).getValues()[0];

  let colIndex = header.indexOf('StockAplicado'); // 0-based
  if (colIndex === -1) {
    colIndex = header.length;
    shRemitos.getRange(1, colIndex + 1).setValue('StockAplicado');
  }
  return colIndex + 1; // 1-based
}

/**
 * Fila (1-based) de un remito en REMITOS buscando por NroRemito (columna A).
 * Devuelve -1 si no existe.
 */
function findRemitoRowByNro_(shRemitos, nro) {
  const buscado = String(nro || '').trim();
  if (!buscado) return -1;
  const last = shRemitos.getLastRow();
  if (last < 2) return -1;
  const vals = shRemitos.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === buscado) return i + 2;
  }
  return -1;
}

/**
 * Asegura que exista la cabecera PDF_URL y devuelve columna (1-based)
 */
function ensurePdfUrlHeader_(shRemitos) {
  const lastCol = Math.max(1, shRemitos.getLastColumn());
  const header  = shRemitos.getRange(1,1,1,lastCol).getValues()[0];

  let colIndex = header.indexOf('PDF_URL'); // 0-based
  if (colIndex === -1) {
    colIndex = header.length; // siguiente col libre
    shRemitos.getRange(1, colIndex + 1).setValue('PDF_URL');
  }
  return colIndex + 1; // 1-based
}

/**
 * Reemplaza el PDF de un remito YA CREADO (nuevo modelo con plantilla,
 * armado en el navegador una vez que se conoce el Nº de remito real).
 * Busca la fila por NroRemito (columna A) y pisa PDF_URL con el nuevo archivo.
 */
function handleAttachPdf_(shRemitos, data) {
  const nro = String(data.nroRemito || '').trim();
  const pdfBase64 = data.pdfBase64 || '';
  const pdfName = data.pdfName || ('Remito_' + nro + '.pdf');

  if (!nro) return jsonOut({ ok:false, error:'Falta nroRemito' }, 400);
  if (!pdfBase64) return jsonOut({ ok:false, error:'Falta pdfBase64' }, 400);

  const row = findRemitoRowByNro_(shRemitos, nro);
  if (row === -1) return jsonOut({ ok:false, error:'Remito no encontrado: ' + nro }, 404);

  const url = savePdfToDrive_(pdfBase64, pdfName);
  const colPdf = ensurePdfUrlHeader_(shRemitos);
  shRemitos.getRange(row, colPdf).setValue(url);

  return jsonOut({ ok:true, nroRemito: nro, pdfUrl: url });
}

/* ==========================================================
 * PDF con plantilla de Slides (imagen de remito real + campos)
 * ========================================================== */

/**
 * SETUP — correr UNA SOLA VEZ a mano desde el editor de Apps Script.
 * Crea la presentación de Slides "plantilla maestra": trae la imagen del
 * remito ya publicada en el sitio, la pone de fondo a página completa, y
 * agrega los cuadros de texto con los placeholders {{...}} en las mismas
 * coordenadas que usaba el modelo jsPDF anterior (1024x1536 pt = 1:1 con
 * los px de la imagen). Al terminar, loguea el ID: copiarlo y pegarlo en
 * la constante TEMPLATE_SLIDE_ID de arriba.
 */
function crearPlantillaRemitoSlides_() {
  // El servicio básico de Slides no permite crear una presentación con
  // tamaño de página custom (no existe Presentation.setPageSize). Por eso
  // la plantilla maestra se crea A MANO una sola vez: Diapositiva nueva en
  // slides.new, Archivo > Configuración de página > Personalizado > Puntos
  // > 1024 x 1536, y se pega acá el ID. A partir de ahí las coordenadas
  // son 1:1 con los px de la imagen (sin escalar). Comparte la misma
  // plantilla maestra que remiteraTrade (misma imagen de referencia).
  const MASTER_ID = '1JtTPvAWSvgR61WMxcu__oIQdnkCjPhqc-JiBvVlH9ug';
  const IMG_URL = 'https://remiteraoff.netlify.app/remito-template.jpg';
  const blob = UrlFetchApp.fetch(IMG_URL).getBlob();

  const pres = SlidesApp.openById(MASTER_ID);
  const slide = pres.getSlides()[0];
  slide.getShapes().forEach(function (sh) {
    try { sh.remove(); } catch (eRm) { /* placeholders por defecto */ }
  });

  slide.insertImage(blob, 0, 0, 1024, 1536);

  function addBox(text, left, top, width, height, size, align) {
    const box = slide.insertTextBox(text, left, top, width, height);
    const tr = box.getText();
    tr.getTextStyle().setFontFamily('Arial').setFontSize(size).setForegroundColor('#141414');
    tr.getParagraphStyle().setParagraphAlignment(
      align === 'CENTER' ? SlidesApp.ParagraphAlignment.CENTER : SlidesApp.ParagraphAlignment.START
    );
    box.getFill().setTransparent();
    box.getBorder().setTransparent();
    return box;
  }

  addBox('{{NUMERO}}', 750, 62, 180, 24, 15, 'LEFT');
  addBox('{{DD}}', 603, 160, 60, 20, 11, 'CENTER');
  addBox('{{MM}}', 661, 160, 60, 20, 11, 'CENTER');
  addBox('{{YYYY}}', 726, 160, 80, 20, 11, 'CENTER');
  addBox('{{CLIENTE_L1}}', 150, 326, 470, 20, 12, 'LEFT');
  addBox('{{CLIENTE_L2}}', 145, 362, 480, 18, 11, 'LEFT');
  addBox('{{CLIENTE_L3}}', 145, 397, 480, 18, 11, 'LEFT');
  addBox('{{OBS}}', 150, 1016, 740, 50, 11, 'LEFT');

  const ROW_TOPS = [575, 606, 633, 660, 687, 714, 741, 768, 795, 822, 849, 876, 901, 928, 955, 982];
  for (let i = 0; i < ROW_TOPS.length; i++) {
    const y = ROW_TOPS[i] + 4;
    const n = i + 1;
    addBox(String(n), 50, y, 30, 18, 10, 'CENTER');
    addBox('{{SKU_' + n + '}}', 100, y, 110, 18, 10, 'LEFT');
    addBox('{{DESC_' + n + '}}', 225, y, 490, 18, 10, 'LEFT');
    addBox('{{CANT_' + n + '}}', 748, y, 50, 18, 10, 'CENTER');
  }

  Logger.log('TEMPLATE_SLIDE_ID = ' + pres.getId());
  return pres.getId();
}

/**
 * Copia la plantilla, reemplaza los placeholders con los datos del remito,
 * exporta a PDF, guarda la copia oficial en Drive y borra la copia de
 * trabajo de Slides. Devuelve { pdfUrl, pdfBase64 }.
 */
function buildRemitoPdfFromTemplate_(nroRemito, header, lines) {
  const copyFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID).makeCopy('Remito_' + nroRemito + '_tmp');
  const pres = SlidesApp.openById(copyFile.getId());
  const slide = pres.getSlides()[0];

  const partesFecha = String(header.fecha || '').split('-'); // yyyy-mm-dd
  const yyyy = partesFecha[0] || '';
  const mm = partesFecha[1] || '';
  const dd = partesFecha[2] || '';
  const partesCliente = String(header.punto_venta || '').split(' — ');

  slide.replaceAllText('{{NUMERO}}', String(nroRemito));
  slide.replaceAllText('{{DD}}', dd);
  slide.replaceAllText('{{MM}}', mm);
  slide.replaceAllText('{{YYYY}}', yyyy);
  slide.replaceAllText('{{CLIENTE_L1}}', partesCliente[0] || '');
  slide.replaceAllText('{{CLIENTE_L2}}', partesCliente[1] || '');
  slide.replaceAllText('{{CLIENTE_L3}}', partesCliente[2] || '');
  slide.replaceAllText('{{OBS}}', String(header.obs || ''));

  for (let i = 0; i < 16; i++) {
    const n = i + 1;
    const l = lines[i];
    slide.replaceAllText('{{SKU_' + n + '}}', l ? String(l.sku || '') : '');
    slide.replaceAllText('{{DESC_' + n + '}}', l ? String(l.descripcion || '') : '');
    slide.replaceAllText('{{CANT_' + n + '}}', l ? String(l.cantidad || '') : '');
  }

  pres.saveAndClose();

  const pdfBlob = DriveApp.getFileById(copyFile.getId()).getAs('application/pdf');
  const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

  const folder = DriveApp.getFolderById(FOLDER_ID_PDF);
  const finalFile = folder.createFile(pdfBlob).setName('Remito_' + nroRemito + '.pdf');
  finalFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  DriveApp.getFileById(copyFile.getId()).setTrashed(true);

  return { pdfUrl: finalFile.getUrl(), pdfBase64: pdfBase64 };
}

/**
 * Handler de la acción 'generarPdfRemito': arma el PDF desde la plantilla
 * de Slides, lo linkea en la fila del remito (PDF_URL) y devuelve el
 * base64 para que el navegador lo descargue al toque.
 */
function handleGenerarPdfRemito_(shRemitos, data) {
  const nro = String(data.nroRemito || '').trim();
  if (!nro) return jsonOut({ ok:false, error:'Falta nroRemito' }, 400);

  const header = data.header || {};
  const lines = Array.isArray(data.lines) ? data.lines : [];

  let result;
  try {
    result = buildRemitoPdfFromTemplate_(nro, header, lines);
  } catch (eGen) {
    return jsonOut({ ok:false, error:'Error generando PDF: ' + eGen }, 500);
  }

  const row = findRemitoRowByNro_(shRemitos, nro);
  if (row !== -1) {
    const colPdf = ensurePdfUrlHeader_(shRemitos);
    shRemitos.getRange(row, colPdf).setValue(result.pdfUrl);
  }

  return jsonOut({ ok:true, nroRemito: nro, pdfUrl: result.pdfUrl, pdfBase64: result.pdfBase64 });
}

/**
 * Busca la última fila de un remito por Fecha + PuntoVenta + Usuario
 * usado en modo soloPdf
 */
function findLastRemitoRowByHeader_(shRemitos, fecha, punto_venta, usuario) {
  const last = shRemitos.getLastRow();
  if (last < 2) return 0;
  const vals = shRemitos.getRange(2,1,last-1,6).getValues(); // A..F

  const F = String(fecha||'').trim();
  const P = String(punto_venta||'').trim();
  const U = String(usuario||'').trim();

  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    if (String(r[1]).trim() === F &&
        String(r[2]).trim() === P &&
        String(r[3]).trim() === U) {
      return i + 2; // +2 porque vals arranca en fila 2
    }
  }
  return 0;
}

/* ==========================================================
 * GENERADOR DE PDF (desde Sheets)
 * ========================================================== */
function colorHex_(hex) { return hex || '#000000'; }

function generarPDFDesdeRemito(nroRemito) {
  const ss  = getSS_();
  const shR = ss.getSheetByName(HOJA_REMITOS);
  const shD = ss.getSheetByName(HOJA_DETALLE);

  if (!shR || !shD) {
    throw new Error('No se encontraron las hojas REMITOS o DETALLEREMITOS');
  }

  const dataR = shR.getDataRange().getValues();
  const nroBuscado = String(nroRemito).trim();

  // Buscar el remito SOLO desde la fila 2 (ignora encabezado)
  let idx = -1;
  for (let i = 1; i < dataR.length; i++) {
    if (String(dataR[i][0]).trim() === nroBuscado) {
      idx = i;
      break;
    }
  }
  if (idx === -1) {
    throw new Error('Remito no encontrado: ' + nroRemito);
  }

  const H = dataR[idx]; // [NroRemito, Fecha, PuntoVenta, Usuario, Obs, Timestamp, PDF_URL]

  // ----- FECHA FORMATEADA (dd/MM/yyyy HH:mm) -----
  let fechaStr = '';
  const fechaObj = toJsDate_(H[1]);
  if (fechaObj) {
    fechaStr = Utilities.formatDate(
      fechaObj,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy HH:mm'
    );
  }

  // Detalle del remito
  const dataD  = shD.getDataRange().getValues();
  const detVals = [];
  for (let i = 1; i < dataD.length; i++) {
    if (String(dataD[i][0]).trim() === nroBuscado) {
      detVals.push(dataD[i]); // [NroRemito, SKU, Descripción, Tipo, Marca, Cantidad]
    }
  }

  // ==========================
  // Crear documento sobrio
  // ==========================
  const doc = DocumentApp.create('Remito ' + nroBuscado);
  const body = doc.getBody();

  body.setAttributes({
    FONT_FAMILY: BRAND.FONT_FAMILY || 'Arial',
    FONT_SIZE: 11
  });

  // Título
  body.appendParagraph('Remito ' + nroBuscado)
      .setBold(true)
      .setFontSize(16)
      .setForegroundColor(colorHex_(BRAND.COLOR_PRIMARY));

  body.appendParagraph(''); // espacio

  // Datos principales (ya con fecha formateada)
  body.appendParagraph('Fecha: '          + (fechaStr || ''));
  body.appendParagraph('Punto de venta: ' + String(H[2] || ''));
  body.appendParagraph('Usuario: '        + String(H[3] || ''));
  body.appendParagraph('Observaciones: '  + String(H[4] || ''));

  body.appendParagraph(''); // espacio
  body.appendParagraph('Detalle:')
      .setBold(true)
      .setSpacingBefore(10);

  body.appendParagraph(''); // espacio antes de la tabla

  // ==========================
  // Tabla de detalle
  // ==========================
  const tableData = [['SKU', 'Cant', 'Descripción']];

  detVals.forEach(d => {
    const sku   = d[1];
    const cant  = d[5];
    const desc  = d[2];
    tableData.push([sku, cant, desc]);
  });

  const tbl = body.appendTable(tableData);
  tbl.setBorderWidth(0.5);

  // Encabezado
  const headRow = tbl.getRow(0);
  for (let c = 0; c < headRow.getNumCells(); c++) {
    headRow.getCell(c)
      .setBold(true)
      .setBackgroundColor('#f0f0f0');
  }

  // Alinear la columna "Cant" (columna 1) al centro
  for (let r = 1; r < tbl.getNumRows(); r++) {
    const cell = tbl.getRow(r).getCell(1); // columna Cant
    for (let p = 0; p < cell.getNumChildren(); p++) {
      const child = cell.getChild(p);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        child.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    }
  }

  // Total de ítems
  const total = detVals.reduce((acc, d) => acc + Number(d[5] || 0), 0);
  body.appendParagraph('')
      .setSpacingBefore(12);
  body.appendParagraph('Total de ítems: ' + total)
      .setBold(true)
      .setForegroundColor(colorHex_(BRAND.COLOR_ACCENT));

  doc.saveAndClose();

  // ==========================
  // Exportar a PDF + guardar URL
  // ==========================
  const pdfBlob = doc.getAs('application/pdf');
  const folder  = DriveApp.getFolderById(FOLDER_ID_PDF);
  const pdfName = 'Remito_' + nroBuscado + '_' + (fechaStr || '') + '_' + (H[2] || '') + '.pdf';
  const file    = folder.createFile(pdfBlob).setName(pdfName);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = file.getUrl();

  // Guardar URL en REMITOS!PDF_URL
  const colPdf = ensurePdfUrlHeader_(shR);
  shR.getRange(idx + 1, colPdf).setValue(url);

  // Limpiamos el Doc fuente para no llenar el Drive
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return url;
}



/**
 * Arma el PDF con el modelo NUEVO (plantilla de Slides) para un remito ya
 * cargado en REMITOS, leyendo su header y detalle desde las hojas. Es lo
 * que usan los dos accesos del menú de Sheets — reemplaza el viejo
 * generarPDFDesdeRemito() (Google Docs), que queda solo como referencia.
 */
function generarPDFDesdeRemitoPlantilla_(nroRemito) {
  const ss  = getSS_();
  const shR = ss.getSheetByName(HOJA_REMITOS);
  const shD = ss.getSheetByName(HOJA_DETALLE);
  if (!shR || !shD) throw new Error('No se encontraron las hojas REMITOS o DETALLEREMITOS');

  const dataR = shR.getDataRange().getValues();
  const nroBuscado = String(nroRemito).trim();
  let idx = -1;
  for (let i = 1; i < dataR.length; i++) {
    if (String(dataR[i][0]).trim() === nroBuscado) { idx = i; break; }
  }
  if (idx === -1) throw new Error('Remito no encontrado: ' + nroRemito);

  const H = dataR[idx]; // [NroRemito, Fecha, PuntoVenta, Usuario, Obs, Timestamp, PDF_URL]
  const fechaObj = toJsDate_(H[1]);
  const fecha = fechaObj
    ? Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(H[1] || '');

  const header = { fecha: fecha, punto_venta: String(H[2] || ''), obs: String(H[4] || '') };

  const dataD = shD.getDataRange().getValues();
  const lines = [];
  for (let i = 1; i < dataD.length; i++) {
    if (String(dataD[i][0]).trim() === nroBuscado) {
      lines.push({ sku: dataD[i][1], descripcion: dataD[i][2], cantidad: dataD[i][5] });
    }
  }

  const result = buildRemitoPdfFromTemplate_(nroBuscado, header, lines);

  const colPdf = ensurePdfUrlHeader_(shR);
  shR.getRange(idx + 1, colPdf).setValue(result.pdfUrl);

  return result.pdfUrl;
}

/* ==========================================================
 * GENERAR TODOS LOS PDFs PENDIENTES
 * ========================================================== */
function generarPDFsPendientes() {
  const ss  = getSS_();
  const shR = ss.getSheetByName(HOJA_REMITOS);
  const vals = shR.getDataRange().getValues();
  const colPdf = ensurePdfUrlHeader_(shR);   // 1-based
  const pdfIdx = colPdf - 1;                 // 0-based

  for (let i = 1; i < vals.length; i++) {
    const nro = vals[i][0];
    const url = vals[i][pdfIdx];
    if (!nro) continue;
    if (!url) {
      try {
        const u = generarPDFDesdeRemitoPlantilla_(nro);
        Logger.log('✅ ' + nro + ' → ' + u);
      } catch (e) {
        Logger.log('❌ ' + nro + ' → ' + e);
      }
    }
  }
}

/**
 * Red de seguridad del PDF. Instalar UNA VEZ desde el editor de Apps Script
 * (o desde el menú 📦 Remitos de la planilla): deja un trigger por tiempo
 * que cada 5 minutos completa el PDF de cualquier remito que haya quedado
 * con PDF_URL vacío.
 *
 * Por qué hace falta: el PDF lo dispara el navegador después de registrar
 * el remito. Si el operario cierra la pestaña, se queda sin señal justo en
 * ese momento, o la cola offline sincroniza con nadie mirando la pantalla,
 * esa segunda llamada nunca sale y el remito queda sin PDF para siempre.
 * Con el trigger, el PDF aparece igual sin que nadie haga nada.
 *
 * Es idempotente: borra el trigger anterior de esta misma función antes de
 * crear el nuevo, así no se acumulan duplicados si se corre varias veces.
 */
function instalarTriggerPdfsPendientes() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'generarPDFsPendientes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('generarPDFsPendientes').timeBased().everyMinutes(5).create();
  Logger.log('Trigger de generarPDFsPendientes instalado (cada 5 minutos).');
}

/* ==========================================================
 * MENÚ EN GOOGLE SHEETS
 * ========================================================== */
function onOpen() {
  const ss = getSS_();
  ensureSheets_(ss);
  SpreadsheetApp.getUi()
    .createMenu('📦 Remitos')
    .addItem('Generar PDF del remito seleccionado', 'menuGenerarPDFSeleccionado')
    .addItem('Generar TODOS los PDFs pendientes', 'generarPDFsPendientes')
    .addSeparator()
    .addItem('Instalar auto-generación de PDFs (cada 5 min)', 'instalarTriggerPdfsPendientes')
    .addToUi();
}

function menuGenerarPDFSeleccionado() {
  const ss  = getSS_();
  ensureSheets_(ss);
  const shR = ss.getSheetByName(HOJA_REMITOS);
  const ui  = SpreadsheetApp.getUi();

  try {
    const range = shR.getActiveRange();
    const row   = range.getRow();

    if (row === 1) {
      ui.alert('Seleccioná una fila de datos (no el encabezado).');
      return;
    }

    const nro = String(shR.getRange(row, 1).getValue()).trim(); // Columna A = NroRemito

    if (!nro || nro.toLowerCase().indexOf('nro') === 0) {
      ui.alert('La fila seleccionada no tiene un número de remito válido en la columna A.');
      return;
    }

    const url = generarPDFDesdeRemitoPlantilla_(nro);
    ui.alert('PDF generado:\n' + url);

  } catch (e) {
    ui.alert('ERROR al generar el PDF desde el menú:\n' + e);
  }
}

/* ==========================================================
 * ENDPOINTS WebApp (Netlify)
 * ========================================================== */
function doOptions() {
  return jsonOut({}, 204);
}

function doGet(e) {
  const ss = getSS_();
  ensureSheets_(ss);

  const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : '';
  const token  = (e && e.parameter && e.parameter.token) ? String(e.parameter.token) : '';

  if (!action) {
    return jsonOut({
      ok: true,
      info: 'Remitera activa. Usar POST con { token, header, lines } o GET ?action=stock|historial&token=...'
    });
  }

  if (TOKEN_ESPERADO && token !== TOKEN_ESPERADO) {
    return jsonOut({ ok:false, error:'Token inválido' }, 401);
  }

  if (action === 'stock') {
    return handleGetStock_(ss);
  }

  if (action === 'historial') {
    return handleGetHistorial_(ss);
  }

  if (action === 'pedidos') {
    return handleGetPedidos_(ss);
  }

  if (action === 'clientes') {
    return handleGetClientes_(ss);
  }

  return jsonOut({ ok:false, error:'Acción desconocida: ' + action }, 400);
}

function handleGetStock_(ss) {
  const shStock = ss.getSheetByName(HOJA_STOCK);
  if (!shStock) return jsonOut({ ok:false, error:'No existe la hoja ' + HOJA_STOCK }, 500);

  const map = getStockMap_(shStock);
  if (map.lastRow < 2) return jsonOut({ ok:true, rows: [] });

  const vals = shStock.getRange(2, 1, map.lastRow - 1, map.lastCol).getValues();

  const rows = vals
    .filter(r => String(r[map.idxSku] || '').trim() !== '')
    .map(r => ({
      sku:          map.idxSku          >= 0 ? String(r[map.idxSku]).trim() : '',
      descripcion:  map.idxDescripcion  >= 0 ? String(r[map.idxDescripcion] || '') : '',
      tipo:         map.idxTipo         >= 0 ? String(r[map.idxTipo] || '') : '',
      marca:        map.idxMarca        >= 0 ? String(r[map.idxMarca] || '') : '',
      stockInicial: map.idxStockInicial >= 0 ? Number(r[map.idxStockInicial] || 0) : 0,
      entregado:    map.idxEntregado    >= 0 ? Number(r[map.idxEntregado] || 0) : 0,
      stockActual:  map.idxStockActual  >= 0 ? Number(r[map.idxStockActual] || 0) : 0,
      imagenUrl:    map.idxImagen       >= 0 ? String(r[map.idxImagen] || '') : '',
      medidas:      map.idxMedidas      >= 0 ? String(r[map.idxMedidas] || '') : ''
    }));

  return jsonOut({ ok:true, headerStockInicial: map.headers[map.idxStockInicial] || 'StockInicial', rows: rows });
}

function handleGetHistorial_(ss) {
  const shRemitos = ss.getSheetByName(HOJA_REMITOS);
  const shDet     = ss.getSheetByName(HOJA_DETALLE);
  if (!shRemitos || !shDet) return jsonOut({ ok:false, error:'Faltan hojas de remitos' }, 500);

  const lastR = shRemitos.getLastRow();
  const remitos = [];

  if (lastR >= 2) {
    const valsR = shRemitos.getRange(2, 1, lastR - 1, Math.max(7, shRemitos.getLastColumn())).getValues();
    valsR.forEach(r => {
      const nro = String(r[0] || '').trim();
      if (!nro) return;
      const fechaDate = toJsDate_(r[1]);
      const tsDate = toJsDate_(r[5]);
      remitos.push({
        nroRemito: nro,
        fecha: fechaDate ? Utilities.formatDate(fechaDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[1] || ''),
        puntoVenta: String(r[2] || ''),
        usuario: String(r[3] || ''),
        obs: String(r[4] || ''),
        timestamp: tsDate ? tsDate.toISOString() : String(r[5] || ''),
        pdfUrl: String(r[6] || ''),
        lineas: []
      });
    });
  }

  const byNro = {};
  remitos.forEach(rm => { byNro[rm.nroRemito] = rm; });

  const lastD = shDet.getLastRow();
  if (lastD >= 2) {
    const valsD = shDet.getRange(2, 1, lastD - 1, 6).getValues();
    valsD.forEach(d => {
      const nro = String(d[0] || '').trim();
      if (!nro || !byNro[nro]) return;
      byNro[nro].lineas.push({
        sku: String(d[1] || ''),
        descripcion: String(d[2] || ''),
        tipo: String(d[3] || ''),
        marca: String(d[4] || ''),
        cantidad: Number(d[5] || 0)
      });
    });
  }

  remitos.forEach(rm => {
    rm.totalItems = rm.lineas.reduce((acc, l) => acc + (Number(l.cantidad) || 0), 0);
  });

  // más recientes primero
  remitos.sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));

  return jsonOut({ ok:true, remitos: remitos });
}

/* ==========================================================
 * PEDIDOS DEL DÍA — respuestas del Google Form de pedido de materiales
 * Viven en una planilla APARTE (PEDIDOS_SHEET_ID), pestaña PEDIDOS_TAB_NAME.
 * Columnas 0 (timestamp) y 9 (estado) del Form vienen ambas tituladas
 * "Columna 1" (mismo nombre repetido), así que esas dos se leen por
 * posición fija; el resto se resuelve por nombre de encabezado.
 * ========================================================== */
function handleGetPedidos_() {
  let sh;
  try {
    sh = SpreadsheetApp.openById(PEDIDOS_SHEET_ID).getSheetByName(PEDIDOS_TAB_NAME);
  } catch (eOpen) {
    return jsonOut({ ok:false, error:'No se pudo abrir la planilla de pedidos: ' + eOpen }, 500);
  }
  if (!sh) return jsonOut({ ok:false, error:'No existe la hoja ' + PEDIDOS_TAB_NAME }, 500);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return jsonOut({ ok:true, pedidos: [] });

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const idx = {
    timestamp:   0,
    solicitante: stockColIndex_(headers, ['repositor']),
    cliente:     stockColIndex_(headers, ['supermercado']),
    direccion:   stockColIndex_(headers, ['direccion','dirección']),
    pedido:      stockColIndex_(headers, ['tipo de material']),
    marca:       stockColIndex_(headers, ['marca']),
    cantidad:    stockColIndex_(headers, ['cantidad']),
    contacto:    stockColIndex_(headers, ['telefono','teléfono']),
    ciudad:      stockColIndex_(headers, ['ciudad']),
    estado:      9,
    observacion: stockColIndex_(headers, ['observacion','observación'])
  };

  const vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const pedidos = vals.map((r, i) => {
    const tsDate = toJsDate_(r[idx.timestamp]);
    const marca = idx.marca >= 0 ? String(r[idx.marca] || '') : '';
    const tipoMaterial = idx.pedido >= 0 ? String(r[idx.pedido] || '') : '';
    return {
      id: 'PED' + (i + 2),
      timestamp: tsDate ? tsDate.toISOString() : '',
      fecha: tsDate ? Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      canal: '',
      solicitante: idx.solicitante >= 0 ? String(r[idx.solicitante] || '') : '',
      cliente: idx.cliente >= 0 ? String(r[idx.cliente] || '') : '',
      codigoCliente: '',
      pedido: tipoMaterial + (marca ? ' — ' + marca : ''),
      cantidad: idx.cantidad >= 0 ? String(r[idx.cantidad] || '') : '',
      contacto: idx.contacto >= 0 ? String(r[idx.contacto] || '') : '',
      direccion: idx.direccion >= 0 ? String(r[idx.direccion] || '') : '',
      ciudad: idx.ciudad >= 0 ? String(r[idx.ciudad] || '') : '',
      estado: idx.estado >= 0 ? String(r[idx.estado] || '') : '',
      observacion: idx.observacion >= 0 ? String(r[idx.observacion] || '') : ''
    };
  }).filter(p => p.solicitante || p.pedido || p.cliente);

  pedidos.sort((a, b) => String(b.timestamp || b.fecha).localeCompare(String(a.timestamp || a.fecha)));

  return jsonOut({ ok:true, pedidos: pedidos });
}

/**
 * Actualiza el estado de un pedido en la planilla externa de pedidos.
 * El id viene como "PED<fila>" (ver handleGetPedidos_), así se sabe
 * exactamente qué fila tocar sin tener que re-buscar por texto.
 */
function handleUpdateEstadoPedido_(data) {
  const id = String(data.id || '').trim();
  const estado = String(data.estado || '').trim();
  if (!id) return jsonOut({ ok:false, error:'Falta id de pedido' }, 400);

  const row = Number(id.replace(/\D/g, ''));
  if (!row || row < 2) return jsonOut({ ok:false, error:'Id de pedido inválido' }, 400);

  let sh;
  try {
    sh = SpreadsheetApp.openById(PEDIDOS_SHEET_ID).getSheetByName(PEDIDOS_TAB_NAME);
  } catch (eOpen) {
    return jsonOut({ ok:false, error:'No se pudo abrir la planilla de pedidos: ' + eOpen }, 500);
  }
  if (!sh) return jsonOut({ ok:false, error:'No existe la hoja ' + PEDIDOS_TAB_NAME }, 500);
  if (row > sh.getLastRow()) return jsonOut({ ok:false, error:'La fila del pedido ya no existe' }, 404);

  sh.getRange(row, 10).setValue(estado); // columna J (índice 9, 1-based=10) = Estado
  return jsonOut({ ok:true, id: id, estado: estado });
}

/* ==========================================================
 * FICHA DE CLIENTES — hoja "CLIENTES" de esta misma planilla (STOCK/REMITOS)
 * ========================================================== */
function handleGetClientes_(ss) {
  const sh = ss.getSheetByName('CLIENTES');
  if (!sh) return jsonOut({ ok:false, error:'No existe la hoja CLIENTES' }, 500);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return jsonOut({ ok:true, clientes: [] });

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const idx = {
    codigo:         stockColIndex_(headers, ['cod. chess','codigo','código']),
    canal:          stockColIndex_(headers, ['canal']),
    responsable:    stockColIndex_(headers, ['responsable']),
    razonSocial:    stockColIndex_(headers, ['razon social','razón social']),
    nombreFantasia: stockColIndex_(headers, ['nombre de fantasia','nombre de fantasía','nombre fantasia']),
    direccion:      stockColIndex_(headers, ['direccion','dirección']),
    localidad:      stockColIndex_(headers, ['localidad'])
  };

  const vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const clientes = vals.map(r => ({
    codigo: idx.codigo >= 0 ? String(r[idx.codigo] || '') : '',
    canal: idx.canal >= 0 ? String(r[idx.canal] || '') : '',
    responsable: idx.responsable >= 0 ? String(r[idx.responsable] || '') : '',
    razonSocial: idx.razonSocial >= 0 ? String(r[idx.razonSocial] || '') : '',
    nombreFantasia: idx.nombreFantasia >= 0 ? String(r[idx.nombreFantasia] || '') : '',
    direccion: idx.direccion >= 0 ? String(r[idx.direccion] || '') : '',
    localidad: idx.localidad >= 0 ? String(r[idx.localidad] || '') : ''
  })).filter(c => c.razonSocial || c.nombreFantasia || c.codigo);

  return jsonOut({ ok:true, clientes: clientes });
}

function doPost(e) {
  try {
    const ss = getSS_();
    ensureSheets_(ss);

    const shStock   = ss.getSheetByName(HOJA_STOCK);
    const shRemitos = ss.getSheetByName(HOJA_REMITOS);
    const shDet     = ss.getSheetByName(HOJA_DETALLE);

    if (!shStock || !shRemitos || !shDet) {
      return jsonOut({ ok:false, error:'Faltan hojas requeridas.' }, 500);
    }

    const data      = readPostBody_(e);
    const token     = data.token || '';
    const action    = data.action || '';

    if (TOKEN_ESPERADO && token !== TOKEN_ESPERADO) {
      return jsonOut({ ok:false, error:'Token inválido' }, 401);
    }

    if (action === 'addProducto') {
      return handleAddProducto_(shStock, data.producto || {});
    }

    if (action === 'uploadImagen') {
      return handleUploadImagen_(shStock, data);
    }

    if (action === 'sumarStock') {
      return handleSumarStock_(shStock, data);
    }

    if (action === 'updateEstadoPedido') {
      return handleUpdateEstadoPedido_(data);
    }

    if (action === 'attachPdf') {
      return handleAttachPdf_(shRemitos, data);
    }

    if (action === 'generarPdfRemito') {
      return handleGenerarPdfRemito_(shRemitos, data);
    }

    const header    = data.header || {};
    const lines     = Array.isArray(data.lines) ? data.lines : [];
    const pdfBase64 = data.pdfBase64 || '';
    const pdfName   = data.pdfName   || 'remito.pdf';
    const soloPdf   = !!data.soloPdf;

    const fecha = header.fecha || Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

    /* -------------------------
     * MODO SOLO PDF
     * ------------------------- */
    if (soloPdf) {
      if (!header.punto_venta)
        return jsonOut({ ok:false, error:'Falta header.punto_venta (soloPdf)' }, 400);
      if (!pdfBase64)
        return jsonOut({ ok:false, error:'Falta pdfBase64 (soloPdf)' }, 400);

      const url = savePdfToDrive_(pdfBase64, pdfName);
      const colPdf = ensurePdfUrlHeader_(shRemitos);
      const row = findLastRemitoRowByHeader_(
        shRemitos,
        fecha,
        header.punto_venta,
        header.usuario || ''
      );

      if (row > 0) {
        shRemitos.getRange(row, colPdf).setValue(url);
        return jsonOut({ ok:true, modo:'soloPdf', row, pdfUrl:url });
      } else {
        let shPDF = ss.getSheetByName('PDFs') || ss.insertSheet('PDFs');
        if (shPDF.getLastRow() < 1) {
          shPDF.appendRow(['Fecha','PuntoVenta','Usuario','Obs','PDF_URL']);
        }
        shPDF.appendRow([fecha, header.punto_venta, header.usuario || '', header.obs || '', url]);
        return jsonOut({
          ok:true,
          modo:'soloPdf',
          aviso:'No se encontró remito. Se registró en "PDFs".',
          pdfUrl:url
        });
      }
    }

    /* -------------------------
     * CREAR REMITO NORMAL
     * ------------------------- */
    if (!header.punto_venta)
      return jsonOut({ ok:false, error:'Falta punto_venta' }, 400);
    if (!lines.length)
      return jsonOut({ ok:false, error:'Sin líneas' }, 400);

    const requestId = String(data.requestId || '').trim();
    const colRequestId = ensureRequestIdHeader_(shRemitos);

    // Validamos las líneas ANTES de tocar REMITOS. Antes este chequeo se
    // hacía después de crear la fila del remito, así que un remito con
    // TODAS las líneas inválidas (SKU inexistente en STOCK, cantidad<=0,
    // etc.) igual quedaba grabado en REMITOS con nro asignado, pero sin
    // ninguna fila en DETALLEREMITOS — un remito "huérfano". Validar antes
    // evita crear el remito si no hay ni una línea válida para guardar.
    const lineasValidas = [];
    const errores = [];

    for (const l of lines) {
      const sku  = String(l.sku || '').trim();
      const cant = Number(l.cantidad || 0);

      if (!sku || !(cant > 0)) {
        errores.push({ sku, error:'SKU vacío o cantidad inválida' });
        continue;
      }
      if (VALIDAR_SKU_EN_STOCK && !skuExisteEnStock_(shStock, sku)) {
        errores.push({ sku, error:'SKU no existe en ' + HOJA_STOCK });
        continue;
      }

      lineasValidas.push({
        sku,
        descripcion: l.descripcion || '',
        tipo: l.tipo || '',
        marca: l.marca || '',
        cantidad: cant
      });
    }

    if (!lineasValidas.length) {
      return jsonOut({ ok:false, error:'Ninguna línea válida', detalle:errores }, 400);
    }

    // Lock para que dos ejecuciones concurrentes (doble tap, reintento de la
    // cola offline solapado con el envío original, etc.) no pisen la
    // verificación de duplicado ni generen dos números de remito a la vez.
    const lock = LockService.getScriptLock();
    const gotLock = lock.tryLock(30000);
    if (!gotLock) {
      return jsonOut({ ok:false, error:'El sistema está ocupado procesando otro remito, reintentá en unos segundos.' }, 429);
    }

    let nro;
    try {
      if (requestId) {
        const existente = buscarNroPorRequestId_(shRemitos, colRequestId, requestId);
        if (existente) {
          // Ya se había registrado este mismo remito antes (reintento tras
          // corte de red). Devolvemos el mismo número sin duplicar nada.
          return jsonOut({ ok:true, nroRemito: existente, deduplicado:true });
        }
      }

      nro = nextRemitoNumber_(shRemitos);
      const ts = new Date();

      const rowValues = [nro, fecha, header.punto_venta, header.usuario || '', header.obs || '', ts];
      // completar columnas intermedias si existieran otras cabeceras entre
      // Timestamp y RequestId (PDF_URL, etc.) para no desalinear la fila
      while (rowValues.length < colRequestId - 1) rowValues.push('');
      rowValues[colRequestId - 1] = requestId;

      shRemitos.appendRow(rowValues);
    } finally {
      lock.releaseLock();
    }

    const vr = lineasValidas.map(function (l) {
      return [nro, l.sku, l.descripcion, l.tipo, l.marca, l.cantidad];
    });

    shDet.getRange(shDet.getLastRow() + 1, 1, vr.length, 6).setValues(vr);

    // Fila real del remito recién creado. NO se usa shRemitos.getLastRow()
    // para escribirle nada: si entre medio entró otro remito, esa última
    // fila ya no es la nuestra y terminaríamos pisando el remito ajeno.
    const rowRemito = findRemitoRowByNro_(shRemitos, nro);

    // Descontar del STOCK lo entregado en este remito (Entregado +cant,
    // StockActual = StockInicial - Entregado). Ver descontarStockPorRemito_.
    try {
      descontarStockPorRemito_(shStock, shRemitos, vr, nro);
    } catch (eStock) {
      Logger.log('Error descontando stock del remito ' + nro + ': ' + eStock);
    }

    /* -------------------------
     * PDF — NO se arma acá adentro.
     *
     * Tentador sería generarlo dentro de este mismo doPost, pero la llamada
     * entra por una Netlify Function y esas cortan a los 10 s: registrar el
     * remito (~2-3 s) + copiar la plantilla de Slides y exportar el PDF
     * (~5-10 s) se pasa del presupuesto, y el timeout haría fallar el
     * REGISTRO del remito, que es lo único que no se puede perder.
     *
     * El PDF se pide aparte (action:'generarPdfRemito'), con reintentos del
     * lado del navegador. Y para que no dependa de que el navegador siga
     * vivo, hay una red de seguridad en el servidor: el trigger de
     * generarPDFsPendientes() (ver instalarTriggerPdfsPendientes) que cada
     * pocos minutos completa cualquier remito que haya quedado sin PDF_URL.
     * ------------------------- */
    let pdfUrl = '';

    if (pdfBase64) {
      const colPdf = ensurePdfUrlHeader_(shRemitos);
      pdfUrl = savePdfToDrive_(pdfBase64, pdfName);
      if (rowRemito > 0) shRemitos.getRange(rowRemito, colPdf).setValue(pdfUrl);
    }

    const resp = { ok:true, nroRemito:nro };
    if (pdfUrl) resp.pdfUrl = pdfUrl;
    if (errores.length) resp.detalle = errores;

    return jsonOut(resp);

  } catch (err) {
    return jsonOut({ ok:false, error:String(err) }, 500);
  }
}

/* ==========================================================
 * ALTA DE PRODUCTO NUEVO EN STOCK
 * NUNCA pisa filas existentes: solo hace appendRow.
 * Respeta el orden real de columnas de la hoja (las ubica por nombre).
 * ========================================================== */
function handleAddProducto_(shStock, producto) {
  let sku = String(producto.sku || '').trim();
  const descripcion = String(producto.descripcion || '').trim();

  if (!descripcion) return jsonOut({ ok:false, error:'Falta Descripción' }, 400);

  if (!sku) {
    // Sin SKU manual: se genera automáticamente (máximo código numérico en STOCK + 1)
    sku = nextStockSku_(shStock);
  } else if (skuExisteEnStock_(shStock, sku)) {
    return jsonOut({ ok:false, error:'El SKU ' + sku + ' ya existe en ' + HOJA_STOCK }, 400);
  }

  const map = getStockMap_(shStock);
  const stockInicial = Number(producto.stockInicial || 0);
  const entregado = 0;

  const newRow = new Array(map.lastCol).fill('');
  if (map.idxSku          >= 0) newRow[map.idxSku] = sku;
  if (map.idxDescripcion  >= 0) newRow[map.idxDescripcion] = descripcion;
  if (map.idxTipo         >= 0) newRow[map.idxTipo] = String(producto.tipo || '').trim();
  if (map.idxMarca        >= 0) newRow[map.idxMarca] = String(producto.marca || '').trim();
  if (map.idxStockInicial >= 0) newRow[map.idxStockInicial] = stockInicial;
  if (map.idxEntregado    >= 0) newRow[map.idxEntregado] = entregado;
  if (map.idxStockActual  >= 0) newRow[map.idxStockActual] = stockInicial - entregado;

  const targetRow = map.lastRow + 1;
  shStock.getRange(targetRow, 1, 1, map.lastCol).setValues([newRow]);

  let imagenUrl = '';
  let imagenError = '';
  if (producto.imagenBase64) {
    try {
      imagenUrl = saveImagenToDrive_(producto.imagenBase64, producto.imagenName || ('sku_' + sku + '.jpg'));
      const colImagen = ensureImagenUrlHeader_(shStock);
      shStock.getRange(targetRow, colImagen).setValue(imagenUrl);
    } catch (eImg) {
      imagenError = String(eImg);
      Logger.log('Error subiendo imagen para ' + sku + ': ' + eImg);
    }
  }

  const medidas = String(producto.medidas || '').trim();
  if (medidas) {
    const colMedidas = ensureMedidasHeader_(shStock);
    shStock.getRange(targetRow, colMedidas).setValue(medidas);
  }

  return jsonOut({ ok:true, sku: sku, row: targetRow, imagenUrl: imagenUrl, imagenError: imagenError, medidas: medidas });
}

/* ==========================================================
 * SUBIR/REEMPLAZAR FOTO DE UN PRODUCTO YA EXISTENTE EN STOCK
 * ========================================================== */
function handleUploadImagen_(shStock, data) {
  const sku = String(data.sku || '').trim();
  const imagenBase64 = data.imagenBase64 || '';

  if (!sku) return jsonOut({ ok:false, error:'Falta SKU' }, 400);
  if (!imagenBase64) return jsonOut({ ok:false, error:'Falta la imagen' }, 400);

  const row = findStockRowBySku_(shStock, sku);
  if (row === -1) return jsonOut({ ok:false, error:'El SKU ' + sku + ' no existe en ' + HOJA_STOCK }, 404);

  let imagenUrl;
  try {
    imagenUrl = saveImagenToDrive_(imagenBase64, data.imagenName || ('sku_' + sku + '.jpg'));
  } catch (eImg) {
    return jsonOut({ ok:false, error:'No se pudo subir la imagen: ' + eImg }, 500);
  }

  const colImagen = ensureImagenUrlHeader_(shStock);
  shStock.getRange(row, colImagen).setValue(imagenUrl);

  return jsonOut({ ok:true, sku: sku, imagenUrl: imagenUrl });
}

/* ==========================================================
 * SUMAR STOCK A UN PRODUCTO YA EXISTENTE (reposición)
 * Incrementa StockInicial y StockActual; nunca toca Entregado.
 * Opcionalmente también permite subir/actualizar la foto en la misma acción.
 * ========================================================== */
function handleSumarStock_(shStock, data) {
  const sku = String(data.sku || '').trim();
  const cantidad = Number(data.cantidad || 0);
  const medidas = String(data.medidas || '').trim();

  if (!sku) return jsonOut({ ok:false, error:'Falta SKU' }, 400);
  if (!(cantidad > 0) && !medidas && !data.imagenBase64) {
    return jsonOut({ ok:false, error:'Nada para actualizar: indicá cantidad, medidas o una foto' }, 400);
  }

  const row = findStockRowBySku_(shStock, sku);
  if (row === -1) return jsonOut({ ok:false, error:'El SKU ' + sku + ' no existe en ' + HOJA_STOCK }, 404);

  const map = getStockMap_(shStock);

  if (cantidad > 0) {
    // Mismo criterio que descontarStockPorRemito_: se leen StockInicial,
    // Entregado y StockActual ANTES de escribir nada, y StockActual se
    // recalcula absoluto (StockInicial - Entregado). Antes se escribía
    // StockInicial y recién después se leía StockActual, así que si esa
    // columna tenía fórmula ya venía recalculada y se le volvía a sumar
    // la cantidad: el ingreso impactaba dos veces.
    const colInicial   = map.idxStockInicial >= 0 ? map.idxStockInicial + 1 : -1;
    const colEntregado = map.idxEntregado    >= 0 ? map.idxEntregado    + 1 : -1;
    const colActual    = map.idxStockActual  >= 0 ? map.idxStockActual  + 1 : -1;
    const colsLeer = [colInicial, colEntregado, colActual].filter(function (c) { return c > 0; });

    if (colsLeer.length) {
      const readStart = Math.min.apply(null, colsLeer);
      const readEnd   = Math.max.apply(null, colsLeer);
      const leidos = shStock.getRange(row, readStart, 1, readEnd - readStart + 1).getValues()[0];
      const val = function (col) { return Number(leidos[col - readStart] || 0); };

      const inicialNuevo = colInicial > 0 ? val(colInicial) + cantidad : null;
      let actualNuevo = null;
      if (colActual > 0) {
        if (colInicial > 0 && colEntregado > 0) actualNuevo = inicialNuevo - val(colEntregado);
        else actualNuevo = val(colActual) + cantidad;
      }

      if (inicialNuevo !== null) shStock.getRange(row, colInicial).setValue(inicialNuevo);
      if (actualNuevo  !== null) shStock.getRange(row, colActual).setValue(actualNuevo);
    }
  }

  let imagenUrl = '';
  let imagenError = '';
  if (data.imagenBase64) {
    try {
      imagenUrl = saveImagenToDrive_(data.imagenBase64, data.imagenName || ('sku_' + sku + '.jpg'));
      const colImagen = ensureImagenUrlHeader_(shStock);
      shStock.getRange(row, colImagen).setValue(imagenUrl);
    } catch (eImg) {
      imagenError = String(eImg);
      Logger.log('Error subiendo imagen (sumarStock) ' + sku + ': ' + eImg);
    }
  }

  if (medidas) {
    const colMedidas = ensureMedidasHeader_(shStock);
    shStock.getRange(row, colMedidas).setValue(medidas);
  }

  return jsonOut({ ok:true, sku: sku, row: row, imagenUrl: imagenUrl, imagenError: imagenError, medidas: medidas });
}
