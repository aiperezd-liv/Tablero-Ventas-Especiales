/**
 * TABLERO COLOCACIÓN / PERFIL - BACKEND (Google Apps Script)
 *  - Ventas_Especiales_Acumulado -> vistas Trimestral, Resumen Tipo de Venta y Perfil   (rows)
 *  - Ventas_Especiales           -> Vista Mensual (tabla y gráficos)                    (mensual)
 * Los cortes y las métricas se calculan en el front.
 */

const SHEET_ACUMULADO = ["Ventas_Especiales_Acumulado"];
const SHEET_MENSUAL = ["Ventas_Especiales"];

const MESES_ABR = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
// Prefijos aceptados al leer meses escritos (español / inglés)
const MES_PREFIJOS = [['ene', 'jan'], ['feb'], ['mar'], ['abr', 'apr'], ['may'], ['jun'], ['jul'],
                      ['ago', 'aug'], ['sep'], ['oct'], ['nov'], ['dic', 'dec']];

// Columnas de texto (encabezado en minúsculas -> campo del front). Todo lo demás se lee como número.
const TEXT_ACUMULADO = {
  'anio': 'anio', 'mes': 'mesNum', 'trimestre': 'trimestre', 'producto': 'producto',
  'segmento': 'segmento', 'temp_venta': 'venta', 'tipo_de_venta': 'tipoVentaOrig'
};
const REQ_ACUMULADO = ['anio', 'mes', 'producto', 'segmento', 'temp_venta', 'tipo_de_venta', 'solicitudes', 'cuentas'];
const REQ_MENSUAL = ['producto', 'segmento', 'venta', 'tipo_de_venta', 'mes', 'solicitudes', 'cuentas'];

/** Llamada desde el front: google.script.run.getDataFromSheet() */
function getDataFromSheet() {
  if (SHEET_ACUMULADO[0] === SHEET_MENSUAL[0]) throw new Error('SHEET_ACUMULADO y SHEET_MENSUAL apuntan a la misma pestaña');
  const acumulado = readAcumulado_();
  const mensual = readMensual_();
  return {
    rows: acumulado.rows,
    mensual: mensual.rows,
    controlMensual: mensual.control,   // incluye la pestaña fuente y totales para cuadrar contra Looker
    fuenteAcumulado: acumulado.control.fuente,
    timestamp: new Date().toISOString()
  };
}

// ---------- Hoja acumulada ----------
function readAcumulado_() {
  return readRows_(SHEET_ACUMULADO, TEXT_ACUMULADO, REQ_ACUMULADO, function (obj) {
    obj.anio = String(obj.anio).trim().replace(/\.0+$/, '');
    obj.mesNum = parseInt(obj.mesNum, 10) || 0;
    return obj.anio && obj.producto;
  });
}

// ---------- Hoja mensual (Ventas_Especiales) ----------
// Para ser rápido: lee SOLO las columnas que usa la vista mensual y envía filas ya agrupadas por
// año / mes / venta / tipo de venta / producto / segmento (en vez de las +75,000 filas originales).
const MENSUAL_COLS = ['producto', 'segmento', 'venta', 'tipo_de_venta', 'mes', 'tienda',
                      'dias_venta_especial', 'solicitudes', 'aprobadas', 'cuentas', 'sum_linea_onus'];

function readMensual_() {
  const t0 = Date.now();
  const sheet = findSheet_(SHEET_MENSUAL);
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  const control = { fuente: sheet.getName(), filasHoja: 0, filasEnviadas: 0, solicitudesHoja: 0, cuentasHoja: 0,
                    celdasNoNumericas: 0, sinAnio: 0, sinMes: 0, ms: 0 };
  if (lastRow < 2) return { rows: [], control: control };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  const missing = REQ_MENSUAL.filter(h => headers.indexOf(h) === -1);
  if (missing.length) {
    throw new Error("Faltan columnas en '" + sheet.getName() + "': " + missing.join(', ') +
                    ". Encabezados detectados: " + headers.join(', '));
  }

  const n = lastRow - 1;
  const col = {};
  MENSUAL_COLS.forEach(h => {
    const i = headers.indexOf(h);
    col[h] = i === -1 ? null : sheet.getRange(2, i + 1, n, 1).getValues();   // una columna a la vez
  });
  const get = (h, r) => col[h] ? col[h][r][0] : '';

  const periodoCache = {};
  const groups = {};
  for (let r = 0; r < n; r++) {
    const producto = String(get('producto', r)).trim();
    const venta = String(get('venta', r)).trim();
    const mesRaw = get('mes', r);
    const sol = toNumber_(get('solicitudes', r), control);
    const cta = toNumber_(get('cuentas', r), control);
    if (!producto && !venta && mesRaw === '' && !sol && !cta) continue;     // fila vacía

    control.filasHoja++;
    control.solicitudesHoja += sol;
    control.cuentasHoja += cta;

    const ck = (mesRaw instanceof Date ? mesRaw.getTime() : String(mesRaw)) + '|' + venta;
    const p = periodoCache[ck] || (periodoCache[ck] = parsePeriodo_(mesRaw, venta));
    const segmento = String(get('segmento', r)).trim();
    const tipoVenta = /bau/i.test(String(get('tipo_de_venta', r))) ? 'BAU' : 'Especial';
    const key = [p.anio, p.mes, venta, tipoVenta, producto, segmento].join('|');

    let g = groups[key];
    if (!g) {
      g = groups[key] = {
        anio: p.anio ? String(p.anio) : '', mesNum: p.mes, mes: MESES_ABR[p.mes - 1] || '',
        trimestre: p.mes ? 'T' + (Math.floor((p.mes - 1) / 3) + 1) : '',
        venta: venta, tipoVenta: tipoVenta, producto: producto, segmento: segmento,
        solicitudes: 0, aprobadas: 0, cuentas: 0, sum_linea_onus: 0, dias_venta_especial: 0, _t: {}
      };
      if (!g.anio) control.sinAnio++;
      if (!g.mesNum) control.sinMes++;
    }
    g.solicitudes += sol;
    g.cuentas += cta;
    g.aprobadas += toNumber_(get('aprobadas', r), control);
    g.sum_linea_onus += toNumber_(get('sum_linea_onus', r), control);
    const d = toNumber_(get('dias_venta_especial', r), control);
    if (d > g.dias_venta_especial) g.dias_venta_especial = d;
    const t = String(get('tienda', r)).trim();
    if (t) g._t[t] = 1;
  }

  const rows = Object.keys(groups).map(k => {
    const g = groups[k];
    g.tiendas = Object.keys(g._t).join(',');   // códigos de tienda distintos del grupo
    delete g._t;
    return g;
  });
  control.filasEnviadas = rows.length;
  control.ms = Date.now() - t0;
  return { rows: rows, control: control };
}

let _tzHoja = null;
function tzHoja_() { return _tzHoja || (_tzHoja = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()); }

/** Obtiene año y mes desde MES (fecha, 202605, "2026-05", "mayo 2026", 5...) y, si falta el año, desde el nombre de Venta */
function parsePeriodo_(raw, venta) {
  let year = 0, month = 0;

  if (raw instanceof Date) {
    // Se lee con la zona horaria de la HOJA (no la del script) para que 12/1/2024 no se corra a noviembre
    const tz = tzHoja_();
    year = parseInt(Utilities.formatDate(raw, tz, 'yyyy'), 10);
    month = parseInt(Utilities.formatDate(raw, tz, 'M'), 10);
  } else if (typeof raw === 'number') {
    if (raw >= 190001 && raw <= 299912) { year = Math.floor(raw / 100); month = raw % 100; }
    else if (raw >= 1 && raw <= 12) month = raw;
  } else if (raw !== '' && raw != null) {
    const s = String(raw).trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    let m;
    if ((m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/))) {            // texto m/d/aaaa (ej. 12/1/2024)
      year = +m[3]; month = (+m[1] > 12) ? +m[2] : +m[1];
    }
    else if ((m = s.match(/^(\d{4})[-\/.]?(\d{1,2})/))) { year = +m[1]; month = +m[2]; }
    else if ((m = s.match(/^(\d{1,2})[-\/.](\d{4})/))) { month = +m[1]; year = +m[2]; }
    else if (/^\d{1,2}$/.test(s)) { month = +s; }
    else {
      const y = s.match(/(19|20)\d{2}/);
      if (y) year = +y[0];
      for (let i = 0; i < 12; i++) {
        if (MES_PREFIJOS[i].some(pf => s.indexOf(pf) !== -1)) { month = i + 1; break; }
      }
    }
  }

  if (!year && venta) {
    const v = String(venta).trim();
    let m;
    if ((m = v.match(/(20\d{2})\s*$/))) year = +m[1];
    else if ((m = v.match(/(\d{2})\s*$/))) year = 2000 + (+m[1]);
  }
  if (month < 1 || month > 12) month = 0;
  return { anio: year, mes: month };
}

// ---------- Lector genérico ----------
function readRows_(names, textMap, required, fixFn) {
  const sheet = findSheet_(names);
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { rows: [], control: { fuente: sheet.getName(), filasHoja: 0, filasEnviadas: 0, solicitudesHoja: 0, cuentasHoja: 0, celdasNoNumericas: 0 } };

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());

  const missing = required.filter(h => headers.indexOf(h) === -1);
  if (missing.length) {
    throw new Error("Faltan columnas en '" + sheet.getName() + "': " + missing.join(', ') +
                    ". Encabezados detectados: " + headers.join(', '));
  }

  const out = [];
  const control = { fuente: sheet.getName(), filasHoja: 0, filasEnviadas: 0, solicitudesHoja: 0, cuentasHoja: 0, celdasNoNumericas: 0 };
  for (let r = 1; r < values.length; r++) {
    const obj = {};
    let hasData = false;
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      if (!h) continue;
      let v = values[r][c];
      if (v !== '' && v !== null) hasData = true;
      if (textMap[h]) {
        if (h === 'mes') obj[textMap[h]] = v;          // se interpreta después (puede ser fecha o número)
        else if (v instanceof Date) obj[textMap[h]] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        else obj[textMap[h]] = (v === '' || v === null) ? '' : v;
      } else {
        obj[h] = toNumber_(v, control);                // acepta "1,234", "$1,234", "12%"; vacío => 0
      }
    }
    if (!hasData) continue;                            // fila totalmente vacía
    control.filasHoja++;
    control.solicitudesHoja += obj.solicitudes || 0;
    control.cuentasHoja += obj.cuentas || 0;
    if (!fixFn(obj)) continue;

    obj.mes = MESES_ABR[obj.mesNum - 1] || '';
    obj.trimestre = obj.trimestre ? String(obj.trimestre).trim() : (obj.mesNum ? 'T' + (Math.floor((obj.mesNum - 1) / 3) + 1) : '');
    ['producto', 'segmento', 'venta'].forEach(k => obj[k] = String(obj[k]).trim());
    obj.tipoVenta = /bau/i.test(obj.tipoVentaOrig) ? 'BAU' : 'Especial';   // el front solo distingue BAU vs Especial
    out.push(obj);
  }
  control.filasEnviadas = out.length;
  return { rows: out, control: control };
}

/** Convierte a número texto como "1,234" o "$1,234"; cuenta las celdas con texto que no pudo convertir */
function toNumber_(v, control) {
  if (typeof v === 'number') return v;
  if (v === '' || v === null || v === undefined || v instanceof Date) return 0;
  const n = parseFloat(String(v).replace(/[$,\s%]/g, ''));
  if (isNaN(n)) { control.celdasNoNumericas++; return 0; }
  return n;
}

function findSheet_(names) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const n of names) {
    const s = ss.getSheetByName(n);
    if (s) return s;
  }
  throw new Error("No se encontró la pestaña. Busqué: " + names.join(' / ') +
                  ". Pestañas existentes: " + ss.getSheets().map(s => s.getName()).join(', '));
}

/** Sirve la app como Web App. El archivo HTML del proyecto debe llamarse "index". */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Tablero Colocación')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Prueba desde el editor: Ejecutar > probarDatos y revisa el registro */
function probarDatos() {
  const d = getDataFromSheet();
  Logger.log('Acumulado (' + d.fuenteAcumulado + '): ' + d.rows.length + ' filas | Mensual (' + d.controlMensual.fuente + '): ' + d.mensual.length + ' filas');
  Logger.log('Control mensual: ' + JSON.stringify(d.controlMensual));
  Logger.log('Ejemplo mensual: ' + JSON.stringify(d.mensual[0]));
}

/**
 * DIAGNÓSTICO: Ejecutar > diagnostico y abrir "Registro de ejecución".
 * Suma la hoja Ventas_Especiales directo (sin front) para comparar contra Looker.
 */
function diagnostico() {
  const m = readMensual_();
  const rows = m.rows;
  const sumBy = function (keyFn) {
    const o = {};
    rows.forEach(function (r) {
      const k = keyFn(r) || '(vacío)';
      if (!o[k]) o[k] = { solicitudes: 0, aprobadas: 0, cuentas: 0 };
      o[k].solicitudes += r.solicitudes || 0;
      o[k].aprobadas += r.aprobadas || 0;
      o[k].cuentas += r.cuentas || 0;
    });
    return o;
  };
  const imprimir = function (titulo, o) {
    Logger.log('--- ' + titulo + ' ---');
    Object.keys(o).sort().forEach(function (k) {
      Logger.log(k + ' | solicitudes ' + o[k].solicitudes + ' | aprobadas ' + o[k].aprobadas + ' | cuentas ' + o[k].cuentas);
    });
  };
  Logger.log('CONTROL: ' + JSON.stringify(m.control));
  imprimir('TOTAL', sumBy(function () { return 'TODO'; }));
  ['tipoVenta', 'producto', 'segmento'].forEach(function (dim) {
    imprimir('Por ' + dim, sumBy(function (r) { return String(r[dim]); }));
  });
  imprimir('Por año-mes', sumBy(function (r) { return r.anio + '-' + String(r.mesNum).padStart(2, '0'); }));
  imprimir('Por año-mes-venta', sumBy(function (r) { return r.anio + '-' + String(r.mesNum).padStart(2, '0') + ' | ' + r.venta; }));
}

/**
 * COMPARAR HOJAS: Ejecutar > compararHojas y abrir "Registro de ejecución".
 * Suma Solicitudes y Cuentas de cada pestaña, para ver cuál coincide con Looker
 * y si la extracción o la hoja con fórmulas está perdiendo filas.
 * Ajusta los nombres si tus pestañas se llaman distinto.
 */
function compararHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombres = ['Extract_1', 'Ventas_Especiales', 'Extract_2', 'Ventas_Especiales_Acumulado', 'Ventas_Especiales_Acumualdo'];

  nombres.forEach(function (nombre) {
    const sh = ss.getSheetByName(nombre);
    if (!sh) { Logger.log(nombre + ' -> (no existe)'); return; }
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2) { Logger.log(nombre + ' -> vacía'); return; }

    const vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const heads = vals[0].map(function (h) { return String(h).trim().toLowerCase(); });
    const iSol = heads.indexOf('solicitudes'), iCta = heads.indexOf('cuentas'), iApr = heads.indexOf('aprobadas');
    let filas = 0, sol = 0, cta = 0, apr = 0;
    for (let r = 1; r < vals.length; r++) {
      if (vals[r].every(function (v) { return v === '' || v === null; })) continue;
      filas++;
      if (iSol > -1) sol += Number(vals[r][iSol]) || 0;
      if (iCta > -1) cta += Number(vals[r][iCta]) || 0;
      if (iApr > -1) apr += Number(vals[r][iApr]) || 0;
    }
    Logger.log(nombre + ' -> filas: ' + filas + ' | solicitudes: ' + sol + ' | aprobadas: ' + apr + ' | cuentas: ' + cta +
               ' | columnas: ' + heads.length + (iSol === -1 ? ' | (sin columna Solicitudes)' : ''));
  });
}
