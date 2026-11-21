/* Panel de Seguimiento — PWA estática
   Los datos viven solo en localStorage de este dispositivo.
   Detección del tipo de archivo por COLUMNAS, no por nombre. */
'use strict';

var LS_KEY = 'panelSeguimiento.v1';
var RANGOS = ['0-3', '4-7', '8-15', '16-30', '+30'];
var COLORES_RANGO = ['#1a7f37', '#2e7d32', '#e8a13a', '#d95f02', '#b3261e'];

/* ---------------- utilidades ---------------- */
function num(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v).replace(/[$,\s]/g, '').replace(/ /g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function str(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function pad2(n) { return (n < 10 ? '0' : '') + n; }
function isoDeFecha(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
function hoyISO() { return isoDeFecha(new Date()); }

/* Acepta: Date, serial de Excel, texto DD.MM.AAAA, DD/MM/AAAA, AAAA-MM-DD */
function parseFecha(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v.getTime())) return isoDeFecha(v);
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    var d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isoDeFecha(d);
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/);
  if (m) return m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]);
  m = s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);
  if (m) return m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]);
  var d2 = new Date(s);
  if (!isNaN(d2.getTime())) return isoDeFecha(d2);
  return null;
}
function diasEntre(aISO, bISO) {
  if (!aISO || !bISO) return null;
  var a = new Date(aISO + 'T00:00:00');
  var b = new Date(bISO + 'T00:00:00');
  return Math.round((b - a) / 86400000);
}
function rangoDe(dias) {
  if (dias === null || dias === undefined) return '+30';
  if (dias <= 3) return '0-3';
  if (dias <= 7) return '4-7';
  if (dias <= 15) return '8-15';
  if (dias <= 30) return '16-30';
  return '+30';
}
function fmtDinero(v) {
  return '$' + num(v).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtNum(v, dec) {
  return num(v).toLocaleString('es-MX', { minimumFractionDigits: dec || 0, maximumFractionDigits: dec || 0 });
}
function fmtFecha(iso) {
  if (!iso) return '—';
  var p = iso.split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}
function esc(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/* ---------------- lectura de hojas ---------------- */
function filasDeHoja(wb, nombreHoja) {
  var ws = wb.Sheets[nombreHoja];
  if (!ws) return [];
  var rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
  return rows.map(function (r) {
    var o = {};
    Object.keys(r).forEach(function (k) { o[String(k).trim()] = r[k]; });
    return o;
  });
}
function encabezadosDe(wb, nombreHoja) {
  var ws = wb.Sheets[nombreHoja];
  if (!ws) return [];
  var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (!rows.length) return [];
  return rows[0].map(function (c) { return String(c).trim(); });
}
function tiene(cols, nombre) {
  return cols.some(function (c) { return c.toLowerCase() === nombre.toLowerCase(); });
}

/* ---------------- detección por columnas ---------------- */
var COLUMNAS_ESPERADAS =
  'MERMA: Centro, Fecha de entrada, Documento material, Material, Texto breve material, Ctd.en UM entrada, Costo Unitario, Costo Total.\n' +
  'TRANSFERENCIAS (DESTINO): Doc.Interno, Fecha Embarque, Tda.Origen, Descripción, Destino, Descripción, Piezas Embarcadas, Piezas Recibidas, Dif. Final, Costo Total, Dias Transcurridos.\n' +
  'TRANSFERENCIAS (SAP): hoja REPORTE con ZONA, SUCURSAL ORIGEN, DOCUMENTO CONTABLE, FECHA CONTABLE, DIAS TRANSCURRIDOS, IMPORTE, SUCURSAL DESTINO, ESTATUS.\n' +
  'BLOQUEADO (CONCENTRADO): hoja CONCENTRADO con Centro, Material, EAN/UPC, Denomin., Unidad, Jerarquía 2, Almacén, Bloqueado, V. Stock bloq. (y hoja RESUMEN con Centro, Importe bloqueado).\n' +
  'BLOQUEADO (original): Hoja1 con ZONA, Material, Centro, Sucursal, Almacen, Libre Util, Stock Bloq, Cto, Importe, FORMATO, Descripción, Departamento, Clase.';

function detectarTipo(wb) {
  var hojas = wb.SheetNames || [];
  // 1) Bloqueado CONCENTRADO: hoja CONCENTRADO con V. Stock bloq.
  if (hojas.indexOf('CONCENTRADO') >= 0) {
    var cCols = encabezadosDe(wb, 'CONCENTRADO');
    if (tiene(cCols, 'V. Stock bloq.') && tiene(cCols, 'Bloqueado')) return { tipo: 'bloq', formato: 'concentrado' };
  }
  if (hojas.indexOf('RESUMEN') >= 0) {
    var rCols = encabezadosDe(wb, 'RESUMEN');
    if (tiene(rCols, 'Importe bloqueado')) return { tipo: 'bloq', formato: 'concentrado' };
  }
  // 2) Transferencias SAP: hoja REPORTE con DOCUMENTO CONTABLE (antes que DESTINO,
  //    porque la Hoja1 del SAP trae el layout destino a nivel artículo)
  if (hojas.indexOf('REPORTE') >= 0) {
    var repCols = encabezadosDe(wb, 'REPORTE');
    if (tiene(repCols, 'DOCUMENTO CONTABLE') && tiene(repCols, 'IMPORTE')) return { tipo: 'transf', formato: 'sap' };
  }
  if (hojas.indexOf('RESUMEN_SUCURSAL') >= 0) return { tipo: 'transf', formato: 'sap' };
  // 3) Por columnas en cualquier hoja
  for (var i = 0; i < hojas.length; i++) {
    var cols = encabezadosDe(wb, hojas[i]);
    if (!cols.length) continue;
    if (tiene(cols, 'Stock Bloq') && tiene(cols, 'Importe') && tiene(cols, 'Libre Util')) {
      return { tipo: 'bloq', formato: 'original', hoja: hojas[i] };
    }
    if (tiene(cols, 'Fecha de entrada') && tiene(cols, 'Costo Total') && tiene(cols, 'Documento material')) {
      return { tipo: 'merma', formato: 'merma', hoja: hojas[i] };
    }
    if (tiene(cols, 'Doc.Interno') && tiene(cols, 'Destino') && tiene(cols, 'Dif. Final') &&
        tiene(cols, 'Costo Total') && tiene(cols, 'Dias Transcurridos') && !tiene(cols, 'Articulo')) {
      return { tipo: 'transf', formato: 'destino', hoja: hojas[i] };
    }
    if (tiene(cols, 'DOCUMENTO CONTABLE')) return { tipo: 'transf', formato: 'sap', hoja: hojas[i] };
    if (tiene(cols, 'Articulo') && tiene(cols, 'Doc.Interno')) return { tipo: 'transf', formato: 'sap', hoja: hojas[i] };
  }
  return { tipo: null, formato: null, error: 'No reconocí el archivo. Columnas esperadas por formato:\n' + COLUMNAS_ESPERADAS };
}

/* ---------------- parseo por formato ---------------- */
function parseMerma(filas) {
  var out = [];
  filas.forEach(function (r) {
    if (!str(r['Centro']) && !str(r['Material'])) return;
    out.push({
      centro: str(r['Centro']),
      fechaEntrada: parseFecha(r['Fecha de entrada']),
      docMaterial: str(r['Documento material']),
      material: str(r['Material']),
      texto: str(r['Texto breve material']),
      cantidad: num(r['Ctd.en UM entrada']),
      costoUnit: num(r['Costo Unitario']),
      costoTotal: num(r['Costo Total']),
      usuario: str(r['Nombre del usuario'])
    });
  });
  return out;
}

function parseTransfDestino(filas) {
  var out = [];
  filas.forEach(function (r) {
    if (!str(r['Doc.Interno'])) return;
    var desc = r['Descripción'];
    var origenDesc = '', destinoDesc = '';
    if (Array.isArray(desc)) { origenDesc = str(desc[0]); destinoDesc = str(desc[1]); }
    else { origenDesc = str(desc); destinoDesc = str(r['Descripción_1']); /* columna duplicada: SheetJS la nombra Descripción_1 */ }
    out.push({
      doc: str(r['Doc.Interno']),
      fechaEmb: parseFecha(r['Fecha Embarque']),
      origen: str(r['Tda.Origen']),
      origenDesc: origenDesc,
      destino: str(r['Destino']),
      destinoDesc: destinoDesc,
      piezasEmb: num(r['Piezas Embarcadas']),
      piezasRec: num(r['Piezas Recibidas']),
      difFinal: num(r['Dif. Final']),
      costoTotal: num(r['Costo Total']),
      dias: num(r['Dias Transcurridos'])
    });
  });
  return out;
}

function parseTransfSAP(wb) {
  var filas = [];
  if (wb.SheetNames.indexOf('REPORTE') >= 0) {
    filasDeHoja(wb, 'REPORTE').forEach(function (r) {
      if (!str(r['DOCUMENTO CONTABLE'])) return;
      filas.push({
        doc: str(r['DOCUMENTO CONTABLE']),
        fechaCont: parseFecha(r['FECHA CONTABLE']),
        origen: str(r['SUCURSAL ORIGEN']),
        origenDesc: str(r['DESCRIPCION SUC. ORIGEN']),
        destino: str(r['SUCURSAL DESTINO']),
        destinoDesc: str(r['DESCRIPCION SUC. DESTINO']),
        importe: num(r['IMPORTE']),
        dias: num(r['DIAS TRANSCURRIDOS']),
        anio: str(r['ANIO']),
        mes: str(r['MES']),
        estatus: str(r['ESTATUS'])
      });
    });
  }
  if (!filas.length) {
    // Fallback: Hoja1 con detalle por artículo, agrupar por documento
    var det = filasDeHoja(wb, 'Hoja1');
    var porDoc = {};
    det.forEach(function (r) {
      var doc = str(r['Doc.Interno']);
      if (!doc) return;
      if (!porDoc[doc]) {
        porDoc[doc] = {
          doc: doc, fechaCont: parseFecha(r['Fecha Embarque']),
          origen: str(r['Tda.Origen']), origenDesc: str(r['Descripción']),
          destino: str(r['Destino']), destinoDesc: '',
          importe: 0, dias: num(r['Dias Transcurridos']), anio: '', mes: '', estatus: 'PENDIENTE'
        };
      }
      porDoc[doc].importe += num(r['Costo']);
      if (num(r['Dias Transcurridos']) > porDoc[doc].dias) porDoc[doc].dias = num(r['Dias Transcurridos']);
    });
    Object.keys(porDoc).forEach(function (k) { filas.push(porDoc[k]); });
  }
  return filas;
}

function parseBloqConcentrado(wb) {
  var filas = [];
  filasDeHoja(wb, 'CONCENTRADO').forEach(function (r) {
    if (!str(r['Centro']) && !str(r['Material'])) return;
    filas.push({
      centro: str(r['Centro']),
      material: str(r['Material']),
      ean: str(r['EAN/UPC']),
      denom: str(r['Denomin.']),
      unidad: str(r['Unidad']),
      jerarquia: str(r['Jerarquía 2']),
      almacen: str(r['Almacén']),
      bloqueado: num(r['Bloqueado']),
      valor: num(r['V. Stock bloq.'])
    });
  });
  var resumen = null;
  if (wb.SheetNames.indexOf('RESUMEN') >= 0) {
    var tot = 0, n = 0;
    filasDeHoja(wb, 'RESUMEN').forEach(function (r) {
      if (str(r['Centro'])) { tot += num(r['Importe bloqueado']); n++; }
    });
    if (n) resumen = { total: tot, centros: n };
  }
  return { filas: filas, resumen: resumen };
}

function parseBloqOriginal(filas) {
  var out = [];
  filas.forEach(function (r) {
    if (!str(r['Material']) && !str(r['Centro'])) return;
    out.push({
      zona: str(r['ZONA']),
      material: str(r['Material']),
      centro: str(r['Centro']),
      sucursal: str(r['Sucursal']),
      almacen: str(r['Almacen']),
      libreUtil: num(r['Libre Util']),
      stockBloq: num(r['Stock Bloq']),
      cto: num(r['Cto']),
      importe: num(r['Importe']),
      formato: str(r['FORMATO']),
      descripcion: str(r['Descripción']),
      departamento: str(r['Departamento']),
      clase: str(r['Clase']),
      descClase: str(r['Descr. Clase'])
    });
  });
  return out;
}

/* Fecha de corte sugerida: fecha en el nombre del archivo, si no la fecha máxima del archivo, si no hoy */
function sugerirCorte(nombre, fechasISO) {
  var m = nombre.match(/(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{4})/);
  if (m) return m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]);
  m = nombre.match(/(\d{4})[.\-_](\d{1,2})[.\-_](\d{1,2})/);
  if (m) return m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]);
  var validas = (fechasISO || []).filter(function (f) { return !!f; }).sort();
  if (validas.length) return validas[validas.length - 1];
  return hoyISO();
}

/* ---------------- cálculo de paneles ---------------- */
function nivelMerma(dias, monto) {
  if (dias !== null && dias > 30) return 'CRITICO';
  if (monto >= 4000) return 'CRITICO';
  if ((dias !== null && dias >= 8) || monto >= 500) return 'SEGUIMIENTO';
  return 'CONTROLADO';
}
function accionNivel(nivel) {
  if (nivel === 'CRITICO') return 'Liberación inmediata y escalamiento';
  if (nivel === 'SEGUIMIENTO') return 'Seguimiento diario';
  return 'Mantener control';
}

function computeMerma(filas, fechaCorte) {
  var total = 0, porCentro = {}, ant = {}, evol = {};
  RANGOS.forEach(function (r) { ant[r] = { n: 0, monto: 0 }; });
  filas.forEach(function (f) {
    total += f.costoTotal;
    var dias = diasEntre(f.fechaEntrada, fechaCorte);
    var rg = rangoDe(dias);
    ant[rg].n++; ant[rg].monto += f.costoTotal;
    if (f.fechaEntrada) evol[f.fechaEntrada] = (evol[f.fechaEntrada] || 0) + f.costoTotal;
    var c = porCentro[f.centro] || { centro: f.centro, monto: 0, registros: 0, entradaMasAntigua: f.fechaEntrada };
    c.monto += f.costoTotal; c.registros++;
    if (f.fechaEntrada && (!c.entradaMasAntigua || f.fechaEntrada < c.entradaMasAntigua)) c.entradaMasAntigua = f.fechaEntrada;
    porCentro[f.centro] = c;
  });
  var lista = Object.keys(porCentro).map(function (k) {
    var c = porCentro[k];
    c.dias = diasEntre(c.entradaMasAntigua, fechaCorte);
    c.nivel = nivelMerma(c.dias, c.monto);
    c.accion = accionNivel(c.nivel);
    c.pct = total ? (c.monto / total * 100) : 0;
    return c;
  }).sort(function (a, b) { return b.monto - a.monto; });
  var evolucion = Object.keys(evol).sort().map(function (f) { return { fecha: f, monto: evol[f] }; });
  return { total: total, registros: filas.length, porCentro: lista, antiguedad: ant, evolucion: evolucion };
}

function computeTransfDestino(filas) {
  var docs = filas.length, piezasEmb = 0, piezasPend = 0, costoTotal = 0;
  var porOrigen = {}, porDestino = {}, ant = {}, masAntiguo = null;
  RANGOS.forEach(function (r) { ant[r] = { n: 0, piezas: 0 }; });
  filas.forEach(function (f) {
    piezasEmb += f.piezasEmb; piezasPend += f.difFinal; costoTotal += f.costoTotal;
    var rg = rangoDe(f.dias);
    ant[rg].n++; ant[rg].piezas += f.difFinal;
    var o = porOrigen[f.origen] || { clave: f.origen, nombre: f.origenDesc, docs: 0, piezas: 0 };
    o.docs++; o.piezas += f.difFinal; porOrigen[f.origen] = o;
    var d = porDestino[f.destino] || { clave: f.destino, nombre: f.destinoDesc, docs: 0, piezas: 0 };
    d.docs++; d.piezas += f.difFinal; porDestino[f.destino] = d;
    if (!masAntiguo || f.dias > masAntiguo.dias) masAntiguo = f;
  });
  function aLista(m) { return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return b.piezas - a.piezas; }); }
  return { documentos: docs, piezasEmb: piezasEmb, piezasPend: piezasPend, costoTotal: costoTotal,
    porOrigen: aLista(porOrigen), porDestino: aLista(porDestino), antiguedad: ant, masAntiguo: masAntiguo };
}

function computeTransfSAP(filas) {
  var total = 0, porOrigen = {}, ant = {}, masAntiguo = null;
  RANGOS.forEach(function (r) { ant[r] = { n: 0, monto: 0 }; });
  filas.forEach(function (f) {
    total += f.importe;
    var rg = rangoDe(f.dias);
    ant[rg].n++; ant[rg].monto += f.importe;
    var o = porOrigen[f.origen] || { clave: f.origen, nombre: f.origenDesc, docs: 0, importe: 0 };
    o.docs++; o.importe += f.importe; porOrigen[f.origen] = o;
    if (!masAntiguo || f.dias > masAntiguo.dias) masAntiguo = f;
  });
  var lista = Object.keys(porOrigen).map(function (k) { return porOrigen[k]; }).sort(function (a, b) { return b.importe - a.importe; });
  return { documentos: filas.length, importeTotal: total, porOrigen: lista, antiguedad: ant, masAntiguo: masAntiguo };
}

function computeBloqConc(filas, resumenHoja) {
  var total = 0, unidades = 0, porCentro = {}, porJer = {};
  filas.forEach(function (f) {
    total += f.valor; unidades += f.bloqueado;
    var c = porCentro[f.centro] || { clave: f.centro, registros: 0, unidades: 0, valor: 0 };
    c.registros++; c.unidades += f.bloqueado; c.valor += f.valor; porCentro[f.centro] = c;
    var j = porJer[f.jerarquia || 'Sin jerarquía'] || { clave: f.jerarquia || 'Sin jerarquía', registros: 0, valor: 0 };
    j.registros++; j.valor += f.valor; porJer[f.jerarquia || 'Sin jerarquía'] = j;
  });
  var top = filas.slice().sort(function (a, b) { return Math.abs(b.valor) - Math.abs(a.valor); }).slice(0, 10);
  var advertencia = null;
  if (resumenHoja && Math.abs(resumenHoja.total - total) > 0.01) {
    advertencia = 'La hoja RESUMEN suma ' + fmtDinero(resumenHoja.total) + ' y el concentrado suma ' + fmtDinero(total) + '. Revisa el archivo.';
  }
  function aLista(m) { return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return b.valor - a.valor; }); }
  return { total: total, registros: filas.length, unidades: unidades,
    porCentro: aLista(porCentro), porJerarquia: aLista(porJer), top: top, advertencia: advertencia };
}

function computeBloqOrig(filas) {
  var neto = 0, impacto = 0, pos = 0, neg = 0, posN = 0, negN = 0;
  var porSuc = {}, porDepto = {};
  filas.forEach(function (f) {
    neto += f.importe; impacto += Math.abs(f.importe);
    if (f.stockBloq > 0) { pos += f.stockBloq; posN++; }
    if (f.stockBloq < 0) { neg += f.stockBloq; negN++; }
    var s = porSuc[f.centro] || { clave: f.centro, nombre: f.sucursal, registros: 0, importe: 0, impacto: 0 };
    s.registros++; s.importe += f.importe; s.impacto += Math.abs(f.importe); porSuc[f.centro] = s;
    var d = porDepto[f.departamento || 'Sin departamento'] || { clave: f.departamento || 'Sin departamento', registros: 0, importe: 0 };
    d.registros++; d.importe += f.importe; porDepto[f.departamento || 'Sin departamento'] = d;
  });
  var top = filas.slice().sort(function (a, b) { return Math.abs(b.importe) - Math.abs(a.importe); }).slice(0, 10);
  function aLista(m, key) { return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return b[key] - a[key]; }); }
  return { neto: neto, impacto: impacto, registros: filas.length,
    stockPos: pos, stockPosN: posN, stockNeg: neg, stockNegN: negN,
    porSucursal: aLista(porSuc, 'impacto'), porDepartamento: aLista(porDepto, 'importe'), top: top };
}

/* Exportación para Node (verificación) */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseFecha: parseFecha, num: num, detectarTipo: detectarTipo,
    parseMerma: parseMerma, parseTransfDestino: parseTransfDestino,
    parseTransfSAP: parseTransfSAP, parseBloqConcentrado: parseBloqConcentrado,
    parseBloqOriginal: parseBloqOriginal, sugerirCorte: sugerirCorte,
    computeMerma: computeMerma, computeTransfDestino: computeTransfDestino,
    computeTransfSAP: computeTransfSAP, computeBloqConc: computeBloqConc,
    computeBloqOrig: computeBloqOrig, filasDeHoja: filasDeHoja
  };
}

/* ================= UI (solo navegador) ================= */
if (typeof window !== 'undefined' && typeof document !== 'undefined') {

var state = { cargas: [], seguimiento: {} };
var pendientes = [];
var tabActual = 'merma';

function cargarEstado() {
  try {
    var raw = localStorage.getItem(LS_KEY);
    if (raw) { var s = JSON.parse(raw); if (s && s.cargas) state = s; }
  } catch (e) { state = { cargas: [], seguimiento: {} }; }
  if (!state.seguimiento) state.seguimiento = {};
  if (!state.cargas) state.cargas = [];
}
function guardarEstado() {
  try { localStorage.setItem(LS_KEY, JSON.stringify(state)); }
  catch (e) { toast('No se pudo guardar en este dispositivo (espacio lleno). Borra cortes viejos en Cargas.'); }
}
function toast(msg) {
  var t = document.getElementById('toast');
  t.textContent = msg; t.classList.remove('hidden');
  setTimeout(function () { t.classList.add('hidden'); }, 4000);
}
function $ (id) { return document.getElementById(id); }
function ultimaCarga(tipo) {
  var cs = state.cargas.filter(function (c) { return c.tipo === tipo; });
  cs.sort(function (a, b) { return (a.fechaCorte < b.fechaCorte) ? 1 : -1; });
  return cs.length ? cs[0] : null;
}
function cortesDe(tipo) {
  return state.cargas.filter(function (c) { return c.tipo === tipo; })
    .sort(function (a, b) { return (a.fechaCorte < b.fechaCorte) ? -1 : 1; });
}
function nombreTipo(tipo) {
  return tipo === 'merma' ? 'Merma' : tipo === 'transf' ? 'Transferencias' : 'Bloqueado';
}

/* ---- gráficas simples ---- */
function dibujarBarras(canvas, items, fmt) {
  if (!canvas) return;
  var dpr = window.devicePixelRatio || 1;
  var W = canvas.clientWidth || 320, H = 180;
  canvas.width = W * dpr; canvas.height = H * dpr;
  var ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.clearRect(0, 0, W, H);
  if (!items.length) { ctx.fillStyle = '#5a6b7c'; ctx.font = '13px sans-serif'; ctx.fillText('Sin datos', 10, 30); return; }
  var max = Math.max.apply(null, items.map(function (i) { return i.valor; })) || 1;
  var bw = Math.min(48, (W - 20) / items.length - 10);
  items.forEach(function (it, i) {
    var h = Math.max(2, (it.valor / max) * (H - 52));
    var x = 10 + i * ((W - 20) / items.length) + ((W - 20) / items.length - bw) / 2;
    ctx.fillStyle = '#0b2a4a';
    ctx.fillRect(x, H - 30 - h, bw, h);
    ctx.fillStyle = '#1c2b3a'; ctx.font = '11px sans-serif'; ctx.textAlign = 'center';
    ctx.fillText(fmt ? fmt(it.valor) : fmtNum(it.valor), x + bw / 2, H - 34 - h);
    ctx.fillText(it.etiqueta, x + bw / 2, H - 12);
  });
}
function donutHTML(ant, total, esPiezas) {
  var acc = 0, partes = [], leyenda = '';
  RANGOS.forEach(function (r, i) {
    var v = ant[r];
    var val = esPiezas ? v.piezas : v.monto;
    var pct = total ? (val / total * 100) : 0;
    var ini = acc; acc += pct;
    partes.push(COLRES_RANGO_SAFE(i) + ' ' + ini.toFixed(1) + '% ' + acc.toFixed(1) + '%');
    leyenda += '<div><span class="punto" style="background:' + COLORES_RANGO[i] + '"></span>' + r + ' días: ' +
      (esPiezas ? fmtNum(val) + ' piezas' : fmtDinero(val)) + ' (' + pct.toFixed(1) + '%)</div>';
  });
  var bg = partes.length ? 'conic-gradient(' + partes.join(', ') + ')' : '#e8eef4';
  return '<div class="donut-wrap"><div class="donut" style="background:' + bg + '"></div><div class="leyenda">' + leyenda + '</div></div>';
}
function COLRES_RANGO_SAFE(i) { return COLORES_RANGO[i]; }
function barrasHTML(titulo, items, fmt) {
  if (!items.length) return '';
  var max = Math.max.apply(null, items.map(function (i) { return i.valor; })) || 1;
  var html = '<div class="grafica"><h3>' + esc(titulo) + '</h3>';
  items.forEach(function (it) {
    var pct = (it.valor / max * 100).toFixed(1);
    html += '<div class="barra-fila"><span class="blabel">' + esc(it.etiqueta) + '</span>' +
      '<span class="btrack"><span class="bfill" style="display:block;width:' + pct + '%"></span></span>' +
      '<span class="bvalor">' + (fmt ? fmt(it.valor) : fmtNum(it.valor)) + '</span></div>';
  });
  return html + '</div>';
}
function tablaHTML(headers, rowsHTML, totalRowHTML) {
  var h = headers.map(function (x) { return '<th class="' + (x.num ? 'num' : '') + '">' + esc(x.t) + '</th>'; }).join('');
  return '<div class="tabla-scroll"><table><thead><tr>' + h + '</tr></thead><tbody>' + rowsHTML +
    (totalRowHTML ? '<tr class="total-row">' + totalRowHTML + '</tr>' : '') + '</tbody></table></div>';
}
function segControl(tipo, clave, estadoActual, notaActual) {
  var s = state.seguimiento[clave] || {};
  var est = s.estado || estadoActual || 'Pendiente';
  var nota = (s.nota !== undefined) ? s.nota : (notaActual || '');
  var opts = ['Pendiente', 'En seguimiento', 'Liberado', 'Escalado'].map(function (o) {
    return '<option value="' + o + '"' + (o === est ? ' selected' : '') + '>' + o + '</option>';
  }).join('');
  return '<select class="estado" data-seg="' + esc(clave) + '" data-campo="estado">' + opts + '</select>' +
    '<br><input class="nota" data-seg="' + esc(clave) + '" data-campo="nota" placeholder="Nota…" value="' + esc(nota) + '">';
}
function variacionHTML(actual, anterior, invertida) {
  if (anterior === null || anterior === undefined) return '<div class="variacion igual">Primer corte</div>';
  var dif = actual - anterior;
  if (Math.abs(dif) < 0.005) return '<div class="variacion igual">Sin cambio vs corte anterior</div>';
  var pct = anterior ? (dif / Math.abs(anterior) * 100) : 0;
  var baja = dif < 0;
  var buena = invertida ? baja : !baja;
  var cls = buena ? 'baja' : 'sube';
  return '<div class="variacion ' + cls + '">' + (baja ? '▼' : '▲') + ' ' + fmtDinero(Math.abs(dif)) +
    ' (' + Math.abs(pct).toFixed(1) + '%) vs corte anterior</div>';
}

/* ---- render de paneles ---- */
function renderMerma() {
  var c = ultimaCarga('merma');
  var el = $('merma-panel'), info = $('merma-corte');
  if (!c) { info.textContent = ''; el.innerHTML = '<p class="vacio">Aún no hay merma cargada. Sube el archivo de merma arriba.</p>'; return; }
  var p = computeMerma(c.filas, c.fechaCorte);
  var cortes = cortesDe('merma');
  var ant = cortes.length > 1 ? computeMerma(cortes[cortes.length - 2].filas, cortes[cortes.length - 2].fechaCorte).total : null;
  info.textContent = 'Corte: ' + fmtFecha(c.fechaCorte) + ' · Archivo: ' + c.nombre + ' · ' + c.filas.length + ' registros';
  var html = '<div class="totales">' +
    '<div class="total-card destacado"><div class="etiqueta">Monto pendiente</div><div class="valor">' + fmtDinero(p.total) + '</div>' + variacionHTML(p.total, ant, true) + '</div>' +
    '<div class="total-card"><div class="etiqueta">Registros</div><div class="valor">' + p.registros + '</div></div>' +
    '<div class="total-card"><div class="etiqueta">Sucursales</div><div class="valor">' + p.porCentro.length + '</div></div>' +
    '<div class="total-card"><div class="etiqueta">Críticos</div><div class="valor">' + p.porCentro.filter(function (x) { return x.nivel === 'CRITICO'; }).length + '</div></div></div>';
  html += '<div class="grafica"><h3>Evolución por fecha de entrada</h3><canvas id="cv-merma"></canvas></div>';
  html += '<div class="grafica"><h3>Distribución por antigüedad</h3>' + donutHTML(p.antiguedad, p.total, false) + '</div>';
  html += barrasHTML('Top sucursales por monto', p.porCentro.slice(0, 6).map(function (x) { return { etiqueta: x.centro, valor: x.monto }; }), fmtDinero);
  var rows = p.porCentro.map(function (x) {
    var clave = 'merma:' + x.centro;
    return '<tr><td>' + esc(x.centro) + '</td><td class="num">' + fmtDinero(x.monto) + '</td><td class="num">' + x.pct.toFixed(1) + '%</td>' +
      '<td class="num">' + x.registros + '</td><td>' + fmtFecha(x.entradaMasAntigua) + '</td><td class="num">' + (x.dias === null ? '—' : x.dias) + '</td>' +
      '<td><span class="nivel ' + x.nivel.toLowerCase() + '">' + x.nivel + '</span></td><td>' + esc(x.accion) + '</td>' +
      '<td>' + segControl('merma', clave) + '</td></tr>';
  }).join('');
  html += '<h3>Detalle por sucursal</h3>' + tablaHTML(
    [{ t: 'Sucursal' }, { t: 'Monto', num: true }, { t: '%', num: true }, { t: 'Reg.', num: true }, { t: 'Entrada más antigua' }, { t: 'Días', num: true }, { t: 'Nivel' }, { t: 'Acción' }, { t: 'Seguimiento' }],
    rows, '<td>Total</td><td class="num">' + fmtDinero(p.total) + '</td><td class="num">100%</td><td class="num">' + p.registros + '</td><td></td><td></td><td></td><td></td><td></td>');
  el.innerHTML = html;
  dibujarBarras($('cv-merma'), p.evolucion.map(function (e) { return { etiqueta: e.fecha.slice(5), valor: e.monto }; }), fmtDinero);
  var ev = cortes.map(function (cc) { return { fecha: cc.fechaCorte, monto: computeMerma(cc.filas, cc.fechaCorte).total }; });
  if (ev.length > 1) {
    el.insertAdjacentHTML('beforeend', '<div class="grafica"><h3>Evolución por corte</h3><canvas id="cv-merma-cortes"></canvas></div>');
    dibujarBarras($('cv-merma-cortes'), ev.map(function (e) { return { etiqueta: e.fecha.slice(5), valor: e.monto }; }), fmtDinero);
  }
}

function renderTransf() {
  var c = ultimaCarga('transf');
  var el = $('transf-panel'), info = $('transf-corte');
  if (!c) { info.textContent = ''; el.innerHTML = '<p class="vacio">Aún no hay transferencias cargadas. Sube el archivo de transferencias arriba.</p>'; return; }
  info.textContent = 'Corte: ' + fmtFecha(c.fechaCorte) + ' · Archivo: ' + c.nombre + ' · Formato: ' + (c.formato === 'destino' ? 'Destino' : 'SAP');
  var html = '';
  if (c.formato === 'destino') {
    var p = computeTransfDestino(c.filas);
    html += '<div class="totales">' +
      '<div class="total-card destacado"><div class="etiqueta">Documentos pendientes</div><div class="valor">' + p.documentos + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Piezas embarcadas</div><div class="valor">' + fmtNum(p.piezasEmb) + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Piezas pendientes</div><div class="valor">' + fmtNum(p.piezasPend) + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Importe</div><div class="valor">' + (p.costoTotal ? fmtDinero(p.costoTotal) : '—') + '</div><div class="sub">' + (p.costoTotal ? '' : 'El archivo trae Costo Total en 0') + '</div></div></div>';
    html += '<div class="grafica"><h3>Distribución por antigüedad (piezas pendientes)</h3>' + donutHTML(p.antiguedad, p.piezasPend, true) + '</div>';
    html += barrasHTML('Pendiente por sucursal origen (piezas)', p.porOrigen.slice(0, 6).map(function (x) { return { etiqueta: x.clave, valor: x.piezas }; }), fmtNum);
    html += barrasHTML('Pendiente por sucursal destino (piezas)', p.porDestino.slice(0, 6).map(function (x) { return { etiqueta: x.clave, valor: x.piezas }; }), fmtNum);
    var rows = c.filas.slice().sort(function (a, b) { return b.dias - a.dias; }).map(function (f) {
      var clave = 'transf:' + f.doc;
      return '<tr><td>' + esc(f.doc) + '</td><td>' + fmtFecha(f.fechaEmb) + '</td><td>' + esc(f.origen) + ' ' + esc(f.origenDesc) + '</td>' +
        '<td>' + esc(f.destino) + ' ' + esc(f.destinoDesc) + '</td><td class="num">' + fmtNum(f.piezasEmb) + '</td><td class="num">' + fmtNum(f.difFinal) + '</td>' +
        '<td class="num">' + f.dias + '</td><td>' + segControl('transf', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Detalle por documento</h3>' + tablaHTML(
      [{ t: 'Documento' }, { t: 'Fecha embarque' }, { t: 'Origen' }, { t: 'Destino' }, { t: 'Piezas emb.', num: true }, { t: 'Piezas pend.', num: true }, { t: 'Días', num: true }, { t: 'Seguimiento' }],
      rows, '<td>Total</td><td></td><td></td><td></td><td class="num">' + fmtNum(p.piezasEmb) + '</td><td class="num">' + fmtNum(p.piezasPend) + '</td><td></td><td></td>');
    if (p.masAntiguo) html += '<p class="hint">Documento más antiguo: ' + esc(p.masAntiguo.doc) + ' de ' + esc(p.masAntiguo.origenDesc || p.masAntiguo.origen) + ' (' + p.masAntiguo.dias + ' días).</p>';
  } else {
    var q = computeTransfSAP(c.filas);
    html += '<div class="totales">' +
      '<div class="total-card destacado"><div class="etiqueta">Importe pendiente</div><div class="valor">' + fmtDinero(q.importeTotal) + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Documentos</div><div class="valor">' + q.documentos + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Sucursales origen</div><div class="valor">' + q.porOrigen.length + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Más antiguo</div><div class="valor">' + (q.masAntiguo ? q.masAntiguo.dias + ' días' : '—') + '</div></div></div>';
    html += '<div class="grafica"><h3>Distribución por antigüedad</h3>' + donutHTML(q.antiguedad, q.importeTotal, false) + '</div>';
    html += barrasHTML('Pendiente por sucursal origen', q.porOrigen.slice(0, 6).map(function (x) { return { etiqueta: x.clave, valor: x.importe }; }), fmtDinero);
    var rows2 = c.filas.slice().sort(function (a, b) { return b.dias - a.dias; }).map(function (f) {
      var clave = 'transf:' + f.doc;
      return '<tr><td>' + esc(f.doc) + '</td><td>' + fmtFecha(f.fechaCont) + '</td><td>' + esc(f.origen) + ' ' + esc(f.origenDesc) + '</td>' +
        '<td>' + esc(f.destino) + ' ' + esc(f.destinoDesc) + '</td><td class="num">' + fmtDinero(f.importe) + '</td>' +
        '<td class="num">' + f.dias + '</td><td>' + esc(f.estatus || '') + '</td><td>' + segControl('transf', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Detalle por documento</h3>' + tablaHTML(
      [{ t: 'Documento' }, { t: 'Fecha contable' }, { t: 'Origen' }, { t: 'Destino' }, { t: 'Importe', num: true }, { t: 'Días', num: true }, { t: 'Estatus' }, { t: 'Seguimiento' }],
      rows2, '<td>Total</td><td></td><td></td><td></td><td class="num">' + fmtDinero(q.importeTotal) + '</td><td></td><td></td><td></td>');
  }
  el.innerHTML = html;
}

function renderBloq() {
  var c = ultimaCarga('bloq');
  var el = $('bloq-panel'), info = $('bloq-corte');
  if (!c) { info.textContent = ''; el.innerHTML = '<p class="vacio">Aún no hay stock bloqueado cargado. Sube el archivo de bloqueados arriba.</p>'; return; }
  info.textContent = 'Corte: ' + fmtFecha(c.fechaCorte) + ' · Archivo: ' + c.nombre + ' · Formato: ' + (c.formato === 'concentrado' ? 'Concentrado' : 'Bloq original');
  var html = '';
  if (c.advertencia) html += '<div class="alerta">' + esc(c.advertencia) + '</div>';
  if (c.formato === 'concentrado') {
    var p = computeBloqConc(c.filas, c.resumenHoja);
    html += '<div class="totales">' +
      '<div class="total-card destacado"><div class="etiqueta">Importe bloqueado</div><div class="valor">' + fmtDinero(p.total) + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Registros</div><div class="valor">' + p.registros + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Unidades bloqueadas</div><div class="valor">' + fmtNum(p.unidades, 3) + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Centros</div><div class="valor">' + p.porCentro.length + '</div></div></div>';
    html += barrasHTML('Importe por centro', p.porCentro.map(function (x) { return { etiqueta: x.clave, valor: x.valor }; }), fmtDinero);
    html += barrasHTML('Por Jerarquía 2', p.porJerarquia.map(function (x) { return { etiqueta: x.clave, valor: x.valor }; }), fmtDinero);
    var rows = p.porCentro.map(function (x) {
      var clave = 'bloq:centro:' + x.clave;
      return '<tr><td>' + esc(x.clave) + '</td><td class="num">' + x.registros + '</td><td class="num">' + fmtNum(x.unidades, 3) + '</td><td class="num">' + fmtDinero(x.valor) + '</td><td>' + segControl('bloq', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Detalle por centro</h3>' + tablaHTML(
      [{ t: 'Centro' }, { t: 'Registros', num: true }, { t: 'Unidades', num: true }, { t: 'Importe', num: true }, { t: 'Seguimiento' }],
      rows, '<td>Total</td><td class="num">' + p.registros + '</td><td class="num">' + fmtNum(p.unidades, 3) + '</td><td class="num">' + fmtDinero(p.total) + '</td><td></td>');
    var topRows = p.top.map(function (f) {
      var clave = 'bloq:' + f.centro + ':' + f.material;
      return '<tr><td>' + esc(f.centro) + '</td><td>' + esc(f.material) + '</td><td>' + esc(f.denom) + '</td><td>' + esc(f.jerarquia) + '</td><td class="num">' + fmtNum(f.bloqueado, 3) + '</td><td class="num">' + fmtDinero(f.valor) + '</td><td>' + segControl('bloq', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Artículos más críticos</h3>' + tablaHTML(
      [{ t: 'Centro' }, { t: 'Material' }, { t: 'Denominación' }, { t: 'Jerarquía' }, { t: 'Bloqueado', num: true }, { t: 'Valor', num: true }, { t: 'Seguimiento' }], topRows, '');
  } else {
    var q = computeBloqOrig(c.filas);
    html += '<div class="totales">' +
      '<div class="total-card destacado"><div class="etiqueta">Impacto absoluto</div><div class="valor">' + fmtDinero(q.impacto) + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Importe neto</div><div class="valor">' + fmtDinero(q.neto) + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Registros</div><div class="valor">' + q.registros + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Stock + / −</div><div class="valor" style="font-size:17px">' + fmtNum(q.stockPos, 1) + ' / ' + fmtNum(q.stockNeg, 1) + '</div><div class="sub">' + q.stockPosN + ' positivos · ' + q.stockNegN + ' negativos</div></div></div>';
    html += barrasHTML('Impacto por sucursal', q.porSucursal.slice(0, 6).map(function (x) { return { etiqueta: x.clave, valor: x.impacto }; }), fmtDinero);
    html += barrasHTML('Por departamento', q.porDepartamento.slice(0, 6).map(function (x) { return { etiqueta: x.clave, valor: Math.abs(x.importe) }; }), fmtDinero);
    var rows2 = q.porSucursal.map(function (x) {
      var clave = 'bloq:centro:' + x.clave;
      return '<tr><td>' + esc(x.clave) + ' ' + esc(x.nombre) + '</td><td class="num">' + x.registros + '</td><td class="num">' + fmtDinero(x.importe) + '</td><td class="num">' + fmtDinero(x.impacto) + '</td><td>' + segControl('bloq', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Detalle por sucursal</h3>' + tablaHTML(
      [{ t: 'Sucursal' }, { t: 'Registros', num: true }, { t: 'Importe neto', num: true }, { t: 'Impacto', num: true }, { t: 'Seguimiento' }],
      rows2, '<td>Total</td><td class="num">' + q.registros + '</td><td class="num">' + fmtDinero(q.neto) + '</td><td class="num">' + fmtDinero(q.impacto) + '</td><td></td>');
    var topRows2 = q.top.map(function (f) {
      var clave = 'bloq:' + f.centro + ':' + f.material;
      return '<tr><td>' + esc(f.centro) + '</td><td>' + esc(f.material) + '</td><td>' + esc(f.descripcion) + '</td><td>' + esc(f.departamento) + '</td><td class="num">' + fmtNum(f.stockBloq, 2) + '</td><td class="num">' + fmtDinero(f.importe) + '</td><td>' + segControl('bloq', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Artículos más críticos</h3>' + tablaHTML(
      [{ t: 'Centro' }, { t: 'Material' }, { t: 'Descripción' }, { t: 'Departamento' }, { t: 'Stock bloq.', num: true }, { t: 'Importe', num: true }, { t: 'Seguimiento' }], topRows2, '');
  }
  el.innerHTML = html;
}

function renderCargas() {
  var el = $('cargas-lista');
  if (!state.cargas.length) { el.innerHTML = '<p class="vacio">No hay cargas todavía.</p>'; return; }
  var cs = state.cargas.slice().sort(function (a, b) { return (a.fechaCorte < b.fechaCorte) ? 1 : -1; });
  el.innerHTML = cs.map(function (c) {
    return '<div class="carga"><h3>' + nombreTipo(c.tipo) + ' — corte ' + fmtFecha(c.fechaCorte) + '</h3>' +
      '<p>Archivo: ' + esc(c.nombre) + '</p>' +
      '<p>Formato: ' + esc(c.formato) + ' · ' + c.filas.length + ' registros · Cargado: ' + esc(new Date(c.fechaCarga).toLocaleString('es-MX')) + '</p>' +
      '<div class="acciones"><button class="btn btn-danger" data-borrar="' + c.id + '" type="button">Borrar corte</button></div></div>';
  }).join('');
}

function renderTodo() { renderMerma(); renderTransf(); renderBloq(); renderCargas(); }

/* ---- subida e importación ---- */
function leerArchivo(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        var det = detectarTipo(wb);
        if (!det.tipo) { resolve({ nombre: file.name, error: det.error }); return; }
        var filas = [], fechas = [], advertencia = null, resumenHoja = null;
        if (det.tipo === 'merma') {
          filas = parseMerma(filasDeHoja(wb, det.hoja));
          fechas = filas.map(function (f) { return f.fechaEntrada; });
        } else if (det.tipo === 'transf' && det.formato === 'destino') {
          filas = parseTransfDestino(filasDeHoja(wb, det.hoja));
          fechas = filas.map(function (f) { return f.fechaEmb; });
        } else if (det.tipo === 'transf') {
          filas = parseTransfSAP(wb);
          fechas = filas.map(function (f) { return f.fechaCont; });
        } else if (det.tipo === 'bloq' && det.formato === 'concentrado') {
          var pc = parseBloqConcentrado(wb);
          filas = pc.filas; resumenHoja = pc.resumen;
          var tot = filas.reduce(function (a, f) { return a + f.valor; }, 0);
          if (resumenHoja && Math.abs(resumenHoja.total - tot) > 0.01) {
            advertencia = 'La hoja RESUMEN suma ' + fmtDinero(resumenHoja.total) + ' y el concentrado suma ' + fmtDinero(tot) + '. Revisa el archivo.';
          }
        } else {
          filas = parseBloqOriginal(filasDeHoja(wb, det.hoja));
        }
        if (!filas.length) { resolve({ nombre: file.name, error: 'El archivo se reconoció como ' + nombreTipo(det.tipo) + ' (' + det.formato + '), pero no trae renglones con datos.' }); return; }
        resolve({
          id: 'p' + Date.now() + Math.floor(Math.random() * 1000),
          nombre: file.name, tipo: det.tipo, formato: det.formato,
          fechaCorte: sugerirCorte(file.name, fechas),
          filas: filas, resumenHoja: resumenHoja, advertencia: advertencia
        });
      } catch (err) { reject(err); }
    };
    reader.onerror = function () { reject(new Error('No se pudo leer el archivo.')); };
    reader.readAsArrayBuffer(file);
  });
}

function resumenPendiente(p) {
  if (p.tipo === 'merma') return 'Total: ' + fmtDinero(p.filas.reduce(function (a, f) { return a + f.costoTotal; }, 0)) + ' · ' + p.filas.length + ' registros';
  if (p.tipo === 'transf' && p.formato === 'destino') {
    var pe = p.filas.reduce(function (a, f) { return a + f.piezasEmb; }, 0);
    var pp = p.filas.reduce(function (a, f) { return a + f.difFinal; }, 0);
    return p.filas.length + ' documentos · ' + fmtNum(pe) + ' piezas embarcadas · ' + fmtNum(pp) + ' pendientes';
  }
  if (p.tipo === 'transf') return 'Total: ' + fmtDinero(p.filas.reduce(function (a, f) { return a + f.importe; }, 0)) + ' · ' + p.filas.length + ' documentos';
  if (p.formato === 'concentrado') return 'Total: ' + fmtDinero(p.filas.reduce(function (a, f) { return a + f.valor; }, 0)) + ' · ' + p.filas.length + ' registros';
  return 'Impacto: ' + fmtDinero(p.filas.reduce(function (a, f) { return a + Math.abs(f.importe); }, 0)) + ' · ' + p.filas.length + ' registros';
}

function renderPendientes() {
  var el = $('pendientes');
  if (!pendientes.length) { el.innerHTML = ''; return; }
  el.innerHTML = pendientes.map(function (p) {
    if (p.error) {
      return '<div class="pendiente"><h3>' + esc(p.nombre) + '</h3><div class="error">' + esc(p.error).replace(/\n/g, '<br>') + '</div>' +
        '<div class="acciones"><button class="btn" data-quitar="' + p.id + '" type="button">Quitar</button></div></div>';
    }
    return '<div class="pendiente"><h3>' + esc(p.nombre) + '</h3>' +
      '<div class="tipo">Detectado: ' + nombreTipo(p.tipo) + ' · formato ' + esc(p.formato) + '</div>' +
      '<div class="totales">' + esc(resumenPendiente(p)) + '</div>' +
      (p.advertencia ? '<div class="alerta">' + esc(p.advertencia) + '</div>' : '') +
      '<label>Fecha de corte: <input type="date" data-corte="' + p.id + '" value="' + p.fechaCorte + '"></label>' +
      '<div class="acciones"><button class="btn btn-primary" data-importar="' + p.id + '" type="button">Importar como corte</button>' +
      '<button class="btn" data-quitar="' + p.id + '" type="button">Cancelar</button></div></div>';
  }).join('');
}

function importarPendiente(id) {
  var p = pendientes.find(function (x) { return x.id === id; });
  if (!p || p.error) return;
  var input = document.querySelector('[data-corte="' + id + '"]');
  var fechaCorte = (input && input.value) ? input.value : p.fechaCorte;
  state.cargas.push({
    id: 'c' + Date.now() + Math.floor(Math.random() * 1000),
    tipo: p.tipo, formato: p.formato, nombre: p.nombre,
    fechaCorte: fechaCorte, fechaCarga: new Date().toISOString(),
    filas: p.filas, resumenHoja: p.resumenHoja || null, advertencia: p.advertencia || null
  });
  // Máximo 12 cortes por tipo para no llenar el localStorage
  ['merma', 'transf', 'bloq'].forEach(function (t) {
    var cs = state.cargas.filter(function (c) { return c.tipo === t; })
      .sort(function (a, b) { return (a.fechaCorte < b.fechaCorte) ? -1 : 1; });
    if (cs.length > 12) {
      var viejos = cs.slice(12).map(function (c) { return c.id; });
      state.cargas = state.cargas.filter(function (c) { return viejos.indexOf(c.id) < 0; });
    }
  });
  pendientes = pendientes.filter(function (x) { return x.id !== id; });
  guardarEstado(); renderPendientes(); renderTodo();
  toast(nombreTipo(p.tipo) + ' importado: corte ' + fmtFecha(fechaCorte) + '.');
}

async function procesarArchivos(files) {
  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    if (!/\.xlsx?$/i.test(f.name)) { toast('El archivo ' + f.name + ' no es Excel (.xlsx).'); continue; }
    try {
      var p = await leerArchivo(f);
      p.id = p.id || ('p' + Date.now() + i);
      pendientes.push(p);
    } catch (e) { pendientes.push({ id: 'p' + Date.now() + i, nombre: f.name, error: 'No se pudo leer el archivo: ' + e.message }); }
  }
  renderPendientes();
}

/* ---- exportación a Excel ---- */
function panelAOA(tipo, carga) {
  var L = [];
  if (tipo === 'merma') {
    var p = computeMerma(carga.filas, carga.fechaCorte);
    L.push(['MERMA PENDIENTE DE LIBERAR', '', '', '', '', '', '', '']);
    L.push(['Corte', carga.fechaCorte, 'Archivo', carga.nombre]);
    L.push(['Monto total pendiente', p.total, 'Registros', p.registros]);
    L.push([]);
    L.push(['DISTRIBUCIÓN POR ANTIGÜEDAD']);
    L.push(['Rango (días)', 'Registros', 'Monto']);
    RANGOS.forEach(function (r) { L.push([r, p.antiguedad[r].n, p.antiguedad[r].monto]); });
    L.push([]);
    L.push(['DETALLE POR SUCURSAL']);
    L.push(['Sucursal', 'Monto pendiente', '% total', 'Registros', 'Entrada más antigua', 'Días', 'Nivel', 'Acción']);
    p.porCentro.forEach(function (x) { L.push([x.centro, x.monto, x.pct / 100, x.registros, x.entradaMasAntigua, x.dias, x.nivel, x.accion]); });
    L.push(['TOTAL', p.total, 1, p.registros, '', '', '', '']);
  } else if (tipo === 'transf') {
    if (carga.formato === 'destino') {
      var q = computeTransfDestino(carga.filas);
      L.push(['TRANSFERENCIAS PENDIENTES (DESTINO)']);
      L.push(['Corte', carga.fechaCorte, 'Archivo', carga.nombre]);
      L.push(['Documentos', q.documentos, 'Piezas embarcadas', q.piezasEmb, 'Piezas pendientes', q.piezasPend, 'Importe', q.costoTotal]);
      L.push([]);
      L.push(['ANTIGÜEDAD (piezas pendientes)']);
      L.push(['Rango (días)', 'Documentos', 'Piezas']);
      RANGOS.forEach(function (r) { L.push([r, q.antiguedad[r].n, q.antiguedad[r].piezas]); });
      L.push([]);
      L.push(['POR SUCURSAL ORIGEN']);
      L.push(['Origen', 'Nombre', 'Documentos', 'Piezas pendientes']);
      q.porOrigen.forEach(function (x) { L.push([x.clave, x.nombre, x.docs, x.piezas]); });
      L.push([]);
      L.push(['POR SUCURSAL DESTINO']);
      L.push(['Destino', 'Nombre', 'Documentos', 'Piezas pendientes']);
      q.porDestino.forEach(function (x) { L.push([x.clave, x.nombre, x.docs, x.piezas]); });
    } else {
      var s = computeTransfSAP(carga.filas);
      L.push(['TRANSFERENCIAS PENDIENTES (SAP)']);
      L.push(['Corte', carga.fechaCorte, 'Archivo', carga.nombre]);
      L.push(['Importe total', s.importeTotal, 'Documentos', s.documentos]);
      L.push([]);
      L.push(['ANTIGÜEDAD']);
      L.push(['Rango (días)', 'Documentos', 'Importe']);
      RANGOS.forEach(function (r) { L.push([r, s.antiguedad[r].n, s.antiguedad[r].monto]); });
      L.push([]);
      L.push(['POR SUCURSAL ORIGEN']);
      L.push(['Origen', 'Nombre', 'Documentos', 'Importe']);
      s.porOrigen.forEach(function (x) { L.push([x.clave, x.nombre, x.docs, x.importe]); });
    }
  } else {
    if (carga.formato === 'concentrado') {
      var b = computeBloqConc(carga.filas, carga.resumenHoja);
      L.push(['STOCK BLOQUEADO (CONCENTRADO)']);
      L.push(['Corte', carga.fechaCorte, 'Archivo', carga.nombre]);
      L.push(['Importe bloqueado', b.total, 'Registros', b.registros, 'Unidades', b.unidades]);
      L.push([]);
      L.push(['POR CENTRO']);
      L.push(['Centro', 'Registros', 'Unidades', 'Importe']);
      b.porCentro.forEach(function (x) { L.push([x.clave, x.registros, x.unidades, x.valor]); });
      L.push([]);
      L.push(['POR JERARQUÍA 2']);
      L.push(['Jerarquía', 'Registros', 'Importe']);
      b.porJerarquia.forEach(function (x) { L.push([x.clave, x.registros, x.valor]); });
      L.push([]);
      L.push(['TOP ARTÍCULOS']);
      L.push(['Centro', 'Material', 'Denominación', 'Jerarquía', 'Bloqueado', 'Valor']);
      b.top.forEach(function (f) { L.push([f.centro, f.material, f.denom, f.jerarquia, f.bloqueado, f.valor]); });
    } else {
      var o = computeBloqOrig(carga.filas);
      L.push(['STOCK BLOQUEADO']);
      L.push(['Corte', carga.fechaCorte, 'Archivo', carga.nombre]);
      L.push(['Impacto absoluto', o.impacto, 'Importe neto', o.neto, 'Registros', o.registros]);
      L.push([]);
      L.push(['POR SUCURSAL']);
      L.push(['Centro', 'Sucursal', 'Registros', 'Importe neto', 'Impacto']);
      o.porSucursal.forEach(function (x) { L.push([x.clave, x.nombre, x.registros, x.importe, x.impacto]); });
      L.push([]);
      L.push(['POR DEPARTAMENTO']);
      L.push(['Departamento', 'Registros', 'Importe']);
      o.porDepartamento.forEach(function (x) { L.push([x.clave, x.registros, x.importe]); });
      L.push([]);
      L.push(['TOP ARTÍCULOS']);
      L.push(['Centro', 'Material', 'Descripción', 'Departamento', 'Stock bloq.', 'Importe']);
      o.top.forEach(function (f) { L.push([f.centro, f.material, f.descripcion, f.departamento, f.stockBloq, f.importe]); });
    }
  }
  return L;
}
function detalleJSON(tipo, carga) {
  if (tipo === 'merma') return carga.filas.map(function (f) {
    return { 'Centro': f.centro, 'Fecha de entrada': f.fechaEntrada, 'Documento material': f.docMaterial, 'Material': f.material,
      'Texto breve material': f.texto, 'Ctd.en UM entrada': f.cantidad, 'Costo Unitario': f.costoUnit, 'Costo Total': f.costoTotal, 'Usuario': f.usuario };
  });
  if (tipo === 'transf' && carga.formato === 'destino') return carga.filas.map(function (f) {
    return { 'Doc.Interno': f.doc, 'Fecha Embarque': f.fechaEmb, 'Tda.Origen': f.origen, 'Descripción origen': f.origenDesc,
      'Destino': f.destino, 'Descripción destino': f.destinoDesc, 'Piezas Embarcadas': f.piezasEmb, 'Piezas Recibidas': f.piezasRec,
      'Dif. Final': f.difFinal, 'Costo Total': f.costoTotal, 'Dias Transcurridos': f.dias };
  });
  if (tipo === 'transf') return carga.filas.map(function (f) {
    return { 'Documento': f.doc, 'Fecha contable': f.fechaCont, 'Sucursal origen': f.origen, 'Descripción origen': f.origenDesc,
      'Sucursal destino': f.destino, 'Descripción destino': f.destinoDesc, 'Importe': f.importe, 'Días transcurridos': f.dias, 'Estatus': f.estatus };
  });
  if (carga.formato === 'concentrado') return carga.filas.map(function (f) {
    return { 'Centro': f.centro, 'Material': f.material, 'EAN/UPC': f.ean, 'Denomin.': f.denom, 'Unidad': f.unidad,
      'Jerarquía 2': f.jerarquia, 'Almacén': f.almacen, 'Bloqueado': f.bloqueado, 'V. Stock bloq.': f.valor };
  });
  return carga.filas.map(function (f) {
    return { 'Zona': f.zona, 'Material': f.material, 'Centro': f.centro, 'Sucursal': f.sucursal, 'Almacén': f.almacen,
      'Libre Util': f.libreUtil, 'Stock Bloq': f.stockBloq, 'Cto': f.cto, 'Importe': f.importe, 'Formato': f.formato,
      'Descripción': f.descripcion, 'Departamento': f.departamento, 'Clase': f.clase };
  });
}
function exportar(tipo) {
  if (typeof XLSX === 'undefined') { toast('No se pudo cargar la librería de Excel. Revisa tu conexión e inténtalo de nuevo.'); return; }
  var wb = XLSX.utils.book_new();
  var tipos = tipo ? [tipo] : ['merma', 'transf', 'bloq'];
  var resumen = [['PANEL DE SEGUIMIENTO — RESUMEN'], ['Generado', new Date().toLocaleString('es-MX')], [], ['Reporte', 'Corte', 'Total principal', 'Registros']];
  var hubo = false;
  tipos.forEach(function (t) {
    var c = ultimaCarga(t);
    if (!c) return;
    hubo = true;
    var etiqueta = nombreTipo(t);
    var totalTxt = '';
    if (t === 'merma') totalTxt = fmtDinero(computeMerma(c.filas, c.fechaCorte).total);
    if (t === 'transf' && c.formato === 'destino') totalTxt = fmtNum(computeTransfDestino(c.filas).piezasPend) + ' piezas pendientes';
    if (t === 'transf' && c.formato !== 'destino') totalTxt = fmtDinero(computeTransfSAP(c.filas).importeTotal);
    if (t === 'bloq' && c.formato === 'concentrado') totalTxt = fmtDinero(computeBloqConc(c.filas, c.resumenHoja).total);
    if (t === 'bloq' && c.formato !== 'concentrado') totalTxt = fmtDinero(computeBloqOrig(c.filas).impacto) + ' impacto';
    resumen.push([etiqueta, c.fechaCorte, totalTxt, c.filas.length]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(panelAOA(t, c)), 'Panel ' + etiqueta);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detalleJSON(t, c)), 'Detallado ' + etiqueta);
  });
  if (!hubo) { toast('No hay datos cargados para exportar.'); return; }
  if (!tipo) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumen), 'Resumen');
  var nombre = tipo ? ('reporte-' + tipo + '-' + (ultimaCarga(tipo) || {}).fechaCorte + '.xlsx')
    : ('reporte-completo-' + hoyISO() + '.xlsx');
  XLSX.writeFile(wb, nombre);
  toast('Excel descargado: ' + nombre);
}

/* ---- eventos ---- */
function init() {
  cargarEstado();
  var dz = $('dropzone'), fi = $('file-input');
  $('btn-elegir').addEventListener('click', function (e) { e.stopPropagation(); fi.click(); });
  dz.addEventListener('click', function () { fi.click(); });
  dz.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); } });
  fi.addEventListener('change', function () { procesarArchivos(fi.files); fi.value = ''; });
  ['dragenter', 'dragover'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('drag'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('drag'); }); });
  dz.addEventListener('drop', function (e) { procesarArchivos(e.dataTransfer.files); });
  document.querySelectorAll('.tab').forEach(function (btn) {
    btn.addEventListener('click', function () {
      tabActual = btn.getAttribute('data-tab');
      document.querySelectorAll('.tab').forEach(function (b) { b.classList.toggle('active', b === btn); });
      ['merma', 'transf', 'bloq', 'cargas'].forEach(function (t) { $('tab-' + t).classList.toggle('hidden', t !== tabActual); });
      if (tabActual === 'cargas') renderCargas();
    });
  });
  document.addEventListener('click', function (e) {
    var imp = e.target.closest('[data-importar]');
    if (imp) { importarPendiente(imp.getAttribute('data-importar')); return; }
    var q = e.target.closest('[data-quitar]');
    if (q) { pendientes = pendientes.filter(function (x) { return x.id !== q.getAttribute('data-quitar'); }); renderPendientes(); return; }
    var b = e.target.closest('[data-borrar]');
    if (b) {
      var id = b.getAttribute('data-borrar');
      state.cargas = state.cargas.filter(function (c) { return c.id !== id; });
      guardarEstado(); renderCargas(); renderTodo();
      toast('Corte borrado.');
    }
  });
  document.addEventListener('change', function (e) {
    var s = e.target.closest('[data-seg]');
    if (!s) return;
    var clave = s.getAttribute('data-seg');
    if (!state.seguimiento[clave]) state.seguimiento[clave] = { estado: 'Pendiente', nota: '' };
    state.seguimiento[clave][s.getAttribute('data-campo')] = s.value;
    guardarEstado();
  });
  $('exp-todo').addEventListener('click', function () { exportar(null); });
  $('exp-merma').addEventListener('click', function () { exportar('merma'); });
  $('exp-transf').addEventListener('click', function () { exportar('transf'); });
  $('exp-bloq').addEventListener('click', function () { exportar('bloq'); });
  renderTodo();
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(function () {});
  }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();
}
