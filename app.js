/* Panel de Seguimiento — PWA estática
   Los datos viven solo en localStorage de este dispositivo.
   Detección del tipo de archivo por COLUMNAS, no por nombre. */
'use strict';

var LS_KEY = 'panelSeguimiento.v1';
var RANGOS = ['0-3', '4-7', '8-15', '16-30', '+30'];
var COLORES_RANGO = ['#1d4ed8', '#60a5fa', '#93c5fd', '#fb923c', '#ea580c'];

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
  'TRANSFERENCIAS (DETALLADO): Doc.Interno, Fecha Embarque, Tda.Origen, Descripción, DOC.CL.MOVTO.303, Fecha Doc.303, Articulo, Destino, Descripción, Texto, Costo, Costo, Fecha Doc.305, Piezas Embarcadas, Piezas Recibidas, Dif.de Piezas, DOC.CL.MOVTO.304, Dif. Final, Dias Transcurridos. En renglones viejos (fecha como texto DD.MM.AAAA) las piezas embarcadas vienen en la columna Fecha Doc.305 y las piezas pendientes en DOC.CL.MOVTO.304; el monto siempre es el primer Costo.\n' +
  'TRANSFERENCIAS (SAP): hoja REPORTE con ZONA, SUCURSAL ORIGEN, DOCUMENTO CONTABLE, FECHA CONTABLE, DIAS TRANSCURRIDOS, IMPORTE, SUCURSAL DESTINO, ESTATUS.\n' +
  'BLOQUEADO (CONCENTRADO): hoja CONCENTRADO con Centro, Material, EAN/UPC, Denomin., Unidad, Jerarquía 2, Almacén, Bloqueado, V. Stock bloq. (y hoja RESUMEN con Centro, Importe bloqueado).\n' +
  'BLOQUEADO (original): Hoja1 con ZONA, Material, Centro, Sucursal, Almacen, Libre Util, Stock Bloq, Cto, Importe, FORMATO, Descripción, Departamento, Clase.\n' +
  'ROTATIVOS: SUC, ART, TEO, FIS, DIFU, DIFP, FCN, ROL (solo se siguen las sucursales 127, 331, 333, 334, 336 y 1060; la semana va de jueves a miércoles).';

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
    // Rotativos: conteos por sucursal (SUC, ART, TEO, FIS, DIFU, DIFP, FCN, ROL).
    // La hoja que sube el usuario ya no trae Hoja2; todo se calcula de esta hoja.
    if (tiene(cols, 'SUC') && tiene(cols, 'ART') && tiene(cols, 'FCN') && tiene(cols, 'ROL') &&
        tiene(cols, 'DIFU') && tiene(cols, 'DIFP')) {
      return { tipo: 'rot', formato: 'rotativos', hoja: hojas[i] };
    }
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
    // Transferencias DETALLADO: detalle por artículo en una sola hoja (el archivo
    // TRANSFERENCIAS_PENDIENTES_DESTINO nuevo). Tiene Articulo + Dias Transcurridos
    // y NO trae hoja REPORTE (esa ya se descartó arriba). Va antes del SAP genérico.
    if (tiene(cols, 'Articulo') && tiene(cols, 'Doc.Interno') && tiene(cols, 'Dias Transcurridos')) {
      return { tipo: 'transf', formato: 'detallado', hoja: hojas[i] };
    }
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

/* Devuelve el número si el valor es numérico (número o texto puramente numérico);
   null en cualquier otro caso (vacío, texto, fecha). Sirve para distinguir las
   columnas desplazadas del formato detallado. */
function valorNumerico(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var s = String(v).trim();
  if (/^-?\d+(\.\d+)?$/.test(s)) return parseFloat(s);
  return null;
}

/* Transferencias DETALLADO (archivo nuevo, detalle por artículo).
   QUIRK de columnas desplazadas en los renglones viejos (fecha como texto
   DD.MM.AAAA): las piezas embarcadas vienen en la columna "Fecha Doc.305" y
   las piezas pendientes en "DOC.CL.MOVTO.304", y el monto del renglón viene
   en la columna "Dif. Final". En los renglones nuevos (fecha real de Excel)
   todo va en sus columnas. Reglas por renglón:
   - monto = primer Costo (siempre)
   - piezas embarcadas = Fecha Doc.305 si tiene valor numérico, si no Piezas Embarcadas
   - piezas pendientes = DOC.CL.MOVTO.304 si tiene valor numérico, si no Dif. Final */
function parseTransfDetallado(filas) {
  var out = [];
  filas.forEach(function (r) {
    if (!str(r['Doc.Interno'])) return;
    var desc = r['Descripción'];
    var origenDesc = '', destinoDesc = '';
    if (Array.isArray(desc)) { origenDesc = str(desc[0]); destinoDesc = str(desc[1]); }
    else { origenDesc = str(desc); destinoDesc = str(r['Descripción_1']); /* columna duplicada */ }
    var piezasEmb = valorNumerico(r['Fecha Doc.305']);
    if (piezasEmb === null) piezasEmb = num(r['Piezas Embarcadas']);
    var piezasPend = valorNumerico(r['DOC.CL.MOVTO.304']);
    if (piezasPend === null) piezasPend = num(r['Dif. Final']);
    out.push({
      doc: str(r['Doc.Interno']),
      fechaEmb: parseFecha(r['Fecha Embarque']),
      origen: str(r['Tda.Origen']),
      origenDesc: origenDesc,
      articulo: str(r['Articulo']),
      destino: str(r['Destino']),
      destinoDesc: destinoDesc,
      texto: str(r['Texto']),
      costo: num(r['Costo']),
      piezasEmb: piezasEmb,
      piezasRec: num(r['Piezas Recibidas']),
      piezasPend: piezasPend,
      dias: num(r['Dias Transcurridos'])
    });
  });
  return out;
}

function computeTransfDetallado(filas) {
  var porDoc = {}, importeTotal = 0, piezasEmb = 0, piezasPend = 0;
  filas.forEach(function (f) {
    importeTotal += f.costo; piezasEmb += f.piezasEmb; piezasPend += f.piezasPend;
    var d = porDoc[f.doc] || { doc: f.doc, fechaEmb: f.fechaEmb, origen: f.origen, origenDesc: f.origenDesc,
      destino: f.destino, destinoDesc: f.destinoDesc, articulos: 0, importe: 0, piezasEmb: 0, piezasPend: 0, dias: f.dias };
    d.articulos++; d.importe += f.costo; d.piezasEmb += f.piezasEmb; d.piezasPend += f.piezasPend;
    if (f.dias > d.dias) d.dias = f.dias;
    if (f.fechaEmb && (!d.fechaEmb || f.fechaEmb < d.fechaEmb)) d.fechaEmb = f.fechaEmb;
    porDoc[f.doc] = d;
  });
  var docs = Object.keys(porDoc).map(function (k) { return porDoc[k]; });
  var ant = {}, porOrigen = {}, porDestino = {}, masAntiguo = null;
  RANGOS.forEach(function (r) { ant[r] = { n: 0, monto: 0, piezas: 0 }; });
  docs.forEach(function (d) {
    var rg = rangoDe(d.dias);
    ant[rg].n++; ant[rg].monto += d.importe; ant[rg].piezas += d.piezasPend;
    var o = porOrigen[d.origen] || { clave: d.origen, nombre: nombreTienda(d.origen, d.origenDesc), docs: 0, importe: 0, piezas: 0 };
    o.docs++; o.importe += d.importe; o.piezas += d.piezasPend; porOrigen[d.origen] = o;
    var t = porDestino[d.destino] || { clave: d.destino, nombre: nombreTienda(d.destino, d.destinoDesc), docs: 0, importe: 0, piezas: 0 };
    t.docs++; t.importe += d.importe; t.piezas += d.piezasPend; porDestino[d.destino] = t;
    if (!masAntiguo || d.dias > masAntiguo.dias) masAntiguo = d;
  });
  function aLista(m) { return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return b.importe - a.importe; }); }
  docs.sort(function (a, b) { return b.dias - a.dias; });
  var topArticulos = filas.slice().sort(function (a, b) { return b.costo - a.costo; }).slice(0, 10);
  return { documentos: docs.length, registros: filas.length, importeTotal: importeTotal,
    piezasEmb: piezasEmb, piezasPend: piezasPend, porDoc: docs,
    porOrigen: aLista(porOrigen), porDestino: aLista(porDestino),
    antiguedad: ant, masAntiguo: masAntiguo, topArticulos: topArticulos };
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
    var o = porOrigen[f.origen] || { clave: f.origen, nombre: nombreTienda(f.origen, f.origenDesc), docs: 0, piezas: 0 };
    o.docs++; o.piezas += f.difFinal; porOrigen[f.origen] = o;
    var d = porDestino[f.destino] || { clave: f.destino, nombre: nombreTienda(f.destino, f.destinoDesc), docs: 0, piezas: 0 };
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
    var o = porOrigen[f.origen] || { clave: f.origen, nombre: nombreTienda(f.origen, f.origenDesc), docs: 0, importe: 0 };
    o.docs++; o.importe += f.importe; porOrigen[f.origen] = o;
    if (!masAntiguo || f.dias > masAntiguo.dias) masAntiguo = f;
  });
  var lista = Object.keys(porOrigen).map(function (k) { return porOrigen[k]; }).sort(function (a, b) { return b.importe - a.importe; });
  return { documentos: filas.length, importeTotal: total, porOrigen: lista, antiguedad: ant, masAntiguo: masAntiguo };
}

/* ---- fuera de política (transferencias) ---- */
/* Política: una transferencia debe liberarse en máximo 3 días.
   Documento fuera de política: Dias Transcurridos > 3. */
var POLITICA_DIAS_MAX = 3;

/* Normaliza el fuera de política para los 3 formatos de transferencias:
   { documentos, piezas, importe, tienePiezas, tieneImporte,
     porDestino: [{clave, nombre, docs, piezas, importe}],
     detalle: [{doc, fechaEmb, origen, origenDesc, destino, destinoDesc, dias, piezas, importe}] }
   El detalle va ordenado del más antiguo al más reciente (más días primero).
   SAP no trae piezas (tienePiezas false); DESTINO trae Costo Total en 0, por
   lo que tieneImporte es false cuando el importe total es 0. */
function fueraPoliticaTransf(carga) {
  var docs = [], tienePiezas = true;
  if (carga.formato === 'detallado') {
    var d = computeTransfDetallado(carga.filas);
    docs = d.porDoc.filter(function (x) { return x.dias > POLITICA_DIAS_MAX; }).map(function (x) {
      return { doc: x.doc, fechaEmb: x.fechaEmb, origen: x.origen, origenDesc: x.origenDesc,
        destino: x.destino, destinoDesc: x.destinoDesc, dias: x.dias, piezas: x.piezasPend, importe: x.importe };
    });
  } else if (carga.formato === 'destino') {
    docs = carga.filas.filter(function (f) { return f.dias > POLITICA_DIAS_MAX; }).map(function (f) {
      return { doc: f.doc, fechaEmb: f.fechaEmb, origen: f.origen, origenDesc: f.origenDesc,
        destino: f.destino, destinoDesc: f.destinoDesc, dias: f.dias, piezas: f.difFinal, importe: f.costoTotal };
    });
    docs.sort(function (a, b) { return b.dias - a.dias || ((a.fechaEmb || '') < (b.fechaEmb || '') ? -1 : 1); });
  } else {
    tienePiezas = false;
    docs = carga.filas.filter(function (f) { return f.dias > POLITICA_DIAS_MAX; }).map(function (f) {
      return { doc: f.doc, fechaEmb: f.fechaCont, origen: f.origen, origenDesc: f.origenDesc,
        destino: f.destino, destinoDesc: f.destinoDesc, dias: f.dias, piezas: 0, importe: f.importe };
    });
    docs.sort(function (a, b) { return b.dias - a.dias || ((a.fechaEmb || '') < (b.fechaEmb || '') ? -1 : 1); });
  }
  var porDestino = {};
  docs.forEach(function (x) {
    var t = porDestino[x.destino] || { clave: x.destino, nombre: nombreTienda(x.destino, x.destinoDesc), docs: 0, piezas: 0, importe: 0 };
    t.docs++; t.piezas += x.piezas; t.importe += x.importe; porDestino[x.destino] = t;
  });
  function aLista(m) { return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return (b.importe - a.importe) || (b.piezas - a.piezas); }); }
  var piezas = 0, importe = 0;
  docs.forEach(function (x) { piezas += x.piezas; importe += x.importe; });
  return { documentos: docs.length, piezas: piezas, importe: importe,
    tienePiezas: tienePiezas, tieneImporte: importe > 0,
    porDestino: aLista(porDestino), detalle: docs };
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
    var s = porSuc[f.centro] || { clave: f.centro, nombre: nombreTienda(f.centro, f.sucursal), registros: 0, importe: 0, impacto: 0 };
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

/* ---------------- rotativos (conteos semanales por sucursal) ---------------- */
var ROT_SUC = ['127', '331', '333', '334', '336', '1060'];
var ROT_META = { '127': 200, '331': 50, '333': 50, '334': 100, '336': 50, '1060': 50 };

/* Catálogo de tiendas del usuario (catalogo_tiendas.xlsx). Se usa el nombre
   completo en todos los lugares donde el panel muestra una sucursal. */
var CATALOGO_TIENDAS = {
  '127': 'CHEDRAUI SELECTO LA PAZ COLIMA',
  '331': 'SUPERCITO MAS ARAMBURO LA PAZ 16 DE SEP',
  '333': 'SUPERCITO SELECTO ARAMBURO LA PAZ MADERO',
  '334': 'SUPER CHEDRAUI ARAMBURO LA PAZ OCAMPO',
  '336': 'SUPERCITO MAS ARAMBURO LA PAZ MARQUES DE LEON',
  '1060': 'SUPERCITO LA PAZ EL CENTENARIO'
};
function nombreTienda(codigo, fallback) {
  var c = sucNormalizada(codigo);
  if (CATALOGO_TIENDAS[c]) return CATALOGO_TIENDAS[c];
  return (fallback !== undefined && fallback !== null && fallback !== '') ? fallback : str(codigo);
}

function sucNormalizada(v) {
  var s = str(v).replace(/^0+/, '');
  return s === '' ? '0' : s;
}

/* La semana de rotativos va de JUEVES a MIÉRCOLES. Devuelve el jueves (ISO)
   de la semana que contiene la fecha dada. */
function juevesDe(iso) {
  var d = new Date(iso + 'T00:00:00');
  var desdeJueves = (d.getDay() - 4 + 7) % 7; /* getDay: domingo=0, jueves=4 */
  d.setDate(d.getDate() - desdeJueves);
  return isoDeFecha(d);
}
function miercolesDe(juevesISO) {
  var d = new Date(juevesISO + 'T00:00:00');
  d.setDate(d.getDate() + 6);
  return isoDeFecha(d);
}
function etiquetaSemana(juevesISO) {
  var m = miercolesDe(juevesISO);
  return 'Semana ' + juevesISO.slice(8, 10) + '/' + juevesISO.slice(5, 7) +
    '–' + m.slice(8, 10) + '/' + m.slice(5, 7) + '/' + m.slice(0, 4);
}

/* Solo se siguen las 6 sucursales de ROT_SUC; los demás SUC del archivo se ignoran.
   FCN puede venir como fecha de Excel o como texto DD.MM.AAAA (parseFecha lo cubre). */
function parseRotativos(filas) {
  var out = [];
  filas.forEach(function (r) {
    var suc = sucNormalizada(r['SUC']);
    if (ROT_SUC.indexOf(suc) < 0) return;
    out.push({
      suc: suc,
      fecha: parseFecha(r['FCN']),
      rol: str(r['ROL']),
      articulo: str(r['ART']),
      descripcion: str(r['DSC']),
      teo: num(r['TEO']),
      fis: num(r['FIS']),
      difu: num(r['DIFU']),
      difp: num(r['DIFP'])
    });
  });
  return out;
}

/* Agrupa por semana (jueves-miércoles) y sucursal. Las sucursales sin renglones
   en una semana que otras sí tienen aparecen con 0 artículos y estatus Pendiente. */
function computeRotativos(filas) {
  var porSemana = {}; /* juevesISO -> suc -> stats */
  var rolesTotales = {};
  var totalArticulos = 0;
  filas.forEach(function (f) {
    if (!f.fecha) return;
    totalArticulos++;
    if (f.rol) rolesTotales[f.rol] = 1;
    var sem = juevesDe(f.fecha);
    if (!porSemana[sem]) porSemana[sem] = {};
    var s = porSemana[sem][f.suc] || { articulos: 0, roles: {}, difu: 0, difp: 0, conDif: 0 };
    s.articulos++;
    if (f.rol) s.roles[f.rol] = 1;
    s.difu += f.difu; s.difp += f.difp;
    if (f.difu !== 0) s.conDif++;
    porSemana[sem][f.suc] = s;
  });
  var semanas = Object.keys(porSemana).sort().reverse();
  var rows = [];
  semanas.forEach(function (sem) {
    ROT_SUC.forEach(function (suc) {
      var s = porSemana[sem][suc] || { articulos: 0, roles: {}, difu: 0, difp: 0, conDif: 0 };
      var meta = ROT_META[suc];
      var cumpl = Math.min(100, s.articulos / meta * 100);
      rows.push({
        semana: sem, etiqueta: etiquetaSemana(sem), suc: suc,
        articulos: s.articulos, conteos: Object.keys(s.roles).length,
        meta: meta, cumplimiento: cumpl,
        estatus: cumpl >= 100 ? 'Cumplido' : 'Pendiente',
        difu: s.difu, difp: s.difp, conDif: s.conDif,
        exactitud: s.articulos ? ((s.articulos - s.conDif) / s.articulos * 100) : 0
      });
    });
  });
  return { filas: rows, semanas: semanas, totalArticulos: totalArticulos,
    totalConteos: Object.keys(rolesTotales).length };
}

/* Mayores diferencias del último corte de rotativos: top 10 faltantes
   (diferencia más negativa) y top 10 sobrantes (diferencia más positiva),
   en dinero (DIFP) y en piezas (DIFU). Si hay menos de 10, salen los que hay. */
function computeTopDiferencias(filas) {
  var faltDin = filas.filter(function (f) { return f.difp < 0; }).sort(function (a, b) { return a.difp - b.difp; }).slice(0, 10);
  var sobDin = filas.filter(function (f) { return f.difp > 0; }).sort(function (a, b) { return b.difp - a.difp; }).slice(0, 10);
  var faltPzas = filas.filter(function (f) { return f.difu < 0; }).sort(function (a, b) { return a.difu - b.difu; }).slice(0, 10);
  var sobPzas = filas.filter(function (f) { return f.difu > 0; }).sort(function (a, b) { return b.difu - a.difu; }).slice(0, 10);
  return { faltantesDinero: faltDin, sobrantesDinero: sobDin,
    faltantesPiezas: faltPzas, sobrantesPiezas: sobPzas };
}

/* Agregación semanal de TODOS los cortes de rotativos. Por cada par
   (semana, sucursal) manda el corte más nuevo que tenga renglones de ese
   par; si ningún corte trae renglones del par, queda en 0/Pendiente. Así
   las semanas no se duplican entre cortes y no se pierde el historial. */
function computeRotativosMulti(cortes) {
  var porPar = {};
  var semanasSet = {};
  var orden = (cortes || []).slice().sort(function (a, b) { return (a.fechaCorte < b.fechaCorte) ? -1 : 1; });
  orden.forEach(function (c) {
    var rp = computeRotativos(c.filas);
    rp.filas.forEach(function (x) {
      if (x.articulos > 0) porPar[x.semana + '|' + x.suc] = x;
      semanasSet[x.semana] = 1;
    });
  });
  var semanas = Object.keys(semanasSet).sort().reverse();
  var filas = [];
  semanas.forEach(function (sem) {
    ROT_SUC.forEach(function (suc) {
      var x = porPar[sem + '|' + suc];
      if (x) filas.push(x);
      else filas.push({ semana: sem, etiqueta: etiquetaSemana(sem), suc: suc,
        articulos: 0, conteos: 0, meta: ROT_META[suc], cumplimiento: 0,
        estatus: 'Pendiente', difu: 0, difp: 0, conDif: 0, exactitud: 0 });
    });
  });
  return { filas: filas, semanas: semanas };
}

/* Sección fuera de política para la hoja Panel de la exportación de transferencias. */
/* DIFP con signo explícito: "-$14,045.24", "+$7,466.10", "$0.00". */
function fmtDifp(v) {
  v = num(v);
  var signo = v < 0 ? '-$' : v > 0 ? '+$' : '$';
  return signo + Math.abs(v).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
/* Exportación para Node (verificación) */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseFecha: parseFecha, num: num, valorNumerico: valorNumerico, detectarTipo: detectarTipo,
    parseMerma: parseMerma, parseTransfDestino: parseTransfDestino,
    parseTransfDetallado: parseTransfDetallado, computeTransfDetallado: computeTransfDetallado,
    parseTransfSAP: parseTransfSAP, parseBloqConcentrado: parseBloqConcentrado,
    parseBloqOriginal: parseBloqOriginal, sugerirCorte: sugerirCorte,
    computeMerma: computeMerma, computeTransfDestino: computeTransfDestino,
    computeTransfSAP: computeTransfSAP, computeBloqConc: computeBloqConc,
    computeBloqOrig: computeBloqOrig, filasDeHoja: filasDeHoja,
    parseRotativos: parseRotativos, computeRotativos: computeRotativos,
    computeTopDiferencias: computeTopDiferencias, computeRotativosMulti: computeRotativosMulti,
    nombreTienda: nombreTienda, CATALOGO_TIENDAS: CATALOGO_TIENDAS,
    juevesDe: juevesDe, etiquetaSemana: etiquetaSemana,
    fmtDifp: fmtDifp,
    fueraPoliticaTransf: fueraPoliticaTransf,
    POLITICA_DIAS_MAX: POLITICA_DIAS_MAX
  };
}

/* ================= UI (solo navegador) ================= */
if (typeof window !== 'undefined' && typeof document !== 'undefined') {

var state = { cargas: [], seguimiento: {} };
var pendientes = [];
var tabActual = 'merma';
var soloFPTransf = false; /* filtro del detalle por artículo de transferencias: solo fuera de política */

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
  return tipo === 'merma' ? 'Merma' : tipo === 'transf' ? 'Transferencias' : tipo === 'rot' ? 'Rotativos' : 'Bloqueado';
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
  var grad = ctx.createLinearGradient(0, 0, 0, H);
  grad.addColorStop(0, '#1d4ed8');
  grad.addColorStop(1, '#f97316');
  items.forEach(function (it, i) {
    var h = Math.max(2, (it.valor / max) * (H - 52));
    var x = 10 + i * ((W - 20) / items.length) + ((W - 20) / items.length - bw) / 2;
    ctx.fillStyle = grad;
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
    return '<tr><td>' + esc(x.centro) + ' ' + esc(nombreTienda(x.centro)) + '</td><td class="num">' + fmtDinero(x.monto) + '</td><td class="num">' + x.pct.toFixed(1) + '%</td>' +
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

/* Sección "Fuera de política (más de 3 días)" al inicio del panel de
   Transferencias: totales, desglose por sucursal destino y detalle de
   documentos, del más antiguo al más reciente. */
function fueraPoliticaHTML(c) {
  var fp = fueraPoliticaTransf(c);
  var totalDocs = c.formato === 'detallado' ? computeTransfDetallado(c.filas).documentos : c.filas.length;
  var html = '<div class="fp-box"><h3 class="fp-titulo">⚠ Fuera de política (más de 3 días)</h3>';
  html += '<div class="totales">' +
    '<div class="total-card fp-destacado"><div class="etiqueta">Documentos fuera de política</div><div class="valor">' + fp.documentos + '</div><div class="sub">de ' + totalDocs + ' documentos</div></div>' +
    '<div class="total-card"><div class="etiqueta">Piezas pendientes</div><div class="valor">' + (fp.tienePiezas ? fmtNum(fp.piezas) : '—') + '</div></div>' +
    '<div class="total-card"><div class="etiqueta">Importe fuera de política</div><div class="valor">' + (fp.tieneImporte ? fmtDinero(fp.importe) : '—') + '</div>' + ((c.formato === 'destino' && !fp.tieneImporte) ? '<div class="sub">El archivo trae Costo Total en 0</div>' : '') + '</div></div>';
  if (!fp.documentos) {
    html += '<p class="hint">Ningún documento fuera de política. Todos están dentro de los 3 días.</p></div>';
    return html;
  }
  var bdHeaders = [{ t: 'Destino' }, { t: 'Documentos', num: true }];
  if (fp.tienePiezas) bdHeaders.push({ t: 'Piezas pend.', num: true });
  if (fp.tieneImporte) bdHeaders.push({ t: 'Importe', num: true });
  var bdRows = fp.porDestino.map(function (x) {
    return '<tr><td>' + esc(x.clave) + ' ' + esc(x.nombre) + '</td><td class="num">' + x.docs + '</td>' +
      (fp.tienePiezas ? '<td class="num">' + fmtNum(x.piezas) + '</td>' : '') +
      (fp.tieneImporte ? '<td class="num">' + fmtDinero(x.importe) + '</td>' : '') + '</tr>';
  }).join('');
  var bdTotal = '<td>Total</td><td class="num">' + fp.documentos + '</td>' +
    (fp.tienePiezas ? '<td class="num">' + fmtNum(fp.piezas) + '</td>' : '') +
    (fp.tieneImporte ? '<td class="num">' + fmtDinero(fp.importe) + '</td>' : '');
  html += '<h3>Por sucursal destino</h3>' + tablaHTML(bdHeaders, bdRows, bdTotal);
  var detHeaders = [{ t: 'Documento' }, { t: 'Fecha embarque' }, { t: 'Origen' }, { t: 'Destino' }, { t: 'Días', num: true }];
  if (fp.tienePiezas) detHeaders.push({ t: 'Piezas pend.', num: true });
  if (fp.tieneImporte) detHeaders.push({ t: 'Importe', num: true });
  detHeaders.push({ t: 'Seguimiento' });
  var detRows = fp.detalle.map(function (x) {
    var clave = 'transf:' + x.doc;
    return '<tr><td>' + esc(x.doc) + '</td><td>' + fmtFecha(x.fechaEmb) + '</td><td>' + esc(x.origen) + ' ' + esc(nombreTienda(x.origen, x.origenDesc)) + '</td>' +
      '<td>' + esc(x.destino) + ' ' + esc(nombreTienda(x.destino, x.destinoDesc)) + '</td><td class="num">' + x.dias + '</td>' +
      (fp.tienePiezas ? '<td class="num">' + fmtNum(x.piezas) + '</td>' : '') +
      (fp.tieneImporte ? '<td class="num">' + fmtDinero(x.importe) + '</td>' : '') +
      '<td>' + segControl('transf', clave) + '</td></tr>';
  }).join('');
  var detTotal = '<td>Total</td><td></td><td></td><td></td><td></td>' +
    (fp.tienePiezas ? '<td class="num">' + fmtNum(fp.piezas) + '</td>' : '') +
    (fp.tieneImporte ? '<td class="num">' + fmtDinero(fp.importe) + '</td>' : '') + '<td></td>';
  html += '<h3>Detalle de documentos fuera de política</h3>' + tablaHTML(detHeaders, detRows, detTotal);
  html += '</div>';
  return html;
}

function renderTransf() {
  var c = ultimaCarga('transf');
  var el = $('transf-panel'), info = $('transf-corte');
  if (!c) { info.textContent = ''; el.innerHTML = '<p class="vacio">Aún no hay transferencias cargadas. Sube el archivo de transferencias arriba.</p>'; return; }
  info.textContent = 'Corte: ' + fmtFecha(c.fechaCorte) + ' · Archivo: ' + c.nombre + ' · Formato: ' + (c.formato === 'destino' ? 'Destino' : c.formato === 'detallado' ? 'Detallado' : 'SAP');
  var html = fueraPoliticaHTML(c);
  if (c.formato === 'detallado') {
    var d = computeTransfDetallado(c.filas);
    html += '<div class="totales">' +
      '<div class="total-card destacado"><div class="etiqueta">Importe pendiente</div><div class="valor">' + fmtDinero(d.importeTotal) + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Documentos</div><div class="valor">' + d.documentos + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Piezas pendientes</div><div class="valor">' + fmtNum(d.piezasPend) + '</div></div>' +
      '<div class="total-card"><div class="etiqueta">Más antiguo</div><div class="valor">' + (d.masAntiguo ? d.masAntiguo.dias + ' días' : '—') + '</div></div></div>';
    html += '<div class="grafica"><h3>Distribución por antigüedad</h3>' + donutHTML(d.antiguedad, d.importeTotal, false) + '</div>';
    html += barrasHTML('Pendiente por sucursal origen', d.porOrigen.slice(0, 6).map(function (x) { return { etiqueta: x.clave, valor: x.importe }; }), fmtDinero);
    html += barrasHTML('Pendiente por sucursal destino', d.porDestino.slice(0, 6).map(function (x) { return { etiqueta: x.clave, valor: x.importe }; }), fmtDinero);
    var topRows = d.topArticulos.map(function (f) {
      var clave = 'transf:art:' + f.doc + ':' + f.articulo;
      return '<tr><td>' + esc(f.articulo) + '</td><td>' + esc(f.texto) + '</td><td>' + esc(f.doc) + '</td><td class="num">' + fmtDinero(f.costo) + '</td><td class="num">' + fmtNum(f.piezasPend) + '</td><td class="num">' + f.dias + '</td><td>' + segControl('transf', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Artículos más críticos</h3>' + tablaHTML(
      [{ t: 'Artículo' }, { t: 'Texto' }, { t: 'Documento' }, { t: 'Costo', num: true }, { t: 'Piezas pend.', num: true }, { t: 'Días', num: true }, { t: 'Seguimiento' }], topRows, '');
    var docRows = d.porDoc.map(function (x) {
      var clave = 'transf:' + x.doc;
      return '<tr><td>' + esc(x.doc) + '</td><td>' + fmtFecha(x.fechaEmb) + '</td><td>' + esc(x.origen) + ' ' + esc(nombreTienda(x.origen, x.origenDesc)) + '</td>' +
        '<td>' + esc(x.destino) + ' ' + esc(nombreTienda(x.destino, x.destinoDesc)) + '</td><td class="num">' + x.articulos + '</td><td class="num">' + fmtDinero(x.importe) + '</td>' +
        '<td class="num">' + fmtNum(x.piezasPend) + '</td><td class="num">' + x.dias + '</td><td>' + segControl('transf', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Detalle por documento</h3>' + tablaHTML(
      [{ t: 'Documento' }, { t: 'Fecha embarque' }, { t: 'Origen' }, { t: 'Destino' }, { t: 'Artículos', num: true }, { t: 'Importe', num: true }, { t: 'Piezas pend.', num: true }, { t: 'Días', num: true }, { t: 'Seguimiento' }],
      docRows, '<td>Total</td><td></td><td></td><td></td><td class="num">' + d.registros + '</td><td class="num">' + fmtDinero(d.importeTotal) + '</td><td class="num">' + fmtNum(d.piezasPend) + '</td><td></td><td></td>');
    var artFilas = c.filas.slice().sort(function (a, b) { return b.costo - a.costo; });
    if (soloFPTransf) artFilas = artFilas.filter(function (f) { return f.dias > POLITICA_DIAS_MAX; });
    var artRows = artFilas.map(function (f) {
      var clave = 'transf:art:' + f.doc + ':' + f.articulo;
      return '<tr><td>' + esc(f.articulo) + '</td><td>' + esc(f.texto) + '</td><td>' + esc(f.doc) + '</td><td>' + esc(nombreTienda(f.origen, f.origenDesc)) + ' → ' + esc(nombreTienda(f.destino, f.destinoDesc)) + '</td><td class="num">' + fmtDinero(f.costo) + '</td><td class="num">' + fmtNum(f.piezasEmb) + '</td><td class="num">' + fmtNum(f.piezasPend) + '</td><td class="num">' + f.dias + '</td><td>' + segControl('transf', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Detalle por artículo' + (soloFPTransf ? ' — solo fuera de política (' + artFilas.length + ' de ' + c.filas.length + ' renglones)' : '') + '</h3>' +
      '<div class="filtro-fp"><label><input type="checkbox" data-filtro-fp="transf"' + (soloFPTransf ? ' checked' : '') + '> Solo fuera de política (más de 3 días)</label></div>' + tablaHTML(
      [{ t: 'Artículo' }, { t: 'Texto' }, { t: 'Documento' }, { t: 'Origen → Destino' }, { t: 'Costo', num: true }, { t: 'Piezas emb.', num: true }, { t: 'Piezas pend.', num: true }, { t: 'Días', num: true }, { t: 'Seguimiento' }],
      artRows, '');
    if (d.masAntiguo) html += '<p class="hint">Documento más antiguo: ' + esc(d.masAntiguo.doc) + ' de ' + esc(nombreTienda(d.masAntiguo.origen, d.masAntiguo.origenDesc)) + ' a ' + esc(nombreTienda(d.masAntiguo.destino, d.masAntiguo.destinoDesc)) + ' (' + d.masAntiguo.dias + ' días, ' + fmtDinero(d.masAntiguo.importe) + ', ' + fmtNum(d.masAntiguo.piezasPend) + ' piezas).</p>';
  } else if (c.formato === 'destino') {
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
      return '<tr><td>' + esc(f.doc) + '</td><td>' + fmtFecha(f.fechaEmb) + '</td><td>' + esc(f.origen) + ' ' + esc(nombreTienda(f.origen, f.origenDesc)) + '</td>' +
        '<td>' + esc(f.destino) + ' ' + esc(nombreTienda(f.destino, f.destinoDesc)) + '</td><td class="num">' + fmtNum(f.piezasEmb) + '</td><td class="num">' + fmtNum(f.difFinal) + '</td>' +
        '<td class="num">' + f.dias + '</td><td>' + segControl('transf', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Detalle por documento</h3>' + tablaHTML(
      [{ t: 'Documento' }, { t: 'Fecha embarque' }, { t: 'Origen' }, { t: 'Destino' }, { t: 'Piezas emb.', num: true }, { t: 'Piezas pend.', num: true }, { t: 'Días', num: true }, { t: 'Seguimiento' }],
      rows, '<td>Total</td><td></td><td></td><td></td><td class="num">' + fmtNum(p.piezasEmb) + '</td><td class="num">' + fmtNum(p.piezasPend) + '</td><td></td><td></td>');
    if (p.masAntiguo) html += '<p class="hint">Documento más antiguo: ' + esc(p.masAntiguo.doc) + ' de ' + esc(nombreTienda(p.masAntiguo.origen, p.masAntiguo.origenDesc)) + ' (' + p.masAntiguo.dias + ' días).</p>';
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
      return '<tr><td>' + esc(f.doc) + '</td><td>' + fmtFecha(f.fechaCont) + '</td><td>' + esc(f.origen) + ' ' + esc(nombreTienda(f.origen, f.origenDesc)) + '</td>' +
        '<td>' + esc(f.destino) + ' ' + esc(nombreTienda(f.destino, f.destinoDesc)) + '</td><td class="num">' + fmtDinero(f.importe) + '</td>' +
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
      return '<tr><td>' + esc(x.clave) + ' ' + esc(nombreTienda(x.clave)) + '</td><td class="num">' + x.registros + '</td><td class="num">' + fmtNum(x.unidades, 3) + '</td><td class="num">' + fmtDinero(x.valor) + '</td><td>' + segControl('bloq', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Detalle por centro</h3>' + tablaHTML(
      [{ t: 'Centro' }, { t: 'Registros', num: true }, { t: 'Unidades', num: true }, { t: 'Importe', num: true }, { t: 'Seguimiento' }],
      rows, '<td>Total</td><td class="num">' + p.registros + '</td><td class="num">' + fmtNum(p.unidades, 3) + '</td><td class="num">' + fmtDinero(p.total) + '</td><td></td>');
    var topRows = p.top.map(function (f) {
      var clave = 'bloq:' + f.centro + ':' + f.material;
      return '<tr><td>' + esc(f.centro) + ' ' + esc(nombreTienda(f.centro)) + '</td><td>' + esc(f.material) + '</td><td>' + esc(f.denom) + '</td><td>' + esc(f.jerarquia) + '</td><td class="num">' + fmtNum(f.bloqueado, 3) + '</td><td class="num">' + fmtDinero(f.valor) + '</td><td>' + segControl('bloq', clave) + '</td></tr>';
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
      return '<tr><td>' + esc(f.centro) + ' ' + esc(nombreTienda(f.centro)) + '</td><td>' + esc(f.material) + '</td><td>' + esc(f.descripcion) + '</td><td>' + esc(f.departamento) + '</td><td class="num">' + fmtNum(f.stockBloq, 2) + '</td><td class="num">' + fmtDinero(f.importe) + '</td><td>' + segControl('bloq', clave) + '</td></tr>';
    }).join('');
    html += '<h3>Artículos más críticos</h3>' + tablaHTML(
      [{ t: 'Centro' }, { t: 'Material' }, { t: 'Descripción' }, { t: 'Departamento' }, { t: 'Stock bloq.', num: true }, { t: 'Importe', num: true }, { t: 'Seguimiento' }], topRows2, '');
  }
  el.innerHTML = html;
}

function tablaRotativos(rows) {
  var html = rows.map(function (x) {
    return '<tr><td>' + esc(x.etiqueta) + '</td><td>' + esc(x.suc) + ' ' + esc(nombreTienda(x.suc)) + '</td>' +
      '<td class="num">' + x.articulos + '</td><td class="num">' + x.meta + '</td>' +
      '<td class="num">' + Math.round(x.cumplimiento) + '%</td>' +
      '<td class="num">' + x.conteos + '</td>' +
      '<td class="num">' + fmtNum(x.difu) + '</td><td class="num">' + fmtDinero(x.difp) + '</td>' +
      '<td class="num">' + x.conDif + '</td>' +
      '<td class="num">' + (x.articulos ? x.exactitud.toFixed(1) + '%' : '—') + '</td>' +
      '<td><span class="nivel ' + (x.estatus === 'Cumplido' ? 'controlado' : 'seguimiento') + '">' + x.estatus + '</span></td></tr>';
  }).join('');
  return tablaHTML(
    [{ t: 'Semana' }, { t: 'Sucursal' }, { t: 'Artículos', num: true }, { t: 'Meta', num: true }, { t: '% cumpl.', num: true },
     { t: 'Conteos', num: true }, { t: 'DIFU', num: true }, { t: 'DIFP', num: true }, { t: 'Art. con dif.', num: true },
     { t: 'Exactitud', num: true }, { t: 'Estatus' }],
    html, '');
}

function tablaTopDiferencias(titulo, filasTop) {
  if (!filasTop.length) return '<h3>' + esc(titulo) + '</h3><p class="hint">Sin renglones en este corte.</p>';
  var rows = filasTop.map(function (f) {
    return '<tr><td>' + esc(f.suc) + ' ' + esc(nombreTienda(f.suc)) + '</td>' +
      '<td>' + esc(f.articulo) + '</td><td>' + esc(f.descripcion) + '</td>' +
      '<td class="num">' + fmtNum(f.teo) + '</td><td class="num">' + fmtNum(f.fis) + '</td>' +
      '<td class="num">' + fmtNum(f.difu) + '</td><td class="num">' + fmtDifp(f.difp) + '</td></tr>';
  }).join('');
  return '<h3>' + esc(titulo) + '</h3>' + tablaHTML(
    [{ t: 'Sucursal' }, { t: 'Artículo' }, { t: 'Descripción' }, { t: 'TEO', num: true }, { t: 'FIS', num: true }, { t: 'Dif. pzas', num: true }, { t: 'Dif. $', num: true }],
    rows, '');
}

function renderRotativos() {
  var c = ultimaCarga('rot');
  var el = $('rot-panel'), info = $('rot-corte');
  if (!c) { info.textContent = ''; el.innerHTML = '<p class="vacio">Aún no hay rotativos cargados. Sube el archivo de rotativos arriba.</p>'; return; }
  var p = computeRotativos(c.filas);
  var multi = computeRotativosMulti(cortesDe('rot'));
  var top = computeTopDiferencias(c.filas);
  info.textContent = 'Corte: ' + fmtFecha(c.fechaCorte) + ' · Archivo: ' + c.nombre + ' · ' + p.totalArticulos + ' artículos en ' + p.totalConteos + ' conteos · Semana de jueves a miércoles';
  var html = '';
  if (multi.semanas.length) {
    var ult = multi.semanas[0];
    var ultRows = multi.filas.filter(function (x) { return x.semana === ult; });
    html += '<h3>Última semana: ' + esc(etiquetaSemana(ult)) + '</h3><div class="totales">';
    ultRows.forEach(function (x) {
      var color = x.cumplimiento >= 100 ? '#1d4ed8' : '#f97316';
      html += '<div class="total-card semana-blanca">' +
        '<div class="etiqueta">' + esc(x.suc) + ' ' + esc(nombreTienda(x.suc)) + '</div>' +
        '<div class="valor">' + Math.round(x.cumplimiento) + '%</div>' +
        '<div class="barra-fila"><span class="btrack"><span class="bfill" style="display:block;width:' + Math.round(x.cumplimiento) + '%;background:' + color + '"></span></span></div>' +
        '<div class="sub">' + x.articulos + ' / ' + x.meta + ' artículos · ' + x.conteos + ' conteos</div>' +
        '<div class="sub">DIFU ' + fmtNum(x.difu) + ' · DIFP ' + fmtDinero(x.difp) + '</div>' +
        '<div class="sub"><span class="nivel neutro">' + x.estatus + '</span></div></div>';
    });
    html += '</div>';
  }
  html += '<h3>Mayores diferencias en $</h3>';
  html += tablaTopDiferencias('Faltantes — Top 10', top.faltantesDinero);
  html += tablaTopDiferencias('Sobrantes — Top 10', top.sobrantesDinero);
  html += '<h3>Mayores diferencias en piezas</h3>';
  html += tablaTopDiferencias('Faltantes — Top 10', top.faltantesPiezas);
  html += tablaTopDiferencias('Sobrantes — Top 10', top.sobrantesPiezas);
  html += '<h3>Productividad semanal por sucursal</h3>' + tablaRotativos(multi.filas);
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

function renderTodo() { renderMerma(); renderTransf(); renderBloq(); renderRotativos(); renderCargas(); }

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
        } else if (det.tipo === 'transf' && det.formato === 'detallado') {
          filas = parseTransfDetallado(filasDeHoja(wb, det.hoja));
          fechas = filas.map(function (f) { return f.fechaEmb; });
        } else if (det.tipo === 'rot') {
          filas = parseRotativos(filasDeHoja(wb, det.hoja));
          fechas = filas.map(function (f) { return f.fecha; });
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
  if (p.tipo === 'rot') {
    var roles = {};
    p.filas.forEach(function (f) { if (f.rol) roles[f.rol] = 1; });
    return p.filas.length + ' artículos · ' + Object.keys(roles).length + ' conteos · sucursales 127, 331, 333, 334, 336 y 1060';
  }
  if (p.tipo === 'merma') return 'Total: ' + fmtDinero(p.filas.reduce(function (a, f) { return a + f.costoTotal; }, 0)) + ' · ' + p.filas.length + ' registros';
  if (p.tipo === 'transf' && p.formato === 'destino') {
    var pe = p.filas.reduce(function (a, f) { return a + f.piezasEmb; }, 0);
    var pp = p.filas.reduce(function (a, f) { return a + f.difFinal; }, 0);
    return p.filas.length + ' documentos · ' + fmtNum(pe) + ' piezas embarcadas · ' + fmtNum(pp) + ' pendientes';
  }
  if (p.tipo === 'transf' && p.formato === 'detallado') {
    var im = p.filas.reduce(function (a, f) { return a + f.costo; }, 0);
    var pd = p.filas.reduce(function (a, f) { return a + f.piezasPend; }, 0);
    var docs = {};
    p.filas.forEach(function (f) { docs[f.doc] = 1; });
    return Object.keys(docs).length + ' documentos · ' + fmtDinero(im) + ' · ' + fmtNum(pd) + ' piezas pendientes · ' + p.filas.length + ' renglones';
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
  ['merma', 'transf', 'bloq', 'rot'].forEach(function (t) {
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
      ['merma', 'transf', 'bloq', 'rot', 'cargas'].forEach(function (t) { $('tab-' + t).classList.toggle('hidden', t !== tabActual); });
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
    var fpf = e.target.closest('[data-filtro-fp]');
    if (fpf) { soloFPTransf = fpf.checked; renderTransf(); return; }
    var s = e.target.closest('[data-seg]');
    if (!s) return;
    var clave = s.getAttribute('data-seg');
    if (!state.seguimiento[clave]) state.seguimiento[clave] = { estado: 'Pendiente', nota: '' };
    state.seguimiento[clave][s.getAttribute('data-campo')] = s.value;
    guardarEstado();
  });
  renderTodo();
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(function () {});
  }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();
}
