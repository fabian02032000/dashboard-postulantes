/**
 * API DE CONVERSIONES DE META - RECLUTAMIENTO MULTICAMPAÑA (Netcall)
 *
 * Qué hace: cada 5 minutos revisa TODAS las hojas y avisa a Meta la tipificación de cada
 * postulante (EN PROCESO, APTO, NO APTO, Ingreso a Capacitación = Si), para que sus
 * anuncios busquen más gente como los postulantes buenos.
 * Con el menú "3. Enviar todo lo ya tipificado" manda también lo que ya estaba en las hojas.
 *
 * En el mismo ciclo de 5 minutos, también revisa si algún postulante cayó por encima de las
 * cabeceras (el bug de la hoja) y lo mueve justo debajo, copiando el formato y el menú
 * desplegable de una fila de datos que ya esté bien ubicada. No se toca ninguna otra celda.
 *
 * También actualiza la hoja "RESUMEN POSTULANTES" con un cuadro por mes y por campaña:
 * cuántos postulantes, cuántos aptos, no aptos, en proceso, sin tipificar, ingresaron a
 * capacitación, y algunos porcentajes útiles. Se recalcula solo, cada 5 minutos.
 *
 * Además publica en GitHub los datos del dashboard (sin nombres, teléfonos ni correos),
 * cada vez que algo cambia. Se configura una vez con el menú "Dashboard: configurar GitHub".
 *
 * Lo único que crea son dos hojas ocultas: "_envios_meta" (registro de envíos a Meta) y
 * "_log_orden" (registro de postulantes reubicados).
 */

// ---------- AJUSTES (normalmente no hace falta tocar nada aquí) ----------
const VERSION_API = 'v26.0';
const NOMBRE_FUENTE = 'Google Sheets Netcall';

// Cada línea = un aviso a Meta. Se envía cuando la columna tiene uno de esos valores.
const EVENTOS = [
  { columna: 'Estatus de R&S',         valores: ['EN PROCESO', 'APTO'], evento: 'Postulante en proceso' },
  { columna: 'Estatus de R&S',         valores: ['APTO'],               evento: 'Postulante apto' },
  { columna: 'Estatus de R&S',         valores: ['NO APTO'],            evento: 'Postulante no apto' },
  { columna: 'Ingreso a Capacitación', valores: ['SI'],                 evento: 'Ingreso a capacitacion' }
];

const HOJA_REGISTRO = '_envios_meta';
const HOJA_LOG_ORDEN = '_log_orden';
const HOJA_RESUMEN = 'RESUMEN POSTULANTES';
const MESES_ES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const MAX_INTENTOS = 3;
const LOTE = 50;            // eventos por envío en el historial
const DIAS_MAX = 6;         // Meta solo acepta eventos de los últimos 7 días

// ---------- MENÚ ----------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Meta')
    .addItem('1. Configurar datos de Meta', 'configurarMeta')
    .addItem('2. Enviar prueba a Meta', 'enviarPrueba')
    .addItem('3. Enviar todo lo ya tipificado (historial)', 'enviarHistorial')
    .addItem('4. Activar envío automático (cada 5 min)', 'activarAutomatico')
    .addSeparator()
    .addItem('Enviar ahora (revisar todo)', 'revisarYEnviar')
    .addItem('Desactivar envío automático', 'desactivarAutomatico')
    .addItem('Revisar configuración', 'revisarConfiguracion')
    .addSeparator()
    .addItem('Arreglar cabeceras ahora', 'arreglarCabecerasAhora')
    .addItem('Actualizar Resumen Postulantes ahora', 'actualizarResumenAhora')
    .addSeparator()
    .addItem('Dashboard: configurar GitHub', 'configurarDashboard')
    .addItem('Dashboard: publicar ahora', 'publicarDashboardAhora')
    .addToUi();
}

// ---------- CONFIGURACIÓN ----------
function configurarMeta() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const a = ui.prompt('Paso 1 de 3', 'Pega el ID del conjunto de datos (solo números):', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const b = ui.prompt('Paso 2 de 3', 'Pega el token de acceso (texto largo):', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  const c = ui.prompt('Paso 3 de 3',
    'Pega el código de prueba (empieza con TEST...). Si no tienes, déjalo vacío:', ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK) return;

  props.setProperty('DATASET_ID', a.getResponseText().replace(/\D/g, ''));
  props.setProperty('TOKEN', limpiarToken_(b.getResponseText()));
  props.setProperty('CODIGO_PRUEBA', c.getResponseText().replace(/\s+/g, ''));
  ui.alert('Listo. Datos guardados. Ahora usa "2. Enviar prueba a Meta".');
}

// Quita espacios, saltos de línea y comillas que se cuelan al pegar el token.
function limpiarToken_(t) {
  return String(t || '').replace(/\s+/g, '').replace(/^["']+|["']+$/g, '');
}

// Muestra qué quedó guardado, sin enseñar el token completo.
function revisarConfiguracion() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('DATASET_ID') || '';
  const tk = limpiarToken_(p.getProperty('TOKEN'));
  const cod = p.getProperty('CODIGO_PRUEBA') || '';
  SpreadsheetApp.getUi().alert('Configuración guardada',
    'ID del conjunto de datos: ' + (id || '(vacío)') + '  (' + id.length + ' números)\n' +
    'Token: ' + (tk ? 'empieza con "' + tk.substring(0, 3) + '", tiene ' + tk.length + ' caracteres' : '(vacío)') + '\n' +
    'Código de prueba: ' + (cod || '(vacío)') + '\n\n' +
    'Lo normal: ID de 17 números, token que empieza con EAA y de más de 150 caracteres.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function leerConfig_() {
  const p = PropertiesService.getScriptProperties();
  const cfg = {
    datasetId: p.getProperty('DATASET_ID'),
    token: limpiarToken_(p.getProperty('TOKEN')),
    codigoPrueba: p.getProperty('CODIGO_PRUEBA')
  };
  if (!cfg.datasetId || !cfg.token) {
    throw new Error('Faltan los datos de Meta. Usa el menú Meta > 1. Configurar datos de Meta.');
  }
  return cfg;
}

// ---------- PRUEBA ----------
function enviarPrueba() {
  const ui = SpreadsheetApp.getUi();
  try {
    const cfg = leerConfig_();
    if (!cfg.codigoPrueba) {
      ui.alert('Falta el código de prueba. Ve a Meta > 1. Configurar y pégalo en el paso 3.');
      return;
    }
    const filas = leerPostulantes_();
    if (!filas.length) { ui.alert('No encontré postulantes en las hojas.'); return; }
    const p = filas[0];
    const res = enviarEvento_(cfg, p, 'Prueba de conexion', true);
    ui.alert(res.ok
      ? 'Prueba enviada. Ahora abre en Meta: Administrador de eventos > Probar eventos y confirma que aparece "Prueba de conexion".'
      : 'Meta respondió con error:\n' + res.detalle);
  } catch (e) {
    ui.alert('Error: ' + e.message);
  }
}

// ---------- AUTOMÁTICO ----------
function activarAutomatico() {
  const ui = SpreadsheetApp.getUi();
  try { leerConfig_(); } catch (e) { ui.alert(e.message); return; }

  desactivarAutomatico(true);
  ScriptApp.newTrigger('revisarYEnviar').timeBased().everyMinutes(5).create();
  ui.alert('Activado. Desde ahora se revisa cada 5 minutos.');
}

function desactivarAutomatico(silencioso) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'revisarYEnviar') ScriptApp.deleteTrigger(t);
  });
  if (silencioso !== true) SpreadsheetApp.getUi().alert('Envío automático desactivado.');
}

// ---------- REVISIÓN PRINCIPAL ----------
function revisarYEnviar() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return; // ya hay una revisión corriendo
  try {
    try { arreglarCabecerasTodasLasHojas_(); } catch (e) { /* si falla, igual seguimos con el envío a Meta */ }
    try { actualizarResumenPostulantes_(); } catch (e) { /* si falla, igual seguimos con el envío a Meta */ }
    try { publicarDashboard_(false); } catch (e) { /* si falla, igual seguimos con el envío a Meta */ }

    const cfg = leerConfig_();
    const registro = leerRegistro_();
    const filas = leerPostulantes_();
    const nuevos = [];

    filas.forEach(function (p) {
      EVENTOS.forEach(function (ev) {
        const valor = norm_(p.valores[norm_(ev.columna)]);
        if (ev.valores.map(norm_).indexOf(valor) === -1) return;
        const clave = p.leadId + '|' + ev.evento;
        if (registro.hechos[clave]) return;
        if ((registro.errores[clave] || 0) >= MAX_INTENTOS) {
          nuevos.push({ fila: [new Date(), clave, p.hoja, 'ABANDONADO', 'Falló ' + MAX_INTENTOS + ' veces'] });
          return;
        }
        const res = enviarEvento_(cfg, p, ev.evento, false);
        nuevos.push({ fila: [new Date(), clave, p.hoja, res.ok ? 'ENVIADO' : 'ERROR', res.detalle] });
      });
    });

    if (nuevos.length) escribirRegistro_(nuevos.map(function (n) { return n.fila; }));
  } finally {
    lock.releaseLock();
  }
}

// Envía a Meta TODO lo que ya está tipificado en las hojas (una sola vez por postulante y estado).
// Meta solo acepta eventos de los últimos 7 días, así que las fechas más antiguas se ajustan a
// "hace 6 días". Si quedan pendientes (por tiempo), vuelve a tocar la opción hasta que diga 0.
function enviarHistorial() {
  const ui = SpreadsheetApp.getUi();
  const inicio = Date.now();
  try {
    const cfg = leerConfig_();
    const registro = leerRegistro_();
    const pendientes = [];
    leerPostulantes_().forEach(function (p) {
      EVENTOS.forEach(function (ev) {
        if (ev.valores.map(norm_).indexOf(norm_(p.valores[norm_(ev.columna)])) === -1) return;
        const clave = p.leadId + '|' + ev.evento;
        if (registro.hechos[clave] || (registro.errores[clave] || 0) >= MAX_INTENTOS) return;
        pendientes.push({ p: p, ev: ev, clave: clave });
      });
    });

    if (!pendientes.length) { ui.alert('No hay nada pendiente por enviar.'); return; }
    let enviados = 0, fallidos = 0, hechos = 0;
    for (let i = 0; i < pendientes.length; i += LOTE) {
      if (Date.now() - inicio > 4.5 * 60 * 1000) break; // límite de tiempo de Google
      const lote = pendientes.slice(i, i + LOTE);
      hechos += lote.length;
      const eventos = lote.map(function (x) { return armarEvento_(x.p, x.ev.evento, tiempoHistorico_(x.p)); });
      const res = postearEventos_(cfg, eventos, false);
      const filas = [];
      if (res.ok) {
        lote.forEach(function (x) { filas.push([new Date(), x.clave, x.p.hoja, 'ENVIADO', 'Historial']); });
        enviados += lote.length;
      } else {
        // si el lote falla, se prueba de a uno para aislar el problema
        lote.forEach(function (x) {
          const r1 = postearEventos_(cfg, [armarEvento_(x.p, x.ev.evento, tiempoHistorico_(x.p))], false);
          filas.push([new Date(), x.clave, x.p.hoja, r1.ok ? 'ENVIADO' : 'ERROR', r1.detalle]);
          if (r1.ok) enviados++; else fallidos++;
        });
      }
      escribirRegistro_(filas);
    }
    const faltan = pendientes.length - hechos;
    ui.alert('Historial: ' + enviados + ' enviados, ' + fallidos + ' con error, ' + faltan +
      ' pendientes.' + (faltan > 0 ? '\nVuelve a tocar esta opción para seguir.' : ''));
  } catch (e) {
    ui.alert('Error: ' + e.message);
  }
}

// Fecha del evento para el historial: "Fecha Gestion" (o la fecha de creación), ajustada
// a los últimos días que Meta acepta.
function tiempoHistorico_(p) {
  const ahora = Date.now();
  const d = parseFecha_(p.valores['fecha gestion']) || parseFecha_(p.valores['created time']);
  let t = d ? d.getTime() : ahora;
  const minimo = ahora - DIAS_MAX * 86400000;
  if (t < minimo) t = minimo;
  if (t > ahora) t = ahora;
  return Math.floor(t / 1000);
}

function parseFecha_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v) ? null : v;
  const t = String(v).trim();
  const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); // dd/mm/aaaa
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0);
  const d = new Date(t);
  return isNaN(d) ? null : d;
}

// ---------- ARREGLAR POSTULANTES QUE CAEN ENCIMA DE LAS CABECERAS ----------
// Revisa una hoja: si hay filas por encima de la cabecera (fila con "id"), las clasifica:
//   - Postulante real (id empieza con "l:")  -> se mueve justo debajo de la cabecera,
//     copiando el formato y el menú desplegable de una fila de datos ya existente.
//   - Fila en blanco -> se elimina (no tenía datos).
//   - Cualquier otra cosa -> no se toca nada en esa hoja, por seguridad, y se avisa.
function procesarHojaCabecera_(hoja) {
  const datos = hoja.getDataRange().getValues();
  let filaCab = -1, colId = -1;
  for (let i = 0; i < datos.length && filaCab === -1; i++) {
    for (let j = 0; j < datos[i].length; j++) {
      if (norm_(datos[i][j]) === 'id') { filaCab = i; colId = j; break; }
    }
  }
  if (filaCab <= 0) return null; // no hay cabecera, o ya está en la primera fila: nada que hacer

  const numCols = datos[filaCab].length;
  const buenas = [];
  let blancos = 0;
  let desconocida = false;

  for (let r = 0; r < filaCab; r++) {
    const fila = datos[r];
    const id = String(fila[colId] || '').trim();
    if (id.indexOf('l:') === 0) {
      buenas.push(fila);
    } else if (fila.every(function (c) { return String(c === null || c === undefined ? '' : c).trim() === ''; })) {
      blancos++;
    } else {
      desconocida = true;
    }
  }

  if (desconocida) {
    return {
      hoja: hoja.getName(), ok: false,
      mensaje: 'Hay una fila encima de las cabeceras que no reconozco como postulante. No la toqué; revísala a mano.'
    };
  }

  const totalMalas = filaCab; // cantidad de filas por encima de la cabecera
  if (buenas.length > 0) {
    hoja.insertRowsAfter(filaCab + 1, buenas.length); // filaCab+1 = número de fila (1-based) de la cabecera
    const destino = hoja.getRange(filaCab + 2, 1, buenas.length, numCols);
    const huboDatosOriginales = datos.length > filaCab + 1; // ya había al menos una fila de datos

    if (huboDatosOriginales) {
      const filaPlantilla = filaCab + 2 + buenas.length; // la fila de datos original, ahora desplazada
      const plantilla = hoja.getRange(filaPlantilla, 1, 1, numCols);
      plantilla.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      plantilla.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    destino.setValues(buenas);
  }
  hoja.deleteRows(1, totalMalas);

  const partes = [];
  if (buenas.length) partes.push(buenas.length + ' postulante(s) reubicados debajo de las cabeceras');
  if (blancos) partes.push(blancos + ' fila(s) en blanco eliminadas');
  return { hoja: hoja.getName(), ok: true, mensaje: partes.join('; ') };
}

function arreglarCabecerasTodasLasHojas_() {
  const resultados = [];
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (hoja) {
    if (hoja.getName() === HOJA_REGISTRO || hoja.getName() === HOJA_LOG_ORDEN) return;
    const r = procesarHojaCabecera_(hoja);
    if (r) resultados.push(r);
  });
  if (resultados.length) {
    const h = hojaLogOrden_();
    const filas = resultados.map(function (r) { return [new Date(), r.hoja, r.ok ? 'OK' : 'AVISO', r.mensaje]; });
    h.getRange(h.getLastRow() + 1, 1, filas.length, 4).setValues(filas);
  }
  return resultados;
}

function hojaLogOrden_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA_LOG_ORDEN);
  if (!h) {
    h = ss.insertSheet(HOJA_LOG_ORDEN);
    h.appendRow(['Fecha', 'Hoja', 'Estado', 'Detalle']);
    h.hideSheet();
  }
  return h;
}

// Revisión manual, con aviso en pantalla (para probarlo cuando quieras, sin esperar los 5 min).
function arreglarCabecerasAhora() {
  const ui = SpreadsheetApp.getUi();
  try {
    const resultados = arreglarCabecerasTodasLasHojas_();
    if (!resultados.length) { ui.alert('Todo en orden. No encontré postulantes fuera de lugar.'); return; }
    const texto = resultados.map(function (r) { return r.hoja + ': ' + r.mensaje; }).join('\n');
    ui.alert('Revisión de cabeceras', texto, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error: ' + e.message);
  }
}

// ---------- RESUMEN POR MES Y POR CAMPAÑA ----------
// Busca la hoja "RESUMEN POSTULANTES" (por nombre exacto o parecido). Si no existe, no hace nada.
function hojaResumenPostulantes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const directa = ss.getSheetByName(HOJA_RESUMEN);
  if (directa) return directa;
  const parecida = ss.getSheets().filter(function (h) { return norm_(h.getName()) === norm_(HOJA_RESUMEN); });
  return parecida.length ? parecida[0] : null;
}

// A partir de una fecha (texto ISO o Date), da el mes en que cae: clave para ordenar,
// texto para mostrar (ej. "jun-2026") y un número para ordenar cronológicamente.
function mesDe_(valor) {
  if (!valor) return null;
  let d;
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    if (isNaN(valor)) return null;
    d = valor;
  } else {
    d = new Date(String(valor));
    if (isNaN(d)) return null;
  }
  const anio = d.getFullYear();
  const mesIdx = d.getMonth(); // 0 a 11
  return { texto: MESES_ES[mesIdx] + '-' + anio, orden: anio * 100 + (mesIdx + 1) };
}

function nuevoAcumuladorResumen_() {
  return { total: 0, aptos: 0, noAptos: 0, enProceso: 0, sinTipificar: 0, capacitacion: 0 };
}

function sumarAcumuladorResumen_(destino, g) {
  destino.total += g.total;
  destino.aptos += g.aptos;
  destino.noAptos += g.noAptos;
  destino.enProceso += g.enProceso;
  destino.sinTipificar += g.sinTipificar;
  destino.capacitacion += g.capacitacion;
}

// Recorre todas las hojas de postulantes y arma los grupos por mes + campaña.
// hojas: por defecto, todas las del documento (menos las ocultas y el propio resumen).
function calcularGruposResumen_(hojas) {
  const grupos = {};
  hojas.forEach(function (hoja) {
    const nombre = hoja.getName();
    if (nombre === HOJA_REGISTRO || nombre === HOJA_LOG_ORDEN || norm_(nombre) === norm_(HOJA_RESUMEN)) return;

    const datos = hoja.getDataRange().getValues();
    let filaCab = -1, colId = -1;
    for (let i = 0; i < datos.length && filaCab === -1; i++) {
      for (let j = 0; j < datos[i].length; j++) {
        if (norm_(datos[i][j]) === 'id') { filaCab = i; colId = j; break; }
      }
    }
    if (filaCab === -1) return; // hoja sin postulantes todavía (por ejemplo, una campaña recién creada)

    const cab = datos[filaCab].map(norm_);
    const colCreado = cab.indexOf('created time');
    const colEstatus = cab.indexOf(norm_('Estatus de R&S'));
    const colCapacitacion = cab.indexOf(norm_('Ingreso a Capacitación'));

    for (let i = 0; i < datos.length; i++) {
      if (i === filaCab) continue;
      const id = String(datos[i][colId] || '').trim();
      if (id.indexOf('l:') !== 0) continue;

      const mes = colCreado !== -1 ? mesDe_(datos[i][colCreado]) : null;
      if (!mes) continue; // sin fecha reconocible: no se puede ubicar en un mes

      const estatus = colEstatus !== -1 ? norm_(datos[i][colEstatus]).toUpperCase() : '';
      const capacita = colCapacitacion !== -1 && norm_(datos[i][colCapacitacion]) === 'si';

      const clave = mes.orden + '|' + nombre;
      if (!grupos[clave]) {
        grupos[clave] = Object.assign({ mesTexto: mes.texto, mesOrden: mes.orden, campana: nombre }, nuevoAcumuladorResumen_());
      }
      const g = grupos[clave];
      g.total++;
      if (estatus === 'APTO') g.aptos++;
      else if (estatus === 'NO APTO') g.noAptos++;
      else if (estatus === 'EN PROCESO') g.enProceso++;
      else g.sinTipificar++;
      if (capacita) g.capacitacion++;
    }
  });
  return grupos;
}

function filaResumen_(mesTexto, campana, g) {
  const tipificados = g.total - g.sinTipificar;
  return [
    mesTexto, campana, g.total, g.sinTipificar, g.enProceso, g.aptos, g.noAptos, g.capacitacion,
    g.total ? tipificados / g.total : 0,
    tipificados ? g.aptos / tipificados : 0,
    g.total ? g.aptos / g.total : 0,
    g.aptos ? g.capacitacion / g.aptos : 0
  ];
}

const ANCHOS_COLUMNAS_RESUMEN = [95, 210, 105, 120, 105, 85, 95, 150, 120, 150, 120, 165];

// Escribe el cuadro en la hoja de resumen: un título, una fila por mes+campaña, un subtotal
// por mes y un total general al final. Los porcentajes se colorean solos (rojo a verde) para
// verlos de un vistazo, y la columna "Sin tipificar" se resalta cuando hay pendientes.
// Devuelve cuántas combinaciones mes+campaña se pintaron.
function pintarResumen_(hoja, grupos) {
  const lista = Object.keys(grupos).map(function (k) { return grupos[k]; });
  lista.sort(function (a, b) {
    if (a.mesOrden !== b.mesOrden) return a.mesOrden - b.mesOrden;
    return a.campana.localeCompare(b.campana, 'es');
  });

  const encabezado = ['Mes', 'Campaña', 'Postulantes', 'Sin tipificar', 'En proceso', 'Aptos', 'No aptos',
    'Ingreso a\nCapacitación', '% Tipificados', '% Aptos /\ntipificados', '% Aptos /\ntotal', '% Capacitación /\naptos'];
  const numCols = encabezado.length;

  hoja.clear();
  hoja.clearConditionalFormatRules();
  if (!lista.length) {
    hoja.getRange(1, 1).setValue('Todavía no hay postulantes tipificados para mostrar.');
    return 0;
  }

  const filas = [];
  let mesActual = null;
  let acumMes = nuevoAcumuladorResumen_();
  const totalGeneral = nuevoAcumuladorResumen_();

  lista.forEach(function (g) {
    if (mesActual !== null && g.mesTexto !== mesActual) {
      filas.push(filaResumen_('Subtotal ' + mesActual, '', acumMes));
      acumMes = nuevoAcumuladorResumen_();
    }
    mesActual = g.mesTexto;
    filas.push(filaResumen_(g.mesTexto, g.campana, g));
    sumarAcumuladorResumen_(acumMes, g);
    sumarAcumuladorResumen_(totalGeneral, g);
  });
  if (mesActual !== null) filas.push(filaResumen_('Subtotal ' + mesActual, '', acumMes));
  filas.push(filaResumen_('TOTAL GENERAL', '', totalGeneral));

  // Fila 1: título. Fila 2: encabezado. Desde la fila 3: datos.
  const filaEncabezado = 2;
  const filaPrimerDato = 3;
  const totalFilasDatos = filas.length;

  hoja.getRange(1, 1, 1, numCols).merge()
    .setValue('Resumen de postulantes — por mes y por campaña')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#0b2545').setFontColor('#ffffff');
  hoja.setRowHeight(1, 34);

  hoja.getRange(filaEncabezado, 1, 1, numCols).setValues([encabezado]);
  hoja.getRange(filaPrimerDato, 1, totalFilasDatos, numCols).setValues(filas);

  hoja.getRange(filaEncabezado, 1, 1, numCols)
    .setFontWeight('bold').setBackground('#1c3d5a').setFontColor('#ffffff')
    .setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  hoja.setRowHeight(filaEncabezado, 38);
  hoja.setFrozenRows(filaEncabezado);

  ANCHOS_COLUMNAS_RESUMEN.forEach(function (ancho, i) { hoja.setColumnWidth(i + 1, ancho); });

  const rangoDatos = hoja.getRange(filaPrimerDato, 1, totalFilasDatos, numCols);
  rangoDatos.setHorizontalAlignment('center').setVerticalAlignment('middle');
  hoja.getRange(filaPrimerDato, 1, totalFilasDatos, 2).setHorizontalAlignment('left'); // Mes y Campaña, a la izquierda

  const rangoConteos = hoja.getRange(filaPrimerDato, 3, totalFilasDatos, 6);
  rangoConteos.setNumberFormat('#,##0');
  const rangoPorcentajes = hoja.getRange(filaPrimerDato, 9, totalFilasDatos, 4);
  rangoPorcentajes.setNumberFormat('0%');

  // Bordes finos en toda la tabla, para que se lea como cuadro y no como texto suelto.
  hoja.getRange(filaEncabezado, 1, totalFilasDatos + 1, numCols)
    .setBorder(true, true, true, true, true, true, '#c9c9c9', SpreadsheetApp.BorderStyle.SOLID);

  // Filas de subtotal y total: negrita, fondo distinto y una línea arriba para separarlas.
  for (let r = 0; r < filas.length; r++) {
    const esResumenFila = String(filas[r][0]).indexOf('Subtotal') === 0 || filas[r][0] === 'TOTAL GENERAL';
    if (esResumenFila) {
      const filaHoja = filaPrimerDato + r;
      hoja.getRange(filaHoja, 1, 1, numCols)
        .setFontWeight('bold').setBackground(filas[r][0] === 'TOTAL GENERAL' ? '#c6dcf0' : '#e8eef7')
        .setBorder(true, null, null, null, null, null, '#6b8cae', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  // Colores automáticos: verde = mejor, rojo = peor, en los 4 porcentajes.
  const reglaPorcentajes = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint('#f4a7a2')
    .setGradientMidpoint('#ffe599')
    .setGradientMaxpoint('#93c47d')
    .setRanges([rangoPorcentajes])
    .build();

  // "Sin tipificar" en naranja cuando hay postulantes pendientes de revisar.
  const rangoSinTipificar = hoja.getRange(filaPrimerDato, 4, totalFilasDatos, 1);
  const reglaSinTipificar = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#fce5cd')
    .setRanges([rangoSinTipificar])
    .build();

  hoja.setConditionalFormatRules([reglaPorcentajes, reglaSinTipificar]);

  hoja.setFrozenColumns(2); // Mes y Campaña siempre visibles al mover la tabla hacia la derecha
  return lista.length;
}

function actualizarResumenPostulantes_() {
  const hojaResumen = hojaResumenPostulantes_();
  if (!hojaResumen) return; // si no existe esa hoja, no se crea sola: la debes tener ya hecha
  const grupos = calcularGruposResumen_(SpreadsheetApp.getActiveSpreadsheet().getSheets());
  pintarResumen_(hojaResumen, grupos);
}

// Revisión manual, con aviso en pantalla.
function actualizarResumenAhora() {
  const ui = SpreadsheetApp.getUi();
  try {
    const hojaResumen = hojaResumenPostulantes_();
    if (!hojaResumen) {
      ui.alert('No encontré una hoja llamada "' + HOJA_RESUMEN + '". Créala primero (puede estar vacía) y vuelve a intentar.');
      return;
    }
    const grupos = calcularGruposResumen_(SpreadsheetApp.getActiveSpreadsheet().getSheets());
    const n = pintarResumen_(hojaResumen, grupos);
    ui.alert('Resumen actualizado: ' + n + ' combinaciones de mes y campaña.');
  } catch (e) {
    ui.alert('Error: ' + e.message);
  }
}

// ---------- DASHBOARD EN GITHUB ----------
// Publica en GitHub un archivo "data.json" con los postulantes SIN datos personales
// (solo fecha, hora, oferta, anuncio, plataforma, estado, motivo de descarte y capacitación).
// La página del dashboard (index.html en GitHub Pages) lee ese archivo y dibuja los gráficos.
// Solo sube un archivo nuevo cuando algo cambió, así no llena el repositorio de versiones.
const GH_ARCHIVO = 'data.json';
const GH_RAMA = 'main';

function configurarDashboard() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const u = ui.prompt('Dashboard 1 de 3', 'Tu usuario de GitHub (solo el nombre, sin @):', ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  const r = ui.prompt('Dashboard 2 de 3', 'Nombre del repositorio (ej. dashboard-postulantes):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const t = ui.prompt('Dashboard 3 de 3', 'Pega el token de GitHub (empieza con github_pat_ o ghp_):', ui.ButtonSet.OK_CANCEL);
  if (t.getSelectedButton() !== ui.Button.OK) return;

  const usuario = u.getResponseText().replace(/[\s@]/g, '');
  const repo = r.getResponseText().replace(/\s/g, '');
  const token = t.getResponseText().replace(/[\s"']/g, '');
  if (!usuario || !repo || !token) { ui.alert('Falta algún dato. Vuelve a intentar.'); return; }
  props.setProperties({ GH_USUARIO: usuario, GH_REPO: repo, GH_TOKEN: token });
  props.deleteProperty('GH_HASH');
  ui.alert('Listo. Datos guardados. Ahora usa "Dashboard: publicar ahora" para probar.');
}

// Arma la lista de postulantes SIN datos personales.
function construirDatosDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const filas = [];
  ss.getSheets().forEach(function (hoja) {
    const nombre = hoja.getName();
    if (nombre === HOJA_REGISTRO || nombre === HOJA_LOG_ORDEN || norm_(nombre) === norm_(HOJA_RESUMEN)) return;

    const datos = hoja.getDataRange().getValues();
    let filaCab = -1, colId = -1;
    for (let i = 0; i < datos.length && filaCab === -1; i++) {
      for (let j = 0; j < datos[i].length; j++) {
        if (norm_(datos[i][j]) === 'id') { filaCab = i; colId = j; break; }
      }
    }
    if (filaCab === -1) return;

    const cab = datos[filaCab].map(norm_);
    const cCre = cab.indexOf(norm_('created_time'));
    const cAnu = cab.indexOf(norm_('ad_name'));
    const cPla = cab.indexOf(norm_('platform'));
    const cEst = cab.indexOf(norm_('Estatus de R&S'));
    const cMot = cab.indexOf(norm_('Motivo de descarte'));
    const cCap = cab.indexOf(norm_('Ingreso a Capacitación'));
    if (cCre === -1) return;
    // Preguntas del formulario (no todas las hojas las tienen: las presenciales no preguntan por internet)
    const cInt = cab.findIndex(function (c) { return c.indexOf('internet') !== -1; });
    const cCv = cab.findIndex(function (c) { return c.indexOf('dejanos tu cv') !== -1; });
    const cDni = cab.findIndex(function (c) { return c.indexOf('dni') === 0; });

    for (let i = 0; i < datos.length; i++) {
      if (i === filaCab) continue;
      if (String(datos[i][colId] || '').trim().indexOf('l:') !== 0) continue;
      if (datos[i].join('|').indexOf('<test lead') !== -1) continue; // pruebas de Meta, no son postulantes reales

      let v = datos[i][cCre];
      if (Object.prototype.toString.call(v) === '[object Date]') {
        if (isNaN(v)) continue;
        v = Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ss");
      }
      v = String(v);
      if (!/^\d{4}-\d{2}-\d{2}/.test(v)) continue;
      const hora = parseInt(v.substring(11, 13), 10);

      const est = cEst !== -1 ? norm_(datos[i][cEst]).toUpperCase() : '';
      const estado = (est === 'APTO' || est === 'NO APTO' || est === 'EN PROCESO') ? est : 'SIN TIPIFICAR';

      // Requisitos de las primeras preguntas: 1 = cumple, 0 = no cumple, -1 = la hoja no tiene esa pregunta
      let internet = -1;
      if (cInt !== -1) {
        const t = norm_(datos[i][cInt]);
        if (t !== '') internet = t.indexOf('no cuento') !== -1 ? 0 : 1;
      }
      let cv = -1;
      if (cCv !== -1) cv = /^https?:\/\//i.test(String(datos[i][cCv]).trim()) ? 1 : 0;
      let dni = -1;
      if (cDni !== -1) {
        let dig = String(datos[i][cDni]).replace(/\D/g, '');
        if (typeof datos[i][cDni] === 'number') dig = dig.padStart(8, '0'); // el 0 inicial se pierde si la celda es número
        dni = dig.length === 8 ? 1 : 0;
      }

      filas.push({
        f: v.substring(0, 10),
        h: isNaN(hora) ? -1 : hora,
        o: nombre,
        a: cAnu !== -1 ? String(datos[i][cAnu] || '').trim().substring(0, 80) : '',
        p: cPla !== -1 ? String(datos[i][cPla] || '').trim().toLowerCase() : '',
        e: estado,
        m: (estado === 'NO APTO' && cMot !== -1) ? String(datos[i][cMot] || '').trim().substring(0, 80) : '',
        k: (cCap !== -1 && norm_(datos[i][cCap]) === 'si') ? 1 : 0,
        i: internet,
        v: cv,
        d: dni
      });
    }
  });
  filas.sort(function (a, b) { return a.f < b.f ? -1 : a.f > b.f ? 1 : (a.o < b.o ? -1 : a.o > b.o ? 1 : 0); });
  return filas;
}

// Sube data.json a GitHub. Devuelve { ok, detalle }.
function publicarDashboard_(forzar) {
  const props = PropertiesService.getScriptProperties();
  const usuario = props.getProperty('GH_USUARIO'), repo = props.getProperty('GH_REPO'), token = props.getProperty('GH_TOKEN');
  if (!usuario || !repo || !token) return { ok: false, detalle: 'Falta configurar GitHub (menú Meta > Dashboard: configurar).' };

  const textoFilas = JSON.stringify(construirDatosDashboard_());
  const huella = hash_(textoFilas);
  if (!forzar && props.getProperty('GH_HASH') === huella) return { ok: true, detalle: 'Sin cambios desde la última publicación.' };

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const json = '{"actualizado":' + JSON.stringify(Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ssXXX")) + ',"filas":' + textoFilas + '}';

  const base = 'https://api.github.com/repos/' + usuario + '/' + repo + '/contents/' + GH_ARCHIVO;
  const cab = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };
  try {
    const g = UrlFetchApp.fetch(base + '?ref=' + GH_RAMA, { method: 'get', headers: cab, muteHttpExceptions: true });
    let sha = null;
    if (g.getResponseCode() === 200) sha = JSON.parse(g.getContentText()).sha;
    else if (g.getResponseCode() !== 404) return { ok: false, detalle: errorGithub_(g) };

    const cuerpo = {
      message: 'Actualización automática de postulantes',
      content: Utilities.base64Encode(Utilities.newBlob(json).getBytes()),
      branch: GH_RAMA
    };
    if (sha) cuerpo.sha = sha;
    const r = UrlFetchApp.fetch(base, { method: 'put', headers: cab, contentType: 'application/json', payload: JSON.stringify(cuerpo), muteHttpExceptions: true });
    const c = r.getResponseCode();
    if (c === 200 || c === 201) {
      props.setProperty('GH_HASH', huella);
      return { ok: true, detalle: 'Publicado: ' + JSON.parse(textoFilas).length + ' postulantes.' };
    }
    return { ok: false, detalle: errorGithub_(r) };
  } catch (e) {
    return { ok: false, detalle: 'Error de conexión: ' + e.message };
  }
}

function errorGithub_(resp) {
  const c = resp.getResponseCode();
  if (c === 401) return 'GitHub dice que el token no sirve (venció o está mal pegado). Crea otro y repite "Dashboard: configurar".';
  if (c === 403) return 'GitHub rechazó el permiso. Revisa que el token tenga "Contents: Read and write" en ese repositorio.';
  if (c === 404) return 'No encuentro el repositorio. Revisa el usuario y el nombre, y que el token tenga acceso a ese repositorio.';
  if (c === 409 || c === 422) return 'GitHub no pudo guardar el archivo (HTTP ' + c + '). Revisa que el repositorio tenga la rama "' + GH_RAMA + '". ' + resp.getContentText().substring(0, 150);
  return 'HTTP ' + c + ' ' + resp.getContentText().substring(0, 200);
}

function publicarDashboardAhora() {
  const ui = SpreadsheetApp.getUi();
  const res = publicarDashboard_(true);
  ui.alert(res.ok ? 'Listo. ' + res.detalle : 'No se pudo publicar. ' + res.detalle);
}

// ---------- LECTURA DE HOJAS (solo lectura) ----------
// Busca en cada hoja la fila de cabeceras (la que tiene "id" en alguna celda) y toma
// como postulante toda fila cuyo id empiece con "l:". Así funciona aunque las cabeceras
// varíen entre hojas, y aunque un postulante haya caído por encima de las cabeceras.
function leerPostulantes_() {
  const salida = [];
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (hoja) {
    if (hoja.getName() === HOJA_REGISTRO || hoja.getName() === HOJA_LOG_ORDEN) return;
    const datos = hoja.getDataRange().getValues();
    let filaCab = -1, colId = -1;
    for (let i = 0; i < datos.length && filaCab === -1; i++) {
      for (let j = 0; j < datos[i].length; j++) {
        if (norm_(datos[i][j]) === 'id') { filaCab = i; colId = j; break; }
      }
    }
    if (filaCab === -1) return;

    const cab = datos[filaCab].map(norm_);
    for (let i = 0; i < datos.length; i++) {
      if (i === filaCab) continue;
      const id = String(datos[i][colId] || '').trim();
      if (id.indexOf('l:') !== 0) continue;
      const valores = {};
      cab.forEach(function (nombre, k) { if (nombre) valores[nombre] = datos[i][k]; });
      salida.push({
        hoja: hoja.getName(),
        leadId: id.substring(2),
        correo: buscar_(valores, ['correo']),
        telefono: buscar_(valores, ['numero de telefono', 'numero de whatsapp']),
        valores: valores
      });
    }
  });
  return salida;
}

function buscar_(valores, pistas) {
  for (let k = 0; k < pistas.length; k++) {
    for (const nombre in valores) {
      if (nombre.indexOf(pistas[k]) !== -1 && String(valores[nombre]).trim() !== '') return String(valores[nombre]);
    }
  }
  return '';
}

// ---------- ENVÍO A META ----------
function armarEvento_(p, nombreEvento, tiempo) {
  const usuario = { lead_id: String(p.leadId) };
  const em = p.correo ? hash_(p.correo.trim().toLowerCase()) : '';
  const digitos = String(p.telefono || '').replace(/\D/g, '');
  const ph = digitos ? hash_(digitos) : '';
  if (em) usuario.em = [em];
  if (ph) usuario.ph = [ph];
  return {
    event_name: nombreEvento,
    event_time: tiempo || Math.floor(Date.now() / 1000),
    event_id: p.leadId + '_' + nombreEvento,
    action_source: 'system_generated',
    user_data: usuario,
    custom_data: { event_source: 'crm', lead_event_source: NOMBRE_FUENTE }
  };
}

function enviarEvento_(cfg, p, nombreEvento, esPrueba) {
  return postearEventos_(cfg, [armarEvento_(p, nombreEvento)], esPrueba);
}

function postearEventos_(cfg, eventos, esPrueba) {
  const cuerpo = { data: eventos };
  if (esPrueba && cfg.codigoPrueba) cuerpo.test_event_code = cfg.codigoPrueba;

  const url = 'https://graph.facebook.com/' + VERSION_API + '/' + cfg.datasetId + '/events?access_token=' + encodeURIComponent(cfg.token);
  try {
    const r = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(cuerpo),
      muteHttpExceptions: true
    });
    const codigo = r.getResponseCode();
    const texto = r.getContentText();
    if (codigo === 200) return { ok: true, detalle: 'OK ' + texto.substring(0, 150) };
    return { ok: false, detalle: 'HTTP ' + codigo + ' ' + texto.substring(0, 300) };
  } catch (e) {
    return { ok: false, detalle: 'Error de conexión: ' + e.message };
  }
}

function hash_(texto) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// ---------- REGISTRO (hoja oculta) ----------
function hojaRegistro_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA_REGISTRO);
  if (!h) {
    h = ss.insertSheet(HOJA_REGISTRO);
    h.appendRow(['Fecha', 'Clave (lead|evento)', 'Hoja', 'Estado', 'Detalle']);
    h.hideSheet();
  }
  return h;
}

function leerRegistro_() {
  const h = hojaRegistro_();
  const datos = h.getDataRange().getValues();
  const hechos = {}, errores = {};
  for (let i = 1; i < datos.length; i++) {
    const clave = datos[i][1], estado = datos[i][3];
    if (estado === 'ENVIADO' || estado === 'BASE' || estado === 'ABANDONADO') hechos[clave] = true;
    else if (estado === 'ERROR') errores[clave] = (errores[clave] || 0) + 1;
  }
  return { hechos: hechos, errores: errores };
}

function escribirRegistro_(filas) {
  const h = hojaRegistro_();
  h.getRange(h.getLastRow() + 1, 1, filas.length, 5).setValues(filas);
}

// ---------- UTILIDAD ----------
function norm_(s) {
  return String(s === null || s === undefined ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[_\\]/g, ' ').replace(/\s+/g, ' ').trim();
}
