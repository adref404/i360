/** Utilidades compartidas. Todas terminan en "_" para que el navegador no pueda llamarlas. */

let _BASE = null;
function ss_() {
  if (!_BASE) _BASE = CONFIG.baseId ? SpreadsheetApp.openById(CONFIG.baseId) : SpreadsheetApp.getActive();
  if (!_BASE) throw new Error('No se encontró el Google Sheet base. Completa CONFIG.baseId en Config.gs.');
  return _BASE;
}

/** Devuelve la hoja; si no existe y se dan encabezados, la crea. */
function hoja_(nombre, encabezados) {
  const ss = ss_();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    if (!encabezados) return null;
    sh = ss.insertSheet(nombre);
  }
  if (encabezados && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Color de Sheets ("#FFF", "#ffffff", "white"…) → "#rrggbb" en minúsculas, o "" si no es válido. */
function colorHex_(c) {
  let s = String(c || '').trim().toLowerCase();
  if (s === 'white') return '#ffffff';
  if (s === 'black') return '#000000';
  if (/^#[0-9a-f]{3}$/.test(s)) s = '#' + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
  return /^#[0-9a-f]{6}$/.test(s) ? s : '';
}

/** Agrega al final de la fila 1 los encabezados que falten (no mueve columnas existentes). */
function asegurarColumnas_(sh, nombres) {
  const lc = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lc).getValues()[0].map(normTxt_);
  const faltan = nombres.filter(h => head.indexOf(normTxt_(h)) < 0);
  if (faltan.length) sh.getRange(1, lc + 1, 1, faltan.length).setValues([faltan]);
  return sh;
}

/** Escribe en una fila los valores indicados por nombre de encabezado. */
function escribirFila_(sh, fila, valores) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(normTxt_);
  Object.keys(valores).forEach(nombre => {
    const i = head.indexOf(normTxt_(nombre));
    if (i < 0) throw new Error('Falta la columna "' + nombre + '" en la hoja ' + sh.getName());
    sh.getRange(fila, i + 1).setValue(valores[nombre]);
  });
}

/** Lee una hoja completa: encabezados normalizados y filas de valores. */
function filas_(sh) {
  if (!sh) return { head: [], rows: [] };
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return { head: [], rows: [] };
  const v = sh.getRange(1, 1, lr, lc).getValues();
  return { head: v[0].map(normTxt_), rows: v.slice(1) };
}

/** Índice de columna por nombre de encabezado (exacto primero, luego "empieza con"). */
function col_(head) {
  const nombres = Array.prototype.slice.call(arguments, 1).map(normTxt_);
  for (const n of nombres) { const i = head.indexOf(n); if (i >= 0) return i; }
  for (const n of nombres) { const i = head.findIndex(h => h && h.indexOf(n) === 0); if (i >= 0) return i; }
  return -1;
}
function celda_(row, i) { return i >= 0 && i < row.length ? row[i] : ''; }

function normTxt_(s) {
  return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
}
function esFecha_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }

/** BP como texto: 2713993, "2713993.0" o " 2713993 " -> "2713993". */
function bpTxt_(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return String(Math.round(v));
  return String(v).trim().replace(/\.0+$/, '').replace(/[^0-9A-Za-z]/g, '');
}

function fecha_(d, fmt) { return Utilities.formatDate(d, CONFIG.zona, fmt || 'yyyy-MM-dd'); }
function hoyISO_() { return fecha_(new Date()); }
function mesActual_() { return fecha_(new Date(), 'yyyy-MM'); }
function ahoraISO_() { return new Date().toISOString(); }
function diasDelMes_(mes) { const p = mes.split('-').map(Number); return new Date(p[0], p[1], 0).getDate(); }

/** "2026-11", fecha 01/11/2026 o "11/2026" -> "2026-11". */
function mesTxt_(v) {
  if (esFecha_(v)) return fecha_(v, 'yyyy-MM');
  const s = String(v == null ? '' : v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  m = s.match(/^(\d{1,2})\/(\d{4})$/);
  if (m) return m[2] + '-' + ('0' + m[1]).slice(-2);
  return '';
}

/** Fecha de celda -> "yyyy-MM-dd" (o el texto tal cual, p. ej. "VENCIDO"). */
function fechaISO_(v) {
  if (esFecha_(v)) return fecha_(v);
  const s = String(v == null ? '' : v).trim();
  if (!s || s.toUpperCase() === 'NULL') return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s;
}

function conBloqueo_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}

/** Otro usuario ya resolvió lo que se intenta hacer (la pantalla estaba desactualizada). */
function yaResuelto_(que, accion, quien, cuando) {
  let f = '';
  try { if (cuando) f = Utilities.formatDate(new Date(cuando), CONFIG.zona, "dd/MM 'a las' HH:mm"); } catch (e) { /* sin fecha */ }
  return falla_(que + ' ya fue ' + accion + (quien ? ' por ' + quien : '') + (f ? ' el ' + f : '') + '. Tu pantalla no estaba actualizada: ya la recargamos.', 'YA_RESUELTO');
}

/** Dirección de la app para los correos: CONFIG.appUrl o, si falta, la del despliegue actual. */
function urlApp_() {
  // En el ambiente de prueba, los botones de los correos abren la app de prueba (nunca la real)
  if (enPrueba_()) { const u = String(props_().getProperty('PRUEBA_URL') || ''); if (u) return u; try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }
  if (CONFIG.appUrl) return CONFIG.appUrl;
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}
/** Botón "Abrir i360" para el cuerpo HTML de los correos. */
function botonAppHtml_(texto) {
  const u = urlApp_(); if (!u) return '';
  return '<p style="margin:22px 0"><a href="' + u + '" style="background:#1B0088;color:#ffffff;text-decoration:none;font-weight:bold;padding:11px 20px;border-radius:8px;display:inline-block">' +
    (texto || 'Abrir ' + CONFIG.nombreApp) + '</a></p>';
}

/** Error que sí se muestra al usuario. */
function falla_(mensaje, codigo) {
  const e = new Error(mensaje);
  e.publico = true;
  e.codigo = codigo || '';
  return e;
}

/**
 * Envoltura de todas las funciones públicas: el cliente siempre recibe
 * {ok:true, data} o {ok:false, error, codigo}. Las fechas se convierten a texto
 * (google.script.run no admite objetos Date en la respuesta).
 */
function responder_(fn) {
  try {
    const data = fn();
    return { ok: true, data: data === undefined ? null : JSON.parse(JSON.stringify(data)) };
  } catch (e) {
    if (e && e.publico) return { ok: false, error: e.message, codigo: e.codigo || '' };
    console.error(e && e.stack ? e.stack : e);
    return { ok: false, error: 'Ocurrió un error inesperado. Intenta de nuevo en unos segundos.', codigo: 'INTERNO' };
  }
}

/** Solo para funciones que se ejecutan desde el editor de Apps Script. */
function soloEditor_() {
  const a = Session.getActiveUser().getEmail();
  const e = Session.getEffectiveUser().getEmail();
  if (!a || a !== e) throw new Error('Esta función solo puede ejecutarse desde el editor de Apps Script por la cuenta dueña.');
}

/* ---------------- Ambiente de prueba ----------------
 * Se activa SOLO en el proyecto de prueba con la propiedad del script MODO = PRUEBA.
 * El código es el mismo en ambos proyectos; el real no tiene esa propiedad.
 *  - Correos: llegan solo a PRUEBA_CORREO (o al dueño del script), con "[PRUEBA]" y a quién habrían llegado.
 *  - Rol del mes: los intercambios aprobados se aplican solo en la copia de la app (hoja Roles) y no se repinta
 *    el Google Sheet, salvo que la propiedad PRUEBA_ESCRIBIR = SI (útil con copias de los Sheets de los meses).
 */
function enPrueba_() { try { return String(props_().getProperty('MODO') || '').trim().toUpperCase() === 'PRUEBA'; } catch (e) { return false; } }
function pruebaEscribeRol_() { return String(props_().getProperty('PRUEBA_ESCRIBIR') || '').trim().toUpperCase() === 'SI'; }
function correoPrueba_(dest, asunto, html, texto) {
  dest = [].concat(dest || []).map(String).filter(Boolean);
  if (!enPrueba_()) return { dest: dest, asunto: asunto, html: html, texto: texto };
  let a = String(props_().getProperty('PRUEBA_CORREO') || '');
  if (!a) try { a = Session.getEffectiveUser().getEmail(); } catch (e) { a = ''; }
  const lista = dest.join(', ') || '(nadie)';
  return {
    dest: a.split(/[,;\s]+/).filter(Boolean), asunto: '[PRUEBA] ' + asunto,
    html: '<div style="background:#FFF4D6;border:1px solid #E8B200;border-radius:8px;padding:8px 12px;margin-bottom:12px;font:13px Arial,sans-serif;color:#5C4500">' +
      '<b>Ambiente de prueba.</b> En la app real este correo habría llegado a: ' + escHtml_(lista) + '</div>' + (html || ''),
    texto: '[PRUEBA] En la app real habría llegado a: ' + lista + '\n\n' + (texto || '')
  };
}
/** MailApp.sendEmail que respeta el ambiente de prueba. */
function enviarCorreo_(o) {
  const c = correoPrueba_(String(o.to || '').split(','), o.subject, o.htmlBody || '', o.body || '');
  if (!c.dest.length) return;
  MailApp.sendEmail(Object.assign({}, o, { to: c.dest.join(','), subject: c.asunto, htmlBody: c.html, body: c.texto }));
}
