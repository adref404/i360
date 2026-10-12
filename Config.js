/**
 * Rol de instructores LP — Configuración y punto de entrada de la aplicación web.
 *
 * Estructura del proyecto:
 *   Config.gs    Configuración, doGet() e include()
 *   Util.gs      Utilidades de hojas, fechas y respuestas al cliente
 *   Auth.gs      Inicio de sesión, contraseñas, códigos temporales, sesiones y registro de accesos
 *   Fuentes.gs   Lectura de los Google Sheets mensuales (ROL INS) y sincronización
 *   Datos.gs     Datos para el panel (filtrados por rol) y respuestas de instructores
 *   Setup.gs     instalar() y utilidades que se ejecutan desde el editor
 *   Semilla.gs   Datos iniciales de Solicitudes y Vigencias (BG INS)
 *   Vacaciones.gs        Vacaciones por periodo DGAC: hojas, ventana de solicitud, validación y aprobación
 *   SemillaVacaciones.gs Datos iniciales de vacaciones (Vacaciones_INSTRUCTORES_1.xlsx)
 *   Index, Estilos, Clasificador, App, VacacionesVista, Vistas, Coordinacion, Admin (HTML)   Interfaz
 */

const CONFIG = {
  // ID del Google Sheet base (Usuarios, Roles, Respuestas, Registro, Permisos...).
  // Vacío = el Google Sheet al que está vinculado este script.
  baseId: '',

  zona: 'America/Lima',
  nombreApp: 'Rol de instructores LP',

  hojas: {
    usuarios: 'Usuarios',
    roles: 'Roles',
    respuestas: 'Respuestas',
    registro: 'Registro',
    permisos: 'Permisos',
    fuentes: 'Fuentes',
    solicitudes: 'Solicitudes',
    vigencias: 'Vigencias',
    vacaciones: 'Vacaciones',
    vacDias: 'VacacionesDias',
    vacBloques: 'VacacionesBloques',
    solVac: 'SolicitudesVacaciones'
  },

  // Sesiones
  sesionHoras: 12,          // duración de una sesión normal
  sesionRecordarDias: 30,   // duración con "Mantener mi sesión iniciada"

  // Protección contra fuerza bruta
  intentosMax: 5,           // intentos fallidos antes de bloquear
  bloqueoMin: 15,           // minutos de bloqueo

  // Códigos temporales (primer ingreso o restablecer contraseña)
  codigoHoras: 48,

  // Iteraciones del hash de contraseñas (HMAC-SHA256 con sal y pimienta)
  hashIter: 200,

  // true solo si vas a incrustar la app en Google Sites u otra página.
  permitirIframe: false,

  // Google Sheets mensuales que se registran en la hoja Fuentes al ejecutar instalar().
  fuentesIniciales: [
    { mes: '2026-11', enlace: 'https://docs.google.com/spreadsheets/d/1B0IizPnZybvm3cm1nnRhKSUHGl0WcEAmeiK3gvEaqxw/edit?gid=681297197#gid=681297197', hoja: 'ROL INS' }
  ]
};

const VERSION_APP = '2.4';
const ARCHIVOS_HTML_ = ['Estilos', 'Clasificador', 'App', 'VacacionesVista', 'Vistas', 'Coordinacion', 'Admin'];

function doGet(e) {
  // https://…/exec?diag=1  ->  diagnóstico en texto (sin datos personales)
  if (e && e.parameter && e.parameter.diag) {
    return ContentService.createTextOutput(diagnostico_()).setMimeType(ContentService.MimeType.TEXT);
  }
  const t = HtmlService.createTemplateFromFile('Index');
  t.version = VERSION_APP;
  const out = t.evaluate()
    .setTitle(CONFIG.nombreApp)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  if (CONFIG.permitirIframe) out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}

/** Revisa la instalación paso a paso. Solo muestra conteos y estados, nunca nombres ni contraseñas. */
function diagnostico_() {
  const L = ['Rol de instructores LP · diagnóstico · versión ' + VERSION_APP, 'Fecha del servidor: ' + new Date().toISOString(), ''];
  const paso = (nombre, fn) => { try { L.push('OK    ' + nombre + (function (r) { return r ? ' — ' + r : ''; })(fn())); } catch (err) { L.push('ERROR ' + nombre + ' — ' + (err && err.message)); } };
  paso('Archivo HTML "Index"', () => { const c = HtmlService.createHtmlOutputFromFile('Index').getContent(); if (c.indexOf("include('VacacionesVista')") < 0 || c.indexOf('Index v' + VERSION_APP) < 0) throw new Error('es una versión anterior; vuelve a copiarlo desde el zip v' + VERSION_APP); return Math.round(c.length / 1024) + ' KB'; });
  ARCHIVOS_HTML_.forEach(n => paso('Archivo HTML "' + n + '"', () => revisarArchivoHtml_(n)));
  paso('Google Sheet base', () => ss_().getName());
  Object.keys(CONFIG.hojas).forEach(k => paso('Hoja ' + CONFIG.hojas[k], () => {
    const sh = ss_().getSheetByName(CONFIG.hojas[k]);
    if (!sh) throw new Error('no existe (ejecuta instalar)');
    return Math.max(sh.getLastRow() - 1, 0) + ' filas';
  }));
  paso('Usuarios', () => {
    const u = usuarios_();
    const act = u.filter(puedeEntrar_);
    const adm = act.filter(x => x.rol === 'admin');
    if (!adm.length) throw new Error('no hay ningún Administrador activo en la hoja Usuarios');
    return act.length + ' activos, ' + adm.length + ' administrador(es), ' + act.filter(x => { const c = cred_(x.bp); return c && c.h; }).length + ' con contraseña creada';
  });
  paso('instalar() ejecutado', () => { if (!props_().getProperty('PIMIENTA')) throw new Error('no; ejecútalo desde el editor'); return 'sí'; });
  paso('Meses en Roles', () => { const m = leerRoles_().map(x => x.mes); if (!m.length) throw new Error('ninguno'); return m.join(' '); });
  paso('Fuentes (Google Sheets)', () => leerFuentes_().map(f => (f.mes || '?') + ':' + (f.activo ? (f.estado || 'sin leer').slice(0, 40) : 'pausada')).join(' | ') || 'ninguna');
  paso('Permisos de correo (MailApp)', () => { try { return MailApp.getRemainingDailyQuota() + ' correos disponibles hoy'; } catch (err) { throw new Error('falta autorizar. En el editor ejecuta autorizarPermisos y acepta los permisos (solo afecta el envío de códigos por correo).'); } });
  paso('Llamada de prueba apiInicio', () => { const r = apiInicio(null); if (!r.ok) throw new Error(r.error); return 'responde'; });
  L.push('', 'Si todo sale OK y la app sigue en "Cargando…", abre la app en una ventana de incógnito con UNA sola cuenta de Google.');
  return L.join('\n');
}

/** Revisa que un archivo HTML exista, esté completo y que su JavaScript no tenga errores de sintaxis. */
function revisarArchivoHtml_(n) {
  let c;
  try { c = HtmlService.createHtmlOutputFromFile(n).getContent(); }
  catch (err) { throw new Error('no existe un archivo HTML con este nombre exacto'); }
  const kb = Math.round(c.length / 1024) + ' KB';
  const t = c.trim();
  if (n === 'Estilos') {
    if (t.indexOf('<style>') !== 0) throw new Error(kb + ', no empieza con <style>');
    if (t.indexOf('#fin-estilos') < 0 || !/<\/style>$/.test(t)) throw new Error(kb + ', INCOMPLETO: falta el final del archivo; vuelve a pegarlo completo');
    return kb + ', completo';
  }
  if (t.indexOf('<script>') !== 0) throw new Error(kb + ', no empieza con <script>');
  if (t.indexOf('window.__archivos["' + n + '"]=') < 0 || !/<\/script>$/.test(t)) throw new Error(kb + ', INCOMPLETO: falta el final del archivo; vuelve a pegarlo completo');
  const js = t.replace(/^<script>/, '').replace(/<\/script>$/, '');
  try { new Function(js); } catch (err) { throw new Error(kb + ', ERROR DE SINTAXIS: ' + err.message + '. El texto quedó alterado al copiarlo: borra todo el contenido del archivo y pégalo de nuevo desde el zip v' + VERSION_APP + '.'); }
  if (c.indexOf('window.__archivos["' + n + '"]="v' + VERSION_APP + '"') < 0) throw new Error(kb + ', es de otra versión (se esperaba la ' + VERSION_APP + '); vuelve a copiarlo');
  return kb + ', completo y sin errores de sintaxis';
}

/** Inserta un archivo HTML del proyecto dentro de Index.html. */
function include(nombre) {
  if (ARCHIVOS_HTML_.indexOf(nombre) < 0) return '';
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}
