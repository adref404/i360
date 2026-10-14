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
 *   Vacaciones.gs        Vacaciones por periodo DGAC: hojas, ventana de solicitud, validación, aprobación y migración
 *   SemillaVacaciones.gs Datos iniciales de vacaciones (Vacaciones_INSTRUCTORES_1.xlsx)
 *   Intercambios.gs      Intercambio de actividades entre instructores y meses ocultos a instructores
 *   Perfiles.gs          Perfiles de acceso de coordinación (Sin acceso / Ver / Gestionar por sección)
 *   Colores.gs           Catálogo de tipos de actividad y colores (TiposActividad); repintar meses
 *   Index, Estilos, Clasificador, App, VacacionesVista, IntercambiosVista, Vistas, Coordinacion, Admin (HTML)   Interfaz
 */

const CONFIG = {
  // ID del Google Sheet base (Usuarios, Roles, Respuestas, Registro, Permisos...).
  // Vacío = el Google Sheet al que está vinculado este script.
  baseId: '',

  zona: 'America/Lima',
  nombreApp: 'INS 360',

  // Dirección de la aplicación (implementación /exec). Se usa en el botón "Abrir INS 360" de los correos.
  appUrl: 'https://script.google.com/a/macros/latam.com/s/AKfycbyYb4597iw3lMveVzdRRQKjFOcZimIs49Pteq7o2Bj_B6VP6CtW4IbU0CWGOrOrkQ9T3A/exec',

  // Imagen de fondo de la pantalla de inicio (archivo de Drive compartido como "Cualquier persona con el enlace").
  // Formato: 'https://lh3.googleusercontent.com/d/1SwfuSrn0j517bGUMy6RD1Eho42OQDUx9'. Vacío o si no carga = fondo de color según el tema.
  fondoInicioUrl: 'https://lh3.googleusercontent.com/d/1SwfuSrn0j517bGUMy6RD1Eho42OQDUx9',

  // Logos de LATAM (archivos de Drive compartidos como "Cualquier persona con el enlace").
  // simbolo: sin letras (barra lateral y barra superior) · completo: con letras (inicio de sesión y pantalla de carga).
  // light = para fondos claros · dark = para fondos oscuros. Si no cargan, se muestra el ícono de avión.
  logos: {
    simbolo: { light: 'https://lh3.googleusercontent.com/d/11abTjcpbU5KzruNghEI1SdIHu9zNAVVQ', dark: 'https://lh3.googleusercontent.com/d/1bHAo7BH2DZ0clMZ1nuZjzWh8K5aNE-Hf' },
    completo: { light: 'https://lh3.googleusercontent.com/d/1MaAElOghT2pFg670-FsP5L74zBEyaWfi', dark: 'https://lh3.googleusercontent.com/d/1c3N3kX7-R3plMW9enBaAq3N0sj_D5W_q' }
  },

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
    vacPeriodos: 'VacacionesPeriodos',
    vacProgramado: 'VacacionesProgramado',
    solVac: 'SolicitudesVacaciones',
    intercambios: 'Intercambios',
    tipos: 'TiposActividad',
    perfiles: 'Perfiles'
  },

  // Administradores principales (BP). Siempre son administradores activos aunque alguien cambie su fila
  // en la hoja Usuarios, solo ellos pueden crear o desactivar administradores, y nadie más puede
  // restablecerles la contraseña ni cerrarles las sesiones. Solo se cambia aquí, en el código.
  adminsPrincipales: [
    '6019028'   // Fernando Celadita · soporte
    , '31096'   // Jefe de instructores: reemplaza por su BP y quita las barras //
  ],

  // Vacaciones: en estos meses, los días de VacacionesProgramado que no tienen fecha exacta
  // (en VacacionesDias o VacacionesBloques) se cuentan igual como programados.
  vacSinFecha: { desde: '2026-11', hasta: '2026-12' },

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

  // Google Sheets mensuales que se registran en la hoja Fuentes al ejecutar instalar() o registrarFuentes().
  // El mes se indica siempre: el nombre del archivo no es confiable (el de septiembre se llama "OCTUBRE (old)").
  fuentesIniciales: [
    { mes: '2026-01', enlace: 'https://docs.google.com/spreadsheets/d/1vvFcP8U0LIzJtWTtEr4b7IeQBNvVOgsJrIAnjBaj27c/edit#gid=1274318071', hoja: 'ENERO' },
    { mes: '2026-02', enlace: 'https://docs.google.com/spreadsheets/d/1LflRo5hBQ4YUw5D7nohJj3gwyOu5IQ8aP9y92EiY-C8/edit#gid=811083904', hoja: 'FEBRERO' },
    { mes: '2026-03', enlace: 'https://docs.google.com/spreadsheets/d/1kSOOAA_RgT0jWur6dd_IpqQS3IaN8EPr3RIunHQNQN8/edit#gid=268835517', hoja: 'MARZO' },
    { mes: '2026-04', enlace: 'https://docs.google.com/spreadsheets/d/1hFMLM2VcADDFuYTgpw5TLIQ8rEyX7zVyOkNS51WCixg/edit#gid=1265676691', hoja: 'ABRIL' },
    { mes: '2026-05', enlace: 'https://docs.google.com/spreadsheets/d/1hyJz4NPcdr93YqIxFjSxarhxzZnj-o8kFxnj3hE4480/edit#gid=663790086', hoja: 'MAYO' },
    { mes: '2026-06', enlace: 'https://docs.google.com/spreadsheets/d/1nITVu410LbnxfMwdc_QH1N8VoS8X5IzUjPmyrzbFJI4/edit#gid=1220769573', hoja: 'Junio' },
    { mes: '2026-07', enlace: 'https://docs.google.com/spreadsheets/d/1wri6PfqlDrAgpgV3Yq8G0OFEXuKu3vAgIJ4TUgG3jOw/edit#gid=563239649', hoja: 'JULIO' },
    { mes: '2026-08', enlace: 'https://docs.google.com/spreadsheets/d/1shsSrSpqSEh2XDsfxBLU4vJaw7aqmYAYtZuZh_fjcd8/edit#gid=820704924', hoja: 'AGOSTO' },
    { mes: '2026-09', enlace: 'https://docs.google.com/spreadsheets/d/1YiRvlYAZPXNbWm0sr8UqZ7WocaUr1PFdq3Wo75FvIQ8/edit#gid=681297197', hoja: 'ROL INS' },
    { mes: '2026-10', enlace: 'https://docs.google.com/spreadsheets/d/1C_KTGpXYoN4fruqzzZrqAfthiFojlq-nDcPhLGUxSys/edit#gid=681297197', hoja: 'ROL INS' },
    { mes: '2026-11', enlace: 'https://docs.google.com/spreadsheets/d/1B0IizPnZybvm3cm1nnRhKSUHGl0WcEAmeiK3gvEaqxw/edit?gid=681297197#gid=681297197', hoja: 'ROL INS' }
  ]
};

const VERSION_APP = '2.4';
const ARCHIVOS_HTML_ = ['Estilos', 'FuentesLatam', 'Clasificador', 'App', 'VacacionesVista', 'IntercambiosVista', 'ColoresVista', 'Vistas', 'Coordinacion', 'Admin'];

function doGet(e) {
  // https://…/exec?diag=1  ->  diagnóstico en texto (sin datos personales)
  if (e && e.parameter && e.parameter.diag) {
    return ContentService.createTextOutput(diagnostico_()).setMimeType(ContentService.MimeType.TEXT);
  }
  const t = HtmlService.createTemplateFromFile('Index');
  t.version = VERSION_APP;
  t.logos = JSON.stringify(CONFIG.logos || {});
  t.logoCarga = logoHtml_('completo', 'tema');
  t.logoBarra = logoHtml_('simbolo', 'tema');
  t.nombreApp = CONFIG.nombreApp;
  t.fondo = CONFIG.fondoInicioUrl || '';
  const out = t.evaluate()
    .setTitle(CONFIG.nombreApp)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  if (CONFIG.permitirIframe) out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}

/**
 * Logo LATAM con sus dos variantes. modo 'tema': cambia con el tema claro/oscuro; 'oscuro': siempre la versión
 * para fondo oscuro. Si las imágenes no cargan queda el ícono de avión. (La página tiene la misma función: logoHTML.)
 */
function logoHtml_(tipo, modo) {
  const L = (CONFIG.logos || {})[tipo] || {};
  const esc = s => String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const img = (u, cls) => u ? '<img class="' + cls + '" src="' + esc(u) + '" alt="LATAM" onerror="this.parentNode.classList.add(\'err\');this.remove()">' : '';
  const imgs = modo === 'oscuro' ? img(L.dark, 'lg-d solo') : img(L.light, 'lg-l') + img(L.dark, 'lg-d');
  return '<span class="lg lg-' + tipo + ' ' + modo + (imgs ? '' : ' err') + '">' + imgs + '<svg class="lg-alt" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M21 16v-2l-8-5V3.5a1.5 1.5 0 0 0-3 0V9l-8 5v2l8-2.5V19l-2 1.5V22l3.5-1 3.5 1v-1.5L13 19v-5.5z"/></svg></span>';
}

/** Revisa la instalación paso a paso. Solo muestra conteos y estados, nunca nombres ni contraseñas. */
function diagnostico_() {
  const L = ['INS 360 LP · diagnóstico · versión ' + VERSION_APP, 'Fecha del servidor: ' + new Date().toISOString(), ''];
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
  paso('Administradores principales', () => {
    const conf = (CONFIG.adminsPrincipales || []).map(bpTxt_).filter(Boolean);
    if (!conf.length) throw new Error('no hay ninguno en CONFIG.adminsPrincipales');
    const falta = conf.filter(bp => !usuarios_().some(u => u.bp === bp));
    if (falta.length) throw new Error('el BP ' + falta.join(', ') + ' está en CONFIG.adminsPrincipales pero no en la hoja Usuarios');
    return conf.length + ' configurado(s)';
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
  if (n === 'Estilos' || n === 'FuentesLatam') {
    if (t.indexOf('<style>') !== 0) throw new Error(kb + ', no empieza con <style>');
    if (t.indexOf(n === 'Estilos' ? '#fin-estilos' : '#fin-fuentes') < 0 || !/<\/style>$/.test(t)) throw new Error(kb + ', INCOMPLETO: falta el final del archivo; vuelve a pegarlo completo');
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
