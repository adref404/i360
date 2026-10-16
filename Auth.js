/**
 * Autenticación.
 *
 * - Quién puede entrar y con qué rol lo decide la hoja Usuarios (columnas Rol y Activo).
 * - Usuario: BP o correo corporativo. Contraseña propia de cada persona.
 * - Las contraseñas NO se guardan en la hoja: se guardan como hash (HMAC-SHA256 con sal
 *   por usuario y una "pimienta" secreta) en las propiedades del script, que nadie ve
 *   desde el Google Sheet.
 * - Primer ingreso o contraseña olvidada: código temporal de un solo uso (48 h) que genera
 *   coordinación desde la pestaña "Usuarios y accesos" o que llega por correo si la persona
 *   tiene Correo corporativo.
 * - Sesiones con token aleatorio (el servidor solo guarda su hash), vencimiento y
 *   revocación. En cada llamada se vuelve a leer Usuarios: si alguien pasa a Activo = NO
 *   o cambia de rol, el cambio aplica de inmediato.
 * - 5 intentos fallidos bloquean el usuario 15 minutos. Todo queda en la hoja Registro.
 */

const props_ = () => PropertiesService.getScriptProperties();
const cache_ = () => CacheService.getScriptCache();
const ALFABETO_CODIGO_ = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

/* ---------------- Criptografía ---------------- */

function pimienta_() {
  const p = props_();
  let v = p.getProperty('PIMIENTA');
  if (!v) { v = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('PIMIENTA', v); }
  return v;
}
function bytes_(s) { return Utilities.newBlob(String(s)).getBytes(); }
function hashClave_(clave, sal) {
  const k = bytes_(sal + '|' + pimienta_());
  let v = Utilities.computeHmacSha256Signature(bytes_(clave), k);
  for (let i = 1; i < CONFIG.hashIter; i++) v = Utilities.computeHmacSha256Signature(v, k);
  return Utilities.base64Encode(v);
}
function sha_(s) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(d).replace(/=+$/, '');
}
function igual_(a, b) {
  a = String(a || ''); b = String(b || '');
  let r = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) r |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return r === 0;
}
function aleatorio_(n, alfabeto) {
  let out = '';
  while (out.length < n) {
    const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid());
    for (let i = 0; i < b.length && out.length < n; i++) {
      const u = (b[i] + 256) % 256;
      if (u < 256 - (256 % alfabeto.length)) out += alfabeto[u % alfabeto.length];
    }
  }
  return out;
}
function nuevaSal_() { return Utilities.getUuid().replace(/-/g, ''); }
function nuevoToken_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); }

/* ---------------- Usuarios (hoja Usuarios) ---------------- */

let _USU = null;
function rolDe_(v) {
  const t = normTxt_(v);
  if (/^(admin|coordina)/.test(t)) return 'admin';
  if (/^inst/.test(t)) return 'inst';
  return '';
}
function usuarios_() {
  if (_USU) return _USU;
  const t = filas_(hoja_(CONFIG.hojas.usuarios));
  const h = t.head;
  const c = {
    bp: col_(h, 'bp'),
    nom: col_(h, 'nombre (como aparece en el rol)', 'nombre'),
    comp: col_(h, 'nombre completo'),
    mail: col_(h, 'correo corporativo', 'correo', 'email'),
    rol: col_(h, 'rol', 'perfil'),
    flota: col_(h, 'flota'),
    cat: col_(h, 'categoria'),
    est: col_(h, 'estado en dotacion'),
    act: col_(h, 'activo'),
    cargo: col_(h, 'cargo'), area: col_(h, 'area'), foto: col_(h, 'foto'), perfil: col_(h, 'perfil')
  };
  _USU = t.rows.map((r, i) => ({
    fila: i + 2,
    bp: bpTxt_(celda_(r, c.bp)),
    nombre: String(celda_(r, c.nom) || '').trim(),
    nombreCompleto: String(celda_(r, c.comp) || '').trim(),
    correo: String(celda_(r, c.mail) || '').trim().toLowerCase(),
    rolTxt: String(celda_(r, c.rol) || '').trim(),
    rol: rolDe_(celda_(r, c.rol)),
    flota: String(celda_(r, c.flota) || '').trim(),
    categoria: String(celda_(r, c.cat) || '').trim(),
    estado: String(celda_(r, c.est) || '').trim(),
    activo: c.act < 0 || normTxt_(celda_(r, c.act)) !== 'no',
    cargo: String(celda_(r, c.cargo) || '').trim(), area: String(celda_(r, c.area) || '').trim(), foto: String(celda_(r, c.foto) || '').trim(),
    perfil: String(celda_(r, c.perfil) || '').trim()
  })).filter(u => u.bp);
  // Los administradores principales siempre entran como administradores, diga lo que diga su fila.
  _USU.forEach(u => { if (esPrincipal_(u.bp)) { u.principal = true; u.rol = 'admin'; u.activo = true; } });
  return _USU;
}
function buscarUsuario_(texto) {
  const t = String(texto || '').trim().toLowerCase();
  if (!t) return null;
  if (t.indexOf('@') >= 0) return usuarios_().find(u => u.correo && u.correo === t) || null;
  const bp = bpTxt_(t);
  return usuarios_().find(u => u.bp === bp) || null;
}
function esPrincipal_(bp) { return (CONFIG.adminsPrincipales || []).map(bpTxt_).indexOf(bpTxt_(bp)) >= 0; }
/** A un administrador principal solo él mismo puede restablecerle el acceso o cerrarle las sesiones. */
/** Restablecer acceso o cerrar sesiones de otro: nunca a un principal ni a alguien con más permisos (evita suplantarlo). */
function exigirPuedeGestionar_(admin, u) {
  if (u.bp === admin.bp) return;
  if (u.principal) throw falla_(nombreDe_(u) + ' es administrador principal: solo él puede restablecer su acceso o cerrar sus sesiones.');
  exigirPuedeAdministrar_(admin, u);
}
function puedeEntrar_(u) { return !!(u && u.activo && u.rol); }
function nombreDe_(u) { return u.nombre || u.nombreCompleto || ('BP ' + u.bp); }

/* ---------------- Permisos (hoja Permisos) ---------------- */

const VISTAS_ORDEN_ = ['ind', 'rol', 'dia', 'inst', 'vac', 'int', 'cargar', 'colores', 'usuarios', 'mi'];
const VISTAS_EQUIPO_ = ['ind', 'rol', 'dia', 'inst', 'vac'];
function claveVista_(etiqueta, rol) {
  const t = normTxt_(etiqueta);
  if (t.indexOf('vista por instructor') === 0 || t === 'por instructor' || t === 'mi rol') return rol === 'inst' ? 'mi' : 'inst';
  if (t.indexOf('indicadores') === 0) return 'ind';
  if (t.indexOf('seguimiento') === 0) return 'seg';
  if (t.indexOf('rol completo') === 0) return 'rol';
  if (t.indexOf('equipo por dia') === 0) return 'dia';
  if (t.indexOf('actualizar rol') === 0 || t.indexOf('fuentes') === 0) return 'cargar';
  if (t.indexOf('usuarios') === 0) return 'usuarios';
  if (t.indexOf('vacaciones') === 0) return 'vac';
  if (t.indexOf('intercambio') === 0) return 'int';
  if (t.indexOf('tipos') === 0 || t.indexOf('colores') === 0) return 'colores';
  return '';
}
function vistasDe_(rol) {
  const def = rol === 'admin' ? ['ind', 'rol', 'dia', 'inst', 'vac', 'int', 'cargar', 'colores', 'usuarios'] : ['mi'];
  const t = filas_(hoja_(CONFIG.hojas.permisos));
  if (!t.rows.length) return def;
  const cRol = rol === 'admin' ? col_(t.head, 'administrador', 'admin') : col_(t.head, 'instructor');
  const cVista = col_(t.head, 'vista');
  if (cRol < 0 || cVista < 0) return def;
  const set = {};
  def.forEach(k => set[k] = true);
  t.rows.forEach(r => {
    const k = claveVista_(r[cVista], rol);
    if (!k) return;
    set[k] = normTxt_(r[cRol]) === 'si';
  });
  if (rol !== 'admin') { delete set.cargar; delete set.usuarios; delete set.vac; delete set.int; delete set.colores; } // nunca para instructores (sus vacaciones están en Mi rol)
  if (rol === 'admin') set.usuarios = true;                         // evita que coordinación se quede fuera
  if (rol === 'inst' && !Object.keys(set).some(k => set[k])) set.mi = true;
  return VISTAS_ORDEN_.filter(k => set[k]);
}
function accesoEquipo_(rol, vistas) { return rol === 'admin' || vistas.some(v => VISTAS_EQUIPO_.indexOf(v) >= 0); }

function perfilPublico_(u) {
  // Coordinación: las secciones salen de su perfil de acceso; instructores: de la hoja Permisos
  const niveles = nivelesDe_(u);
  const vistas = u.rol === 'admin' ? VISTAS_ORDEN_.filter(k => (niveles[k] || 0) >= 1) : vistasDe_(u.rol);
  return {
    bp: u.bp, nombre: nombreDe_(u), nombreCompleto: u.nombreCompleto, correo: u.correo,
    rol: u.rol, rolTxt: u.principal ? 'Administrador principal' : u.rol === 'admin' ? 'Administrador' : 'Instructor',
    principal: !!u.principal, cargo: u.cargo || '', area: u.area || '', foto: u.foto || '',
    perfil: u.rol === 'admin' ? (u.principal ? 'Principal' : (u.perfil || PERFIL_DEF_)) : '', niveles: niveles,
    vistas: vistas, equipo: accesoEquipo_(u.rol, vistas)
  };
}

/* ---------------- Credenciales (propiedades del script) ---------------- */

function cred_(bp) { const v = props_().getProperty('C_' + bp); return v ? JSON.parse(v) : null; }
function guardarCred_(bp, c) { props_().setProperty('C_' + bp, JSON.stringify(c)); }

function validarClave_(clave, u) {
  const p = String(clave || '');
  if (p.length < 8) throw falla_('La contraseña debe tener al menos 8 caracteres.');
  if (p.length > 128) throw falla_('La contraseña es demasiado larga.');
  if (!/[A-Za-zÁÉÍÓÚáéíóúÑñ]/.test(p) || !/\d/.test(p)) throw falla_('La contraseña debe combinar letras y números.');
  if (u && p.indexOf(u.bp) >= 0) throw falla_('La contraseña no puede contener tu BP.');
}
function normCodigo_(c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }

function generarCodigo_(u, quien) {
  const codigo = aleatorio_(8, ALFABETO_CODIGO_);
  const c = cred_(u.bp) || { v: 0 };
  const s = nuevaSal_();
  c.cod = { h: hashClave_(codigo, s), s: s, exp: Date.now() + CONFIG.codigoHoras * 3600e3, por: quien || '' };
  guardarCred_(u.bp, c);
  // Solo coordinación (o el editor) desbloquea al generar un código. Si lo pidió cualquiera con
  // "¿Olvidaste tu contraseña?", limpiar los fallos permitiría saltarse el bloqueo por intentos.
  if (quien !== 'correo') limpiarFallos_(u.bp);
  return { codigo: codigo.slice(0, 4) + '-' + codigo.slice(4), vence: new Date(c.cod.exp).toISOString() };
}

function enviarCodigoPorCorreo_(u, codigo, vence) {
  const venceTxt = Utilities.formatDate(new Date(vence), CONFIG.zona, "dd/MM/yyyy 'a las' HH:mm");
  enviarCorreo_({
    to: u.correo,
    subject: CONFIG.nombreApp + ': tu código de acceso',
    body: 'Hola ' + nombreDe_(u) + ',\n\nTu código de acceso es: ' + codigo + '\n\n' +
      'Entra a la aplicación, elige "Genera código y activa tu cuenta", escribe tu BP, este código y tu nueva contraseña.\n' +
      'El código sirve una sola vez y vence el ' + venceTxt + '.\n\n' + (urlApp_() ? 'Entra aquí: ' + urlApp_() + '\n\n' : '') + 'Si no lo pediste, ignora este correo.',
    htmlBody: '<p>Hola ' + nombreDe_(u) + ',</p><p>Tu código de acceso es:</p>' +
      '<p style="font:700 24px monospace;letter-spacing:3px">' + codigo + '</p>' +
      '<p>Entra a la aplicación, elige <b>Genera código y activa tu cuenta</b>, escribe tu BP, este código y tu nueva contraseña.<br>' +
      'El código sirve una sola vez y vence el ' + venceTxt + '.</p>' + botonAppHtml_('Abrir ' + CONFIG.nombreApp + ' y activar mi cuenta') + '<p style="color:#666">Si no lo pediste, ignora este correo.</p>',
    name: CONFIG.nombreApp
  });
}

/* ---------------- Bloqueo por intentos fallidos ---------------- */

function claveFallos_(id) { return 'F_' + sha_(String(id).toLowerCase()); }
function fallos_(id) { return Number(cache_().get(claveFallos_(id)) || 0); }
function revisarBloqueo_(id) {
  if (fallos_(id) >= CONFIG.intentosMax)
    throw falla_('Por seguridad bloqueamos el acceso de este usuario por ' + CONFIG.bloqueoMin + ' minutos tras varios intentos fallidos. Espera o pide a coordinación un código nuevo.', 'BLOQUEADO');
}
function sumarFallo_(id) {
  const n = fallos_(id) + 1;
  cache_().put(claveFallos_(id), String(n), CONFIG.bloqueoMin * 60);
  return n;
}
function limpiarFallos_(id) { cache_().remove(claveFallos_(id)); }

/* ---------------- Registro de accesos (hoja Registro) ---------------- */

const ENC_REGISTRO_ = ['Fecha y hora', 'BP', 'Nombre', 'Correo', 'Rol', 'Evento', 'Resultado', 'Dispositivo / navegador'];
function registrar_(u, evento, resultado, disp, bpTexto) {
  try {
    const sh = hoja_(CONFIG.hojas.registro, ENC_REGISTRO_);
    sh.appendRow([new Date(), u ? u.bp : (bpTexto || ''), u ? nombreDe_(u) : '', u ? u.correo : '',
      u ? (u.rol === 'admin' ? 'Administrador' : u.rol === 'inst' ? 'Instructor' : u.rolTxt) : '',
      evento, resultado || 'OK', String(disp || '').slice(0, 120)]);
  } catch (e) { console.warn('No se pudo escribir en Registro: ' + e); }
}

/* ---------------- Sesiones ---------------- */

function iniciarSesion_(u, recordar, disp) {
  const c = cred_(u.bp) || { v: 0 };
  const token = nuevoToken_();
  const id = 'S_' + sha_(token);
  const ahora = Date.now();
  const exp = ahora + (recordar ? CONFIG.sesionRecordarDias * 864e5 : CONFIG.sesionHoras * 3600e3);
  const s = { bp: u.bp, v: c.v || 0, exp: exp, t0: ahora, rec: !!recordar };
  props_().setProperty(id, JSON.stringify(s));
  cache_().put(id, JSON.stringify(s), 600);
  registrar_(u, 'Ingreso', 'OK', disp);
  limpiarSesionesVencidas_();
  return { token: token, recordar: !!recordar, vence: new Date(exp).toISOString(), usuario: perfilPublico_(u) };
}

/** Valida el token. opciones.admin exige rol Administrador. */
function sesion_(token, opciones) {
  opciones = opciones || {};
  if (!token || typeof token !== 'string' || token.length < 40) throw falla_('Tu sesión terminó. Vuelve a ingresar.', 'SESION');
  const id = 'S_' + sha_(token);
  let raw = cache_().get(id);
  if (!raw) raw = props_().getProperty(id);
  if (!raw) throw falla_('Tu sesión terminó. Vuelve a ingresar.', 'SESION');
  const s = JSON.parse(raw);
  const fin = () => { props_().deleteProperty(id); cache_().remove(id); };
  if (s.exp < Date.now()) { fin(); throw falla_('Tu sesión venció. Vuelve a ingresar.', 'SESION'); }
  const u = usuarios_().find(x => x.bp === s.bp);
  if (!puedeEntrar_(u)) { fin(); throw falla_('Tu acceso fue desactivado. Consulta con coordinación.', 'SESION'); }
  const c = cred_(u.bp) || { v: 0 };
  if ((c.v || 0) !== s.v) { fin(); throw falla_('Tu sesión se cerró porque cambió tu contraseña o coordinación cerró tus sesiones.', 'SESION'); }
  if ((opciones.admin || opciones.seccion) && u.rol !== 'admin') throw falla_('Esta acción es solo para coordinación.', 'PERMISO');
  if (opciones.seccion) exigirNivel_(u, opciones.seccion, opciones.nivel || 1);
  cache_().put(id, raw, 600);
  return { u: u, s: s, id: id };
}

function limpiarSesionesVencidas_() {
  if (Math.random() > 0.2) return;
  const p = props_(), ahora = Date.now();
  p.getKeys().filter(k => k.indexOf('S_') === 0).forEach(k => {
    try { const s = JSON.parse(p.getProperty(k)); if (!s || s.exp < ahora) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); }
  });
}

/* =================== Funciones públicas (llamadas desde la página) =================== */

/** Al abrir la página: ¿hay sesión válida? Además indica si la cuenta de Google coincide con un usuario. */
function apiInicio(token) {
  return responder_(() => {
    let usuario = null;
    if (token) { try { usuario = perfilPublico_(sesion_(token).u); } catch (e) { if (!e.publico) throw e; } }
    let google = '';
    try {
      const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
      if (email && puedeEntrar_(buscarUsuario_(email))) google = email;
    } catch (e) { /* sin acceso al correo del visitante */ }
    return { usuario: usuario, google: google, app: CONFIG.nombreApp };
  });
}

function apiLogin(usuario, clave, recordar, disp) {
  return responder_(() => {
    const id = String(usuario || '').trim().toLowerCase();
    if (!id || !clave) throw falla_('Escribe tu usuario y tu contraseña.');
    const u = buscarUsuario_(id);
    const claveBloqueo = u ? u.bp : id;
    revisarBloqueo_(claveBloqueo);
    const c = u ? cred_(u.bp) : null;
    let ok = false;
    if (puedeEntrar_(u) && c && c.h) ok = igual_(hashClave_(String(clave), c.s), c.h);
    else hashClave_(String(clave), 'x'); // mismo tiempo de respuesta exista o no el usuario
    if (!ok) {
      const motivo = !u ? 'Usuario no registrado' : !u.activo ? 'Usuario inactivo' : !u.rol ? 'Sin rol asignado' : !(c && c.h) ? 'Cuenta sin activar' : 'Contraseña incorrecta';
      const n = sumarFallo_(claveBloqueo);
      registrar_(u, 'Intento fallido', motivo, disp, u ? '' : id);
      if (n >= CONFIG.intentosMax) { registrar_(u, 'Bloqueo temporal', CONFIG.bloqueoMin + ' min', disp, u ? '' : id); revisarBloqueo_(claveBloqueo); }
      throw falla_('Usuario o contraseña incorrectos. Si es tu primer ingreso, usa "Genera código y activa tu cuenta".', 'CREDENCIALES');
    }
    limpiarFallos_(claveBloqueo);
    return iniciarSesion_(u, !!recordar, disp);
  });
}

/** Primer ingreso o restablecer: usuario + código temporal + nueva contraseña. */
function apiActivar(usuario, codigo, nueva, recordar, disp) {
  return responder_(() => {
    const id = String(usuario || '').trim().toLowerCase();
    if (!id || !codigo) throw falla_('Escribe tu usuario y el código que recibiste.');
    const u = buscarUsuario_(id);
    const claveBloqueo = u ? u.bp : id;
    revisarBloqueo_(claveBloqueo);
    const c = u ? cred_(u.bp) : null;
    const cod = normCodigo_(codigo);
    const ok = puedeEntrar_(u) && c && c.cod && c.cod.exp > Date.now() && igual_(hashClave_(cod, c.cod.s), c.cod.h);
    if (!ok) {
      sumarFallo_(claveBloqueo);
      registrar_(u, 'Intento fallido', 'Código inválido o vencido', disp, u ? '' : id);
      throw falla_('El usuario o el código no son correctos, o el código ya venció. Pide uno nuevo a coordinación.', 'CODIGO');
    }
    validarClave_(nueva, u);
    conBloqueo_(() => {
      const s = nuevaSal_();
      const primera = !c.h;
      c.h = hashClave_(String(nueva), s); c.s = s; c.t = ahoraISO_();
      delete c.cod;
      c.v = (c.v || 0) + 1; // cierra cualquier otra sesión abierta
      guardarCred_(u.bp, c);
      registrar_(u, primera ? 'Activación de cuenta' : 'Restablecimiento de contraseña', 'OK', disp);
    });
    limpiarFallos_(claveBloqueo);
    return iniciarSesion_(u, !!recordar, disp);
  });
}

/** "Olvidé mi contraseña": si el usuario tiene correo, se le envía un código. Respuesta siempre genérica. */
function apiSolicitarCodigo(usuario, disp) {
  return responder_(() => {
    const id = String(usuario || '').trim().toLowerCase();
    const generico = 'Si tu usuario tiene un correo registrado, te enviamos un código que vence en ' + CONFIG.codigoHoras +
      ' horas. Si no te llega, pide el código a coordinación.';
    if (!id) throw falla_('Escribe tu BP o tu correo.');
    const rk = 'R_' + sha_(id);
    if (cache_().get(rk)) return { mensaje: generico };
    cache_().put(rk, '1', 300);
    const u = buscarUsuario_(id);
    if (puedeEntrar_(u) && u.correo) {
      try {
        const g = generarCodigo_(u, 'correo');
        enviarCodigoPorCorreo_(u, g.codigo, g.vence);
        registrar_(u, 'Código enviado por correo', 'OK', disp);
      } catch (e) { registrar_(u, 'Código enviado por correo', 'Error: ' + e.message, disp); }
    }
    return { mensaje: generico };
  });
}

/** Ingreso con la cuenta de Google del visitante (solo si su correo está en Usuarios). */
function apiLoginGoogle(recordar, disp) {
  return responder_(() => {
    const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
    const u = email ? buscarUsuario_(email) : null;
    if (!puedeEntrar_(u)) throw falla_('Tu cuenta de Google no está registrada como usuario activo.');
    return iniciarSesion_(u, !!recordar, disp);
  });
}

function apiCambiarClave(token, actual, nueva, disp) {
  return responder_(() => {
    const ses = sesion_(token), u = ses.u;
    revisarBloqueo_(u.bp);
    const c = cred_(u.bp);
    if (!c || !c.h || !igual_(hashClave_(String(actual || ''), c.s), c.h)) {
      sumarFallo_(u.bp);
      registrar_(u, 'Intento fallido', 'Contraseña actual incorrecta (cambio)', disp);
      throw falla_('La contraseña actual no es correcta.');
    }
    if (String(actual) === String(nueva)) throw falla_('La nueva contraseña debe ser distinta de la actual.');
    validarClave_(nueva, u);
    conBloqueo_(() => {
      const s = nuevaSal_();
      c.h = hashClave_(String(nueva), s); c.s = s; c.t = ahoraISO_(); c.v = (c.v || 0) + 1;
      guardarCred_(u.bp, c);
    });
    props_().deleteProperty(ses.id); cache_().remove(ses.id);
    registrar_(u, 'Cambio de contraseña', 'OK', disp);
    return iniciarSesion_(u, !!ses.s.rec, disp);
  });
}

function apiLogout(token, disp) {
  return responder_(() => {
    if (!token) return true;
    const id = 'S_' + sha_(token);
    const raw = props_().getProperty(id);
    props_().deleteProperty(id); cache_().remove(id);
    if (raw) { const s = JSON.parse(raw); registrar_(usuarios_().find(x => x.bp === s.bp), 'Salida', 'OK', disp); }
    return true;
  });
}

/* ---------------- Administración de accesos (solo coordinación) ---------------- */

function apiUsuarios(token) {
  return responder_(() => {
    sesion_(token, { seccion: 'usuarios', nivel: 1 });
    // Resumen de Registro por BP
    const reg = {};
    const t = filas_(hoja_(CONFIG.hojas.registro));
    const cF = col_(t.head, 'fecha y hora', 'fecha'), cB = col_(t.head, 'bp'), cE = col_(t.head, 'evento');
    t.rows.forEach(r => {
      const bp = bpTxt_(celda_(r, cB)); if (!bp) return;
      const x = reg[bp] || (reg[bp] = { ingresos: 0, ultimo: null, fallidos: 0 });
      const ev = normTxt_(celda_(r, cE)), f = celda_(r, cF);
      if (ev === 'ingreso') { x.ingresos++; if (esFecha_(f) && (!x.ultimo || f > x.ultimo)) x.ultimo = f; }
      if (ev === 'intento fallido') x.fallidos++;
    });
    const ahora = Date.now();
    const yo = sesion_(token, { seccion: 'usuarios', nivel: 1 }).u;
    return {
      baseUrl: ss_().getUrl(), soyPrincipal: !!yo.principal, miBp: yo.bp, misNiveles: nivelesDe_(yo), puedoGestionar: nivel_(yo, 'usuarios') >= 2,
      principalesFaltantes: (CONFIG.adminsPrincipales || []).map(bpTxt_).filter(bp => bp && !usuarios_().some(u => u.bp === bp)),
      usuarios: usuarios_().map(u => {
        const c = cred_(u.bp) || {};
        const r = reg[u.bp] || {};
        return {
          bp: u.bp, nombre: nombreDe_(u), nombreCorto: u.nombre, nombreCompleto: u.nombreCompleto, correo: u.correo,
          rolTxt: u.principal ? 'Administrador principal' : u.rol === 'admin' ? 'Administrador' : u.rol === 'inst' ? 'Instructor' : (u.rolTxt || 'Sin rol'),
          rol: u.rol, activo: u.activo, principal: !!u.principal, cargo: u.cargo, area: u.area, foto: u.foto,
          perfil: u.rol === 'admin' ? (u.principal ? 'Principal' : (u.perfil || PERFIL_DEF_)) : '',
          administrable: u.bp === yo.bp || (() => { try { exigirPuedeGestionar_(yo, u); return true; } catch (e) { return false; } })(),
          cuenta: c.h ? 'activa' : 'sin activar',
          claveDesde: c.t || null,
          codigoVence: c.cod && c.cod.exp > ahora ? new Date(c.cod.exp).toISOString() : null,
          bloqueado: fallos_(u.bp) >= CONFIG.intentosMax,
          ingresos: r.ingresos || 0, ultimoIngreso: r.ultimo ? r.ultimo.toISOString() : null, fallidos: r.fallidos || 0
        };
      })
    };
  });
}

function apiGenerarCodigo(token, bp, enviarCorreo) {
  return responder_(() => {
    const admin = sesion_(token, { seccion: 'usuarios', nivel: 2 }).u;
    const u = usuarios_().find(x => x.bp === bpTxt_(bp));
    if (!u) throw falla_('No encontramos ese BP en la hoja Usuarios.');
    if (!puedeEntrar_(u)) throw falla_('Ese usuario está inactivo o no tiene rol. Revísalo en la hoja Usuarios.');
    exigirPuedeGestionar_(admin, u);
    const g = conBloqueo_(() => generarCodigo_(u, admin.bp));
    let enviado = false, errorCorreo = '';
    if (enviarCorreo) {
      if (!u.correo) errorCorreo = 'No tiene Correo corporativo en la hoja Usuarios.';
      else { try { enviarCodigoPorCorreo_(u, g.codigo, g.vence); enviado = true; } catch (e) { errorCorreo = e.message; } }
    }
    registrar_(u, 'Código temporal generado', 'Por ' + nombreDe_(admin) + (enviado ? ' (enviado por correo)' : ''), '');
    return { codigo: g.codigo, vence: g.vence, enviado: enviado, correo: u.correo, errorCorreo: errorCorreo, nombre: nombreDe_(u) };
  });
}

function apiCerrarSesiones(token, bp) {
  return responder_(() => {
    const admin = sesion_(token, { seccion: 'usuarios', nivel: 2 }).u;
    const u = usuarios_().find(x => x.bp === bpTxt_(bp));
    if (!u) throw falla_('No encontramos ese BP.');
    exigirPuedeGestionar_(admin, u);
    conBloqueo_(() => { const c = cred_(u.bp) || { v: 0 }; c.v = (c.v || 0) + 1; guardarCred_(u.bp, c); });
    limpiarFallos_(u.bp);
    registrar_(u, 'Sesiones cerradas', 'Por ' + nombreDe_(admin), '');
    return true;
  });
}

/* ---------------- Administradores (solo administradores principales) ---------------- */

/** Crear o activar/desactivar usuarios: principales o quien gestione "Usuarios y accesos". */
function exigirGestorUsuarios_(token) { return sesion_(token, { seccion: 'usuarios', nivel: 2 }).u; }

/** Crea un administrador en la hoja Usuarios. Los instructores vienen del escalafón y se cargan en la hoja. */
function apiCrearAdmin(token, d) {
  return responder_(() => {
    const yo = exigirGestorUsuarios_(token);
    d = d || {};
    const bp = bpTxt_(d.bp), nombre = String(d.nombre || '').trim(), completo = String(d.nombreCompleto || '').trim();
    const correo = String(d.correo || '').trim().toLowerCase();
    if (!/^\d{3,10}$/.test(bp)) throw falla_('El BP debe ser numérico.');
    if (!nombre) throw falla_('Escribe el nombre con el que aparecerá en la app.');
    if (correo && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) throw falla_('El correo no es válido.');
    _PERF = null;
    const perfil = perfilPorNombre_(d.perfil || PERFIL_DEF_);
    if (!perfil) throw falla_('Ese perfil de acceso no existe.');
    exigirSinEscalar_(yo, perfil.niveles);
    return conBloqueo_(() => {
      _USU = null;
      const ya = usuarios_().find(u => u.bp === bp);
      if (ya) throw falla_('El BP ' + bp + ' ya existe en la hoja Usuarios (' + nombreDe_(ya) + ', ' + (ya.rol === 'admin' ? 'administrador' : ya.rol === 'inst' ? 'instructor' : 'sin rol') + ').');
      if (correo && usuarios_().some(u => u.correo === correo)) throw falla_('Ese correo ya está registrado en otro usuario.');
      const sh = hoja_(CONFIG.hojas.usuarios);
      const n = sh.getLastRow() + 1;
      sh.getRange(n, 1).setNumberFormat('@');
      escribirFila_(sh, n, {
        'BP': bp, 'Nombre (como aparece en el rol)': nombre, 'Nombre completo': completo, 'Correo corporativo': correo,
        'Rol': 'Administrador', 'Activo': 'SI', 'Observaciones': 'Creado desde la app por ' + nombreDe_(yo) + ' el ' + hoyISO_()
      });
      asegurarColumnas_(sh, ['Perfil']);
      escribirFila_(sh, n, { 'Perfil': perfil.nombre });
      _USU = null;
      const u = usuarios_().find(x => x.bp === bp);
      registrar_(u, 'Administrador creado', 'Por ' + nombreDe_(yo), '');
      return { bp: bp, nombre: nombre };
    });
  });
}

/**
 * Corrige el nombre, el nombre completo o el correo de un usuario (instructor o administrador).
 * El BP no se edita: es la llave de su contraseña, vacaciones, intercambios y de su fila en el rol.
 */
function apiEditarUsuario(token, bp, d) {
  return responder_(() => {
    const yo = exigirGestorUsuarios_(token);
    d = d || {};
    const nombre = String(d.nombre || '').trim().slice(0, 60), completo = String(d.nombreCompleto || '').trim().slice(0, 120);
    const correo = String(d.correo || '').trim().toLowerCase();
    if (!nombre) throw falla_('Escribe el nombre con el que aparece en la app.');
    if (correo && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) throw falla_('El correo no es válido.');
    return conBloqueo_(() => {
      _USU = null;
      const u = usuarios_().find(x => x.bp === bpTxt_(bp));
      if (!u) throw falla_('Ese BP ya no está en la hoja Usuarios. Recarga la página.', 'YA_RESUELTO');
      exigirPuedeAdministrar_(yo, u);
      if (correo && usuarios_().some(x => x.bp !== u.bp && x.correo === correo)) throw falla_('Ese correo ya lo tiene otro usuario.');
      const cambios = [];
      if (nombre !== u.nombre) cambios.push('nombre: ' + (u.nombre || '—') + ' → ' + nombre);
      if (completo !== u.nombreCompleto) cambios.push('nombre completo');
      if (correo !== u.correo) cambios.push('correo: ' + (u.correo || '—') + ' → ' + (correo || '—'));
      if (!cambios.length) return { bp: u.bp, nombre: nombre, cambios: 0 };
      escribirFila_(hoja_(CONFIG.hojas.usuarios), u.fila, { 'Nombre (como aparece en el rol)': nombre, 'Nombre completo': completo, 'Correo corporativo': correo });
      _USU = null;
      registrar_(u, 'Datos editados', cambios.join(' · ') + ' · por ' + nombreDe_(yo), '');
      return { bp: u.bp, nombre: nombre, cambios: cambios.length };
    });
  });
}

/** Activa o desactiva el acceso de un usuario (no se puede con administradores principales). */
function apiActivarUsuario(token, bp, activo) {
  return responder_(() => {
    const yo = exigirGestorUsuarios_(token);
    return conBloqueo_(() => {
      _USU = null;
      const u = usuarios_().find(x => x.bp === bpTxt_(bp));
      if (!u) throw falla_('No encontramos ese BP.');
      if (u.principal) throw falla_('Un administrador principal no se puede desactivar desde la app.');
      if (u.bp === yo.bp) throw falla_('No puedes desactivar tu propia cuenta.');
      exigirPuedeAdministrar_(yo, u);
      escribirFila_(hoja_(CONFIG.hojas.usuarios), u.fila, { 'Activo': activo ? 'SI' : 'NO' });
      if (!activo) { const c = cred_(u.bp) || { v: 0 }; c.v = (c.v || 0) + 1; guardarCred_(u.bp, c); } // cierra sus sesiones
      _USU = null;
      registrar_(u, activo ? 'Acceso activado' : 'Acceso desactivado', 'Por ' + nombreDe_(yo), '');
      return true;
    });
  });
}

/* ---------------- Perfil corporativo (directorio de Google Workspace) ---------------- */

/**
 * Busca a cada usuario con correo en el directorio corporativo y guarda Cargo, Área y Foto en la hoja Usuarios.
 * Usa People API (servicio avanzado). Solo lee datos públicos del directorio; nunca contraseñas.
 */
function apiSincronizarPerfiles(token) {
  return responder_(() => {
    sesion_(token, { seccion: 'usuarios', nivel: 2 });
    if (typeof People === 'undefined') throw falla_('Falta activar el servicio avanzado People API en el proyecto de Apps Script.');
    const sh = asegurarColumnas_(hoja_(CONFIG.hojas.usuarios), ['Cargo', 'Área', 'Foto']);
    _USU = null;
    const res = { actualizados: 0, sinCorreo: 0, noEncontrados: [], errores: [] };
    usuarios_().forEach(u => {
      if (!u.correo) { res.sinCorreo++; return; }
      try {
        const p = perfilDirectorio_(u.correo);
        if (!p) { res.noEncontrados.push(nombreDe_(u)); return; }
        escribirFila_(sh, u.fila, { 'Cargo': p.cargo, 'Área': p.area, 'Foto': p.foto });
        res.actualizados++;
      } catch (e) {
        res.errores.push(nombreDe_(u) + ': ' + (e && e.message));
      }
    });
    _USU = null;
    if (res.errores.length && !res.actualizados) {
      console.error(res.errores.join(' | '));
      throw falla_('No se pudo consultar el directorio corporativo. Ejecuta autorizarPermisos en el editor y acepta el permiso del directorio. Detalle: ' + res.errores[0]);
    }
    return res;
  });
}

function perfilDirectorio_(correo) {
  const r = People.People.searchDirectoryPeople({
    query: correo, readMask: 'names,photos,organizations,emailAddresses',
    sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE']
  });
  const p = (r.people || []).find(x => (x.emailAddresses || []).some(e => String(e.value || '').toLowerCase() === correo));
  if (!p) return null;
  const org = (p.organizations || [])[0] || {};
  const foto = (p.photos || []).find(f => !f.default);
  return { cargo: org.title || '', area: org.department || '', foto: foto ? foto.url : '' };
}
