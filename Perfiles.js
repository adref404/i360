/**
 * Perfiles de acceso de coordinación (hoja Perfiles).
 *
 * Cada administrador tiene un perfil (columna "Perfil" de la hoja Usuarios). Un perfil define, sección por
 * sección, un nivel: Sin acceso (0) · Ver (1) · Gestionar (2). Los administradores principales siempre
 * tienen todo. Sin perfil asignado = "Coordinación completa" (el comportamiento de antes).
 *
 * Delegación: pueden crear, editar y asignar perfiles los principales y quien tenga "Gestionar" en Usuarios
 * y accesos, pero nunca por encima de sus propios niveles ni sobre un administrador principal.
 *
 * Instructor focal: un instructor con perfil asignado ve además esas secciones de coordinación, siempre en
 * solo lectura ("Gestionar" cuenta como "Ver") y nunca Usuarios y accesos. Sigue siendo instructor para todo
 * lo demás (su rol, intercambios, vacaciones, bloque de días libres). Sin perfil = instructor normal.
 * El servidor valida cada acción (sesion_ con { seccion, nivel }); la página solo oculta botones.
 */

const SECCIONES_ADM_ = [
  ['ind', 'Indicadores'], ['rol', 'Rol completo'], ['dia', 'Equipo por día'],
  ['inst', 'Por instructor'], ['vac', 'Vacaciones'], ['int', 'Intercambios'], ['cargar', 'Actualizar rol'],
  ['colores', 'Tipos y colores'], ['usuarios', 'Usuarios y accesos']
];
const NIVEL_TXT_ = ['Sin acceso', 'Ver', 'Gestionar'];
const PERFIL_DEF_ = 'Coordinación completa';
const ENC_PERFILES_ = ['Perfil', 'Descripción'].concat(SECCIONES_ADM_.map(s => s[1]));
const PERFILES_INICIALES_ = [
  [PERFIL_DEF_, 'Ve y gestiona todo el panel de coordinación.', 2, 2, 2, 2, 2, 2, 2, 2, 2],
  ['Solo lectura', 'Ve el panel completo sin poder aprobar, editar ni cambiar accesos.', 1, 1, 1, 1, 1, 1, 0, 1, 0]
];

const nivelDeTexto_ = t => { const n = normTxt_(t); return /^gest/.test(n) ? 2 : /^ver/.test(n) ? 1 : 0; };
const nombreSeccion_ = k => (SECCIONES_ADM_.find(s => s[0] === k) || [k, k])[1];

function hojaPerfiles_() {
  const sh = hoja_(CONFIG.hojas.perfiles, ENC_PERFILES_);
  asegurarColumnas_(sh, ENC_PERFILES_);
  if (sh.getLastRow() < 2) {
    PERFILES_INICIALES_.forEach(p => sh.appendRow([p[0], p[1]].concat(p.slice(2).map(n => NIVEL_TXT_[n]))));
  }
  return sh;
}

let _PERF = null;
function leerPerfiles_() {
  if (_PERF) return _PERF;
  const t = filas_(ss_().getSheetByName(CONFIG.hojas.perfiles));
  const cP = col_(t.head, 'perfil'), cD = col_(t.head, 'descripcion');
  const lista = t.rows.map((r, i) => {
    const niveles = {};
    SECCIONES_ADM_.forEach(([k, nom]) => { niveles[k] = nivelDeTexto_(celda_(r, col_(t.head, nom))); });
    return { fila: i + 2, nombre: String(celda_(r, cP) || '').trim(), descripcion: String(celda_(r, cD) || '').trim(), niveles: niveles };
  }).filter(p => p.nombre);
  if (!lista.some(p => normTxt_(p.nombre) === normTxt_(PERFIL_DEF_))) {
    const n = {}; SECCIONES_ADM_.forEach(([k]) => { n[k] = 2; });
    lista.unshift({ fila: 0, nombre: PERFIL_DEF_, descripcion: PERFILES_INICIALES_[0][1], niveles: n });
  }
  _PERF = lista;
  return lista;
}
const perfilPorNombre_ = nombre => leerPerfiles_().find(p => normTxt_(p.nombre) === normTxt_(nombre));

/** Niveles efectivos de un usuario: {seccion: 0|1|2}. */
function nivelesDe_(u) {
  const n = {};
  if (u && u.rol === 'inst') {
    const p = u.perfil ? perfilPorNombre_(u.perfil) : null;
    SECCIONES_ADM_.forEach(([k]) => { n[k] = p && k !== 'usuarios' ? Math.min(1, p.niveles[k] || 0) : 0; });
    return n;
  }
  if (!u || u.rol !== 'admin') { SECCIONES_ADM_.forEach(([k]) => { n[k] = 0; }); return n; }
  if (u.principal) { SECCIONES_ADM_.forEach(([k]) => { n[k] = 2; }); return n; }
  // Perfil desconocido (borrado o mal escrito): solo lectura, por seguridad
  const p = perfilPorNombre_(u.perfil || PERFIL_DEF_);
  SECCIONES_ADM_.forEach(([k]) => { n[k] = p ? p.niveles[k] : (k === 'usuarios' || k === 'cargar' ? 0 : 1); });
  return n;
}
const nivel_ = (u, seccion) => nivelesDe_(u)[seccion] || 0;

/** Lanza error si el usuario no llega al nivel pedido. Lo usa sesion_. */
function exigirNivel_(u, seccion, nivel) {
  if (nivel_(u, seccion) >= nivel) return;
  throw falla_('Tu perfil no permite ' + (nivel >= 2 ? 'gestionar' : 'ver') + ' "' + nombreSeccion_(seccion) + '". Pídeselo a un administrador.', 'PERMISO');
}

/** ¿Puede "yo" otorgar estos niveles? Nunca más de lo que él mismo tiene. */
function exigirSinEscalar_(yo, niveles) {
  if (yo.principal) return;
  const mios = nivelesDe_(yo);
  const exceso = SECCIONES_ADM_.filter(([k]) => (niveles[k] || 0) > (mios[k] || 0)).map(s => s[1]);
  if (exceso.length) throw falla_('No puedes otorgar más acceso del que tienes en: ' + exceso.join(', ') + '.', 'PERMISO');
}
/** "yo" puede gestionar a "u" si no es principal y su perfil no lo supera. */
function exigirPuedeAdministrar_(yo, u) {
  if (yo.principal) { if (u.principal && u.bp !== yo.bp) throw falla_('Un administrador principal no se puede modificar desde la app.', 'PERMISO'); return; }
  if (u.principal) throw falla_(nombreDe_(u) + ' es administrador principal: nadie puede modificar su acceso.', 'PERMISO');
  if (u.rol !== 'admin') return;
  const mios = nivelesDe_(yo), suyos = nivelesDe_(u);
  const mas = SECCIONES_ADM_.filter(([k]) => (suyos[k] || 0) > (mios[k] || 0)).map(s => s[1]);
  if (mas.length) throw falla_(nombreDe_(u) + ' tiene más acceso que tu perfil (' + mas.join(', ') + '). Solo puede gestionarlo alguien con igual o mayor acceso.', 'PERMISO');
}

/* ---------------- Funciones públicas ---------------- */

function apiPerfiles(token) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'usuarios', nivel: 1 }).u;
    hojaPerfiles_(); _PERF = null;
    return { secciones: SECCIONES_ADM_, perfiles: leerPerfiles_().map(p => ({ nombre: p.nombre, descripcion: p.descripcion, niveles: p.niveles })), misNiveles: nivelesDe_(yo), soyPrincipal: !!yo.principal };
  });
}

/** Crea o actualiza un perfil. original = nombre anterior (para renombrar) o vacío si es nuevo. */
function apiGuardarPerfil(token, d) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'usuarios', nivel: 2 }).u;
    d = d || {};
    const nombre = String(d.nombre || '').trim().slice(0, 60), desc = String(d.descripcion || '').trim().slice(0, 200), original = String(d.original || '').trim();
    if (!nombre) throw falla_('Escribe el nombre del perfil.');
    const niveles = {}; SECCIONES_ADM_.forEach(([k]) => { const v = Number((d.niveles || {})[k]); niveles[k] = v >= 2 ? 2 : v >= 1 ? 1 : 0; });
    exigirSinEscalar_(yo, niveles);
    return conBloqueo_(() => {
      const sh = hojaPerfiles_(); _PERF = null;
      const previo = original ? perfilPorNombre_(original) : null;
      if (original && !previo) throw falla_('Ese perfil ya no existe. Recarga la página.', 'YA_RESUELTO');
      const otro = perfilPorNombre_(nombre);
      if (otro && otro !== previo) throw falla_('Ya existe un perfil llamado "' + otro.nombre + '".');
      if (previo) exigirSinEscalar_(yo, previo.niveles); // no puede editar un perfil que lo supera
      // No puede quitarse a sí mismo la gestión de usuarios editando su propio perfil (quedaría sin poder deshacerlo)
      if (!yo.principal && previo && normTxt_(yo.perfil || PERFIL_DEF_) === normTxt_(previo.nombre) && niveles.usuarios < 2)
        throw falla_('Ese es tu propio perfil: no puedes quitarte la gestión de usuarios. Pídeselo a otro administrador.');
      const fila = [nombre, desc].concat(SECCIONES_ADM_.map(([k]) => NIVEL_TXT_[niveles[k]]));
      const n = previo && previo.fila ? previo.fila : sh.getLastRow() + 1;
      sh.getRange(n, 1, 1, fila.length).setValues([fila]);
      // Renombrado: actualizar a quienes lo tenían asignado
      if (previo && normTxt_(previo.nombre) !== normTxt_(nombre)) {
        const shU = hoja_(CONFIG.hojas.usuarios); _USU = null;
        usuarios_().filter(u => normTxt_(u.perfil || '') === normTxt_(previo.nombre)).forEach(u => escribirFila_(shU, u.fila, { 'Perfil': nombre }));
        _USU = null;
      }
      _PERF = null;
      registrar_(yo, previo ? 'Perfil de acceso editado' : 'Perfil de acceso creado', nombre, '');
      return true;
    });
  });
}

function apiEliminarPerfil(token, nombre) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'usuarios', nivel: 2 }).u;
    return conBloqueo_(() => {
      _PERF = null; _USU = null;
      const p = perfilPorNombre_(nombre);
      if (!p || !p.fila) throw falla_(p ? 'El perfil por defecto no se puede eliminar.' : 'Ese perfil ya no existe.');
      exigirSinEscalar_(yo, p.niveles);
      const usan = usuarios_().filter(u => (u.rol === 'admin' && !u.principal && normTxt_(u.perfil || PERFIL_DEF_) === normTxt_(p.nombre)) ||
        (u.rol === 'inst' && u.perfil && normTxt_(u.perfil) === normTxt_(p.nombre)));
      if (usan.length) throw falla_('Lo usan ' + usan.map(nombreDe_).join(', ') + '. Asígnales otro perfil antes de eliminarlo.');
      hoja_(CONFIG.hojas.perfiles).deleteRow(p.fila);
      _PERF = null;
      registrar_(yo, 'Perfil de acceso eliminado', p.nombre, '');
      return true;
    });
  });
}

function apiAsignarPerfil(token, bp, nombre) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'usuarios', nivel: 2 }).u;
    return conBloqueo_(() => {
      _USU = null; _PERF = null;
      const u = usuarios_().find(x => x.bp === bpTxt_(bp));
      if (!u) throw falla_('No encontramos ese BP.');
      if (u.rol !== 'admin' && u.rol !== 'inst') throw falla_('Esa persona no tiene rol en la hoja Usuarios.');
      if (u.bp === yo.bp && !yo.principal) throw falla_('No puedes cambiar tu propio perfil. Pídeselo a otro administrador.');
      exigirPuedeAdministrar_(yo, u);
      asegurarColumnas_(hoja_(CONFIG.hojas.usuarios), ['Perfil']);
      // Instructor sin perfil = instructor normal (sin acceso a coordinación)
      if (u.rol === 'inst' && !String(nombre || '').trim()) {
        escribirFila_(hoja_(CONFIG.hojas.usuarios), u.fila, { 'Perfil': '' });
        _USU = null;
        registrar_(u, 'Perfil de acceso quitado', 'Instructor sin acceso a coordinación · por ' + nombreDe_(yo), '');
        return true;
      }
      const p = perfilPorNombre_(nombre);
      if (!p) throw falla_('Ese perfil no existe.');
      exigirSinEscalar_(yo, u.rol === 'inst' ? nivelesDe_(Object.assign({}, u, { perfil: p.nombre })) : p.niveles);
      escribirFila_(hoja_(CONFIG.hojas.usuarios), u.fila, { 'Perfil': p.nombre });
      _USU = null;
      registrar_(u, 'Perfil de acceso asignado', p.nombre + (u.rol === 'inst' ? ' (instructor focal, solo lectura)' : '') + ' · por ' + nombreDe_(yo), '');
      return true;
    });
  });
}
