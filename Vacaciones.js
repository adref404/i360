/**
 * Vacaciones por periodo DGAC.
 *
 * Cada instructor tiene su propio periodo anual, que empieza en su fecha DGAC (día y mes).
 * Hojas del Google Sheet base:
 *   Vacaciones             una fila por instructor: fechas DGAC
 *   VacacionesPeriodos     una fila por instructor y periodo DGAC: derecho y tomados de referencia
 *   VacacionesProgramado   una fila por instructor y mes: días programados (para los que no tienen fecha exacta)
 *   VacacionesDias         una fila por día de vacaciones con fecha (tomadas o programadas)
 *   VacacionesBloques      bloques programados (desde, hasta)
 *   SolicitudesVacaciones  lo que piden los instructores desde la app y la decisión de coordinación
 *
 * Modelo anterior: el derecho, los tomados y lo programado estaban en columnas por año de la hoja
 * Vacaciones (Derecho 2025-26, Derecho 2026-27, Tomados 2025-26, Programado por mes 2026).
 * Mientras no exista la hoja VacacionesPeriodos se siguen leyendo esas columnas; migrarVacaciones()
 * las copia a las hojas nuevas sin borrarlas y verificarVacaciones() compara ambos modelos.
 * La ventana de solicitud (abierta/cerrada, fechas y meses que se pueden pedir) se guarda en
 * las propiedades del script y la cambia coordinación desde la pestaña Vacaciones.
 *
 * El cálculo del periodo es el mismo que hace la página (Vacaciones.html): el servidor lo repite
 * para validar cada solicitud, así nadie puede pedir más días de los que le quedan.
 */

const ENC_VAC_ = ['BP', 'Nombre', 'Fecha DGAC original', 'Fecha DGAC 2025', 'Fecha DGAC 2026'];
const ENC_VAC_LEGADO_ = ['Derecho 2025-26', 'Derecho 2026-27', 'Tomados 2025-26', 'Programado por mes 2026'];
const ENC_VAC_PER_ = ['BP', 'Periodo', 'Derecho', 'Tomados (referencia)'];
const ENC_VAC_PROG_ = ['BP', 'Mes', 'Días'];
const ENC_VAC_DIAS_ = ['BP', 'Fecha'];
const ENC_VAC_BLOQ_ = ['BP', 'Desde', 'Hasta', 'Días', 'Registrado'];
const ENC_SOL_VAC_ = ['ID', 'BP', 'Nombre', 'Días', 'Cantidad', 'Enviada', 'Estado', 'Decidida', 'Decidida por'];
// Columnas agregadas después: se añaden al final de la hoja si faltan (no mueven las anteriores).
const ENC_SOL_VAC_EXTRA_ = ['Comentario instructor', 'Comentario coordinación', 'Respuesta instructor', 'Motivo revisión', 'Respondida', 'Revisión atendida'];
const MAX_COMENTARIO_VAC_ = 500;
const CFG_VAC_DEF_ = { habilitado: false, desde: '', hasta: '', mesDesde: '2027-01', mesHasta: '2027-12', cupo: 3 };

/* ---------------- Lectura de hojas ---------------- */

function numero_(v) { if (v === '' || v == null) return null; const n = Number(v); return isNaN(n) ? null : n; }

/** "4:8; 6:15; 11:15" -> {"4":8,"6":15,"11":15} */
function mensualDeTexto_(t) {
  const out = {};
  String(t || '').split(/[;,|]/).forEach(par => {
    const m = par.trim().match(/^(\d{1,2})\s*[:=]\s*(\d+(?:\.\d+)?)$/);
    if (m) out[String(Number(m[1]))] = Number(m[2]);
  });
  return out;
}
/** Valores del modelo anterior (columnas por año) -> {derechos, tomadosRef, programado} por año de inicio y mes. */
function legadoDeValores_(m25, m26, tom, mensual2026) {
  const o = { derechos: {}, tomadosRef: {}, programado: {} };
  if (m25 != null) o.derechos['2025'] = m25;
  if (m26 != null) o.derechos['2026'] = m26;
  if (tom != null) o.tomadosRef['2025'] = tom;
  Object.keys(mensual2026 || {}).forEach(m => { o.programado['2026-' + pad2_(Number(m))] = mensual2026[m]; });
  return o;
}

/** "2025-26", 2025 o una fecha -> 2025 (año en que empieza el periodo DGAC). */
function anioPeriodo_(v) {
  if (esFecha_(v)) return v.getFullYear();
  const m = String(v == null ? '' : v).match(/^\s*(\d{4})/);
  return m ? Number(m[1]) : null;
}
function etiquetaPer_(Y) { return Y + '-' + pad2_((Y + 1) % 100); }

/**
 * modo: 'legado' lee las columnas por año de la hoja Vacaciones, 'nuevo' lee VacacionesPeriodos y
 * VacacionesProgramado. Sin modo, usa las hojas nuevas solo si existe VacacionesPeriodos.
 */
function leerVacaciones_(modo) {
  const personas = {};
  const nuevo = modo ? modo === 'nuevo' : !!ss_().getSheetByName(CONFIG.hojas.vacPeriodos);
  const t = filas_(hoja_(CONFIG.hojas.vacaciones));
  if (t.rows.length) {
    const c = {
      bp: col_(t.head, 'bp'), nom: col_(t.head, 'nombre'), orig: col_(t.head, 'fecha dgac original'),
      d25: col_(t.head, 'fecha dgac 2025'), d26: col_(t.head, 'fecha dgac 2026'),
      m25: col_(t.head, 'derecho 2025-26'), m26: col_(t.head, 'derecho 2026-27'),
      tom: col_(t.head, 'tomados 2025-26'), men: col_(t.head, 'programado por mes 2026')
    };
    t.rows.forEach(r => {
      const bp = bpTxt_(celda_(r, c.bp)); if (!bp) return;
      const datos = nuevo ? { derechos: {}, tomadosRef: {}, programado: {} }
        : legadoDeValores_(numero_(celda_(r, c.m25)), numero_(celda_(r, c.m26)), numero_(celda_(r, c.tom)), mensualDeTexto_(celda_(r, c.men)));
      personas[bp] = Object.assign({
        nombreVac: String(celda_(r, c.nom) || '').trim(),
        dgacOrig: fechaISO_(celda_(r, c.orig)), dgac25: fechaISO_(celda_(r, c.d25)), dgac26: fechaISO_(celda_(r, c.d26))
      }, datos, { dias: [], bloques: [] });
    });
  }
  if (nuevo) {
    const tp = filas_(hoja_(CONFIG.hojas.vacPeriodos));
    const cp = { bp: col_(tp.head, 'bp'), per: col_(tp.head, 'periodo'), der: col_(tp.head, 'derecho'), tom: col_(tp.head, 'tomados') };
    tp.rows.forEach(r => {
      const p = personas[bpTxt_(celda_(r, cp.bp))]; const Y = anioPeriodo_(celda_(r, cp.per));
      if (!p || !Y) return;
      const der = numero_(celda_(r, cp.der)), tom = numero_(celda_(r, cp.tom));
      if (der != null) p.derechos[Y] = der;
      if (tom != null) p.tomadosRef[Y] = tom;
    });
    const tg = filas_(hoja_(CONFIG.hojas.vacProgramado));
    const cg = { bp: col_(tg.head, 'bp'), mes: col_(tg.head, 'mes'), n: col_(tg.head, 'dias') };
    tg.rows.forEach(r => {
      const p = personas[bpTxt_(celda_(r, cg.bp))]; const mes = mesTxt_(celda_(r, cg.mes)); const n = numero_(celda_(r, cg.n));
      if (p && mes && n != null) p.programado[mes] = (p.programado[mes] || 0) + n;
    });
  }
  const td = filas_(hoja_(CONFIG.hojas.vacDias));
  const cdB = col_(td.head, 'bp'), cdF = col_(td.head, 'fecha');
  td.rows.forEach(r => {
    const p = personas[bpTxt_(celda_(r, cdB))]; const f = fechaISO_(celda_(r, cdF));
    if (p && f && /^\d{4}-\d{2}-\d{2}$/.test(f)) p.dias.push(f);
  });
  const tb = filas_(hoja_(CONFIG.hojas.vacBloques));
  const cb = { bp: col_(tb.head, 'bp'), de: col_(tb.head, 'desde'), a: col_(tb.head, 'hasta'), n: col_(tb.head, 'dias'), reg: col_(tb.head, 'registrado') };
  tb.rows.forEach(r => {
    const p = personas[bpTxt_(celda_(r, cb.bp))]; const de = fechaISO_(celda_(r, cb.de)), a = fechaISO_(celda_(r, cb.a)) || de;
    if (p && de && /^\d{4}-\d{2}-\d{2}$/.test(de)) p.bloques.push({ desde: de, hasta: a, n: numero_(celda_(r, cb.n)), registro: fechaISO_(celda_(r, cb.reg)) });
  });
  Object.keys(personas).forEach(bp => {
    const p = personas[bp];
    p.dias = Array.from(new Set(p.dias)).sort();
    if (!p.dgac25 && !p.dgacOrig && !p.dgac26) delete personas[bp]; // sin fecha DGAC no hay periodo
  });
  let meta = {};
  try { meta = JSON.parse(props_().getProperty('VAC_META') || '{}'); } catch (e) { meta = {}; }
  return { personas: personas, ventanaUltima: meta.ventanaUltima || null, fuente: meta.fuente || 'Hoja Vacaciones' };
}

/** Hoja SolicitudesVacaciones con todas las columnas (agrega al final las que falten). */
function hojaSolVac_() {
  return asegurarColumnas_(hoja_(CONFIG.hojas.solVac, ENC_SOL_VAC_.concat(ENC_SOL_VAC_EXTRA_)), ENC_SOL_VAC_EXTRA_);
}
function escribirSolVac_(sh, fila, valores) { return escribirFila_(sh, fila, valores); }

function comentarioVac_(t) {
  t = String(t == null ? '' : t).trim();
  if (t.length > MAX_COMENTARIO_VAC_) throw falla_('El comentario es muy largo (máximo ' + MAX_COMENTARIO_VAC_ + ' caracteres).');
  return t;
}
const fechaCelda_ = v => esFecha_(v) ? v.toISOString() : String(v || '');

function leerSolicitudesVac_() {
  const t = filas_(hoja_(CONFIG.hojas.solVac, ENC_SOL_VAC_));
  const c = {
    id: col_(t.head, 'id'), bp: col_(t.head, 'bp'), nom: col_(t.head, 'nombre'), dias: col_(t.head, 'dias'), n: col_(t.head, 'cantidad'),
    at: col_(t.head, 'enviada'), es: col_(t.head, 'estado'), dec: col_(t.head, 'decidida'), por: col_(t.head, 'decidida por'),
    ci: col_(t.head, 'comentario instructor'), cc: col_(t.head, 'comentario coordinacion'), re: col_(t.head, 'respuesta instructor'),
    mr: col_(t.head, 'motivo revision'), ra: col_(t.head, 'respondida'), ate: col_(t.head, 'revision atendida')
  };
  return t.rows.map((r, i) => {
    const at = celda_(r, c.at), dec = celda_(r, c.dec);
    const resp = normTxt_(celda_(r, c.re));
    const dias = String(celda_(r, c.dias) || '').split(/[\s,;]+/).map(fechaISO_).filter(x => x && /^\d{4}-\d{2}-\d{2}$/.test(x)).sort();
    return {
      fila: i + 2, id: String(celda_(r, c.id) || ''), bp: bpTxt_(celda_(r, c.bp)), nombre: String(celda_(r, c.nom) || ''),
      dias: dias, n: dias.length, at: esFecha_(at) ? at.toISOString() : String(at || ''),
      estado: normTxt_(celda_(r, c.es)).replace(/^aprobad.*/, 'aprobada').replace(/^rechazad.*/, 'rechazada').replace(/^pendient.*/, 'pendiente') || 'pendiente',
      decididoAt: esFecha_(dec) ? dec.toISOString() : String(dec || ''), decididoPor: String(celda_(r, c.por) || ''),
      comInst: String(celda_(r, c.ci) || ''), comCoord: String(celda_(r, c.cc) || ''),
      respuesta: /^enterad/.test(resp) ? 'enterado' : /^revision/.test(resp) ? 'revision' : '',
      motivoRev: String(celda_(r, c.mr) || ''), respondidaAt: fechaCelda_(celda_(r, c.ra)), revAtendidaAt: fechaCelda_(celda_(r, c.ate))
    };
  }).filter(s => s.id && s.bp);
}

function cfgVac_() {
  let c = {};
  try { c = JSON.parse(props_().getProperty('CFG_VAC') || '{}'); } catch (e) { c = {}; }
  return Object.assign({}, CFG_VAC_DEF_, c);
}
function ventanaAbierta_(c, hoy) { return !!c.habilitado && (!c.desde || hoy >= c.desde) && (!c.hasta || hoy <= c.hasta); }

/* ---------------- Cálculo del periodo (igual que Vacaciones.html) ---------------- */

function pad2_(n) { return ('0' + n).slice(-2); }
function isoD_(d) { return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate()); }
function parseISO_(s) { const p = s.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function hoyFecha_() { const p = hoyISO_().split('-').map(Number); return new Date(p[0], p[1] - 1, p[2], 12); }

function vacPersona_(v, bp) {
  const base = v.dgac25 || v.dgacOrig || v.dgac26;
  const md = base.split('-').map(Number);
  const fechas = new Map();
  v.dias.forEach(d => fechas.set(d, 'registro'));
  v.bloques.forEach(b => { for (let d = parseISO_(b.desde); d <= parseISO_(b.hasta); d.setDate(d.getDate() + 1)) if (!fechas.has(isoD_(d))) fechas.set(isoD_(d), 'bloque'); });
  const sinFecha = sinFechaDe_(v.programado, fechas, CONFIG.vacSinFecha);
  return Object.assign({}, v, { bp: bp, mm: md[1], dd: md[2], fechas: fechas, sinFecha: sinFecha, fechaBase: base, sinDgacPropia: !v.dgac25 });
}
/** "2026-11", "2027-01" -> ["2026-11", "2026-12", "2027-01"] */
function mesesEntre_(a, b) {
  const out = [];
  if (!/^\d{4}-\d{2}$/.test(a || '') || !/^\d{4}-\d{2}$/.test(b || '')) return out;
  let [y, m] = a.split('-').map(Number);
  while (y + '-' + pad2_(m) <= b && out.length < 120) { out.push(y + '-' + pad2_(m)); if (++m > 12) { m = 1; y++; } }
  return out;
}
/**
 * Días de VacacionesProgramado sin fecha exacta, dentro del rango rango.desde–rango.hasta:
 * lo programado en esos meses menos los días con fecha que ya caen en ellos.
 */
function sinFechaDe_(programado, fechas, rango) {
  const meses = mesesEntre_((rango || {}).desde, (rango || {}).hasta);
  const prog = programado || {}, sinFecha = [];
  const enRango = Array.from(fechas.keys()).filter(k => meses.indexOf(k.slice(0, 7)) >= 0);
  let rest = Math.max(0, meses.reduce((s, m) => s + (prog[m] || 0), 0) - enRango.length);
  meses.forEach(m => { const dm = enRango.filter(k => k.indexOf(m) === 0).length; const n = Math.min(rest, Math.max(0, (prog[m] || 0) - dm)); if (n > 0) { sinFecha.push({ mes: m, n: n }); rest -= n; } });
  return sinFecha;
}
function inicioPer_(vp, Y) { return new Date(Y, vp.mm - 1, vp.dd); }
function finPer_(vp, Y) { return new Date(Y + 1, vp.mm - 1, vp.dd - 1); }
/** Derecho del periodo que empieza en Y. Sin dato en la hoja: +1 día por renovación (mínimo 30, máximo 45). */
function derechoPer_(vp, Y) {
  const d = vp.derechos || {};
  if (d[Y] != null) return d[Y];
  const ys = Object.keys(d).map(Number).sort((a, b) => a - b);
  if (!ys.length) return null;
  if (Y < ys[0]) return Math.max(30, d[ys[0]] - (ys[0] - Y));
  const prev = ys.filter(y => y < Y).pop();
  return Math.min(45, d[prev] + (Y - prev));
}
function anioDeFecha_(vp, f) { return f >= inicioPer_(vp, f.getFullYear()) ? f.getFullYear() : f.getFullYear() - 1; }
function periodoVac_(vp, Y, hoy, extra) {
  const ini = inicioPer_(vp, Y), fin = finPer_(vp, Y), a = isoD_(ini), b = isoD_(fin), h = isoD_(hoy);
  let tom = 0, prog = 0, sol = 0;
  for (const k of vp.fechas.keys()) if (k >= a && k <= b) { if (k <= h) tom++; else prog++; }
  extra.aprob.forEach(k => { if (k >= a && k <= b && !vp.fechas.has(k)) prog++; });
  extra.pend.forEach(k => { if (k >= a && k <= b && !vp.fechas.has(k)) sol++; });
  vp.sinFecha.forEach(s => { const mid = s.mes + '-15'; if (mid >= a && mid <= b) prog += s.n; });
  const der = derechoPer_(vp, Y);
  return { Y: Y, der: der, tom: tom, prog: prog, sol: sol, disp: der - tom - prog - sol, etiqueta: etiquetaPer_(Y) };
}
function extraDe_(sols, bp) {
  const aprob = new Set(), pend = new Set();
  sols.filter(s => s.bp === bp).forEach(s => s.dias.forEach(k => { if (s.estado === 'aprobada') aprob.add(k); else if (s.estado === 'pendiente') pend.add(k); }));
  return { aprob: aprob, pend: pend };
}

/* ---------------- Datos para la página ---------------- */

function datosVac_(equipo, bp) {
  const v = leerVacaciones_();
  const personas = {};
  Object.keys(v.personas).forEach(k => { if (equipo || k === bp) personas[k] = v.personas[k]; });
  const sols = leerSolicitudesVac_().filter(s => equipo || s.bp === bp).map(s => { const o = Object.assign({}, s); delete o.fila; return o; });
  return { personas: personas, ventanaUltima: v.ventanaUltima, fuente: v.fuente, cfg: cfgVac_(), solicitudes: sols, mesesSinFecha: CONFIG.vacSinFecha || null, escalafon: equipo ? escalafonVac_() : {} };
}

/* =================== Funciones públicas =================== */

/** El instructor envía los días que eligió. Se valida todo de nuevo en el servidor. */
function apiEnviarSolicitudVac(token, dias, disp, comentario) {
  return responder_(() => {
    const u = sesion_(token).u;
    if (u.rol !== 'inst') throw falla_('Solo los instructores pueden solicitar vacaciones desde la app.');
    const cfg = cfgVac_(), hoy = hoyISO_();
    if (!ventanaAbierta_(cfg, hoy)) throw falla_('La ventana de solicitud de vacaciones no está habilitada.');
    if (!Array.isArray(dias) || !dias.length) throw falla_('Elige al menos un día.');
    if (dias.length > 120) throw falla_('Son demasiados días en una sola solicitud.');
    const com = comentarioVac_(comentario);
    const lista = Array.from(new Set(dias.map(String))).sort();
    lista.forEach(k => {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(k) || isoD_(parseISO_(k)) !== k) throw falla_('Fecha inválida: ' + k);
      if (k <= hoy) throw falla_('Solo puedes pedir días a partir de mañana.');
      const m = k.slice(0, 7);
      if (m < cfg.mesDesde || m > cfg.mesHasta) throw falla_('Solo se pueden pedir días entre ' + cfg.mesDesde + ' y ' + cfg.mesHasta + '.');
    });
    const conRol = {};
    leerRoles_().forEach(r => { if (r.instructores.some(i => i.bp === u.bp)) conRol[r.mes] = true; });
    const mesConRol = lista.find(k => conRol[k.slice(0, 7)]);
    if (mesConRol) throw falla_('El rol de ' + mesConRol.slice(0, 7) + ' ya está publicado: los cambios de ese mes se coordinan directamente con coordinación.');
    return conBloqueo_(() => {
      const v = leerVacaciones_().personas[u.bp];
      if (!v) throw falla_('No tienes datos en la hoja Vacaciones. Consulta con coordinación.');
      const vp = vacPersona_(v, u.bp);
      const sols = leerSolicitudesVac_();
      const ex = extraDe_(sols, u.bp);
      const hoyF = hoyFecha_();
      lista.forEach(k => {
        if (vp.fechas.has(k) || ex.aprob.has(k) || ex.pend.has(k)) throw falla_('El ' + k.split('-').reverse().join('/') + ' ya figura como vacaciones o en otra solicitud.');
      });
      const porY = {};
      lista.forEach(k => { const Y = anioDeFecha_(vp, parseISO_(k)); porY[Y] = (porY[Y] || 0) + 1; });
      Object.keys(porY).forEach(Y => {
        const p = periodoVac_(vp, Number(Y), hoyF, ex);
        if (porY[Y] > p.disp) throw falla_('En el periodo ' + p.etiqueta + ' te quedan ' + Math.max(0, p.disp) + ' días y estás pidiendo ' + porY[Y] + '.');
      });
      const sh = hojaSolVac_();
      const id = 's' + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36);
      const n = sh.getLastRow() + 1;
      sh.getRange(n, 4).setNumberFormat('@');
      sh.getRange(n, 1, 1, ENC_SOL_VAC_.length).setValues([[id, u.bp, nombreDe_(u), lista.join(', '), lista.length, new Date(), 'pendiente', '', '']]);
      if (com) escribirSolVac_(sh, n, { 'Comentario instructor': com });
      registrar_(u, 'Solicitud de vacaciones', lista.length + ' días' + (com ? ' · con comentario' : ''), disp);
      return { id: id, solicitudes: leerSolicitudesVac_().filter(s => s.bp === u.bp).map(s => { delete s.fila; return s; }) };
    });
  });
}

/** El instructor retira una solicitud que aún no se decide. */
function apiCancelarSolicitudVac(token, id) {
  return responder_(() => {
    const u = sesion_(token).u;
    return conBloqueo_(() => {
      const s = leerSolicitudesVac_().find(x => x.id === String(id));
      if (!s || s.bp !== u.bp) throw falla_('No encontramos esa solicitud.');
      if (s.estado !== 'pendiente') throw yaResuelto_('Esa solicitud', s.estado === 'cancelada' ? 'cancelada' : s.estado, s.decididoPor, s.decididoAt);
      hoja_(CONFIG.hojas.solVac).getRange(s.fila, 7, 1, 3).setValues([['cancelada', new Date(), nombreDe_(u)]]);
      return { solicitudes: leerSolicitudesVac_().filter(x => x.bp === u.bp).map(x => { delete x.fila; return x; }) };
    });
  });
}

/** Coordinación aprueba o rechaza. El comentario es opcional y lo ve el instructor. */
function apiDecidirSolicitudVac(token, id, estado, comentario) {
  return responder_(() => {
    const admin = sesion_(token, { seccion: 'vac', nivel: 2 }).u;
    if (['aprobada', 'rechazada'].indexOf(estado) < 0) throw falla_('Decisión inválida.');
    const com = comentarioVac_(comentario);
    return conBloqueo_(() => {
      const s = leerSolicitudesVac_().find(x => x.id === String(id));
      if (!s) throw falla_('No encontramos esa solicitud.');
      if (s.estado !== 'pendiente') throw yaResuelto_('Esa solicitud de vacaciones', s.estado === 'cancelada' ? 'cancelada por el instructor' : s.estado, s.estado === 'cancelada' ? '' : s.decididoPor, s.decididoAt);
      const sh = hojaSolVac_();
      escribirSolVac_(sh, s.fila, { 'Estado': estado, 'Decidida': new Date(), 'Decidida por': nombreDe_(admin), 'Comentario coordinación': com });
      const u = usuarios_().find(x => x.bp === s.bp);
      registrar_(u, 'Vacaciones ' + estado, s.n + ' días · por ' + nombreDe_(admin) + (com ? ' · con comentario' : ''), '');
      return { solicitudes: leerSolicitudesVac_().map(x => { delete x.fila; return x; }) };
    });
  });
}

/**
 * El instructor responde a una decisión: "enterado" cierra el ciclo; "revision" pide a coordinación
 * que la revise (motivo obligatorio). Solo se puede pedir revisión una vez por solicitud.
 */
function apiResponderSolicitudVac(token, id, tipo, motivo, disp) {
  return responder_(() => {
    const u = sesion_(token).u;
    if (['enterado', 'revision'].indexOf(tipo) < 0) throw falla_('Respuesta inválida.');
    const mot = comentarioVac_(motivo);
    if (tipo === 'revision' && !mot) throw falla_('Cuéntale a coordinación por qué pides la revisión.');
    return conBloqueo_(() => {
      const s = leerSolicitudesVac_().find(x => x.id === String(id));
      if (!s || s.bp !== u.bp) throw falla_('No encontramos esa solicitud.');
      if (['aprobada', 'rechazada'].indexOf(s.estado) < 0) throw falla_('Esa solicitud todavía no tiene una decisión.');
      if (s.respuesta) throw falla_('Ya respondiste a esta decisión.');
      if (tipo === 'revision' && s.revAtendidaAt) throw falla_('Coordinación ya revisó esta solicitud; solo puedes marcarla como enterado.');
      const valores = { 'Respuesta instructor': tipo, 'Respondida': new Date() };
      if (tipo === 'revision') valores['Motivo revisión'] = mot;
      escribirSolVac_(hojaSolVac_(), s.fila, valores);
      registrar_(u, tipo === 'revision' ? 'Pidió revisión de vacaciones' : 'Enterado de vacaciones', s.estado + ' · ' + s.n + ' días', disp);
      return { solicitudes: leerSolicitudesVac_().filter(x => x.bp === u.bp).map(x => { delete x.fila; return x; }) };
    });
  });
}

/**
 * Coordinación atiende una revisión pedida: puede mantener o cambiar la decisión y dejar un comentario.
 * La respuesta del instructor se limpia para que vea el resultado y marque "enterado".
 */
function apiAtenderRevisionVac(token, id, estado, comentario) {
  return responder_(() => {
    const admin = sesion_(token, { seccion: 'vac', nivel: 2 }).u;
    if (['aprobada', 'rechazada'].indexOf(estado) < 0) throw falla_('Decisión inválida.');
    const com = comentarioVac_(comentario);
    return conBloqueo_(() => {
      const s = leerSolicitudesVac_().find(x => x.id === String(id));
      if (!s) throw falla_('No encontramos esa solicitud.');
      if (s.revAtendidaAt) throw yaResuelto_('Esa revisión', 'atendida', s.decididoPor, s.revAtendidaAt);
      if (s.respuesta !== 'revision') throw falla_('Esa solicitud no tiene una revisión pendiente.', 'YA_RESUELTO');
      const valores = { 'Revisión atendida': new Date(), 'Respuesta instructor': '', 'Respondida': '' };
      if (com) valores['Comentario coordinación'] = com;
      if (estado !== s.estado) { valores['Estado'] = estado; valores['Decidida'] = new Date(); valores['Decidida por'] = nombreDe_(admin); }
      escribirSolVac_(hojaSolVac_(), s.fila, valores);
      const u = usuarios_().find(x => x.bp === s.bp);
      registrar_(u, 'Revisión de vacaciones atendida', (estado !== s.estado ? 'cambió a ' + estado : 'se mantiene ' + estado) + ' · por ' + nombreDe_(admin), '');
      return { solicitudes: leerSolicitudesVac_().map(x => { delete x.fila; return x; }) };
    });
  });
}

function apiGuardarCfgVac(token, c) {
  return responder_(() => {
    sesion_(token, { seccion: 'vac', nivel: 2 });
    c = c || {};
    const f = s => { s = String(s || '').trim(); if (s && !/^\d{4}-\d{2}-\d{2}$/.test(s)) throw falla_('Fecha inválida: ' + s); return s; };
    const m = (s, def) => { s = String(s || '').trim(); if (!s) return def; if (!/^\d{4}-\d{2}$/.test(s)) throw falla_('Mes inválido: ' + s); return s; };
    const cupo = c.cupo == null || c.cupo === '' ? cfgVac_().cupo : Math.round(Number(c.cupo));
    if (!(cupo >= 1 && cupo <= 20)) throw falla_('El cupo por día debe estar entre 1 y 20 instructores.');
    const cfg = { habilitado: !!c.habilitado, desde: f(c.desde), hasta: f(c.hasta), mesDesde: m(c.mesDesde, CFG_VAC_DEF_.mesDesde), mesHasta: m(c.mesHasta, CFG_VAC_DEF_.mesHasta), cupo: cupo };
    if (cfg.desde && cfg.hasta && cfg.desde > cfg.hasta) throw falla_('La fecha "Abierta desde" es posterior a "Abierta hasta".');
    if (cfg.mesDesde > cfg.mesHasta) throw falla_('El mes inicial es posterior al mes final.');
    props_().setProperty('CFG_VAC', JSON.stringify(cfg));
    return cfg;
  });
}

/* ---------------- Escalafón (prioridad cuando una ventana supera el cupo) ---------------- */

/** {bp: {anio, trim}} cargado por coordinación. Sin dato, la página usa la Fecha DGAC original. */
function escalafonVac_() {
  try { return JSON.parse(props_().getProperty('ESC_VAC') || '{}'); } catch (e) { return {}; }
}

function apiGuardarEscalafonVac(token, mapa) {
  return responder_(() => {
    sesion_(token, { seccion: 'vac', nivel: 2 });
    const validos = {}; usuarios_().forEach(u => { validos[u.bp] = true; });
    const out = {};
    Object.keys(mapa || {}).forEach(k => {
      const bp = bpTxt_(k); if (!validos[bp]) return;
      const anio = Number(mapa[k] && mapa[k].anio) || null, trim = Number(mapa[k] && mapa[k].trim) || null;
      if (anio !== null && !(anio >= 1970 && anio <= 2100)) throw falla_('Año de escalafón inválido para el BP ' + bp + '.');
      if (trim !== null && [1, 2, 3, 4].indexOf(trim) < 0) throw falla_('Trimestre inválido para el BP ' + bp + '.');
      if (anio || trim) out[bp] = { anio: anio, trim: trim };
    });
    props_().setProperty('ESC_VAC', JSON.stringify(out));
    return out;
  });
}

/* ---------------- Instalación y migración ---------------- */

function instalarVacaciones_(log) {
  const sv = hoja_(CONFIG.hojas.vacaciones, ENC_VAC_);
  const sd = hoja_(CONFIG.hojas.vacDias, ENC_VAC_DIAS_);
  const sb = hoja_(CONFIG.hojas.vacBloques, ENC_VAC_BLOQ_);
  hojaSolVac_().getRange('D:D').setNumberFormat('@');
  const P = SEMILLA_VACACIONES.personas;
  const sembrar = sv.getLastRow() < 2;
  if (sembrar) {
    const filas = Object.keys(P).map(bp => { const p = P[bp]; return [bp, p.nombreVac || '', p.dgacOrig || '', p.dgac25 || '', p.dgac26 || '']; });
    sv.getRange('A:A').setNumberFormat('@');
    sv.getRange(2, 1, filas.length, ENC_VAC_.length).setValues(filas);
    log.push('Vacaciones: ' + filas.length + ' instructores.');
  }
  if (sd.getLastRow() < 2) {
    const filas = []; Object.keys(P).forEach(bp => P[bp].dias.forEach(d => filas.push([bp, d])));
    if (filas.length) { sd.getRange(2, 1, filas.length, 2).setValues(filas); log.push('VacacionesDias: ' + filas.length + ' días.'); }
  }
  if (sb.getLastRow() < 2) {
    const filas = []; Object.keys(P).forEach(bp => P[bp].bloques.forEach(b => filas.push([bp, b.desde, b.hasta, b.n, (b.registro || '').slice(0, 10)])));
    if (filas.length) { sb.getRange(2, 1, filas.length, 5).setValues(filas); log.push('VacacionesBloques: ' + filas.length + ' bloques.'); }
  }
  if (!ss_().getSheetByName(CONFIG.hojas.vacPeriodos)) {
    if (sembrar) {
      crearHojasPeriodos_(Object.keys(P).map(bp => Object.assign({ bp: bp }, legadoDeValores_(P[bp].meta2526, P[bp].meta2627, P[bp].tomados2526, P[bp].mensual2026))), log);
    } else {
      migrarVacaciones_(log);
    }
  }
  if (!props_().getProperty('VAC_META')) props_().setProperty('VAC_META', JSON.stringify({ ventanaUltima: SEMILLA_VACACIONES.ventanaUltima, fuente: SEMILLA_VACACIONES.fuente }));
  if (!props_().getProperty('CFG_VAC')) props_().setProperty('CFG_VAC', JSON.stringify(CFG_VAC_DEF_));
}

/**
 * Ejecútala desde el editor (una vez). Copia las columnas por año de la hoja Vacaciones a
 * VacacionesPeriodos y VacacionesProgramado y compara los saldos de ambos modelos.
 * No borra ni cambia los datos de la hoja Vacaciones: solo marca en gris las columnas que dejan de usarse.
 * Para volver al modelo anterior basta con borrar (o renombrar) la hoja VacacionesPeriodos.
 */
function migrarVacaciones() {
  soloEditor_();
  const log = [];
  if (ss_().getSheetByName(CONFIG.hojas.vacPeriodos)) log.push('Ya estaba migrado: existe la hoja ' + CONFIG.hojas.vacPeriodos + '.');
  else migrarVacaciones_(log);
  Logger.log(log.join('\n'));
}

/** Compara los saldos calculados con las columnas antiguas y con las hojas nuevas. Solo lee. */
function verificarVacaciones() {
  soloEditor_();
  Logger.log(verificarVacaciones_());
}

function migrarVacaciones_(log) {
  conBloqueo_(() => {
    const sh = hoja_(CONFIG.hojas.vacaciones);
    const t = filas_(sh);
    const c = { bp: col_(t.head, 'bp'), m25: col_(t.head, 'derecho 2025-26'), m26: col_(t.head, 'derecho 2026-27'), tom: col_(t.head, 'tomados 2025-26'), men: col_(t.head, 'programado por mes 2026') };
    const porBp = {};
    t.rows.forEach(r => {
      const bp = bpTxt_(celda_(r, c.bp)); if (!bp) return;
      porBp[bp] = Object.assign({ bp: bp }, legadoDeValores_(numero_(celda_(r, c.m25)), numero_(celda_(r, c.m26)), numero_(celda_(r, c.tom)), mensualDeTexto_(celda_(r, c.men))));
    });
    crearHojasPeriodos_(Object.keys(porBp).map(k => porBp[k]), log);
    // Las columnas antiguas quedan como respaldo, en gris y con una nota.
    const encLegado = ENC_VAC_LEGADO_.map(normTxt_);
    sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach((h, i) => {
      if (encLegado.indexOf(normTxt_(h)) < 0) return;
      sh.getRange(1, i + 1, sh.getMaxRows(), 1).setFontColor('#9aa0a6');
      sh.getRange(1, i + 1).setNote('Ya no se usa desde el ' + hoyISO_() + ': estos datos se copiaron a ' + CONFIG.hojas.vacPeriodos + ' y ' + CONFIG.hojas.vacProgramado + '. Edita esas hojas. Puedes borrar esta columna cuando confirmes que todo está bien.');
    });
  });
  log.push(verificarVacaciones_());
}

/** Escribe VacacionesProgramado y luego VacacionesPeriodos. La app cambia de modelo cuando aparece VacacionesPeriodos. */
function crearHojasPeriodos_(datos, log) {
  const ss = ss_();
  const per = [], prog = [];
  datos.forEach(p => {
    Object.keys(Object.assign({}, p.derechos, p.tomadosRef)).map(Number).sort((a, b) => a - b).forEach(Y => {
      per.push([p.bp, etiquetaPer_(Y), p.derechos[Y] != null ? p.derechos[Y] : '', p.tomadosRef[Y] != null ? p.tomadosRef[Y] : '']);
    });
    Object.keys(p.programado).sort().forEach(m => prog.push([p.bp, m, p.programado[m]]));
  });
  const sg = hoja_(CONFIG.hojas.vacProgramado, ENC_VAC_PROG_);
  if (sg.getLastRow() > 1) sg.getRange(2, 1, sg.getLastRow() - 1, sg.getLastColumn()).clearContent(); // restos de un intento anterior
  sg.getRange('A:B').setNumberFormat('@');
  if (prog.length) sg.getRange(2, 1, prog.length, ENC_VAC_PROG_.length).setValues(prog);
  // Se llena con otro nombre y se renombra al final, para que la app nunca lea la hoja a medio escribir.
  const tmp = CONFIG.hojas.vacPeriodos + '_migrando';
  if (ss.getSheetByName(tmp)) ss.deleteSheet(ss.getSheetByName(tmp));
  const sp = hoja_(tmp, ENC_VAC_PER_);
  sp.getRange('A:B').setNumberFormat('@');
  if (per.length) sp.getRange(2, 1, per.length, ENC_VAC_PER_.length).setValues(per);
  SpreadsheetApp.flush();
  sp.setName(CONFIG.hojas.vacPeriodos);
  log.push(CONFIG.hojas.vacPeriodos + ': ' + per.length + ' filas. ' + CONFIG.hojas.vacProgramado + ': ' + prog.length + ' filas.');
}

function verificarVacaciones_() {
  if (!ss_().getSheetByName(CONFIG.hojas.vacPeriodos)) return 'Verificación: todavía no existe la hoja ' + CONFIG.hojas.vacPeriodos + ' (ejecuta migrarVacaciones).';
  const a = leerVacaciones_('legado').personas, b = leerVacaciones_('nuevo').personas;
  const hoy = hoyFecha_(), sols = leerSolicitudesVac_(), dif = [];
  const bps = Array.from(new Set(Object.keys(a).concat(Object.keys(b)))).sort();
  bps.forEach(bp => {
    if (!a[bp] || !b[bp]) { dif.push('BP ' + bp + ': solo figura en ' + (a[bp] ? 'las columnas antiguas' : 'las hojas nuevas')); return; }
    const va = vacPersona_(a[bp], bp), vb = vacPersona_(b[bp], bp), ex = extraDe_(sols, bp);
    const d25 = a[bp].derechos['2025'], d26 = a[bp].derechos['2026'];
    for (let Y = 2023; Y <= 2028; Y++) {
      const pa = periodoVac_(va, Y, hoy, ex), pb = periodoVac_(vb, Y, hoy, ex);
      // Regla anterior del derecho (años fijos), para confirmar que la regla general da lo mismo.
      const antes = Y === 2025 ? d25 : Y === 2026 ? d26 : Y < 2025 ? Math.max(30, d25 - (2025 - Y)) : Math.min(45, d26 + (Y - 2026));
      if (JSON.stringify(pa) !== JSON.stringify(pb)) dif.push('BP ' + bp + ' ' + pa.etiqueta + ': columnas antiguas ' + JSON.stringify(pa) + ' · hojas nuevas ' + JSON.stringify(pb));
      else if (antes !== pa.der) dif.push('BP ' + bp + ' ' + pa.etiqueta + ': derecho con la regla anterior ' + antes + ' · ahora ' + pa.der);
    }
  });
  return dif.length
    ? 'VERIFICACIÓN: ' + dif.length + ' diferencia(s). Revisa antes de seguir (para volver atrás, borra la hoja ' + CONFIG.hojas.vacPeriodos + '):\n' + dif.join('\n')
    : 'VERIFICACIÓN OK: ' + bps.length + ' instructores; periodos 2023-24 a 2028-29 iguales en ambos modelos.';
}
