/**
 * Vacaciones por periodo DGAC.
 *
 * Cada instructor tiene su propio periodo anual, que empieza en su fecha DGAC (día y mes).
 * Hojas del Google Sheet base:
 *   Vacaciones             una fila por instructor: fechas DGAC, derecho de cada periodo y
 *                          días programados por mes de 2026 (para los que no tienen fecha exacta)
 *   VacacionesDias         una fila por día de vacaciones con fecha (tomadas o programadas)
 *   VacacionesBloques      bloques programados (desde, hasta)
 *   SolicitudesVacaciones  lo que piden los instructores desde la app y la decisión de coordinación
 * La ventana de solicitud (abierta/cerrada, fechas y meses que se pueden pedir) se guarda en
 * las propiedades del script y la cambia coordinación desde la pestaña Vacaciones.
 *
 * El cálculo del periodo es el mismo que hace la página (Vacaciones.html): el servidor lo repite
 * para validar cada solicitud, así nadie puede pedir más días de los que le quedan.
 */

const ENC_VAC_ = ['BP', 'Nombre', 'Fecha DGAC original', 'Fecha DGAC 2025', 'Fecha DGAC 2026', 'Derecho 2025-26', 'Derecho 2026-27', 'Tomados 2025-26', 'Programado por mes 2026'];
const ENC_VAC_DIAS_ = ['BP', 'Fecha'];
const ENC_VAC_BLOQ_ = ['BP', 'Desde', 'Hasta', 'Días', 'Registrado'];
const ENC_SOL_VAC_ = ['ID', 'BP', 'Nombre', 'Días', 'Cantidad', 'Enviada', 'Estado', 'Decidida', 'Decidida por'];
const CFG_VAC_DEF_ = { habilitado: false, desde: '', hasta: '', mesDesde: '2027-01', mesHasta: '2027-12' };

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
function mensualATexto_(o) { return Object.keys(o || {}).sort((a, b) => a - b).map(k => k + ':' + o[k]).join('; '); }

function leerVacaciones_() {
  const personas = {};
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
      personas[bp] = {
        nombreVac: String(celda_(r, c.nom) || '').trim(),
        dgacOrig: fechaISO_(celda_(r, c.orig)), dgac25: fechaISO_(celda_(r, c.d25)), dgac26: fechaISO_(celda_(r, c.d26)),
        meta2526: numero_(celda_(r, c.m25)), meta2627: numero_(celda_(r, c.m26)), tomados2526: numero_(celda_(r, c.tom)),
        mensual2026: mensualDeTexto_(celda_(r, c.men)), dias: [], bloques: []
      };
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

function leerSolicitudesVac_() {
  const t = filas_(hoja_(CONFIG.hojas.solVac, ENC_SOL_VAC_));
  const c = {
    id: col_(t.head, 'id'), bp: col_(t.head, 'bp'), nom: col_(t.head, 'nombre'), dias: col_(t.head, 'dias'), n: col_(t.head, 'cantidad'),
    at: col_(t.head, 'enviada'), es: col_(t.head, 'estado'), dec: col_(t.head, 'decidida'), por: col_(t.head, 'decidida por')
  };
  return t.rows.map((r, i) => {
    const at = celda_(r, c.at), dec = celda_(r, c.dec);
    const dias = String(celda_(r, c.dias) || '').split(/[\s,;]+/).map(fechaISO_).filter(x => x && /^\d{4}-\d{2}-\d{2}$/.test(x)).sort();
    return {
      fila: i + 2, id: String(celda_(r, c.id) || ''), bp: bpTxt_(celda_(r, c.bp)), nombre: String(celda_(r, c.nom) || ''),
      dias: dias, n: dias.length, at: esFecha_(at) ? at.toISOString() : String(at || ''),
      estado: normTxt_(celda_(r, c.es)).replace(/^aprobad.*/, 'aprobada').replace(/^rechazad.*/, 'rechazada').replace(/^pendient.*/, 'pendiente') || 'pendiente',
      decididoAt: esFecha_(dec) ? dec.toISOString() : String(dec || ''), decididoPor: String(celda_(r, c.por) || '')
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
  // Noviembre y diciembre 2026: días de la hoja VACACIONES que no tienen fecha exacta
  const m26 = v.mensual2026 || {};
  const sinFecha = []; const sumM = (m26['11'] || 0) + (m26['12'] || 0);
  let dated = 0; for (const k of fechas.keys()) if (k >= '2026-11-01' && k <= '2026-12-31') dated++;
  let rest = Math.max(0, sumM - dated);
  ['11', '12'].forEach(m => { const dm = Array.from(fechas.keys()).filter(k => k.indexOf('2026-' + m) === 0).length; const n = Math.min(rest, Math.max(0, (m26[m] || 0) - dm)); if (n > 0) { sinFecha.push({ mes: '2026-' + m, n: n }); rest -= n; } });
  return Object.assign({}, v, { bp: bp, mm: md[1], dd: md[2], fechas: fechas, sinFecha: sinFecha, fechaBase: base, sinDgacPropia: !v.dgac25 });
}
function inicioPer_(vp, Y) { return new Date(Y, vp.mm - 1, vp.dd); }
function finPer_(vp, Y) { return new Date(Y + 1, vp.mm - 1, vp.dd - 1); }
function derechoPer_(vp, Y) {
  if (Y === 2025) return vp.meta2526;
  if (Y === 2026) return vp.meta2627;
  if (Y < 2025) return Math.max(30, vp.meta2526 - (2025 - Y));
  return Math.min(45, vp.meta2627 + (Y - 2026));
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
  return { Y: Y, der: der, tom: tom, prog: prog, sol: sol, disp: der - tom - prog - sol, etiqueta: Y + '-' + pad2_((Y + 1) % 100) };
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
  return { personas: personas, ventanaUltima: v.ventanaUltima, fuente: v.fuente, cfg: cfgVac_(), solicitudes: sols };
}

/* =================== Funciones públicas =================== */

/** El instructor envía los días que eligió. Se valida todo de nuevo en el servidor. */
function apiEnviarSolicitudVac(token, dias, disp) {
  return responder_(() => {
    const u = sesion_(token).u;
    if (u.rol !== 'inst') throw falla_('Solo los instructores pueden solicitar vacaciones desde la app.');
    const cfg = cfgVac_(), hoy = hoyISO_();
    if (!ventanaAbierta_(cfg, hoy)) throw falla_('La ventana de solicitud de vacaciones no está habilitada.');
    if (!Array.isArray(dias) || !dias.length) throw falla_('Elige al menos un día.');
    if (dias.length > 120) throw falla_('Son demasiados días en una sola solicitud.');
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
      const sh = hoja_(CONFIG.hojas.solVac, ENC_SOL_VAC_);
      const id = 's' + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36);
      const n = sh.getLastRow() + 1;
      sh.getRange(n, 4).setNumberFormat('@');
      sh.getRange(n, 1, 1, ENC_SOL_VAC_.length).setValues([[id, u.bp, nombreDe_(u), lista.join(', '), lista.length, new Date(), 'pendiente', '', '']]);
      registrar_(u, 'Solicitud de vacaciones', lista.length + ' días', disp);
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
      if (s.estado !== 'pendiente') throw falla_('Esa solicitud ya fue decidida por coordinación.');
      hoja_(CONFIG.hojas.solVac).getRange(s.fila, 7, 1, 3).setValues([['cancelada', new Date(), nombreDe_(u)]]);
      return { solicitudes: leerSolicitudesVac_().filter(x => x.bp === u.bp).map(x => { delete x.fila; return x; }) };
    });
  });
}

function apiDecidirSolicitudVac(token, id, estado) {
  return responder_(() => {
    const admin = sesion_(token, { admin: true }).u;
    if (['aprobada', 'rechazada'].indexOf(estado) < 0) throw falla_('Decisión inválida.');
    return conBloqueo_(() => {
      const s = leerSolicitudesVac_().find(x => x.id === String(id));
      if (!s) throw falla_('No encontramos esa solicitud.');
      if (s.estado !== 'pendiente') throw falla_('Esa solicitud ya estaba ' + s.estado + '.');
      hoja_(CONFIG.hojas.solVac).getRange(s.fila, 7, 1, 3).setValues([[estado, new Date(), nombreDe_(admin)]]);
      const u = usuarios_().find(x => x.bp === s.bp);
      registrar_(u, 'Vacaciones ' + estado, s.n + ' días · por ' + nombreDe_(admin), '');
      return { solicitudes: leerSolicitudesVac_().map(x => { delete x.fila; return x; }) };
    });
  });
}

function apiGuardarCfgVac(token, c) {
  return responder_(() => {
    sesion_(token, { admin: true });
    c = c || {};
    const f = s => { s = String(s || '').trim(); if (s && !/^\d{4}-\d{2}-\d{2}$/.test(s)) throw falla_('Fecha inválida: ' + s); return s; };
    const m = (s, def) => { s = String(s || '').trim(); if (!s) return def; if (!/^\d{4}-\d{2}$/.test(s)) throw falla_('Mes inválido: ' + s); return s; };
    const cfg = { habilitado: !!c.habilitado, desde: f(c.desde), hasta: f(c.hasta), mesDesde: m(c.mesDesde, CFG_VAC_DEF_.mesDesde), mesHasta: m(c.mesHasta, CFG_VAC_DEF_.mesHasta) };
    if (cfg.desde && cfg.hasta && cfg.desde > cfg.hasta) throw falla_('La fecha "Abierta desde" es posterior a "Abierta hasta".');
    if (cfg.mesDesde > cfg.mesHasta) throw falla_('El mes inicial es posterior al mes final.');
    props_().setProperty('CFG_VAC', JSON.stringify(cfg));
    return cfg;
  });
}

/* ---------------- Instalación ---------------- */

function instalarVacaciones_(log) {
  const sv = hoja_(CONFIG.hojas.vacaciones, ENC_VAC_);
  const sd = hoja_(CONFIG.hojas.vacDias, ENC_VAC_DIAS_);
  const sb = hoja_(CONFIG.hojas.vacBloques, ENC_VAC_BLOQ_);
  hoja_(CONFIG.hojas.solVac, ENC_SOL_VAC_).getRange('D:D').setNumberFormat('@');
  const P = SEMILLA_VACACIONES.personas;
  if (sv.getLastRow() < 2) {
    const filas = Object.keys(P).map(bp => { const p = P[bp]; return [bp, p.nombreVac || '', p.dgacOrig || '', p.dgac25 || '', p.dgac26 || '', p.meta2526, p.meta2627, p.tomados2526, mensualATexto_(p.mensual2026)]; });
    sv.getRange('I:I').setNumberFormat('@'); // evita que "11:15" se convierta en hora
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
  if (!props_().getProperty('VAC_META')) props_().setProperty('VAC_META', JSON.stringify({ ventanaUltima: SEMILLA_VACACIONES.ventanaUltima, fuente: SEMILLA_VACACIONES.fuente }));
  if (!props_().getProperty('CFG_VAC')) props_().setProperty('CFG_VAC', JSON.stringify(CFG_VAC_DEF_));
}
