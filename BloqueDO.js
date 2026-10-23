/**
 * Bloque de 4 días libres (DO).
 *
 * Coordinación abre una ventana para un mes que todavía se está armando (normalmente el siguiente al rol recién
 * publicado). En ella cada instructor puede pedir un bloque de 4 días libres seguidos que incluya sábado y
 * domingo: de jueves a domingo o de sábado a martes (BLOQUES_DO_). No se puede pedir en un mes en que el
 * instructor tiene vacaciones (tomadas, programadas o solicitadas).
 *
 * Como el rol de ese mes todavía no existe en la app, la solicitud se guarda aparte (hoja SolicitudesDO) y
 * coordinación la tiene en cuenta al armar el Google Sheet del mes. Cuando ese rol se publica, la pestaña
 * Vacaciones muestra si cada bloque aprobado ya figura como DO en el rol.
 *
 * La ventana (habilitada, mes, fechas y cupo por día) se guarda en las propiedades del script (CFG_DO) y la
 * cambia coordinación desde la pestaña Vacaciones (mismo permiso que Vacaciones).
 */

const ENC_DO_ = ['ID', 'BP', 'Nombre', 'Mes', 'Desde', 'Hasta', 'Bloque', 'Estado', 'Enviada', 'Decidida', 'Decidida por',
  'Comentario instructor', 'Comentario coordinación'];
const CFG_DO_DEF_ = { habilitado: false, mes: '', desde: '', hasta: '', cupo: 0 };
const DIAS_DO_ = 4;
// Día de la semana en que puede empezar el bloque (0 = domingo) → nombre. Siempre 4 días seguidos con sábado y domingo.
const BLOQUES_DO_ = { 4: 'Jueves a domingo', 6: 'Sábado a martes' };

function hojaDO_() { return asegurarColumnas_(hoja_(CONFIG.hojas.solDO, ENC_DO_), ENC_DO_); }

function cfgDO_() {
  let c = {};
  try { c = JSON.parse(props_().getProperty('CFG_DO') || '{}'); } catch (e) { c = {}; }
  return Object.assign({}, CFG_DO_DEF_, c);
}
function ventanaDOAbierta_(c, hoy) { return !!c.habilitado && !!c.mes && (!c.desde || hoy >= c.desde) && (!c.hasta || hoy <= c.hasta); }

function leerSolicitudesDO_() {
  const sh = ss_().getSheetByName(CONFIG.hojas.solDO);
  if (!sh) return [];
  const t = filas_(sh);
  const c = {}; ENC_DO_.forEach(h => { c[h] = col_(t.head, h); });
  return t.rows.map((r, i) => {
    const g = h => celda_(r, c[h]);
    const est = normTxt_(g('Estado'));
    return {
      fila: i + 2, id: String(g('ID') || ''), bp: bpTxt_(g('BP')), nombre: String(g('Nombre') || ''), mes: mesTxt_(g('Mes')),
      desde: fechaISO_(g('Desde')), hasta: fechaISO_(g('Hasta')), bloque: String(g('Bloque') || ''),
      estado: /^aprob/.test(est) ? 'aprobada' : /^rechaz/.test(est) ? 'rechazada' : /^cancel/.test(est) ? 'cancelada' : 'pendiente',
      at: fechaCelda_(g('Enviada')), decididoAt: fechaCelda_(g('Decidida')), decididoPor: String(g('Decidida por') || ''),
      comInst: String(g('Comentario instructor') || ''), comCoord: String(g('Comentario coordinación') || '')
    };
  }).filter(s => s.id && s.bp && s.mes && /^\d{4}-\d{2}-\d{2}$/.test(s.desde));
}

/** Valida el día de inicio: tiene que ser jueves o sábado y el bloque tiene que quedar completo dentro del mes. */
function bloqueDO_(mes, dia) {
  dia = Number(dia);
  if (!(dia >= 1 && dia + DIAS_DO_ - 1 <= diasDelMes_(mes))) throw falla_('El bloque de ' + DIAS_DO_ + ' días debe quedar completo dentro del mes.');
  const nombre = BLOQUES_DO_[new Date(Number(mes.slice(0, 4)), Number(mes.slice(5, 7)) - 1, dia).getDay()];
  if (!nombre) throw falla_('El bloque debe ser de jueves a domingo o de sábado a martes.');
  const dias = []; for (let k = 0; k < DIAS_DO_; k++) dias.push(dia + k);
  return { desde: mes + '-' + pad2_(dia), hasta: mes + '-' + pad2_(dia + DIAS_DO_ - 1), dias: dias, bloque: nombre };
}
const diasDeSolDO_ = s => { const d0 = Number(s.desde.slice(8, 10)), out = []; for (let k = 0; k < DIAS_DO_; k++) out.push(d0 + k); return out; };

/** Instructores con bloque pedido o aprobado por día del mes: {día: n}. Sin nombres (lo ve el instructor para el cupo). */
function ocupacionDO_(sols, mes) {
  const oc = {};
  sols.filter(s => s.mes === mes && (s.estado === 'pendiente' || s.estado === 'aprobada')).forEach(s => diasDeSolDO_(s).forEach(d => { oc[d] = (oc[d] || 0) + 1; }));
  return oc;
}

/** Motivo por el que no puede pedir bloque ese mes por vacaciones ('' si no tiene). */
function vacacionesEnMesDO_(bp, mes) {
  const nm = mesLargoTxt_(mes);
  const v = leerVacaciones_().personas[bp];
  if (v) {
    let vp = null;
    try { vp = vacPersona_(v, bp); } catch (e) { vp = null; } // sin fecha DGAC no se puede calcular el periodo
    if (vp && Array.from(vp.fechas.keys()).some(k => k.slice(0, 7) === mes)) return 'Tienes vacaciones en ' + nm + '.';
    if (vp && vp.sinFecha.some(s => s.mes === mes && s.n > 0)) return 'Tienes vacaciones programadas en ' + nm + '.';
  }
  const sv = leerSolicitudesVac_().find(s => s.bp === bp && (s.estado === 'aprobada' || s.estado === 'pendiente') && s.dias.some(k => k.slice(0, 7) === mes));
  if (sv) return sv.estado === 'aprobada' ? 'Tienes una solicitud de vacaciones aprobada en ' + nm + '.'
    : 'Tienes una solicitud de vacaciones pendiente en ' + nm + ' (si coordinación la rechaza, podrás pedir el bloque).';
  return '';
}

/** Datos para la página. Si algo falla, devuelve la ventana cerrada: nunca impide cargar la app. */
function datosDO_(equipo, u) {
  try {
    const cfg = cfgDO_(), sols = leerSolicitudesDO_();
    const out = {
      cfg: cfg, abierta: ventanaDOAbierta_(cfg, hoyISO_()), dias: DIAS_DO_,
      solicitudes: sols.filter(s => equipo || s.bp === u.bp).map(s => { delete s.fila; return s; })
    };
    if (u.rol === 'inst' && cfg.mes) { out.ocupacion = ocupacionDO_(sols, cfg.mes); out.vacMotivo = vacacionesEnMesDO_(u.bp, cfg.mes); }
    return out;
  } catch (e) {
    console.warn('Bloques de días libres no disponibles: ' + e.message);
    return { cfg: CFG_DO_DEF_, abierta: false, dias: DIAS_DO_, solicitudes: [] };
  }
}

/* =================== Funciones públicas =================== */

/** El instructor pide un bloque que empieza el día dia del mes abierto. Se valida todo de nuevo en el servidor. */
function apiSolicitarDO(token, dia, comentario, disp) {
  return responder_(() => {
    const u = sesion_(token).u;
    if (u.rol !== 'inst') throw falla_('Solo los instructores pueden pedir un bloque de días libres.');
    const cfg = cfgDO_(), hoy = hoyISO_();
    if (!ventanaDOAbierta_(cfg, hoy)) throw falla_('La solicitud de bloque de días libres no está habilitada.');
    const mes = cfg.mes, nm = mesLargoTxt_(mes);
    const com = comentarioVac_(comentario);
    const b = bloqueDO_(mes, dia);
    if (b.desde <= hoy) throw falla_('Solo puedes pedir bloques a partir de mañana.');
    const motivo = vacacionesEnMesDO_(u.bp, mes);
    if (motivo) throw falla_(motivo + ' No se puede pedir un bloque de días libres en un mes con vacaciones.');
    return conBloqueo_(() => {
      const sols = leerSolicitudesDO_();
      const mia = sols.find(s => s.bp === u.bp && s.mes === mes && (s.estado === 'pendiente' || s.estado === 'aprobada'));
      if (mia) throw falla_('Ya tienes un bloque ' + (mia.estado === 'aprobada' ? 'aprobado' : 'pedido') + ' para ' + nm + ' (del ' + Number(mia.desde.slice(8)) + ' al ' + Number(mia.hasta.slice(8)) + ').' +
        (mia.estado === 'pendiente' ? ' Cancélalo si quieres pedir otro.' : ''));
      if (cfg.cupo > 0) {
        const oc = ocupacionDO_(sols, mes), lleno = b.dias.find(d => (oc[d] || 0) >= cfg.cupo);
        if (lleno) throw falla_('El día ' + lleno + ' ya tiene el cupo completo (' + cfg.cupo + ' instructores con bloque). Elige otro bloque.');
      }
      const sh = hojaDO_();
      const id = 'd' + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36);
      const n = sh.getLastRow() + 1;
      sh.getRange(n, 1, 1, 7).setNumberFormat('@');
      escribirFila_(sh, n, {
        'ID': id, 'BP': u.bp, 'Nombre': nombreDe_(u), 'Mes': mes, 'Desde': b.desde, 'Hasta': b.hasta, 'Bloque': b.bloque,
        'Estado': 'pendiente', 'Enviada': new Date(), 'Comentario instructor': com
      });
      registrar_(u, 'Bloque de días libres solicitado', b.bloque + ' · ' + b.desde + ' a ' + b.hasta, disp);
      return { do: datosDO_(false, u) };
    });
  });
}

/** El instructor retira un bloque que aún no se decide. */
function apiCancelarDO(token, id) {
  return responder_(() => {
    const u = sesion_(token).u;
    return conBloqueo_(() => {
      const s = leerSolicitudesDO_().find(x => x.id === String(id));
      if (!s || s.bp !== u.bp) throw falla_('No encontramos esa solicitud.');
      if (s.estado !== 'pendiente') throw yaResuelto_('Esa solicitud', s.estado, s.decididoPor, s.decididoAt);
      escribirFila_(hojaDO_(), s.fila, { 'Estado': 'cancelada', 'Decidida': new Date(), 'Decidida por': nombreDe_(u) });
      registrar_(u, 'Bloque de días libres cancelado', s.desde + ' a ' + s.hasta, '');
      return { do: datosDO_(false, u) };
    });
  });
}

/** Coordinación aprueba o rechaza. El comentario es opcional y lo ve el instructor. */
function apiDecidirDO(token, id, estado, comentario) {
  return responder_(() => {
    const admin = sesion_(token, { seccion: 'vac', nivel: 2 }).u;
    if (['aprobada', 'rechazada'].indexOf(estado) < 0) throw falla_('Decisión inválida.');
    const com = comentarioVac_(comentario);
    return conBloqueo_(() => {
      const s = leerSolicitudesDO_().find(x => x.id === String(id));
      if (!s) throw falla_('No encontramos esa solicitud.');
      if (s.estado !== 'pendiente') throw yaResuelto_('Esa solicitud de bloque', s.estado === 'cancelada' ? 'cancelada por el instructor' : s.estado, s.estado === 'cancelada' ? '' : s.decididoPor, s.decididoAt);
      escribirFila_(hojaDO_(), s.fila, { 'Estado': estado, 'Decidida': new Date(), 'Decidida por': nombreDe_(admin), 'Comentario coordinación': com });
      registrar_(usuarios_().find(x => x.bp === s.bp), 'Bloque de días libres ' + estado, s.desde + ' a ' + s.hasta + ' · por ' + nombreDe_(admin), '');
      return { do: datosDO_(true, admin) };
    });
  });
}

function apiGuardarCfgDO(token, c) {
  return responder_(() => {
    const admin = sesion_(token, { seccion: 'vac', nivel: 2 }).u;
    c = c || {};
    const f = s => { s = String(s || '').trim(); if (s && !/^\d{4}-\d{2}-\d{2}$/.test(s)) throw falla_('Fecha inválida: ' + s); return s; };
    const mes = String(c.mes || '').trim();
    if (mes && !/^\d{4}-\d{2}$/.test(mes)) throw falla_('Mes inválido: ' + mes);
    const cupo = c.cupo === '' || c.cupo == null ? 0 : Math.round(Number(c.cupo));
    if (!(cupo >= 0 && cupo <= 30)) throw falla_('El cupo por día debe estar entre 0 (sin límite) y 30.');
    const cfg = { habilitado: !!c.habilitado, mes: mes, desde: f(c.desde), hasta: f(c.hasta), cupo: cupo };
    if (cfg.habilitado && !cfg.mes) throw falla_('Elige el mes para el que se piden los bloques.');
    if (cfg.habilitado && cfg.mes < mesActual_()) throw falla_('El mes de los bloques ya pasó.');
    if (cfg.desde && cfg.hasta && cfg.desde > cfg.hasta) throw falla_('La fecha "Abierta desde" es posterior a "Abierta hasta".');
    props_().setProperty('CFG_DO', JSON.stringify(cfg));
    registrar_(admin, 'Ventana de bloques de días libres', (cfg.habilitado ? 'Habilitada' : 'Cerrada') + (cfg.mes ? ' · ' + cfg.mes : ''), '');
    return datosDO_(true, admin);
  });
}
