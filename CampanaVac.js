/**
 * Campaña de vacaciones por prioridades (noviembre 2026 a diciembre 2027).
 * Reglas: "Modelo Vacaciones Prioridades Instructores LP.md". Días mínimos: VACACIONES v2.xlsx, columna L (31-dic-2027).
 *
 * R1 Cada instructor registra hasta 6 bloques de días seguidos, ordenados de P1 (la más importante) a P6.
 * R2 Días calendario: cuenta todos los días del bloque. Si termina un viernes, se suman el sábado y el domingo.
 * R3 Un día elegido en una prioridad no puede estar en otra (los bloques no se cruzan).
 * R4 La suma de los bloques debe llegar al menos a sus días mínimos.
 * R5 Todo dentro de la ventana (cfg.inicio–cfg.fin) y sin pisar vacaciones tomadas, programadas o solicitadas.
 * R6 Coordinación asigna por rondas: la P1 de todos por escalafón (año y trimestre), luego la P2, y así hasta la P6.
 *    Si un bloque supera el cupo diario en algún día, no se asigna y esa persona sigue con su siguiente prioridad.
 *
 * Hojas: VacacionesMinimos (días mínimos por BP; se crea con los datos del Excel) y VacacionesPrioridades (una fila
 * por instructor y prioridad). Los bloques asignados se escriben en VacacionesBloques: desde ahí cuentan como
 * vacaciones programadas en toda la app. La ventana se guarda en las propiedades del script (CFG_CAMP).
 */

const ENC_MIN_VAC_ = ['BP', 'Nombre', 'Próximo aniversario', 'Saldo al 25-oct', 'Meta 2026-27', 'Mínimo exacto', 'Días mínimos', 'Saldo periodo actual'];
const ENC_PRIO_VAC_ = ['ID', 'BP', 'Nombre', 'Prioridad', 'Desde', 'Hasta', 'Días', 'Fin de semana', 'Estado', 'Enviada', 'Comentario',
  'Ronda', 'Resultado', 'Asignada por', 'Asignada'];
const CFG_CAMP_DEF_ = { habilitado: false, desde: '', hasta: '', inicio: '2026-11-01', fin: '2027-12-31', tope: 0, cupo: 0, plazoEnero: '2026-11-08', alta: 'Abril y Octubre' };
const MAX_PRIO_VAC_ = 6;

// VACACIONES v2.xlsx, cruzado por nombre con la hoja Usuarios.
// [BP, nombre, próximo aniversario (C), saldo al 25-oct (D), meta 2026-27 (K), mínimo exacto (L), saldo final del periodo actual (J)]
// El mínimo se reparte en dos tramos: J hasta la renovación y L − J desde la renovación hasta el fin de la ventana.
const SEMILLA_MINIMOS_VAC_ = [
  ['3779550', 'Jennifert Acurio', '2027-07-15', -3, 39, 41.75, 25.5], ['2823133', 'César Campos', '2027-01-03', -14, 41, 37, -4],
  ['2604360', 'Mariella Carrasco', '2027-05-03', -16, 42, 32.42, 7.92], ['3841387', 'Gianfranco Celiz', '2027-01-17', -5, 38, 42.25, 4.25],
  ['2396710', 'Sebastián Correa', '2027-06-15', 39, 45, 90.17, 67.67], ['967092', 'Erika Dávila', '2027-06-01', 3, 45, 55.5, 33],
  ['2440915', 'Judith Fernández', '2027-10-01', -7, 45, 44.5, 37], ['3217561', 'Claudia Flores', '2027-01-24', -7, 41, 44, 3],
  ['29530', 'Jazmín Guerra', '2027-01-24', 22, 41, 69.58, 32], ['3750335', 'Jefferson Méndez', '2026-12-17', 4, 37, 47, 10],
  ['71348', 'Karla Moz', '2027-08-24', 66, 45, 148.5, 103.5], ['2369624', 'Patricia Najar', '2027-04-15', -10, 45, 42, 12],
  ['3796947', 'Rafael Nieto', '2027-09-09', -9, 39, 35.58, 25.83], ['1271571', 'Christian Rondón', '2026-10-20', 46, 45, 98.5, 46],
  ['2713993', 'Fiorella Ruiz', '2027-02-01', 10, 45, 62.5, 25], ['2415373', 'Milagros Salas', '2027-08-01', 1, 45, 52.67, 37.67],
  ['2843319', 'Karen Santa Cruz', '2027-08-01', -10, 43, 39.33, 25], ['3189967', 'Kevin Segovia', '2027-01-03', -7, 41, 40.58, 3],
  ['2369641', 'Elizabeth Torres', '2027-04-25', -2, 45, 50, 20], ['3852423', 'Gabriela Ungaro', '2027-07-15', -15, 38, 28.58, 12.75],
  ['3134911', 'Javier Zapata', '2027-10-11', -15, 42, 33, 26]
];

function cfgCamp_() {
  let c = {};
  try { c = JSON.parse(props_().getProperty('CFG_CAMP') || '{}'); } catch (e) { c = {}; }
  const r = Object.assign({}, CFG_CAMP_DEF_, c);
  // Cupo por día: si no se guardó con la solicitud, se usa el que ya tenía la configuración de Vacaciones
  if (!(Number(r.cupo) > 0)) r.cupo = Math.max(1, Number(cfgVac_().cupo) || 3);
  return r;
}
const campAbierta_ = (c, hoy) => !!c.habilitado && (!c.desde || hoy >= c.desde) && (!c.hasta || hoy <= c.hasta);

/* ---------------- Hojas ---------------- */

/** Días mínimos por BP. Si la hoja no existe, se crea con los datos del Excel (después se edita en el Google Sheet). */
function leerMinimosVac_() {
  let sh = ss_().getSheetByName(CONFIG.hojas.vacMinimos);
  if (!sh) {
    sh = hoja_(CONFIG.hojas.vacMinimos, ENC_MIN_VAC_);
    sh.getRange(2, 1, SEMILLA_MINIMOS_VAC_.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, SEMILLA_MINIMOS_VAC_.length, ENC_MIN_VAC_.length)
      .setValues(SEMILLA_MINIMOS_VAC_.map(r => r.slice(0, 6).concat([Math.ceil(r[5] - 1e-9), r[6]])));
  }
  asegurarColumnas_(sh, ENC_MIN_VAC_);
  const t = filas_(sh);
  const c = { bp: col_(t.head, 'bp'), nom: col_(t.head, 'nombre'), ani: col_(t.head, 'proximo aniversario'), sal: col_(t.head, 'saldo al'),
    meta: col_(t.head, 'meta'), exa: col_(t.head, 'minimo exacto'), min: col_(t.head, 'dias minimos'), act: col_(t.head, 'saldo periodo actual') };
  const semilla = {}; SEMILLA_MINIMOS_VAC_.forEach(r => { semilla[r[0]] = r[6]; });
  const out = {};
  t.rows.forEach((r, i) => {
    const bp = bpTxt_(celda_(r, c.bp)), min = numero_(celda_(r, c.min));
    if (!bp || min == null) return;
    // Hoja creada antes de esta columna: se completa una vez con el dato del Excel
    let act = numero_(celda_(r, c.act));
    if (act == null && semilla[bp] != null) { act = semilla[bp]; sh.getRange(i + 2, c.act + 1).setValue(act); }
    const minimo = Math.ceil(min), actual = act == null ? null : Math.round(act * 100) / 100;
    out[bp] = { bp: bp, nombre: String(celda_(r, c.nom) || ''), aniversario: fechaISO_(celda_(r, c.ani)) || '',
      saldo: numero_(celda_(r, c.sal)), meta: numero_(celda_(r, c.meta)), exacto: numero_(celda_(r, c.exa)), minimo: minimo,
      // Barra del instructor, tal cual la hoja: amarillo = "Saldo periodo actual" (hasta la renovación) y
      // verde = "Días mínimos" (nuevo periodo, de la renovación al siguiente aniversario)
      actual: actual, nuevo: Math.round(min * 100) / 100 };
  });
  return out;
}

function hojaPrioVac_() { return asegurarColumnas_(hoja_(CONFIG.hojas.vacPrioridades, ENC_PRIO_VAC_), ENC_PRIO_VAC_); }

/** Filas vigentes (las reemplazadas por un reenvío no cuentan). */
function leerPrioridadesVac_() {
  const sh = ss_().getSheetByName(CONFIG.hojas.vacPrioridades);
  if (!sh) return [];
  const t = filas_(sh);
  const c = {}; ENC_PRIO_VAC_.forEach(h => { c[h] = col_(t.head, h); });
  return t.rows.map((r, i) => {
    const g = h => celda_(r, c[h]);
    const est = normTxt_(g('Estado'));
    return {
      fila: i + 2, id: String(g('ID') || ''), bp: bpTxt_(g('BP')), nombre: String(g('Nombre') || ''), prioridad: Number(g('Prioridad')) || 0,
      desde: fechaISO_(g('Desde')) || '', hasta: fechaISO_(g('Hasta')) || '', dias: Number(g('Días')) || 0, fds: Number(g('Fin de semana')) || 0,
      estado: /^asign/.test(est) ? 'asignada' : /^no asign/.test(est) ? 'no_asignada' : /^reempl/.test(est) ? 'reemplazada' : /^retir/.test(est) ? 'retirada' : 'enviada',
      enviada: fechaCelda_(g('Enviada')), comentario: String(g('Comentario') || ''), ronda: Number(g('Ronda')) || 0,
      resultado: String(g('Resultado') || ''), asignadaPor: String(g('Asignada por') || ''), asignada: fechaCelda_(g('Asignada'))
    };
  }).filter(p => p.id && p.bp && p.prioridad && /^\d{4}-\d{2}-\d{2}$/.test(p.desde) && p.estado !== 'reemplazada' && p.estado !== 'retirada');
}

/* ---------------- Reglas ---------------- */

const masDias_ = (iso, n) => { const d = parseISO_(iso); d.setDate(d.getDate() + n); return isoD_(d); };
const diasEntre_ = (a, b) => Math.round((parseISO_(b) - parseISO_(a)) / 864e5) + 1;
function diasDeBloque_(desde, hasta) { const out = []; for (let k = desde; k <= hasta; k = masDias_(k, 1)) out.push(k); return out; }
const finDeSemana_ = dias => dias.filter(k => { const w = parseISO_(k).getDay(); return w === 0 || w === 6; }).length;

/** R2: si el bloque termina un viernes se le suman el sábado y el domingo (si siguen dentro de la ventana). */
function extenderViernes_(hasta, fin) {
  if (parseISO_(hasta).getDay() !== 5) return hasta;
  const dom = masDias_(hasta, 2);
  return dom <= fin ? dom : hasta;
}

/** Días ya ocupados por vacaciones de esta persona (tomadas, programadas con fecha o solicitadas). */
function ocupadosVac_(bp, v, sols) {
  const ocup = new Set();
  const p = v.personas[bp];
  if (p) { try { vacPersona_(p, bp).fechas.forEach((t, k) => ocup.add(k)); } catch (e) { /* sin fecha DGAC */ } }
  // Solicitudes del flujo antiguo: solo las aprobadas ocupan días (las pendientes ya no se atienden)
  extraDe_(sols, bp).aprob.forEach(k => ocup.add(k));
  return ocup;
}

/**
 * Valida y normaliza los bloques que envía el instructor (R1–R5).
 * Devuelve [{prioridad, desde, hasta, dias, fds}] en el orden recibido (P1 primero).
 */
function validarPrioridadesVac_(bloques, cfg, minimo, ocupados, hoy) {
  if (!Array.isArray(bloques) || !bloques.length) throw falla_('Elige al menos un bloque.');
  if (bloques.length > MAX_PRIO_VAC_) throw falla_('Son máximo ' + MAX_PRIO_VAC_ + ' prioridades.');
  const usados = {}, out = [];
  bloques.forEach((b, i) => {
    const P = 'P' + (i + 1);
    const desde = String((b && b.desde) || ''), hasta0 = String((b && b.hasta) || '');
    const ok = s => /^\d{4}-\d{2}-\d{2}$/.test(s) && isoD_(parseISO_(s)) === s;
    if (!ok(desde) || !ok(hasta0)) throw falla_(P + ': fecha inválida.');
    if (hasta0 < desde) throw falla_(P + ': la fecha final es anterior a la inicial.');
    if (desde <= hoy) throw falla_(P + ': solo se pueden elegir días a partir de mañana.');
    if (desde < cfg.inicio || hasta0 > cfg.fin) throw falla_(P + ': debe estar entre el ' + cfg.inicio.split('-').reverse().join('/') + ' y el ' + cfg.fin.split('-').reverse().join('/') + '.');
    const hasta = extenderViernes_(hasta0, cfg.fin);
    const dias = diasDeBloque_(desde, hasta);
    dias.forEach(k => {
      const f = k.split('-').reverse().join('/');
      if (usados[k]) throw falla_(P + ': el ' + f + ' ya está en ' + usados[k] + '. Los bloques no pueden cruzarse.');
      if (ocupados.has(k)) throw falla_(P + ': el ' + f + ' ya figura como vacaciones tomadas, programadas o solicitadas.');
      usados[k] = P;
    });
    out.push({ prioridad: i + 1, desde: desde, hasta: hasta, dias: dias.length, fds: finDeSemana_(dias) });
  });
  const total = out.reduce((s, b) => s + b.dias, 0);
  if (total < minimo) throw falla_('Elegiste ' + total + ' de ' + minimo + ' días. Te faltan ' + (minimo - total) + '.');
  if (cfg.tope > 0 && total > cfg.tope) throw falla_('Elegiste ' + total + ' días y el máximo permitido es ' + cfg.tope + '.');
  return out;
}

/* ---------------- Datos para la página ---------------- */

/** Nunca impide cargar la app: si algo falla, la campaña se muestra cerrada. */
function datosCampana_(equipo, u) {
  try {
    const cfg = cfgCamp_(), minimos = leerMinimosVac_();
    const prios = leerPrioridadesVac_().filter(p => equipo || p.bp === u.bp).map(p => { delete p.fila; return p; });
    const min = {}; Object.keys(minimos).forEach(bp => { if (equipo || bp === u.bp) min[bp] = minimos[bp]; });
    return { cfg: cfg, abierta: campAbierta_(cfg, hoyISO_()), max: MAX_PRIO_VAC_, minimos: min, prioridades: prios };
  } catch (e) {
    console.warn('Campaña de vacaciones no disponible: ' + e.message);
    return { cfg: CFG_CAMP_DEF_, abierta: false, max: MAX_PRIO_VAC_, minimos: {}, prioridades: [] };
  }
}

/* =================== Funciones públicas: instructor =================== */

/** Envía (o reemplaza, mientras la campaña siga abierta) las prioridades P1–P6. Se valida todo en el servidor. */
function apiEnviarPrioridadesVac(token, bloques, comentario, disp) {
  return responder_(() => {
    const u = sesion_(token).u;
    if (u.rol !== 'inst') throw falla_('Solo los instructores registran prioridades de vacaciones.');
    const cfg = cfgCamp_(), hoy = hoyISO_();
    if (!campAbierta_(cfg, hoy)) throw falla_('La solicitud de vacaciones no está abierta.');
    const com = comentarioVac_(comentario);
    const m = leerMinimosVac_()[u.bp];
    if (!m) throw falla_('No tienes días de vacaciones cargados para esta solicitud. Consulta con coordinación.');
    return conBloqueo_(() => {
      const actuales = leerPrioridadesVac_().filter(p => p.bp === u.bp);
      if (actuales.some(p => p.estado === 'asignada' || p.estado === 'no_asignada'))
        throw falla_('Coordinación ya asignó tus vacaciones: no se pueden cambiar desde la app.');
      const v = leerVacaciones_(), sols = leerSolicitudesVac_();
      const lista = validarPrioridadesVac_(bloques, cfg, m.minimo, ocupadosVac_(u.bp, v, sols), hoy);
      const sh = hojaPrioVac_();
      actuales.forEach(p => escribirFila_(sh, p.fila, { 'Estado': 'reemplazada' }));
      const base = 'v' + Date.now().toString(36), ahora = new Date();
      const filas = lista.map(b => {
        const fila = {};
        Object.assign(fila, { 'ID': base + '-' + b.prioridad, 'BP': u.bp, 'Nombre': nombreDe_(u), 'Prioridad': b.prioridad, 'Desde': b.desde, 'Hasta': b.hasta,
          'Días': b.dias, 'Fin de semana': b.fds, 'Estado': 'enviada', 'Enviada': ahora, 'Comentario': b.prioridad === 1 ? com : '' });
        return ENC_PRIO_VAC_.map(h => fila[h] == null ? '' : fila[h]);
      });
      const n = sh.getLastRow() + 1;
      sh.getRange(n, 1, filas.length, ENC_PRIO_VAC_.length).setNumberFormat('@');
      sh.getRange(n, 1, filas.length, ENC_PRIO_VAC_.length).setValues(filas.map(f => f.map(x => x instanceof Date ? x : String(x))));
      sh.getRange(n, ENC_PRIO_VAC_.indexOf('Enviada') + 1, filas.length, 1).setNumberFormat('dd/MM/yyyy HH:mm').setValues(filas.map(() => [ahora]));
      registrar_(u, 'Prioridades de vacaciones enviadas', lista.length + ' bloques · ' + lista.reduce((s, b) => s + b.dias, 0) + ' días', disp);
      return { campana: datosCampana_(false, u) };
    });
  });
}

/** Retira las prioridades enviadas (solo mientras la campaña siga abierta y antes de la asignación). */
function apiRetirarPrioridadesVac(token) {
  return responder_(() => {
    const u = sesion_(token).u;
    if (!campAbierta_(cfgCamp_(), hoyISO_())) throw falla_('La solicitud de vacaciones ya cerró.');
    return conBloqueo_(() => {
      const mias = leerPrioridadesVac_().filter(p => p.bp === u.bp);
      if (mias.some(p => p.estado !== 'enviada')) throw falla_('Coordinación ya asignó tus vacaciones.');
      const sh = hojaPrioVac_();
      mias.forEach(p => escribirFila_(sh, p.fila, { 'Estado': 'retirada' }));
      registrar_(u, 'Prioridades de vacaciones retiradas', mias.length + ' bloques', '');
      return { campana: datosCampana_(false, u) };
    });
  });
}

/* =================== Funciones públicas: coordinación =================== */

function apiGuardarCfgCamp(token, c) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'vac', nivel: 2 }).u;
    c = c || {};
    const f = (s, def) => { s = String(s || '').trim(); if (!s) return def; if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw falla_('Fecha inválida: ' + s); return s; };
    const tope = c.tope === '' || c.tope == null ? 0 : Math.round(Number(c.tope));
    if (!(tope >= 0 && tope <= 400)) throw falla_('El tope debe estar entre 0 (sin tope) y 400 días.');
    const cupo = c.cupo === '' || c.cupo == null ? Math.max(1, Number(cfgVac_().cupo) || 3) : Math.round(Number(c.cupo));
    if (!(cupo >= 1 && cupo <= 20)) throw falla_('El cupo por día debe estar entre 1 y 20 instructores.');
    const cfg = { habilitado: !!c.habilitado, desde: f(c.desde, ''), hasta: f(c.hasta, ''), inicio: f(c.inicio, CFG_CAMP_DEF_.inicio), fin: f(c.fin, CFG_CAMP_DEF_.fin), tope: tope, cupo: cupo,
      // Solo para el texto que ve el instructor
      plazoEnero: f(c.plazoEnero, CFG_CAMP_DEF_.plazoEnero), alta: String(c.alta || '').trim().slice(0, 120) || CFG_CAMP_DEF_.alta };
    if (cfg.desde && cfg.hasta && cfg.desde > cfg.hasta) throw falla_('"Abierta desde" es posterior a "Abierta hasta".');
    if (cfg.inicio > cfg.fin) throw falla_('El inicio de la ventana de vacaciones es posterior al fin.');
    // Un tope por debajo del mínimo de alguien impediría que esa persona envíe sus prioridades
    if (tope > 0) {
      const mins = leerMinimosVac_(), mayor = Object.keys(mins).reduce((m, bp) => Math.max(m, mins[bp].minimo || 0), 0);
      if (tope < mayor) throw falla_('El máximo por instructor (' + tope + ') es menor que el mínimo más alto (' + mayor + ' días). Usa 0 (sin tope) o un número de ' + mayor + ' o más. El cupo diario va en «Cupo por día».');
    }
    props_().setProperty('CFG_CAMP', JSON.stringify(cfg));
    registrar_(yo, 'Solicitud de vacaciones', (cfg.habilitado ? 'Abierta' : 'Cerrada') + ' · ' + cfg.inicio + ' a ' + cfg.fin + ' · cupo ' + cfg.cupo, '');
    return datosCampana_(true, yo);
  });
}

/** Orden de escalafón (igual que la pestaña Vacaciones): año y trimestre cargados o, si no, los de la Fecha DGAC original. */
function ordenEscalafonVac_(bps, v) {
  const ov = escalafonVac_();
  const filas = bps.map(bp => {
    const p = v.personas[bp] || {}, o = ov[bp] || {}, f = p.dgacOrig ? parseISO_(p.dgacOrig) : null;
    return { bp: bp, anio: o.anio || (f ? f.getFullYear() : null), trim: o.trim || (f ? Math.floor(f.getMonth() / 3) + 1 : null), f: f ? f.getTime() : Infinity };
  });
  filas.sort((a, b) => (!a.anio !== !b.anio) ? (a.anio ? -1 : 1) : (a.anio - b.anio) || (a.trim - b.trim) || (a.f - b.f));
  return filas.map(r => r.bp);
}

/**
 * R6: asignación por rondas sobre las prioridades enviadas que aún no se asignaron.
 * Cuenta como ocupación de cada día las vacaciones ya registradas de todo el equipo y lo que se va asignando.
 */
function planAsignacionVac_() {
  const v = leerVacaciones_(), sols = leerSolicitudesVac_(), cupo = Math.max(1, Number(cfgCamp_().cupo) || 3);
  const prios = leerPrioridadesVac_().filter(p => p.estado === 'enviada');
  const oc = {};
  Object.keys(v.personas).forEach(bp => ocupadosVac_(bp, v, sols).forEach(k => { oc[k] = (oc[k] || 0) + 1; }));
  const orden = ordenEscalafonVac_(Array.from(new Set(prios.map(p => p.bp))), v);
  const plan = [];
  for (let ronda = 1; ronda <= MAX_PRIO_VAC_; ronda++) {
    orden.forEach((bp, lugar) => {
      const p = prios.find(x => x.bp === bp && x.prioridad === ronda);
      if (!p) return;
      const dias = diasDeBloque_(p.desde, p.hasta);
      const lleno = dias.find(k => (oc[k] || 0) >= cupo);
      if (!lleno) dias.forEach(k => { oc[k] = (oc[k] || 0) + 1; });
      plan.push({ id: p.id, fila: p.fila, bp: bp, nombre: p.nombre, prioridad: ronda, escalafon: lugar + 1, desde: p.desde, hasta: p.hasta, dias: p.dias,
        asignada: !lleno, motivo: lleno ? 'El ' + lleno.split('-').reverse().join('/') + ' ya tiene ' + cupo + ' de vacaciones (cupo)' : '' });
    });
  }
  return { cupo: cupo, plan: plan };
}

function apiSimularAsignacionVac(token) {
  return responder_(() => {
    sesion_(token, { seccion: 'vac', nivel: 1 });
    const r = planAsignacionVac_(); r.plan.forEach(x => { delete x.fila; });
    return r;
  });
}

/** Aplica la asignación: marca cada prioridad y escribe los bloques asignados en VacacionesBloques (quedan programados). */
function apiAplicarAsignacionVac(token) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'vac', nivel: 2 }).u;
    if (campAbierta_(cfgCamp_(), hoyISO_())) throw falla_('Cierra la solicitud antes de asignar: los instructores todavía pueden cambiar sus prioridades.');
    return conBloqueo_(() => {
      const r = planAsignacionVac_();
      if (!r.plan.length) throw falla_('No hay prioridades pendientes de asignar.');
      const sh = hojaPrioVac_(), ahora = new Date();
      r.plan.forEach(x => escribirFila_(sh, x.fila, { 'Estado': x.asignada ? 'asignada' : 'no asignada', 'Ronda': x.prioridad,
        'Resultado': x.asignada ? 'Asignada en la ronda ' + x.prioridad + ' (escalafón #' + x.escalafon + ')' : x.motivo, 'Asignada por': nombreDe_(yo), 'Asignada': ahora }));
      const asignadas = r.plan.filter(x => x.asignada);
      if (asignadas.length) {
        const sb = hoja_(CONFIG.hojas.vacBloques, ENC_VAC_BLOQ_);
        const f12 = iso => { const d = parseISO_(iso); d.setHours(12); return d; };
        sb.getRange(sb.getLastRow() + 1, 1, asignadas.length, 5)
          .setValues(asignadas.map(x => [x.bp, f12(x.desde), f12(x.hasta), x.dias, ahora]));
      }
      avisarAsignacionVac_(r.plan);
      registrar_(yo, 'Solicitud de vacaciones asignada', asignadas.length + ' de ' + r.plan.length + ' bloques asignados', '');
      return datosCampana_(true, yo);
    });
  });
}

/** Un correo por instructor con el resultado de cada prioridad. Un error de correo no detiene la asignación. */
function avisarAsignacionVac_(plan) {
  const usu = usuarios_(), porBp = {};
  plan.forEach(x => { (porBp[x.bp] = porBp[x.bp] || []).push(x); });
  const f = iso => iso.split('-').reverse().join('/');
  Object.keys(porBp).forEach(bp => {
    const u = usu.find(v => v.bp === bp); if (!u || !u.correo) return;
    const l = porBp[bp].sort((a, b) => a.prioridad - b.prioridad);
    const td = 'padding:4px 8px;border:1px solid #ddd';
    try {
      enviarCorreo_({ to: u.correo, name: CONFIG.nombreApp, subject: CONFIG.nombreApp + ' · Resultado de tus prioridades de vacaciones',
        htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0F004F"><p>Hola ' + escHtml_(nombreDe_(u)) + ', coordinación asignó tu solicitud de vacaciones:</p>' +
          '<table style="border-collapse:collapse;font-size:13px"><tr><th style="' + td + '">Prioridad</th><th style="' + td + '">Fechas</th><th style="' + td + '">Días</th><th style="' + td + '">Resultado</th></tr>' +
          l.map(x => '<tr><td style="' + td + '">P' + x.prioridad + '</td><td style="' + td + '">' + f(x.desde) + ' al ' + f(x.hasta) + '</td><td style="' + td + '">' + x.dias + '</td><td style="' + td + '">' +
            (x.asignada ? '<b style="color:#1a7f4b">Asignada</b>' : 'No asignada: ' + escHtml_(x.motivo)) + '</td></tr>').join('') + '</table>' +
          '<p>Las asignadas ya figuran como vacaciones programadas en la app.</p>' + botonAppHtml_() + '</div>',
        body: 'Resultado de tus prioridades de vacaciones:\n' + l.map(x => 'P' + x.prioridad + ' ' + f(x.desde) + ' al ' + f(x.hasta) + ' (' + x.dias + ' días): ' + (x.asignada ? 'asignada' : 'no asignada, ' + x.motivo)).join('\n') });
    } catch (e) { console.warn('Correo de asignación no enviado a ' + bp + ': ' + e.message); }
  });
}
