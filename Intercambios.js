/**
 * Intercambio de actividades entre instructores.
 *
 * Flujo: el solicitante elige un día (o un bloque de viaje) y un compañero → el receptor acepta o rechaza →
 * si acepta, coordinación (supervisor) aprueba o rechaza → al aprobar, se intercambian las celdas en el
 * Google Sheet del mes (texto, colores y formato) y se vuelve a sincronizar.
 *
 * Hoja Intercambios: una fila por solicitud. Los meses ocultos a instructores no admiten intercambios.
 */

const ENC_INT_ = ['ID', 'Mes', 'Días', 'Solicitante BP', 'Solicitante', 'Receptor BP', 'Receptor', 'Actividades solicitante', 'Actividades receptor',
  'Estado', 'Comentario solicitante', 'Comentario receptor', 'Comentario supervisor', 'Creado', 'Respondido receptor', 'Decidido', 'Supervisor',
  'Aplicación', 'Visto solicitante', 'Visto receptor', 'Revisado coordinación', 'Revisor', 'Comentario revisor', 'Hilo correo'];
const EST_INT_ = {
  pendiente: 'Pendiente', en_supervisor: 'En revisión de coordinación', en_jefe: 'En aprobación final',
  rechazado_receptor: 'Rechazado por receptor', rechazado_jefe: 'Rechazado en aprobación final',
  aprobado: 'Aprobado', rechazado_supervisor: 'Rechazado por supervisor', cancelado: 'Cancelado'
};
const MAX_DIAS_INT_ = 7;

/* ---------------- Meses ocultos a instructores ---------------- */

function mesesOcultos_() { try { return JSON.parse(props_().getProperty('MESES_OCULTOS') || '[]'); } catch (e) { return []; } }
const mesVisible_ = mes => mesesOcultos_().indexOf(mes) < 0;

function apiVisibilidadMes(token, mes, visible) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'cargar', nivel: 2 }).u;
    mes = mesTxt_(mes); if (!mes) throw falla_('Mes inválido.');
    return conBloqueo_(() => {
      const set = mesesOcultos_().filter(m => m !== mes);
      if (!visible) set.push(mes);
      props_().setProperty('MESES_OCULTOS', JSON.stringify(set.sort()));
      registrar_(yo, visible ? 'Mes visible para instructores' : 'Mes oculto a instructores', mes, '');
      return set;
    });
  });
}

/* ---------------- Lectura ---------------- */

function hojaInt_() { return asegurarColumnas_(hoja_(CONFIG.hojas.intercambios, ENC_INT_), ENC_INT_); }

function estadoIntDeTexto_(t) {
  const n = normTxt_(t);
  if (/^en espera|^en revision/.test(n)) return 'en_supervisor';
  if (/^en aprobacion final/.test(n)) return 'en_jefe';
  if (/^rechazado en aprobacion final/.test(n)) return 'rechazado_jefe';
  if (/^rechazado por receptor/.test(n)) return 'rechazado_receptor';
  if (/^rechazado por supervisor/.test(n)) return 'rechazado_supervisor';
  if (/^aprobado/.test(n)) return 'aprobado';
  if (/^cancelado/.test(n)) return 'cancelado';
  return 'pendiente';
}
function jsonArr_(v) { try { const a = JSON.parse(v || '[]'); return Array.isArray(a) ? a.map(String) : []; } catch (e) { return []; } }

function leerIntercambios_() {
  const sh = ss_().getSheetByName(CONFIG.hojas.intercambios);
  if (!sh) return [];
  const t = filas_(sh);
  const c = {}; ENC_INT_.forEach(h => { c[normTxt_(h)] = col_(t.head, h); });
  const g = (r, h) => celda_(r, c[normTxt_(h)]);
  const hoy = hoyISO_();
  return t.rows.map((r, i) => {
    const mes = mesTxt_(g(r, 'Mes'));
    const dias = String(g(r, 'Días') || '').split(/[\s,;]+/).map(Number).filter(n => n >= 1 && n <= 31);
    let estado = estadoIntDeTexto_(g(r, 'Estado'));
    const primero = mes && dias.length ? mes + '-' + ('0' + Math.min.apply(null, dias)).slice(-2) : '';
    const vencido = (estado === 'pendiente' || estado === 'en_supervisor' || estado === 'en_jefe') && primero && primero <= hoy;
    return {
      fila: i + 2, id: String(g(r, 'ID') || ''), mes: mes, dias: dias,
      solBp: bpTxt_(g(r, 'Solicitante BP')), solNombre: String(g(r, 'Solicitante') || ''),
      recBp: bpTxt_(g(r, 'Receptor BP')), recNombre: String(g(r, 'Receptor') || ''),
      actSol: jsonArr_(g(r, 'Actividades solicitante')), actRec: jsonArr_(g(r, 'Actividades receptor')),
      estado: vencido ? 'vencido' : estado, estadoGuardado: estado,
      comSol: String(g(r, 'Comentario solicitante') || ''), comRec: String(g(r, 'Comentario receptor') || ''), comSup: String(g(r, 'Comentario supervisor') || ''),
      creado: fechaCelda_(g(r, 'Creado')), respondido: fechaCelda_(g(r, 'Respondido receptor')), decidido: fechaCelda_(g(r, 'Decidido')),
      supervisor: String(g(r, 'Supervisor') || ''), aplicacion: String(g(r, 'Aplicación') || ''),
      hilo: String(g(r, 'Hilo correo') || ''),
      revisado: fechaCelda_(g(r, 'Revisado coordinación')), revisor: String(g(r, 'Revisor') || ''), comRev: String(g(r, 'Comentario revisor') || ''),
      vistoSol: normTxt_(g(r, 'Visto solicitante')), vistoRec: normTxt_(g(r, 'Visto receptor'))
    };
  }).filter(x => x.id && x.mes);
}

/** Lo que ve cada uno: coordinación ve todo; un instructor, solo donde participa. */
function intercambiosPara_(u) {
  return leerIntercambios_().filter(x => u.rol === 'admin' || x.solBp === u.bp || x.recBp === u.bp)
    .map(x => { delete x.fila; return x; });
}

/* ---------------- Validaciones comunes ---------------- */

function rolDelMes_(mes) {
  const m = leerRoles_().find(x => x.mes === mes);
  if (!m) throw falla_('Ese mes no tiene rol publicado.');
  return m;
}
function exigirDiaFuturo_(mes, dia) {
  const f = mes + '-' + ('0' + dia).slice(-2);
  if (f <= hoyISO_()) throw falla_('Solo se pueden intercambiar días a partir de mañana.');
}
const activoInt_ = x => x.estado === 'pendiente' || x.estado === 'en_supervisor' || x.estado === 'en_jefe';
/* Ya aceptado por ambos instructores (compromete esos días hasta que se decida) */
const comprometidoInt_ = x => x.estado === 'en_supervisor' || x.estado === 'en_jefe';

/* ---------------- Funciones públicas: instructores ---------------- */

/** Compañeros del rol y sus actividades alrededor del día elegido (para elegir con quién y detectar bloques de viaje). */
function apiOpcionesIntercambio(token, mes, dia) {
  return responder_(() => {
    const u = sesion_(token).u;
    mes = mesTxt_(mes); dia = Number(dia);
    if (u.rol !== 'inst') throw falla_('Solo los instructores pueden proponer intercambios.');
    if (!mesVisible_(mes)) throw falla_('Ese mes todavía no está publicado para instructores.');
    exigirDiaFuturo_(mes, dia);
    const m = rolDelMes_(mes);
    if (!m.instructores.some(i => i.bp === u.bp)) throw falla_('No figuras en el rol de ese mes.');
    const nd = diasDelMes_(mes), desde = Math.max(1, dia - 3), hasta = Math.min(nd, dia + 3);
    const ocupados = {};
    leerIntercambios_().filter(x => x.mes === mes && comprometidoInt_(x)).forEach(x => x.dias.forEach(d => { (ocupados[x.solBp] = ocupados[x.solBp] || []).push(d); (ocupados[x.recBp] = ocupados[x.recBp] || []).push(d); }));
    return {
      mes: mes, dia: dia, desde: desde, hasta: hasta, yo: u.bp, hoy: hoyISO_(),
      instructores: m.instructores.filter(i => i.bp).map(i => {
        const celdas = {}; for (let d = desde; d <= hasta; d++) celdas[d] = i.dias[d - 1] || '';
        return { bp: i.bp, nombre: i.nombre, celdas: celdas, ocupados: ocupados[i.bp] || [] };
      })
    };
  });
}

function apiCrearIntercambio(token, mes, dias, receptorBp, comentario, disp) {
  return responder_(() => {
    const u = sesion_(token).u;
    mes = mesTxt_(mes); receptorBp = bpTxt_(receptorBp);
    if (u.rol !== 'inst') throw falla_('Solo los instructores pueden proponer intercambios.');
    if (!mesVisible_(mes)) throw falla_('Ese mes todavía no está publicado para instructores.');
    const com = comentarioVac_(comentario);
    const lista = Array.from(new Set((dias || []).map(Number))).sort((a, b) => a - b);
    const nd = diasDelMes_(mes);
    if (!lista.length || lista.length > MAX_DIAS_INT_) throw falla_('Elige entre 1 y ' + MAX_DIAS_INT_ + ' días.');
    if (lista.some(d => !(d >= 1 && d <= nd))) throw falla_('Día inválido.');
    if (lista[lista.length - 1] - lista[0] !== lista.length - 1) throw falla_('Los días del intercambio deben ser seguidos.');
    exigirDiaFuturo_(mes, lista[0]);
    if (receptorBp === u.bp) throw falla_('Elige a otro instructor.');
    return conBloqueo_(() => {
      const m = rolDelMes_(mes);
      const a = m.instructores.find(i => i.bp === u.bp), b = m.instructores.find(i => i.bp === receptorBp);
      if (!a) throw falla_('No figuras en el rol de ese mes.');
      if (!b) throw falla_('Ese instructor no figura en el rol de ese mes.');
      const actA = lista.map(d => a.dias[d - 1] || ''), actB = lista.map(d => b.dias[d - 1] || '');
      if (actA.every((x, k) => normCelda_(x) === normCelda_(actB[k]))) throw falla_('Ambos tienen la misma actividad en esos días.');
      const cruza = x => x.mes === mes && x.dias.some(d => lista.indexOf(d) >= 0);
      const todos = leerIntercambios_();
      // Una sola propuesta enviada en curso por día
      if (todos.some(x => cruza(x) && x.estado === 'pendiente' && x.solBp === u.bp))
        throw falla_('Ya enviaste otra propuesta que incluye esos días. Cancélala antes de proponer una nueva.');
      // Lo que bloquea es un intercambio ya aceptado por ambos (en espera de coordinación)
      const comprometido = todos.find(x => cruza(x) && comprometidoInt_(x) && [x.solBp, x.recBp].some(p => p === u.bp || p === receptorBp));
      if (comprometido) throw falla_((comprometido.solBp === u.bp || comprometido.recBp === u.bp ? 'Ya tienes' : b.nombre + ' ya tiene') + ' un intercambio aceptado esos días, esperando a coordinación.');
      const receptor = usuarios_().find(x => x.bp === receptorBp);
      const sh = hojaInt_();
      const id = 'i' + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36);
      const n = sh.getLastRow() + 1;
      sh.getRange(n, 1, 1, 7).setNumberFormat('@');
      escribirFila_(sh, n, {
        'ID': id, 'Mes': mes, 'Días': lista.join(', '), 'Solicitante BP': u.bp, 'Solicitante': nombreDe_(u),
        'Receptor BP': receptorBp, 'Receptor': receptor ? nombreDe_(receptor) : b.nombre,
        'Actividades solicitante': JSON.stringify(actA), 'Actividades receptor': JSON.stringify(actB),
        'Estado': EST_INT_.pendiente, 'Comentario solicitante': com, 'Creado': new Date(), 'Visto solicitante': 'pendiente'
      });
      registrar_(u, 'Intercambio solicitado', 'Con ' + b.nombre + ' · ' + rangoDiasTxt_(mes, lista), disp);
      notificarInt_(id, nombreDe_(u) + ' te propone intercambiar ' + rangoDiasTxt_(mes, lista) + '.', [receptor, u],
        b.nombre + ': acéptalo o recházalo en "Mi rol" de la app.');
      return { intercambios: intercambiosPara_(u) };
    });
  });
}

/** El receptor acepta o rechaza. */
function apiResponderIntercambio(token, id, acepta, comentario, disp) {
  return responder_(() => {
    const u = sesion_(token).u;
    const com = comentarioVac_(comentario);
    return conBloqueo_(() => {
      const x = leerIntercambios_().find(i => i.id === String(id));
      if (!x || x.recBp !== u.bp) throw falla_('No encontramos ese intercambio.');
      if (x.estado === 'vencido') throw falla_('Este intercambio venció: la fecha ya pasó.');
      if (x.estado !== 'pendiente') throw yaResuelto_('Este intercambio', x.estado === 'cancelado' ? 'cancelado por ' + x.solNombre : (EST_INT_[x.estado] || x.estado).toLowerCase(), '', x.decidido || x.respondido);
      if (acepta) {
        const otro = leerIntercambios_().find(z => z.id !== x.id && z.mes === x.mes && comprometidoInt_(z) && z.dias.some(d => x.dias.indexOf(d) >= 0) &&
          [z.solBp, z.recBp].some(p => p === x.solBp || p === x.recBp));
        if (otro) throw falla_((otro.solBp === u.bp || otro.recBp === u.bp ? 'Ya tienes' : x.solNombre + ' ya tiene') + ' otro intercambio aceptado esos días, esperando a coordinación. Rechaza este o espera esa decisión.');
      }
      const est = acepta ? 'en_supervisor' : 'rechazado_receptor';
      escribirFila_(hojaInt_(), x.fila, { 'Estado': EST_INT_[est], 'Comentario receptor': com, 'Respondido receptor': new Date(), 'Visto receptor': est });
      registrar_(u, acepta ? 'Intercambio aceptado' : 'Intercambio rechazado', 'Con ' + x.solNombre + ' · ' + rangoDiasTxt_(x.mes, x.dias), disp);
      if (acepta) notificarInt_(x.id, x.recNombre + ' aceptó el intercambio. Pasa a revisión de coordinación.', aprobadoresInt_().revision,
        'Coordinación: revísalo en la pestaña "Intercambios de rol" de la app.');
      else notificarInt_(x.id, x.recNombre + ' rechazó el intercambio. La propuesta queda cerrada.', [], '');
      return { intercambios: intercambiosPara_(u) };
    });
  });
}

/** El solicitante retira su propuesta mientras coordinación no la haya decidido. */
function apiCancelarIntercambio(token, id, disp) {
  return responder_(() => {
    const u = sesion_(token).u;
    return conBloqueo_(() => {
      const x = leerIntercambios_().find(i => i.id === String(id));
      if (!x || x.solBp !== u.bp) throw falla_('No encontramos ese intercambio.');
      if (!activoInt_(x)) throw yaResuelto_('Este intercambio', (EST_INT_[x.estado] || x.estado).toLowerCase(), x.supervisor, x.decidido || x.respondido);
      escribirFila_(hojaInt_(), x.fila, { 'Estado': EST_INT_.cancelado, 'Decidido': new Date(), 'Visto solicitante': 'cancelado' });
      registrar_(u, 'Intercambio cancelado', 'Con ' + x.recNombre + ' · ' + rangoDiasTxt_(x.mes, x.dias), disp);
      notificarInt_(x.id, x.solNombre + ' canceló la propuesta de intercambio.', [], '');
      return { intercambios: intercambiosPara_(u) };
    });
  });
}

/** Marca como vistos los cambios de estado (apaga la notificación). */
function apiVistoIntercambios(token, ids) {
  return responder_(() => {
    const u = sesion_(token).u;
    return conBloqueo_(() => {
      const sh = hojaInt_();
      leerIntercambios_().filter(x => (ids || []).indexOf(x.id) >= 0).forEach(x => {
        if (x.solBp === u.bp && x.vistoSol !== x.estadoGuardado) escribirFila_(sh, x.fila, { 'Visto solicitante': x.estadoGuardado });
        if (x.recBp === u.bp && x.vistoRec !== x.estadoGuardado && x.estadoGuardado !== 'pendiente') escribirFila_(sh, x.fila, { 'Visto receptor': x.estadoGuardado });
      });
      return true;
    });
  });
}

/* ---------------- Funciones públicas: coordinación ---------------- */

/**
 * Coordinación decide en dos pasos:
 *  1) Revisión de coordinación (estado en_supervisor): revisores configurados (Antonella o su suplente).
 *     Aprueba → pasa a aprobación final (o se aplica si no hay aprobación final configurada). Rechaza → termina.
 *  2) Aprobación final (estado en_jefe): aprobadores finales (Jorge). Aprueba → se aplica en el Google Sheet.
 * El aprobador final recibe correo solo cuando coordinación resuelve; antes, solo la notificación en la app.
 */
function apiDecidirIntercambio(token, id, aprueba, comentario, pasoEsperado) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'int', nivel: 1 }).u;
    const com = comentarioVac_(comentario);
    return conBloqueo_(() => {
      const x = leerIntercambios_().find(i => i.id === String(id));
      if (!x) throw falla_('No encontramos ese intercambio.');
      if (x.estado === 'vencido') throw falla_('Este intercambio venció: la fecha ya pasó.');
      if (['aprobado', 'rechazado_supervisor', 'rechazado_jefe'].indexOf(x.estado) >= 0) throw yaResuelto_('Este intercambio', (EST_INT_[x.estado] || x.estado).toLowerCase(), x.supervisor, x.decidido);
      if (!comprometidoInt_(x)) throw yaResuelto_('Este intercambio', (EST_INT_[x.estado] || x.estado).toLowerCase(), '', x.decidido || x.respondido);
      const ap = aprobadoresInt_(), paso = x.estado === 'en_supervisor' ? 'revision' : 'final';
      if (pasoEsperado && pasoEsperado !== paso) throw yaResuelto_('Este intercambio', 'revisado en coordinación', x.revisor, x.revisado);
      if (!puedeDecidirInt_(yo, paso, ap)) throw falla_(paso === 'revision'
        ? 'La revisión de coordinación la hacen: ' + (ap.revision.map(nombreDe_).join(', ') || 'quien gestione Intercambios') + '.'
        : 'La aprobación final la hace: ' + ap.final.map(nombreDe_).join(', ') + '.', 'PERMISO');
      const sh = hojaInt_();
      const rango = rangoDiasTxt_(x.mes, x.dias);
      // Paso 1 aprobado y hay aprobación final: pasa a Jorge (no toca el Sheet todavía)
      if (paso === 'revision' && aprueba && ap.final.length) {
        escribirFila_(sh, x.fila, { 'Estado': EST_INT_.en_jefe, 'Revisado coordinación': new Date(), 'Revisor': nombreDe_(yo), 'Comentario revisor': com });
        registrar_(yo, 'Intercambio aprobado por coordinación', x.solNombre + ' ↔ ' + x.recNombre + ' · ' + rango, '');
        notificarInt_(x.id, nombreDe_(yo) + ' aprobó el intercambio en coordinación. Pasa a aprobación final.', ap.final,
          ap.final.map(nombreDe_).join(', ') + ': apruébalo o recházalo en la pestaña "Intercambios de rol" de la app.');
        return { intercambios: intercambiosPara_(yo), aplicacion: '' };
      }
      // Decisión que cierra el intercambio
      let aplicacion = '';
      if (aprueba) aplicacion = aplicarIntercambio_(x, yo); // si falla, lanza el error y no se aprueba
      const est = aprueba ? 'aprobado' : (paso === 'revision' ? 'rechazado_supervisor' : 'rechazado_jefe');
      const valores = { 'Estado': EST_INT_[est], 'Decidido': new Date(), 'Supervisor': nombreDe_(yo), 'Comentario supervisor': com, 'Aplicación': aplicacion };
      if (paso === 'revision') { valores['Revisado coordinación'] = new Date(); valores['Revisor'] = nombreDe_(yo); valores['Comentario revisor'] = com; }
      escribirFila_(sh, x.fila, valores);
      const quien = paso === 'revision' ? 'coordinación' : 'aprobación final';
      registrar_(yo, 'Intercambio ' + (aprueba ? 'aprobado' : 'rechazado') + ' en ' + quien, x.solNombre + ' ↔ ' + x.recNombre + ' · ' + rango, '');
      // Si coordinación rechaza, la aprobación final se suma al hilo solo para quedar informada
      notificarInt_(x.id, nombreDe_(yo) + ' (' + quien + ') ' + (aprueba ? 'aprobó el intercambio. El rol del mes ya muestra el cambio.' : 'rechazó el intercambio. La propuesta queda cerrada.'),
        paso === 'revision' && !aprueba ? ap.final : [], '');
      return { intercambios: intercambiosPara_(yo), aplicacion: aplicacion };
    });
  });
}

/* ---------------- Quién aprueba (configurable desde la app) ---------------- */

/** {revision:[usuarios], final:[usuarios]}: administradores activos guardados en INT_APROB. */
function aprobadoresInt_() {
  let c = {}; try { c = JSON.parse(props_().getProperty('INT_APROB') || '{}'); } catch (e) { c = {}; }
  const lista = bps => (bps || []).map(b => usuarios_().find(u => u.bp === bpTxt_(b))).filter(u => u && u.rol === 'admin' && u.activo);
  return { revision: lista(c.revision), final: lista(c.final) };
}
/** Sin revisores configurados, revisa quien gestione Intercambios. La aprobación final solo la hacen los designados. */
function puedeDecidirInt_(u, paso, ap) {
  ap = ap || aprobadoresInt_();
  if (paso === 'revision') return ap.revision.length ? ap.revision.some(v => v.bp === u.bp) : nivel_(u, 'int') >= 2;
  return ap.final.some(v => v.bp === u.bp);
}
function infoAprobadoresInt_(u) {
  const ap = aprobadoresInt_(), p = v => ({ bp: v.bp, nombre: nombreDe_(v) });
  return { revision: ap.revision.map(p), final: ap.final.map(p), soyRevisor: puedeDecidirInt_(u, 'revision', ap), soyFinal: puedeDecidirInt_(u, 'final', ap),
    puedoConfigurar: !!u.principal || nivel_(u, 'int') >= 2,
    admins: usuarios_().filter(v => v.rol === 'admin' && v.activo).map(p) };
}

function apiAprobadoresIntercambio(token, cfg) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'int', nivel: 2 }).u;
    const validos = {}; usuarios_().filter(v => v.rol === 'admin' && v.activo).forEach(v => { validos[v.bp] = true; });
    const limpia = l => Array.from(new Set((l || []).map(bpTxt_).filter(b => validos[b])));
    const c = { revision: limpia(cfg && cfg.revision), final: limpia(cfg && cfg.final) };
    if (!c.final.length && !c.revision.length) throw falla_('Designa al menos a una persona.');
    props_().setProperty('INT_APROB', JSON.stringify(c));
    registrar_(yo, 'Aprobadores de intercambios', 'Revisión: ' + c.revision.length + ' · Final: ' + c.final.length, '');
    return infoAprobadoresInt_(yo);
  });
}

/* ---------------- Aplicar en el Google Sheet del mes ---------------- */

const normCelda_ = s => String(s == null ? '' : s).replace(/\r/g, '').replace(/\s+/g, ' ').trim();

/**
 * Intercambia las celdas de ambos instructores en el Google Sheet del mes (texto, colores y formato),
 * después de verificar que siguen siendo las mismas actividades que se pidieron. Devuelve un texto para la hoja.
 */
function aplicarIntercambio_(x, yo) {
  const f = leerFuentes_().find(z => z.mes === x.mes && z.activo && z.id);
  const n = x.dias.length, d0 = x.dias[0];
  if (!f) return aplicarEnCopia_(x);
  let ss;
  try { ss = SpreadsheetApp.openById(f.id); } catch (e) { throw falla_('No se pudo abrir el Google Sheet de ' + x.mes + '. La cuenta dueña del script necesita permiso de edición.'); }
  const sh = elegirHoja_(ss, f.hoja, f.gid, f.mes);
  const rango = sh.getDataRange(), vals = rango.getValues(), disp = rango.getDisplayValues();
  const ub = ubicarMatriz_(vals);
  const filaDe = bp => { for (let r = ub.h + 1; r < vals.length; r++) if (bpTxt_(vals[r][0]) === bp) return r; return -1; };
  const ra = filaDe(x.solBp), rb = filaDe(x.recBp);
  if (ra < 0 || rb < 0) throw falla_('No encontramos a ' + (ra < 0 ? x.solNombre : x.recNombre) + ' (por su BP) en el Google Sheet del mes.');
  x.dias.forEach((d, k) => {
    const c = ub.c0 + d - 1;
    if (normCelda_(disp[ra][c]) !== normCelda_(x.actSol[k]) || normCelda_(disp[rb][c]) !== normCelda_(x.actRec[k]))
      throw falla_('El rol del día ' + d + ' cambió en el Google Sheet desde que se pidió el intercambio. Revísalo con ambos instructores antes de aprobar.');
  });
  // getRange usa columnas desde 1: el día d0 está en la columna ub.c0 + d0
  const A = sh.getRange(ra + 1, ub.c0 + d0, 1, n), B = sh.getRange(rb + 1, ub.c0 + d0, 1, n);
  const tmp = ss.insertSheet('_intercambio_' + x.id);
  try {
    A.copyTo(tmp.getRange(1, 1));
    B.copyTo(A);
    tmp.getRange(1, 1, 1, n).copyTo(B);
  } finally { ss.deleteSheet(tmp); }
  // copyTo ya movió las notas con su celda; se agrega una línea con el intercambio
  const anotar = (rg, otro) => rg.setNotes([rg.getNotes()[0].map(p => (p ? p + '\n' : '') + 'Intercambio con ' + otro + ' aprobado por ' + nombreDe_(yo) + ' el ' + hoyISO_() + '.')]);
  anotar(A, x.recNombre); anotar(B, x.solNombre);
  // Si hay catálogo de colores, cada celda queda con el color de su nuevo tipo de actividad
  let pintadas = 0;
  try { pintadas = pintarSegunCatalogo_(A) + pintarSegunCatalogo_(B); } catch (e) { console.warn('No se pudo pintar con el catálogo: ' + e.message); }
  SpreadsheetApp.flush();
  const s = sincronizarFuente_(f);
  return 'Aplicado en el Google Sheet' + (pintadas ? ', pintado según el catálogo' : '') + (s.ok ? ' y sincronizado' : ' (sincroniza el mes: ' + s.error + ')') + ' · ' + hoyISO_();
}

/** Mes cargado desde Excel (sin Google Sheet activo): se intercambia en la copia de la hoja Roles. */
function aplicarEnCopia_(x) {
  const m = rolDelMes_(x.mes);
  const a = m.instructores.find(i => i.bp === x.solBp), b = m.instructores.find(i => i.bp === x.recBp);
  if (!a || !b) throw falla_('Uno de los instructores ya no figura en el rol del mes.');
  x.dias.forEach((d, k) => {
    if (normCelda_(a.dias[d - 1]) !== normCelda_(x.actSol[k]) || normCelda_(b.dias[d - 1]) !== normCelda_(x.actRec[k]))
      throw falla_('El rol del día ' + d + ' cambió desde que se pidió el intercambio. Revísalo antes de aprobar.');
  });
  const cambia = (arrA, arrB, d) => { if (arrA && arrB) { const t = arrA[d - 1]; arrA[d - 1] = arrB[d - 1]; arrB[d - 1] = t; } };
  x.dias.forEach(d => { cambia(a.dias, b.dias, d); cambia(a.bg, b.bg, d); cambia(a.fg, b.fg, d); });
  guardarRol_({ mes: m.mes, fuente: m.fuente, hoja: m.hoja, instructores: m.instructores }, m.origen || 'Excel');
  return 'Aplicado en la copia de la app (el mes se cargó desde Excel; actualiza también el archivo original) · ' + hoyISO_();
}

/** Fila de encabezado (BP | Nombre) y columna del día 1, igual que al leer el rol. */
function ubicarMatriz_(vals) {
  let h = -1;
  for (let r = 0; r < Math.min(vals.length, 20); r++) {
    if (normTxt_(vals[r][1]).indexOf('nombre') >= 0 || normTxt_(vals[r][0]) === 'bp') { h = r; break; }
  }
  if (h < 0) throw falla_('No encontramos la fila de encabezado (BP y Nombre) en el Google Sheet del mes.');
  let c0 = 2;
  for (let c = 2; c < Math.min(vals[h].length, 8); c++) if (esFecha_(vals[h][c])) { c0 = c; break; }
  return { h: h, c0: c0 }; // índice (desde 0) de la columna del día 1: el día d está en c0 + d - 1
}

/* ---------------- Correos ---------------- */

function rangoDiasTxt_(mes, dias) {
  const nm = MESES_ES_[Number(mes.slice(5, 7)) - 1];
  return dias.length > 1 ? 'del ' + dias[0] + ' al ' + dias[dias.length - 1] + ' de ' + nm : 'el ' + dias[0] + ' de ' + nm;
}
function escHtml_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

function detalleCorreo_(nomA, actA, nomB, actB, mes, dias, comentarios) {
  const filas = dias.map((d, k) => '<tr><td style="padding:4px 8px;border:1px solid #ddd">' + d + '</td><td style="padding:4px 8px;border:1px solid #ddd">' +
    escHtml_(actA[k] || '—').replace(/\n/g, '<br>') + '</td><td style="padding:4px 8px;border:1px solid #ddd">' + escHtml_(actB[k] || '—').replace(/\n/g, '<br>') + '</td></tr>').join('');
  const coms = (comentarios || []).filter(c => c[1]).map(c => '<p><b>' + escHtml_(c[0]) + ':</b> ' + escHtml_(c[1]) + '</p>').join('');
  return '<table style="border-collapse:collapse;font-size:13px"><tr><th style="padding:4px 8px;border:1px solid #ddd">Día</th><th style="padding:4px 8px;border:1px solid #ddd">' +
    escHtml_(nomA) + ' tenía</th><th style="padding:4px 8px;border:1px solid #ddd">' + escHtml_(nomB) + ' tenía</th></tr>' + filas + '</table>' + coms;
}

/* ---------------- Correo en hilo por intercambio ----------------
   Todos los correos de un intercambio van en un mismo hilo de Gmail (mismo asunto, respuestas encadenadas).
   Los destinatarios se acumulan: instructores desde el inicio, coordinación al aceptar el receptor y la
   aprobación final cuando coordinación resuelve. Cada correo trae el historial completo, para quien se suma.
   Si la API de Gmail no está disponible, se envía con MailApp (sin hilo). */

function notificarInt_(id, evento, sumar, accion) {
  try {
    const x = leerIntercambios_().find(i => i.id === String(id)); if (!x) return;
    const base = [usuarios_().find(v => v.bp === x.solBp), usuarios_().find(v => v.bp === x.recBp)];
    let h = {}; try { h = JSON.parse(x.hilo || '{}'); } catch (e) { h = {}; }
    const dest = Array.from(new Set((h.dest || []).concat(base.concat(sumar || []).filter(v => v && v.correo).map(v => v.correo.toLowerCase()))));
    if (!dest.length) return;
    const asunto = CONFIG.nombreApp + ' · Intercambio ' + x.solNombre + ' ⇄ ' + x.recNombre + ' · ' + rangoDiasTxt_(x.mes, x.dias);
    const html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0F004F">' +
      '<p style="font-size:15px"><b>' + escHtml_(evento) + '</b></p>' + (accion ? '<p>' + escHtml_(accion) + '</p>' : '') +
      botonAppHtml_() + '<h3 style="margin:18px 0 6px;font-size:14px">Actividades</h3>' + detalleCorreo_(x.solNombre, x.actSol, x.recNombre, x.actRec, x.mes, x.dias, []) +
      '<h3 style="margin:18px 0 6px;font-size:14px">Historial</h3>' + historialIntHtml_(x) +
      '<p style="color:#6C6893;font-size:12px;margin-top:18px">Las decisiones se toman en la app. Responder este correo no cambia el intercambio.</p></div>';
    const texto = evento + (accion ? '\n' + accion : '') + (urlApp_() ? '\n\n' + urlApp_() : '') + '\n\nLas decisiones se toman en la app.';
    let nuevo = null;
    try { nuevo = enviarGmailHilo_(dest, asunto, html, texto, h); }
    catch (e) {
      console.warn('Hilo de Gmail no disponible, se envía sin hilo: ' + e.message);
      MailApp.sendEmail({ to: dest.join(','), subject: asunto, htmlBody: html, body: texto, name: CONFIG.nombreApp });
    }
    escribirFila_(hojaInt_(), x.fila, { 'Hilo correo': JSON.stringify(Object.assign({}, h, nuevo || {}, { dest: dest })) });
  } catch (e) { console.warn('Correo de intercambio no enviado: ' + e.message); }
}

/** Envía con la API de Gmail, respondiendo en el hilo h ({thread, refs}) si existe. Devuelve el hilo actualizado. */
function enviarGmailHilo_(dest, asunto, html, texto, h) {
  const b64 = s => Utilities.base64Encode(s, Utilities.Charset.UTF_8);
  const enc = s => '=?UTF-8?B?' + b64(s) + '?=';
  const yo = Session.getEffectiveUser().getEmail();
  const lim = 'ins360' + Date.now().toString(36);
  const cab = ['From: ' + enc(CONFIG.nombreApp) + ' <' + yo + '>', 'To: ' + dest.join(', '), 'Subject: ' + enc(asunto), 'MIME-Version: 1.0'];
  if (h.ultimo) { cab.push('In-Reply-To: ' + h.ultimo); cab.push('References: ' + (h.refs || h.ultimo)); }
  cab.push('Content-Type: multipart/alternative; boundary="' + lim + '"');
  const parte = (tipo, cont) => '--' + lim + '\r\nContent-Type: ' + tipo + '; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n' + b64(cont).replace(/(.{76})/g, '$1\r\n') + '\r\n';
  const mime = cab.join('\r\n') + '\r\n\r\n' + parte('text/plain', texto) + parte('text/html', html) + '--' + lim + '--';
  const cuerpo = { raw: Utilities.base64EncodeWebSafe(mime, Utilities.Charset.UTF_8) };
  if (h.thread) cuerpo.threadId = h.thread;
  const m = Gmail.Users.Messages.send(cuerpo, 'me');
  let mid = '';
  try {
    const meta = Gmail.Users.Messages.get('me', m.id, { format: 'metadata', metadataHeaders: ['Message-Id'] });
    const hd = (meta.payload && meta.payload.headers || []).find(x => String(x.name).toLowerCase() === 'message-id');
    mid = hd ? hd.value : '';
  } catch (e) { console.warn('No se pudo leer el Message-Id: ' + e.message); }
  return { thread: m.threadId, ultimo: mid || h.ultimo || '', refs: [h.refs, mid].filter(Boolean).join(' ') };
}

/** Línea de tiempo del intercambio para el correo. */
function historialIntHtml_(x) {
  const f = d => { try { return d ? Utilities.formatDate(new Date(d), CONFIG.zona, "dd/MM HH:mm") : ''; } catch (e) { return ''; } };
  const filas = [[x.creado, x.solNombre, 'Propuso el intercambio', x.comSol]];
  if (x.respondido) filas.push([x.respondido, x.recNombre, x.estadoGuardado === 'rechazado_receptor' ? 'Rechazó' : 'Aceptó', x.comRec]);
  if (x.revisado) filas.push([x.revisado, x.revisor, x.estadoGuardado === 'rechazado_supervisor' ? 'Rechazó en coordinación' : 'Aprobó en coordinación', x.comRev]);
  // Paso final separado (no cuando coordinación aprobó directo, sin aprobación final configurada)
  const directo = x.revisor && x.supervisor === x.revisor && f(x.revisado) === f(x.decidido);
  if ((x.estadoGuardado === 'aprobado' || x.estadoGuardado === 'rechazado_jefe') && x.supervisor && !directo)
    filas.push([x.decidido, x.supervisor, x.estadoGuardado === 'aprobado' ? 'Aprobó (aprobación final)' : 'Rechazó en aprobación final', x.comSup]);
  if (x.estadoGuardado === 'cancelado') filas.push([x.decidido, x.solNombre, 'Canceló la propuesta', '']);
  const td = 'padding:5px 8px;border:1px solid #E2E0EF;vertical-align:top';
  return '<table style="border-collapse:collapse;font-size:13px">' + filas.map(r => '<tr><td style="' + td + ';white-space:nowrap;color:#6C6893">' + f(r[0]) + '</td><td style="' + td + '"><b>' +
    escHtml_(r[1] || '') + '</b> · ' + escHtml_(r[2]) + (r[3] ? '<br><span style="color:#3D3770">“' + escHtml_(r[3]) + '”</span>' : '') + '</td></tr>').join('') + '</table>';
}

/** Envía a quienes tengan correo en la hoja Usuarios (sin hilo). Un error de correo nunca detiene el intercambio. */
function correoIntercambio_(usuarios, asunto, texto, detalleHtml, accion) {
  const dest = Array.from(new Set((usuarios || []).filter(v => v && v.correo).map(v => v.correo)));
  if (!dest.length) return;
  const url = urlApp_();
  try {
    MailApp.sendEmail({
      to: dest.join(','), subject: CONFIG.nombreApp + ': ' + asunto, name: CONFIG.nombreApp,
      body: texto + (accion ? '\n\n' + accion : '') + (url ? '\n\n' + url : ''),
      htmlBody: '<p>' + escHtml_(texto) + '</p>' + (detalleHtml || '') + (accion ? '<p>' + escHtml_(accion) + '</p>' : '') + (url ? '<p><a href="' + url + '">Abrir la aplicación</a></p>' : '')
    });
  } catch (e) { console.warn('Correo de intercambio no enviado: ' + e.message); }
}
