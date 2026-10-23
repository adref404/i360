/**
 * Datos del panel. Lo que se envía al navegador depende del rol:
 * - Instructor: solo su fila de cada mes, sus respuestas y sus vigencias.
 * - Coordinación (o instructor con vistas de equipo en Permisos): todo el equipo.
 */

const ENC_RESP_ = ['id', 'json', 'actualizadoAt'];
const ENC_VIGENCIAS_ = ['BP', 'Nombre', 'Experiencia reciente programada', 'Última experiencia reciente'];

function leerVigencias_() {
  const t = filas_(hoja_(CONFIG.hojas.vigencias));
  const out = {};
  if (!t.rows.length) return out;
  const c = { bp: col_(t.head, 'bp'), prog: col_(t.head, 'experiencia reciente programada', 'exp programada'), ult: col_(t.head, 'ultima experiencia reciente', 'ultima exp') };
  t.rows.forEach(r => {
    const bp = bpTxt_(celda_(r, c.bp)); if (!bp) return;
    const p = fechaISO_(celda_(r, c.prog));
    out[bp] = { expProgramada: p && /vencid/i.test(p) ? 'VENCIDO' : p, ultimaExp: fechaISO_(celda_(r, c.ult)) };
  });
  return out;
}

function leerRespuestas_() {
  const sh = hoja_(CONFIG.hojas.respuestas, ENC_RESP_);
  const lr = sh.getLastRow();
  const out = {};
  if (lr < 2) return out;
  sh.getRange(2, 1, lr - 1, 2).getValues().forEach((r, i) => {
    if (!r[0] || !r[1]) return;
    try { const o = JSON.parse(r[1]); o._fila = i + 2; out[String(r[0])] = o; } catch (e) { /* fila dañada */ }
  });
  return out;
}

function guardarRespuesta_(resp) {
  const sh = hoja_(CONFIG.hojas.respuestas, ENC_RESP_);
  const id = resp.mes + '_' + resp.bp;
  const fila = resp._fila;
  delete resp._fila;
  resp.actualizadoAt = ahoraISO_();
  const valores = [id, JSON.stringify(resp), new Date()];
  if (fila) sh.getRange(fila, 1, 1, 3).setValues([valores]);
  else sh.appendRow(valores);
  return resp;
}

/** Todo lo que necesita el panel, según quién pide. */
function apiDatos(token) {
  return responder_(() => {
    const u = sesion_(token).u;
    const perfil = perfilPublico_(u);
    const equipo = perfil.equipo;
    const mio = bp => bp === u.bp;
    const ocultos = mesesOcultos_();
    const meses = leerRoles_().filter(m => u.rol === 'admin' || ocultos.indexOf(m.mes) < 0).map(m => ({
      mes: m.mes, fuente: m.fuente, hoja: m.hoja, origen: m.origen, actualizadoAt: m.actualizadoAt,
      total: m.instructores.length, oculto: ocultos.indexOf(m.mes) >= 0,
      instructores: equipo ? m.instructores : m.instructores.filter(i => mio(i.bp))
    }));
    const resp = leerRespuestas_();
    const respuestas = {};
    Object.keys(resp).forEach(id => {
      const r = resp[id]; delete r._fila;
      if (equipo || mio(String(r.bp))) respuestas[id] = r;
    });
    const vig = leerVigencias_(), vigencias = {};
    Object.keys(vig).forEach(bp => { if (equipo || mio(bp)) vigencias[bp] = vig[bp]; });
    const dotacion = {};
    usuarios_().forEach(x => { if ((equipo || mio(x.bp)) && x.rol === 'inst') dotacion[x.bp] = { flota: x.flota, categoria: x.categoria, estado: x.estado, nombreCompleto: x.nombreCompleto }; });
    return {
      usuario: perfil,
      hoy: hoyISO_(),
      meses: meses,
      respuestas: respuestas,
      vigencias: vigencias,
      dotacion: dotacion,
      vac: datosVac_(equipo, u.bp),
      intercambios: intercambiosPara_(u),
      intAprob: u.rol === 'admin' ? infoAprobadoresInt_(u) : null,
      intConFinal: aprobadoresInt_().final.length > 0, // para dibujar los pasos del intercambio
      baseUrl: u.rol === 'admin' ? ss_().getUrl() : ''
    };
  });
}

/* ---------------- Respuestas del instructor ---------------- */

function exigirMesEditable_(mes) {
  if (!/^\d{4}-\d{2}$/.test(String(mes || ''))) throw falla_('Mes inválido.');
  if (mes < mesActual_()) throw falla_('Solo puedes confirmar o comentar el rol del mes en curso o de los meses siguientes.');
}
function exigirEnRol_(bp, mes) {
  if (!mesVisible_(mes)) throw falla_('Ese mes todavía no está publicado para instructores.');
  const m = leerRoles_().find(x => x.mes === mes);
  if (!m || !m.instructores.some(i => i.bp === bp)) throw falla_('No figuras en el rol de ese mes.');
}
function respuestaDe_(mes, bp) {
  return leerRespuestas_()[mes + '_' + bp] || { mes: mes, bp: bp, confirmado: false, obs: [] };
}

function apiConfirmarRol(token, mes, disp) {
  return responder_(() => {
    const u = sesion_(token).u;
    exigirMesEditable_(mes); exigirEnRol_(u.bp, mes);
    return conBloqueo_(() => {
      const r = respuestaDe_(mes, u.bp);
      r.confirmado = true; r.confirmadoAt = ahoraISO_();
      return guardarRespuesta_(r);
    });
  });
}

function apiObservacion(token, mes, dia, texto) {
  return responder_(() => {
    const u = sesion_(token).u;
    exigirMesEditable_(mes); exigirEnRol_(u.bp, mes);
    texto = String(texto || '').trim();
    if (!texto) throw falla_('Escribe tu observación.');
    if (texto.length > 1500) throw falla_('La observación es muy larga (máximo 1500 caracteres).');
    const d = dia == null || dia === '' ? null : Number(dia);
    if (d !== null && !(d >= 1 && d <= diasDelMes_(mes))) throw falla_('Día inválido.');
    return conBloqueo_(() => {
      const r = respuestaDe_(mes, u.bp);
      r.obs = r.obs || [];
      r.obs.push({ id: Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36), mes: mes, dia: d, texto: texto, at: ahoraISO_(), estado: 'abierta' });
      return guardarRespuesta_(r);
    });
  });
}

function apiResolverObs(token, mes, bp, obsId) {
  return responder_(() => {
    const admin = sesion_(token, { seccion: 'inst', nivel: 2 }).u;
    bp = bpTxt_(bp);
    return conBloqueo_(() => {
      const r = leerRespuestas_()[mes + '_' + bp];
      const o = r && (r.obs || []).find(x => x.id === obsId);
      if (!o) throw falla_('No encontramos esa observación.');
      if (o.estado === 'resuelta') throw yaResuelto_('Esa observación', 'resuelta', o.resueltaPor, o.resueltaAt);
      o.estado = 'resuelta'; o.resueltaAt = ahoraISO_(); o.resueltaPor = nombreDe_(admin);
      return guardarRespuesta_(r);
    });
  });
}

/* ---------------- Carga de un Excel (alternativa a Google Sheets) ---------------- */

function apiPublicarRol(token, rol) {
  return responder_(() => {
    sesion_(token, { seccion: 'cargar', nivel: 2 });
    if (!rol || !/^\d{4}-\d{2}$/.test(String(rol.mes || ''))) throw falla_('Mes del rol inválido.');
    const nd = diasDelMes_(rol.mes);
    const ins = (rol.instructores || []).slice(0, 300).map(i => ({
      bp: bpTxt_(i.bp), nombre: String(i.nombre || '').trim().slice(0, 120),
      dias: Array.from({ length: nd }, (_, k) => String((i.dias || [])[k] == null ? '' : i.dias[k]).slice(0, 2000))
    })).filter(i => i.nombre);
    if (!ins.length) throw falla_('El archivo no tiene instructores.');
    guardarRol_({ mes: rol.mes, fuente: String(rol.archivo || 'Excel').slice(0, 200), hoja: String(rol.hoja || '').slice(0, 100), instructores: ins }, 'Excel');
    // Si el mes se leía de un Google Sheet, se pausa esa fuente para que no reemplace el Excel.
    const f = leerFuentes_().find(x => x.mes === rol.mes && x.activo);
    if (f) conBloqueo_(() => escribirFuente_(f, { activo: false, estado: 'Pausada: se publicó un Excel para este mes' }));
    return { mes: rol.mes, instructores: ins.length, fuentePausada: !!f };
  });
}
