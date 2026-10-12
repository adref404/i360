/**
 * Fuentes de datos: un Google Sheet por mes (como "Rol_Instructores_LP_NOVIEMBRE"),
 * con la hoja ROL INS (fila de encabezado: BP | Nombre | 1er día | 2º día ...).
 *
 * La hoja Fuentes del Google Sheet base lista los archivos a leer. Cada sincronización
 * lee el archivo y guarda una copia del mes en la hoja Roles, que es lo que usa la app
 * (carga rápida y el mes sigue disponible aunque el archivo original cambie de lugar).
 */

const ENC_FUENTES_ = ['Mes', 'Google Sheet (enlace o ID)', 'Hoja', 'Activo', 'Última sincronización', 'Estado', 'Instructores', 'Archivo'];
const ENC_ROLES_ = ['mes', 'json', 'actualizadoAt', 'origen', 'json (continuación)', 'json (continuación)'];
const MESES_ES_ = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const TAM_TROZO_ = 45000; // una celda admite hasta 50.000 caracteres

function idDeEnlace_(s) {
  s = String(s || '').trim();
  const m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  return '';
}
function gidDeEnlace_(s) { const m = String(s || '').match(/[#&?]gid=(\d+)/); return m ? Number(m[1]) : null; }

/* ---------------- Hoja Roles (copias por mes) ---------------- */

function leerRoles_() {
  const sh = hoja_(CONFIG.hojas.roles, ENC_ROLES_.slice(0, 4));
  const lr = sh.getLastRow(), lc = Math.max(sh.getLastColumn(), 4);
  if (lr < 2) return [];
  const vals = sh.getRange(2, 1, lr - 1, lc).getValues();
  const porMes = {};
  vals.forEach((r, i) => {
    let txt = String(r[1] || '');
    for (let k = 4; k < r.length; k++) if (r[k]) txt += String(r[k]);
    if (!txt) return;
    let obj;
    try { obj = JSON.parse(txt); } catch (e) { console.warn('Roles fila ' + (i + 2) + ': JSON inválido'); return; }
    const mes = mesTxt_(r[0]) || mesTxt_(obj.mes);
    if (!mes) return;
    porMes[mes] = {
      mes: mes,
      fuente: obj.fuente || obj.archivo || '',
      hoja: obj.hoja || '',
      origen: String(r[3] || obj.origen || ''),
      actualizadoAt: esFecha_(r[2]) ? r[2].toISOString() : String(r[2] || obj.cargadoAt || ''),
      fila: i + 2,
      instructores: (obj.instructores || []).map(x => ({
        bp: bpTxt_(x.bp), nombre: String(x.nombre || '').trim(),
        dias: (x.dias || []).map(d => String(d == null ? '' : d))
      })).filter(x => x.nombre || x.bp)
    };
  });
  return Object.keys(porMes).sort().map(k => porMes[k]);
}

/** Guarda (o reemplaza) la copia de un mes en Roles. */
function guardarRol_(rol, origen) {
  return conBloqueo_(() => {
    const sh = hoja_(CONFIG.hojas.roles, ENC_ROLES_.slice(0, 4));
    const datos = { mes: rol.mes, fuente: rol.fuente || '', hoja: rol.hoja || '', instructores: rol.instructores };
    const json = JSON.stringify(datos);
    const trozos = [];
    for (let i = 0; i < json.length; i += TAM_TROZO_) trozos.push(json.slice(i, i + TAM_TROZO_));
    const fila = [rol.mes, trozos[0], new Date(), origen || ''].concat(trozos.slice(1));
    const lr = sh.getLastRow();
    let destino = -1;
    if (lr >= 2) {
      const meses = sh.getRange(2, 1, lr - 1, 2).getValues();
      for (let i = 0; i < meses.length; i++) {
        let m = mesTxt_(meses[i][0]);
        if (!m) { const x = String(meses[i][1] || '').match(/"mes"\s*:\s*"(\d{4}-\d{2})"/); m = x ? x[1] : ''; }
        if (m === rol.mes) { destino = i + 2; break; }
      }
    }
    if (destino < 0) destino = Math.max(lr, 1) + 1;
    const ancho = Math.max(fila.length, sh.getLastColumn(), 4);
    while (fila.length < ancho) fila.push('');
    if (ancho > 4 && sh.getRange(1, 5).getValue() === '') sh.getRange(1, 5, 1, ancho - 4).setValues([new Array(ancho - 4).fill('json (continuación)')]);
    sh.getRange(destino, 1).setNumberFormat('@');
    sh.getRange(destino, 1, 1, ancho).setValues([fila]);
    return destino;
  });
}

/* ---------------- Lectura de un Google Sheet mensual ---------------- */

function elegirHoja_(ss, hojaPref, gid, mes) {
  const hojas = ss.getSheets();
  if (hojaPref) {
    const h = ss.getSheetByName(hojaPref) || hojas.find(x => normTxt_(x.getName()) === normTxt_(hojaPref));
    if (!h) throw falla_('El archivo no tiene una hoja llamada "' + hojaPref + '". Hojas disponibles: ' + hojas.map(x => x.getName()).join(', ') + '.');
    return h;
  }
  const rolIns = hojas.find(x => /rol\s*ins/i.test(x.getName()));
  if (rolIns) return rolIns;
  if (gid != null) { const g = hojas.find(x => x.getSheetId() === gid); if (g) return g; }
  if (mes) {
    const nombre = MESES_ES_[Number(mes.slice(5, 7)) - 1];
    const hm = hojas.find(x => normTxt_(x.getName()) === nombre);
    if (hm) return hm;
  }
  return hojas[0];
}

/**
 * Lee un rol mensual. Devuelve {mes, hoja, fuente, instructores:[{bp,nombre,dias}], avisos:[]}.
 * Formato esperado (como ROL INS): una fila de encabezado con "BP" en la columna A, "Nombre"
 * en la B y la fecha del día 1 en la C; debajo, una fila por instructor.
 */
function leerRolDeSheet_(id, hojaPref, gid, mesPref) {
  let ss;
  try { ss = SpreadsheetApp.openById(id); } catch (e) {
    throw falla_('No se pudo abrir el Google Sheet. Revisa el enlace y que la cuenta dueña del script tenga acceso (al menos de lector). ' +
      'Si es un .xlsx subido a Drive, ábrelo y usa Archivo › Guardar como Hojas de cálculo de Google.');
  }
  const sh = elegirHoja_(ss, hojaPref, gid, mesPref);
  const rango = sh.getDataRange();
  const vals = rango.getValues(), disp = rango.getDisplayValues();
  const avisos = [];

  let h = -1;
  for (let r = 0; r < Math.min(vals.length, 20); r++) {
    if (normTxt_(vals[r][1]).indexOf('nombre') >= 0 || normTxt_(vals[r][0]) === 'bp') { h = r; break; }
  }
  if (h < 0) throw falla_('En la hoja "' + sh.getName() + '" no encontramos la fila de encabezado (BP en la columna A y Nombre en la B).');

  // La fecha se interpreta en la zona horaria del archivo de origen (si no, el 01/11 00:00 puede leerse como 31/10).
  let zonaOrigen = CONFIG.zona;
  try { zonaOrigen = ss.getSpreadsheetTimeZone() || CONFIG.zona; } catch (e) { /* zona del proyecto */ }
  let c0 = -1, mesDetectado = '';
  for (let c = 2; c < Math.min(vals[h].length, 8); c++) {
    if (esFecha_(vals[h][c])) {
      c0 = c;
      mesDetectado = Utilities.formatDate(vals[h][c], zonaOrigen, 'yyyy-MM');
      // Respaldo: el texto visible de la celda (p. ej. "1/11/2026") manda si está disponible.
      const iso = fechaISO_(disp[h][c]);
      if (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso)) mesDetectado = iso.slice(0, 7);
      break;
    }
  }
  if (c0 < 0) {
    c0 = 2;
    // Fecha escrita como texto ("01/11/2026")
    const iso = fechaISO_(disp[h][2]);
    if (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso)) mesDetectado = iso.slice(0, 7);
  }
  const mes = mesPref || mesDetectado;
  if (!mes) throw falla_('No pudimos saber de qué mes es el rol: la celda del primer día (columna C del encabezado) no tiene fecha. Indica el mes al conectar el archivo.');
  if (mesPref && mesDetectado && mesPref !== mesDetectado) avisos.push('La fecha del archivo es ' + mesDetectado + ', pero se registró como ' + mesPref + '.');

  const nd = diasDelMes_(mes);
  const instructores = [];
  const vistos = {};
  for (let r = h + 1; r < vals.length; r++) {
    const bp = bpTxt_(vals[r][0]);
    const nombre = String(disp[r][1] || '').trim();
    if (!nombre) continue;
    const dias = [];
    for (let k = 0; k < nd; k++) dias.push(String(disp[r][c0 + k] == null ? '' : disp[r][c0 + k]).trim());
    if (!bp && dias.every(d => !d)) continue;          // leyendas o notas al pie
    if (!bp) avisos.push(nombre + ' no tiene BP en la columna A.');
    if (bp && vistos[bp]) { avisos.push('El BP ' + bp + ' aparece dos veces; se usa la primera fila.'); continue; }
    if (bp) vistos[bp] = true;
    instructores.push({ bp: bp, nombre: nombre, dias: dias });
  }
  if (!instructores.length) throw falla_('La hoja "' + sh.getName() + '" no tiene instructores debajo del encabezado.');
  const vacios = instructores.reduce((a, i) => a + i.dias.filter(d => !d).length, 0);
  if (vacios) avisos.push(vacios + ' días sin asignar.');
  return { mes: mes, hoja: sh.getName(), fuente: ss.getName(), instructores: instructores, avisos: avisos };
}

/* ---------------- Hoja Fuentes ---------------- */

function leerFuentes_() {
  const sh = hoja_(CONFIG.hojas.fuentes, ENC_FUENTES_);
  const t = filas_(sh);
  const c = {
    mes: col_(t.head, 'mes'), enl: col_(t.head, 'google sheet (enlace o id)', 'google sheet', 'enlace'),
    hoja: col_(t.head, 'hoja'), act: col_(t.head, 'activo'), ult: col_(t.head, 'ultima sincronizacion'),
    est: col_(t.head, 'estado'), n: col_(t.head, 'instructores'), arc: col_(t.head, 'archivo')
  };
  return t.rows.map((r, i) => {
    const enlace = String(celda_(r, c.enl) || '').trim();
    const ult = celda_(r, c.ult);
    return {
      fila: i + 2, mes: mesTxt_(celda_(r, c.mes)), enlace: enlace, id: idDeEnlace_(enlace), gid: gidDeEnlace_(enlace),
      hoja: String(celda_(r, c.hoja) || '').trim(), activo: normTxt_(celda_(r, c.act)) !== 'no',
      ultima: esFecha_(ult) ? ult.toISOString() : String(ult || ''), estado: String(celda_(r, c.est) || ''),
      instructores: Number(celda_(r, c.n)) || 0, archivo: String(celda_(r, c.arc) || '')
    };
  }).filter(f => f.enlace);
}

function escribirFuente_(f, cambios) {
  const sh = hoja_(CONFIG.hojas.fuentes, ENC_FUENTES_);
  const fila = [
    cambios.mes != null ? cambios.mes : f.mes,
    cambios.enlace != null ? cambios.enlace : f.enlace,
    cambios.hoja != null ? cambios.hoja : f.hoja,
    (cambios.activo != null ? cambios.activo : f.activo) ? 'SI' : 'NO',
    cambios.ultima != null ? cambios.ultima : (f.ultima ? new Date(f.ultima) : ''),
    cambios.estado != null ? cambios.estado : f.estado,
    cambios.instructores != null ? cambios.instructores : f.instructores,
    cambios.archivo != null ? cambios.archivo : f.archivo
  ];
  const n = f.fila || sh.getLastRow() + 1;
  sh.getRange(n, 1).setNumberFormat('@');
  sh.getRange(n, 1, 1, fila.length).setValues([fila]);
  return n;
}

function sincronizarFuente_(f) {
  try {
    const r = leerRolDeSheet_(f.id, f.hoja, f.gid, f.mes);
    guardarRol_({ mes: r.mes, fuente: r.fuente, hoja: r.hoja, instructores: r.instructores }, 'Google Sheet');
    const estado = 'OK' + (r.avisos.length ? ' · ' + r.avisos.join(' ') : '');
    conBloqueo_(() => escribirFuente_(f, { mes: r.mes, ultima: new Date(), estado: estado.slice(0, 500), instructores: r.instructores.length, archivo: r.fuente + ' · ' + r.hoja }));
    return { ok: true, mes: r.mes, instructores: r.instructores.length, archivo: r.fuente, hoja: r.hoja, avisos: r.avisos };
  } catch (e) {
    const msg = e && e.publico ? e.message : 'Error al leer: ' + (e && e.message);
    if (!(e && e.publico)) console.error(e && e.stack ? e.stack : e);
    try { conBloqueo_(() => escribirFuente_(f, { ultima: new Date(), estado: ('Error: ' + msg).slice(0, 500) })); } catch (x) { /* nada */ }
    return { ok: false, mes: f.mes, error: msg };
  }
}

function sincronizarTodo_(soloMes) {
  return leerFuentes_().filter(f => f.activo && f.id && (!soloMes || f.mes === soloMes)).map(sincronizarFuente_);
}

/* =================== Funciones públicas (coordinación) =================== */

function apiFuentes(token) {
  return responder_(() => {
    sesion_(token, { admin: true });
    const meses = leerRoles_().map(m => ({ mes: m.mes, fuente: m.fuente, hoja: m.hoja, origen: m.origen, actualizadoAt: m.actualizadoAt, instructores: m.instructores.length }));
    const auto = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sincronizacionProgramada');
    return { fuentes: leerFuentes_(), meses: meses, automatica: auto };
  });
}

/** Conecta un Google Sheet mensual: lo lee, valida y guarda el mes. */
function apiAgregarFuente(token, enlace, mes, hoja) {
  return responder_(() => {
    sesion_(token, { admin: true });
    const id = idDeEnlace_(enlace);
    if (!id) throw falla_('Pega el enlace completo del Google Sheet (https://docs.google.com/spreadsheets/d/...).');
    mes = mesTxt_(mes);
    hoja = String(hoja || '').trim();
    const r = leerRolDeSheet_(id, hoja, gidDeEnlace_(enlace), mes);
    guardarRol_({ mes: r.mes, fuente: r.fuente, hoja: r.hoja, instructores: r.instructores }, 'Google Sheet');
    conBloqueo_(() => {
      const existentes = leerFuentes_();
      const previa = existentes.find(f => f.mes === r.mes) || existentes.find(f => f.id === id && !f.mes);
      escribirFuente_(previa || {}, {
        mes: r.mes, enlace: String(enlace).trim(), hoja: hoja || r.hoja, activo: true, ultima: new Date(),
        estado: ('OK' + (r.avisos.length ? ' · ' + r.avisos.join(' ') : '')).slice(0, 500),
        instructores: r.instructores.length, archivo: r.fuente + ' · ' + r.hoja
      });
    });
    return { mes: r.mes, instructores: r.instructores.length, archivo: r.fuente, hoja: r.hoja, avisos: r.avisos };
  });
}

function apiSincronizar(token, mes) {
  return responder_(() => {
    sesion_(token, { admin: true });
    const res = sincronizarTodo_(mesTxt_(mes));
    if (!res.length) throw falla_(mes ? 'Ese mes no tiene un Google Sheet activo en la hoja Fuentes.' : 'No hay Google Sheets activos en la hoja Fuentes.');
    return res;
  });
}

function apiActivarFuente(token, mes, activo) {
  return responder_(() => {
    sesion_(token, { admin: true });
    const f = leerFuentes_().find(x => x.mes === mesTxt_(mes));
    if (!f) throw falla_('No encontramos ese mes en la hoja Fuentes.');
    conBloqueo_(() => escribirFuente_(f, { activo: !!activo, estado: activo ? f.estado : 'Pausada: se conserva la última copia' }));
    return true;
  });
}

function apiSincronizacionAutomatica(token, activar) {
  return responder_(() => {
    sesion_(token, { admin: true });
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sincronizacionProgramada').forEach(t => ScriptApp.deleteTrigger(t));
    if (activar) ScriptApp.newTrigger('sincronizacionProgramada').timeBased().everyHours(1).create();
    return !!activar;
  });
}

/** Disparador horario. Solo hace algo si lo invoca un disparador real del proyecto. */
function sincronizacionProgramada(e) {
  if (!e || !e.triggerUid) return;
  const valido = ScriptApp.getProjectTriggers().some(t => t.getUniqueId() === String(e.triggerUid));
  if (!valido) return;
  sincronizarTodo_();
}
