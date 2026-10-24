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
        dias: (x.dias || []).map(d => String(d == null ? '' : d)),
        bg: Array.isArray(x.bg) ? x.bg : null, fg: Array.isArray(x.fg) ? x.fg : null
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
    const marca = String(Date.now()); // la página la consulta para actualizarse sola (ver apiCambios)
    props_().setProperty('ROLES_CAMBIO', marca);
    CacheService.getScriptCache().put('ROLES_CAMBIO', marca, 21600);
    return destino;
  });
}

/* ---------------- Lectura de un Google Sheet mensual ---------------- */

function elegirHoja_(ss, hojaPref, gid, mes) {
  const hojas = ss.getSheets();
  const porGid = gid != null ? hojas.find(x => x.getSheetId() === gid) : null;
  if (hojaPref) {
    const h = ss.getSheetByName(hojaPref) || hojas.find(x => normTxt_(x.getName()) === normTxt_(hojaPref));
    if (h) return h;
    // La renombraron: el gid del enlace sigue identificando la misma pestaña
    if (porGid) return porGid;
    throw falla_('El archivo no tiene una hoja llamada "' + hojaPref + '". Hojas disponibles: ' + hojas.map(x => x.getName()).join(', ') + '.');
  }
  // Sin nombre (enlace nuevo o editado): la pestaña exacta del enlace, se llame como se llame
  if (porGid) return porGid;
  const rolIns = hojas.find(x => /rol\s*ins/i.test(x.getName()));
  if (rolIns) return rolIns;
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
  const fondos = rango.getBackgrounds(), letras = rango.getFontColors(); // la app muestra los mismos colores que el Sheet
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
    const dias = [], bg = [], fg = [];
    for (let k = 0; k < nd; k++) {
      dias.push(String(disp[r][c0 + k] == null ? '' : disp[r][c0 + k]).trim());
      bg.push(colorHex_((fondos[r] || [])[c0 + k])); fg.push(colorHex_((letras[r] || [])[c0 + k]));
    }
    if (!bp && dias.every(d => !d)) continue;          // leyendas o notas al pie
    if (!bp) avisos.push(nombre + ' no tiene BP en la columna A.');
    if (bp && vistos[bp]) { avisos.push('El BP ' + bp + ' aparece dos veces; se usa la primera fila.'); continue; }
    if (bp) vistos[bp] = true;
    instructores.push({ bp: bp, nombre: nombre, dias: dias, bg: bg, fg: fg });
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

/** excluirId: intercambio que se está aplicando en este momento (no se anula por el cambio que él mismo hizo). */
function sincronizarFuente_(f, excluirId) {
  try {
    const r = leerRolDeSheet_(f.id, f.hoja, f.gid, f.mes);
    guardarRol_({ mes: r.mes, fuente: r.fuente, hoja: r.hoja, instructores: r.instructores }, 'Google Sheet');
    // Los intercambios en curso que ya no coinciden con el rol nuevo se anulan (y se avisa por correo)
    try { anularDesfasados_(r, excluirId); } catch (e) { console.warn('No se pudieron revisar los intercambios de ' + r.mes + ': ' + e.message); }
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

/** Mes ya terminado: su copia queda congelada y no se vuelve a leer del Google Sheet. */
const mesCerrado_ = mes => !!mes && mes < mesActual_();

/** Sincroniza los meses vigentes. Un mes cerrado solo se lee si todavía no tiene copia (primera carga). */
function sincronizarTodo_(soloMes) {
  const conCopia = {}; leerRoles_().forEach(m => { conCopia[m.mes] = true; });
  return leerFuentes_().filter(f => f.activo && f.id && (!soloMes || f.mes === soloMes) && !(mesCerrado_(f.mes) && conCopia[f.mes]))
    .map(f => sincronizarFuente_(f));
}

/* =================== Funciones públicas (coordinación) =================== */

function apiFuentes(token) {
  return responder_(() => {
    sesion_(token, { seccion: 'cargar', nivel: 1 });
    const ocultos = mesesOcultos_();
    const meses = leerRoles_().map(m => ({ mes: m.mes, fuente: m.fuente, hoja: m.hoja, origen: m.origen, actualizadoAt: m.actualizadoAt, instructores: m.instructores.length, oculto: ocultos.indexOf(m.mes) >= 0 }));
    // Sin el permiso de disparadores autorizado, getProjectTriggers() falla: no debe impedir ver las fuentes.
    let auto = false;
    try { auto = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sincronizacionProgramada'); }
    catch (e) { console.error('apiFuentes: no se pudieron leer los disparadores (ejecuta autorizarPermisos): ' + (e && e.message)); }
    let tiempoReal = 0;
    try { tiempoReal = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alCambiarRol').length; } catch (e) { /* sin permiso de disparadores */ }
    return { fuentes: leerFuentes_(), meses: meses, automatica: auto, tiempoReal: tiempoReal };
  });
}

/** Conecta un Google Sheet mensual: lo lee, valida y guarda el mes. */
function apiAgregarFuente(token, enlace, mes, hoja) {
  return responder_(() => {
    sesion_(token, { seccion: 'cargar', nivel: 2 });
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
    refrescarTiempoRealSiAuto_();
    return { mes: r.mes, instructores: r.instructores.length, archivo: r.fuente, hoja: r.hoja, avisos: r.avisos };
  });
}

/**
 * Cambia el Google Sheet de un mes ya conectado (por ejemplo, si se pegó el enlace equivocado).
 * Lee el nuevo enlace (ID + gid identifican archivo y pestaña; el nombre de la hoja se completa solo),
 * lo valida y reemplaza la copia del mes. Si el mes también cambia, la copia del mes anterior se elimina.
 */
function apiEditarFuente(token, mesActual, enlace, mesNuevo, hoja) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'cargar', nivel: 2 }).u;
    mesActual = mesTxt_(mesActual);
    const id = idDeEnlace_(enlace);
    if (!id) throw falla_('Pega el enlace completo del Google Sheet (https://docs.google.com/spreadsheets/d/...).');
    const previa = leerFuentes_().find(f => f.mes === mesActual);
    if (!previa) throw falla_('Ese mes ya no está conectado. Recarga la página.', 'YA_RESUELTO');
    hoja = String(hoja || '').trim();
    const gid = gidDeEnlace_(enlace);
    // Con gid no hace falta el nombre de la hoja: se usa la pestaña exacta del enlace
    const r = leerRolDeSheet_(id, gid != null ? '' : hoja, gid, mesTxt_(mesNuevo) || '');
    if (r.mes !== mesActual && leerFuentes_().some(f => f.mes === r.mes))
      throw falla_(mesLargoTxt_(r.mes) + ' ya tiene otro Google Sheet conectado. Edita esa fila o elige otro mes.');
    if (r.mes !== mesActual) {
      // Cambiar de mes dejaría huérfanas las confirmaciones e intercambios ya registrados
      const conf = Object.keys(leerRespuestas_()).filter(k => k.indexOf(mesActual + '_') === 0).length;
      const ints = leerIntercambios_().filter(x => x.mes === mesActual).length;
      if (conf || ints) throw falla_(mesLargoTxt_(mesActual) + ' ya tiene ' + [conf ? conf + ' confirmación(es)' : '', ints ? ints + ' intercambio(s)' : ''].filter(Boolean).join(' y ') +
        '. No se puede pasar a otro mes: conecta ' + mesLargoTxt_(r.mes) + ' como una fuente nueva y, si este mes está mal, ocúltalo.');
    } else {
      const aplicados = leerIntercambios_().filter(x => x.mes === mesActual && x.estadoGuardado === 'aprobado').length;
      if (aplicados) r.avisos.push(aplicados + ' intercambio(s) aprobado(s) se escribieron en el archivo anterior: revisa que el nuevo los tenga.');
    }
    guardarRol_({ mes: r.mes, fuente: r.fuente, hoja: r.hoja, instructores: r.instructores }, 'Google Sheet');
    conBloqueo_(() => {
      escribirFuente_(leerFuentes_().find(f => f.mes === mesActual), {
        mes: r.mes, enlace: String(enlace).trim(), hoja: r.hoja, activo: true, ultima: new Date(),
        estado: ('OK' + (r.avisos.length ? ' · ' + r.avisos.join(' ') : '')).slice(0, 500),
        instructores: r.instructores.length, archivo: r.fuente + ' · ' + r.hoja
      });
      if (r.mes !== mesActual) {
        borrarRol_(mesActual);
        // Si el mes estaba oculto a instructores, el nuevo también queda oculto
        const ocultos = mesesOcultos_();
        if (ocultos.indexOf(mesActual) >= 0) props_().setProperty('MESES_OCULTOS', JSON.stringify(ocultos.filter(m => m !== mesActual).concat([r.mes]).sort()));
      }
    });
    refrescarTiempoRealSiAuto_();
    registrar_(yo, 'Fuente del rol cambiada', mesActual + (r.mes !== mesActual ? ' → ' + r.mes : '') + ' · ' + r.fuente + ' · ' + r.hoja, '');
    return { mes: r.mes, mesAnterior: mesActual, instructores: r.instructores.length, archivo: r.fuente, hoja: r.hoja, avisos: r.avisos };
  });
}
function mesLargoTxt_(mes) { return MESES_ES_[Number(mes.slice(5, 7)) - 1].replace(/^./, c => c.toUpperCase()) + ' ' + mes.slice(0, 4); }

/** Elimina la copia de un mes en la hoja Roles. */
function borrarRol_(mes) {
  const sh = hoja_(CONFIG.hojas.roles); if (!sh) return;
  const m = leerRoles_().find(x => x.mes === mes);
  if (m && m.fila) sh.deleteRow(m.fila);
}

function apiSincronizar(token, mes) {
  return responder_(() => {
    sesion_(token, { seccion: 'cargar', nivel: 2 });
    mes = mesTxt_(mes);
    if (mes && mesCerrado_(mes) && leerRoles_().some(m => m.mes === mes))
      throw falla_(mesLargoTxt_(mes) + ' ya terminó: su rol queda congelado y no se vuelve a sincronizar.');
    const res = sincronizarTodo_(mes);
    if (!res.length) throw falla_(mes ? 'Ese mes no tiene un Google Sheet activo en la hoja Fuentes.' : 'No hay Google Sheets activos del mes actual en adelante.');
    return res;
  });
}

function apiActivarFuente(token, mes, activo) {
  return responder_(() => {
    sesion_(token, { seccion: 'cargar', nivel: 2 });
    const f = leerFuentes_().find(x => x.mes === mesTxt_(mes));
    if (!f) throw falla_('No encontramos ese mes en la hoja Fuentes.');
    conBloqueo_(() => escribirFuente_(f, { activo: !!activo, estado: activo ? f.estado : 'Pausada: se conserva la última copia' }));
    refrescarTiempoRealSiAuto_();
    return true;
  });
}

function apiSincronizacionAutomatica(token, activar) {
  return responder_(() => {
    sesion_(token, { seccion: 'cargar', nivel: 2 });
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sincronizacionProgramada').forEach(t => ScriptApp.deleteTrigger(t));
    if (activar) ScriptApp.newTrigger('sincronizacionProgramada').timeBased().everyHours(1).create();
    actualizarTiempoReal_(!!activar);
    return !!activar;
  });
}

/** Disparador horario. Solo hace algo si lo invoca un disparador real del proyecto. */
function sincronizacionProgramada(e) {
  if (!e || !e.triggerUid) return;
  const valido = ScriptApp.getProjectTriggers().some(t => t.getUniqueId() === String(e.triggerUid));
  if (!valido) return;
  sincronizarTodo_();
  // Al cambiar de mes, el Sheet del mes que terminó deja de vigilarse y se suma el del mes nuevo
  try { actualizarTiempoReal_(true); } catch (x) { console.warn('Tiempo real: ' + x.message); }
}

/* ---------------- Tiempo real ----------------
   Disparador instalable "al cambiar" (onChange) en cada Google Sheet mensual vigente (mes actual y siguientes):
   cuando alguien edita, pega, cambia colores o inserta filas, Google avisa al instante y se vuelve a leer ese mes.
   Se activa junto con la sincronización automática; el disparador horario queda como respaldo (los cambios que
   llegan por fórmulas o por otros scripts no disparan onChange). Máximo 20 disparadores por usuario y script:
   por eso solo se vigilan los meses vigentes. */

/** Deja un disparador por cada Google Sheet vigente (o ninguno si activar = false). */
function actualizarTiempoReal_(activar) {
  // Se llama desde la app o el disparador horario, donde el Sheet activo sí es el base: se guarda su ID para alCambiarRol
  if (!CONFIG.baseId) props_().setProperty('BASE_ID', ss_().getId());
  const faltan = activar ? Array.from(new Set(leerFuentes_().filter(f => f.activo && f.id && f.mes >= mesActual_()).map(f => f.id))) : [];
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alCambiarRol').forEach(t => {
    const i = faltan.indexOf(t.getTriggerSourceId());
    if (i >= 0) faltan.splice(i, 1); else ScriptApp.deleteTrigger(t); // ya no vigente o duplicado
  });
  faltan.forEach(id => {
    try { ScriptApp.newTrigger('alCambiarRol').forSpreadsheet(id).onChange().create(); }
    catch (e) { console.warn('Tiempo real: no se pudo vigilar el Google Sheet ' + id + ': ' + e.message); }
  });
}
function refrescarTiempoRealSiAuto_() {
  try { if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sincronizacionProgramada')) actualizarTiempoReal_(true); }
  catch (e) { console.warn('Tiempo real: ' + e.message); }
}

/**
 * Disparador "al cambiar" de un Google Sheet mensual. Agrupa ráfagas de ediciones (pegar, varias celdas seguidas):
 * si ya hay una lectura en curso solo deja la marca, y esa lectura vuelve a leer si hubo cambios mientras leía.
 */
function alCambiarRol(e) {
  if (!e || !e.triggerUid) return;
  // Aquí SpreadsheetApp.getActive() es el Sheet del mes que cambió, no el base: se abre el base por su ID
  // (si no, leerFuentes_ buscaría la hoja Fuentes dentro del Sheet del mes)
  const base = CONFIG.baseId || props_().getProperty('BASE_ID');
  if (!base) { console.log('Tiempo real: falta el ID del Google Sheet base. Apaga y prende la sincronización automática en Actualizar rol.'); return; }
  _BASE = SpreadsheetApp.openById(base);
  // El evento "al cambiar" no siempre trae e.source: el Sheet sale del propio disparador (que además valida que sea real)
  const disp = ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === String(e.triggerUid));
  if (!disp) return;
  const id = disp.getTriggerSourceId(), c = CacheService.getScriptCache(), k = 'RT_' + id;
  const meses = leerFuentes_().filter(f => f.activo && f.id === id && !mesCerrado_(f.mes)).map(f => f.mes);
  if (!meses.length) { console.log('Tiempo real: el Google Sheet ' + id + ' no corresponde a un mes vigente activo.'); return; }
  c.put(k + '_cambio', String(Date.now()), 21600);
  if (c.get(k + '_leyendo')) { console.log('Tiempo real: ' + meses.join(', ') + ' ya se está leyendo; se tomará este cambio.'); return; }
  c.put(k + '_leyendo', '1', 300);
  try {
    for (let vuelta = 0; vuelta < 4; vuelta++) {
      Utilities.sleep(2000);
      const inicio = Date.now();
      leerFuentes_().filter(f => f.activo && f.id === id && !mesCerrado_(f.mes)).forEach(f => {
        const r = sincronizarFuente_(f);
        console.log('Tiempo real: ' + f.mes + (r.ok ? ' actualizado (' + r.instructores + ' instructores)' : ' con error: ' + r.error));
      });
      if (Number(c.get(k + '_cambio') || 0) <= inicio) break; // nadie editó mientras se leía
    }
  } finally { c.remove(k + '_leyendo'); }
}

/**
 * Marca de la última vez que cambió algún rol (solo una hora, sin datos). La página la consulta cada minuto para
 * actualizarse sola, así que no valida la sesión ni lee hojas: sale de la caché y, si expiró, de las propiedades.
 */
function apiCambios() {
  return responder_(() => {
    const c = CacheService.getScriptCache();
    let m = c.get('ROLES_CAMBIO');
    if (m === null) { m = props_().getProperty('ROLES_CAMBIO') || ''; c.put('ROLES_CAMBIO', m, 21600); }
    return { roles: m };
  });
}
