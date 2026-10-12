/**
 * Funciones que se ejecutan desde el editor de Apps Script (menú Ejecutar).
 * 1) instalar()                          — una sola vez, después de pegar el código.
 * 2) generarCodigosAdministradores()     — si coordinación necesita un código nuevo para entrar.
 * 3) sincronizarAhora()                  — lee todos los Google Sheets de la hoja Fuentes.
 * 4) autorizarPermisos()                 — vuelve a pedir los permisos que falten (correo, disparadores, otros archivos).
 */

function instalar() {
  soloEditor_();
  const ss = ss_();
  const log = [];

  hoja_(CONFIG.hojas.usuarios, ['BP', 'Nombre (como aparece en el rol)', 'Nombre completo', 'Correo corporativo', 'Rol', 'Flota', 'Categoría', 'Estado en dotación', 'Activo', 'Observaciones']);
  const roles = hoja_(CONFIG.hojas.roles, ENC_ROLES_.slice(0, 4));
  if (!roles.getRange(1, 4).getValue()) roles.getRange(1, 4).setValue('origen');
  roles.getRange('A:A').setNumberFormat('@');
  hoja_(CONFIG.hojas.respuestas, ENC_RESP_);
  hoja_(CONFIG.hojas.registro, ENC_REGISTRO_);

  // Permisos
  const per = hoja_(CONFIG.hojas.permisos, ['Vista', 'Instructor', 'Administrador', 'Alcance']);
  const filasPer = filas_(per).rows.map(r => normTxt_(r[0]));
  [
    ['Vista por instructor', 'SI', 'SI', 'El instructor ve solo su propio rol (según su BP). El administrador puede elegir a cualquiera.'],
    ['Indicadores', 'NO', 'SI', 'Equidad del Freeze, tendencias, alertas.'],
    ['Seguimiento del mes', 'NO', 'SI', 'Confirmaciones y observaciones por instructor.'],
    ['Rol completo', 'NO', 'SI', 'Matriz con todos los instructores del mes.'],
    ['Equipo por día', 'NO', 'SI', 'Quiénes tienen actividad cada día.'],
    ['Actualizar rol', 'NO', 'SI', 'Google Sheets mensuales conectados y carga de Excel. Nunca disponible para instructores.'],
    ['Usuarios y accesos', 'NO', 'SI', 'Códigos de acceso, sesiones e ingresos. Nunca disponible para instructores.']
  ].forEach(f => { if (filasPer.indexOf(normTxt_(f[0])) < 0) { per.appendRow(f); log.push('Permisos: se agregó "' + f[0] + '".'); } });

  // Fuentes (Google Sheets mensuales)
  const fu = hoja_(CONFIG.hojas.fuentes, ENC_FUENTES_);
  fu.getRange('A:A').setNumberFormat('@');
  const existentes = leerFuentes_();
  CONFIG.fuentesIniciales.forEach(f => {
    if (existentes.some(x => x.mes === f.mes)) return;
    escribirFuente_({}, { mes: f.mes, enlace: f.enlace, hoja: f.hoja || '', activo: true, ultima: '', estado: 'Pendiente de sincronizar', instructores: 0, archivo: '' });
    log.push('Fuentes: se registró ' + f.mes + '.');
  });

  // Solicitudes y Vigencias (datos que antes venían incrustados en la página)
  const so = hoja_(CONFIG.hojas.solicitudes, ENC_SOLICITUDES_);
  if (so.getLastRow() < 2 && SEMILLA_SOLICITUDES.length) {
    so.getRange(2, 1, SEMILLA_SOLICITUDES.length, 6).setValues(SEMILLA_SOLICITUDES.map(s => [s.bp, s.nombre, s.desde, s.hasta, s.tipo, s.estado]));
    so.getRange(2, 1, SEMILLA_SOLICITUDES.length, 1).setNumberFormat('0');
    log.push('Solicitudes: ' + SEMILLA_SOLICITUDES.length + ' filas iniciales.');
  }
  const vi = hoja_(CONFIG.hojas.vigencias, ENC_VIGENCIAS_);
  if (vi.getLastRow() < 2) {
    const usu = {}; usuarios_().forEach(u => usu[u.bp] = u.nombre);
    const filas = Object.keys(SEMILLA_VIGENCIAS).map(bp => [bp, usu[bp] || '', SEMILLA_VIGENCIAS[bp].expProgramada || '', SEMILLA_VIGENCIAS[bp].ultimaExp || '']);
    if (filas.length) { vi.getRange(2, 1, filas.length, 4).setValues(filas); log.push('Vigencias: ' + filas.length + ' filas iniciales.'); }
  }

  // Instrucciones
  const ins = ss.getSheetByName('Instrucciones');
  if (ins) {
    const ya = filas_(ins).rows.map(r => normTxt_(r[0])).concat([normTxt_(ins.getRange(1, 1).getValue())]);
    [
      ['Fuentes', 'Un Google Sheet por mes (hoja ROL INS). Se agregan desde la app (Actualizar rol › Conectar Google Sheet) o escribiendo Mes (aaaa-mm) y el enlace. Activo = NO pausa la lectura.'],
      ['Solicitudes', 'Pedidos de jefatura e instructores: BP, Nombre, Desde, Hasta, Tipo y Estado (asignado / no se asignó).'],
      ['Vigencias', 'Experiencia reciente por BP (hoja BG INS). Escribe VENCIDO en la columna programada para generar una alerta.'],
      ['Acceso a la app', 'Usuario = BP o correo corporativo. La primera vez, coordinación genera un código temporal (Usuarios y accesos) y la persona crea su contraseña. Las contraseñas no se guardan en este archivo.']
    ].forEach(f => { if (ya.indexOf(normTxt_(f[0])) < 0) ins.appendRow(f); });
  }

  pimienta_();

  // Primer acceso de coordinación
  usuarios_().filter(u => u.rol === 'admin' && u.activo).forEach(u => {
    const c = cred_(u.bp);
    if (c && c.h) return;
    const g = generarCodigo_(u, 'instalar');
    log.push('CÓDIGO DE PRIMER INGRESO para ' + nombreDe_(u) + ' (BP ' + u.bp + '): ' + g.codigo + '  — vence en ' + CONFIG.codigoHoras + ' h.');
  });

  // Primera lectura de los Google Sheets registrados
  sincronizarTodo_().forEach(r => log.push('Sincronización ' + (r.mes || '') + ': ' + (r.ok ? r.instructores + ' instructores' : r.error)));

  const txt = log.join('\n') || 'Todo estaba instalado.';
  Logger.log(txt);
  return txt;
}

/** Genera un código nuevo para cada administrador activo (sirve también si olvidaron su contraseña). */
function generarCodigosAdministradores() {
  soloEditor_();
  const out = usuarios_().filter(u => u.rol === 'admin' && u.activo).map(u => {
    const g = generarCodigo_(u, 'editor');
    return nombreDe_(u) + ' (BP ' + u.bp + '): ' + g.codigo;
  });
  Logger.log(out.join('\n') || 'No hay administradores activos en la hoja Usuarios.');
}

function sincronizarAhora() {
  soloEditor_();
  sincronizarTodo_().forEach(r => Logger.log((r.mes || '?') + ': ' + (r.ok ? 'OK, ' + r.instructores + ' instructores' + (r.avisos.length ? ' · ' + r.avisos.join(' ') : '') : r.error)));
}

/** Ejecútala desde el editor si el diagnóstico (…/exec?diag=1) muestra un error de permisos. */
function autorizarPermisos() {
  soloEditor_();
  const log = [];
  log.push('Correo: ' + MailApp.getRemainingDailyQuota() + ' envíos disponibles hoy');
  log.push('Disparadores: ' + ScriptApp.getProjectTriggers().length);
  log.push('Google Sheet base: ' + ss_().getName());
  leerFuentes_().forEach(f => { try { log.push('Fuente ' + f.mes + ': ' + SpreadsheetApp.openById(f.id).getName()); } catch (e) { log.push('Fuente ' + f.mes + ': SIN ACCESO (' + e.message + ')'); } });
  Logger.log(log.join('\n'));
}
