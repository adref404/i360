/**
 * Catálogo de tipos de actividad y sus colores (hoja TiposActividad).
 *
 * Es la fuente oficial de colores: pinta las celdas que la app cambia (intercambios), permite repintar
 * un mes completo del Google Sheet con vista previa y, a futuro, servirá para generar la matriz desde la app.
 *
 * Cada tipo tiene palabras clave. Una celda es de un tipo si su texto contiene alguna de sus claves como
 * palabra completa (sin importar mayúsculas ni tildes). Con "^" al inicio, la clave debe estar al comienzo
 * de la celda. Se usa el primer tipo que coincide, en el orden de la hoja (lo más específico, primero).
 */

const ENC_TIPOS_ = ['Orden', 'Tipo', 'Palabras clave', 'Color fondo', 'Color texto', 'Activo', 'Notas'];

function hojaTipos_() { return asegurarColumnas_(hoja_(CONFIG.hojas.tipos, ENC_TIPOS_), ENC_TIPOS_); }

function leerTipos_() {
  const sh = ss_().getSheetByName(CONFIG.hojas.tipos);
  if (!sh) return [];
  const t = filas_(sh);
  const c = { o: col_(t.head, 'orden'), t: col_(t.head, 'tipo'), k: col_(t.head, 'palabras clave'), bg: col_(t.head, 'color fondo'), fg: col_(t.head, 'color texto'), a: col_(t.head, 'activo'), n: col_(t.head, 'notas') };
  return t.rows.map((r, i) => ({
    orden: Number(celda_(r, c.o)) || (i + 1), tipo: String(celda_(r, c.t) || '').trim(),
    claves: String(celda_(r, c.k) || '').split(/[;\n]+/).map(s => s.trim()).filter(Boolean),
    bg: colorHex_(celda_(r, c.bg)), fg: colorHex_(celda_(r, c.fg)) || '#000000',
    activo: normTxt_(celda_(r, c.a)) !== 'no', notas: String(celda_(r, c.n) || '')
  })).filter(x => x.tipo).sort((a, b) => a.orden - b.orden);
}

/* ---------------- Coincidencia de texto ---------------- */

const escRegex_ = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
/** Compila las claves una sola vez: devuelve una función texto → tipo (o null). */
function clasificadorTipos_(tipos) {
  const reglas = tipos.filter(t => t.activo && t.bg && t.claves.length).map(t => ({
    t: t, res: t.claves.map(k => {
      const ini = k.charAt(0) === '^', c = normTxt_(ini ? k.slice(1) : k);
      return c ? new RegExp((ini ? '^' : '(^|[^a-z0-9])') + escRegex_(c) + '($|[^a-z0-9])') : null;
    }).filter(Boolean)
  }));
  return texto => {
    const s = normTxt_(String(texto || '').replace(/\n/g, ' '));
    if (!s) return null;
    for (const r of reglas) if (r.res.some(re => re.test(s))) return r.t;
    return null;
  };
}

/* ---------------- Ubicar la matriz de un mes en su Google Sheet ---------------- */

/** Abre el Google Sheet del mes y devuelve la región de la matriz (filas de instructores × días). */
function matrizDelMes_(mes) {
  const f = leerFuentes_().find(z => z.mes === mes && z.activo && z.id);
  if (!f) throw falla_('Ese mes no tiene un Google Sheet activo conectado (puede haberse cargado desde Excel).');
  let ss;
  try { ss = SpreadsheetApp.openById(f.id); } catch (e) { throw falla_('No se pudo abrir el Google Sheet de ' + mes + '.'); }
  const sh = elegirHoja_(ss, f.hoja, f.gid, f.mes);
  const rango = sh.getDataRange(), vals = rango.getValues();
  const ub = ubicarMatriz_(vals);
  const nd = diasDelMes_(mes);
  let r1 = ub.h + 1, r2 = ub.h;
  for (let r = ub.h + 1; r < vals.length; r++) if (String(vals[r][1] || '').trim() || bpTxt_(vals[r][0])) r2 = r;
  if (r2 < r1) throw falla_('No hay instructores debajo del encabezado en el Google Sheet de ' + mes + '.');
  const region = sh.getRange(r1 + 1, ub.c0 + 1, r2 - r1 + 1, nd);
  return { f: f, sh: sh, region: region, filas: vals.slice(r1, r2 + 1).map(r => ({ bp: bpTxt_(r[0]), nombre: String(r[1] || '').trim() })) };
}

/* ---------------- Funciones públicas ---------------- */

function apiTipos(token) {
  return responder_(() => {
    sesion_(token, { seccion: 'colores', nivel: 1 });
    const meses = leerFuentes_().filter(f => f.activo && f.id && f.mes).map(f => f.mes).sort().reverse();
    return { tipos: leerTipos_(), meses: meses };
  });
}

function apiGuardarTipos(token, lista) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'colores', nivel: 2 }).u;
    const filas = (lista || []).map((t, i) => {
      const tipo = String(t.tipo || '').trim();
      const claves = (Array.isArray(t.claves) ? t.claves : String(t.claves || '').split(/[;\n]+/)).map(s => String(s).trim()).filter(Boolean);
      const bg = colorHex_(t.bg), fg = colorHex_(t.fg) || '#000000';
      if (!tipo) throw falla_('El tipo de la fila ' + (i + 1) + ' no tiene nombre.');
      if (!claves.length) throw falla_('"' + tipo + '" necesita al menos una palabra clave.');
      if (!bg) throw falla_('"' + tipo + '" necesita un color de fondo válido.');
      return [i + 1, tipo, claves.join('; '), bg, fg, t.activo === false ? 'NO' : 'SI', String(t.notas || '').slice(0, 300)];
    });
    return conBloqueo_(() => {
      const sh = hojaTipos_();
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), ENC_TIPOS_.length)).clearContent();
      if (filas.length) sh.getRange(2, 1, filas.length, ENC_TIPOS_.length).setValues(filas);
      registrar_(yo, 'Catálogo de colores actualizado', filas.length + ' tipos', '');
      return leerTipos_();
    });
  });
}

/** Propone tipos a partir de los colores que ya usa la matriz de un mes. */
function apiImportarColores(token, mes) {
  return responder_(() => {
    sesion_(token, { seccion: 'colores', nivel: 1 });
    const m = matrizDelMes_(mesTxt_(mes));
    const txt = m.region.getDisplayValues(), bgs = m.region.getBackgrounds(), fgs = m.region.getFontColors();
    const grupos = {};
    txt.forEach((fila, r) => fila.forEach((t, c) => {
      const linea = String(t || '').split('\n')[0].trim();
      if (!linea) return;
      const bg = colorHex_(bgs[r][c]), fg = colorHex_(fgs[r][c]) || '#000000';
      if (!bg) return;
      const k = bg + '|' + fg, g = grupos[k] || (grupos[k] = { bg: bg, fg: fg, n: 0, textos: {} });
      g.n++; const clave = linea.replace(/\s+/g, ' ').toUpperCase(); g.textos[clave] = (g.textos[clave] || 0) + 1;
    }));
    return Object.keys(grupos).map(k => {
      const g = grupos[k]; const top = Object.keys(g.textos).sort((a, b) => g.textos[b] - g.textos[a]);
      return { bg: g.bg, fg: g.fg, celdas: g.n, textos: top.slice(0, 8).map(t => ({ texto: t, n: g.textos[t] })) };
    }).sort((a, b) => b.celdas - a.celdas);
  });
}

/** Compara la matriz de un mes con el catálogo. aplicar=true pinta las celdas que no coinciden. */
function repintado_(mes, aplicar) {
  const tipos = leerTipos_();
  if (!tipos.some(t => t.activo)) throw falla_('El catálogo está vacío: agrega tipos de actividad primero.');
  const tipoDe = clasificadorTipos_(tipos);
  const m = matrizDelMes_(mes);
  const txt = m.region.getDisplayValues(), bgs = m.region.getBackgrounds(), fgs = m.region.getFontColors();
  const nuevosBg = bgs.map(f => f.slice()), nuevosFg = fgs.map(f => f.slice());
  const sinTipo = {}, muestra = [], porTipo = {};
  let cambios = 0, coinciden = 0, celdas = 0;
  txt.forEach((fila, r) => fila.forEach((t, c) => {
    if (!String(t || '').trim()) return;
    celdas++;
    const tipo = tipoDe(t);
    if (!tipo) { const l = String(t).split('\n')[0].trim().toUpperCase(); sinTipo[l] = (sinTipo[l] || 0) + 1; return; }
    if (colorHex_(bgs[r][c]) === tipo.bg && (colorHex_(fgs[r][c]) || '#000000') === tipo.fg) { coinciden++; return; }
    cambios++; porTipo[tipo.tipo] = (porTipo[tipo.tipo] || 0) + 1;
    nuevosBg[r][c] = tipo.bg; nuevosFg[r][c] = tipo.fg;
    if (muestra.length < 25) muestra.push({ nombre: m.filas[r].nombre, dia: c + 1, texto: String(t).split('\n')[0], antes: colorHex_(bgs[r][c]), despues: tipo.bg, fg: tipo.fg, tipo: tipo.tipo });
  }));
  if (aplicar && cambios) {
    m.region.setBackgrounds(nuevosBg);
    m.region.setFontColors(nuevosFg);
    SpreadsheetApp.flush();
    sincronizarFuente_(m.f);
  }
  return {
    mes: mes, celdas: celdas, coinciden: coinciden, cambios: cambios, porTipo: porTipo, muestra: muestra,
    sinTipo: Object.keys(sinTipo).sort((a, b) => sinTipo[b] - sinTipo[a]).slice(0, 30).map(k => ({ texto: k, n: sinTipo[k] }))
  };
}

function apiVistaRepintar(token, mes) {
  return responder_(() => { sesion_(token, { seccion: 'colores', nivel: 1 }); return repintado_(mesTxt_(mes), false); });
}

function apiRepintarMes(token, mes) {
  return responder_(() => {
    const yo = sesion_(token, { seccion: 'colores', nivel: 2 }).u;
    if (enPrueba_() && !pruebaEscribeRol_()) throw falla_('Ambiente de prueba: no se repinta el Google Sheet real del mes. Para probarlo, conecta una copia del Sheet y pon PRUEBA_ESCRIBIR = SI.');
    return conBloqueo_(() => {
      const r = repintado_(mesTxt_(mes), true);
      registrar_(yo, 'Mes repintado con el catálogo', mes + ' · ' + r.cambios + ' celdas', '');
      return r;
    });
  });
}

/** Pinta según el catálogo las celdas de un rango (las que no tienen tipo se dejan como están). */
function pintarSegunCatalogo_(rango) {
  const tipos = leerTipos_();
  if (!tipos.some(t => t.activo)) return 0;
  const tipoDe = clasificadorTipos_(tipos);
  const txt = rango.getDisplayValues(), bgs = rango.getBackgrounds(), fgs = rango.getFontColors();
  let n = 0;
  txt.forEach((fila, r) => fila.forEach((t, c) => { const tipo = tipoDe(t); if (tipo) { bgs[r][c] = tipo.bg; fgs[r][c] = tipo.fg; n++; } }));
  if (n) { rango.setBackgrounds(bgs); rango.setFontColors(fgs); }
  return n;
}
