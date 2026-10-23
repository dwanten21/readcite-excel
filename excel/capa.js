// @ts-check
// Capa Excel (Office.js): conecta core/ con las tablas reales de ReadCiteExcel.
// Uso desde el panel:  await Excel.run(ctx => crearCita(ctx, formulario, Date.now()))
// Reglas:
// - Solo tablas con nombre (tblCitas…), nunca referencias A1 sueltas.
// - Todas las escrituras de una operación se encolan y se envían en un único context.sync() final.
// - Antes de escribir se leen las fórmulas: si una escritura pisaría una fórmula, se para con un error.
// - Tabla vacía (solo encabezados + la fila en blanco que Excel exige) = 0 filas.
// - Hoja o tabla ausente: las lecturas devuelven [] con un aviso; las escrituras fallan con un error claro.
import * as D from '../core/datos.js';
import { TABLAS, letra } from './esquema.js';

export class ErrorReadCite extends Error {
  /** @param {string} msg @param {string} [id] */
  constructor(msg, id) { super(msg); this.name = 'ErrorReadCite'; this.id = id; }
}

/** @typedef {keyof typeof TABLAS} NombreTabla */
/** @typedef {{nombre:NombreTabla, t:any, encabezados:string[], filas:any[][], formulas:any[][], vacia:boolean, cuerpo:any}} Tabla */

const esVacia = (/** @type {any[]} */ fila) => fila.every(v => v === '' || v == null);
const esFormula = (/** @type {any} */ f) => typeof f === 'string' && f.startsWith('=');

/** Carga una tabla (encabezados, valores y fórmulas) en 2 sync. Devuelve null si no existe.
 * @param {any} ctx @param {NombreTabla} nombre @returns {Promise<Tabla|null>} */
export async function cargarTabla(ctx, nombre) {
  const t = ctx.workbook.tables.getItemOrNullObject(TABLAS[nombre].tabla);
  t.load('isNullObject');
  await ctx.sync();
  if (t.isNullObject) return null;
  const cab = t.getHeaderRowRange().load('values');
  const cuerpo = t.getDataBodyRange().load('values, formulas, rowCount');
  await ctx.sync();
  const encabezados = cab.values[0].map((/** @type {any} */ h) => String(h).trim());
  // vacía = una sola fila sin datos propios (las celdas de columnas calculadas no cuentan)
  const vacia = cuerpo.values.length === 1 && cuerpo.values[0].every((/** @type {any} */ v, /** @type {number} */ j) => esFormula(cuerpo.formulas[0][j]) || v === '' || v == null);
  return { nombre, t, encabezados, filas: vacia ? [] : cuerpo.values, formulas: cuerpo.formulas, vacia, cuerpo };
}

/** Como cargarTabla, pero exige que exista y tenga las columnas obligatorias (para escribir).
 * @param {any} ctx @param {NombreTabla} nombre */
async function tablaParaEscribir(ctx, nombre) {
  const d = TABLAS[nombre];
  const tb = await cargarTabla(ctx, nombre);
  if (!tb) throw new ErrorReadCite(`No encuentro la tabla ${d.tabla} (¿se borró la hoja «${d.hoja}»?). Usa «Reparar estructura» en el panel.`);
  const faltan = d.obligatorias.filter(c => !tb.encabezados.includes(c));
  if (faltan.length) throw new ErrorReadCite(`A la tabla ${d.tabla} le faltan las columnas: ${faltan.join(', ')}.`);
  return tb;
}

/** Comprueba que escribir `fila` en la posición `i` (o en una fila nueva si i<0) no pisa fórmulas.
 * Las celdas que no se tocan van como null (Office.js las deja intactas).
 * @param {Tabla} tb @param {any[]} fila @param {number} i @param {string} id */
function protegerFormulas(tb, fila, i, id) {
  fila.forEach((v, j) => {
    if (v === null) return;
    // fila existente: su propia fórmula; fila nueva: columna calculada (fórmula en la primera fila de datos)
    const fi = i >= 0 ? i : 0;
    const f = tb.formulas[fi]?.[j];
    // Un texto protegido con apóstrofo («=…» guardado como texto) devuelve la misma cadena en values y formulas:
    // solo es fórmula de verdad si su resultado es distinto de su texto.
    if (esFormula(f) && f !== tb.cuerpo.values?.[fi]?.[j]) throw new ErrorReadCite(`Cita ${id}: la columna «${tb.encabezados[j]}» de ${TABLAS[tb.nombre].tabla} tiene una fórmula (${f}); no la sobrescribo.`, id);
  });
}

/** Encola la escritura de una fila existente (un solo rango, una sola asignación). @param {Tabla} tb @param {number} i @param {any[]} fila @param {string} id */
function escribirFila(tb, i, fila, id) {
  protegerFormulas(tb, fila, i, id);
  tb.cuerpo.getRow(i).values = [fila.map(D.textoSeguro)];
}

/** Encola filas nuevas: si la tabla está vacía, la primera ocupa la fila en blanco. @param {Tabla} tb @param {any[][]} filas @param {string} id */
function anadirFilas(tb, filas, id) {
  if (!filas.length) return;
  filas.forEach(f => protegerFormulas(tb, f, -1, id));
  filas = filas.map(f => f.map(D.textoSeguro)); // nunca se escribe texto del usuario que Excel tome por fórmula
  let resto = filas;
  if (tb.vacia) { tb.cuerpo.getRow(0).values = [filas[0]]; resto = filas.slice(1); tb.vacia = false; }
  if (resto.length) tb.t.rows.add(null, resto);
}

/** Índices (en el cuerpo) de las filas cuyo valor en `col` cumple `pred`. @param {Tabla} tb @param {string} col @param {(v:any)=>boolean} pred */
function indices(tb, col, pred) {
  const j = tb.encabezados.indexOf(col);
  return tb.filas.map((f, i) => (pred(f[j]) ? i : -1)).filter(i => i >= 0);
}

/** Encola el borrado de filas (de abajo arriba). Si se borrarían todas, la primera se vacía en vez de borrarse
 * (Excel exige al menos una fila de datos). @param {Tabla} tb @param {number[]} idx */
function borrarFilas(tb, idx) {
  const orden = [...idx].sort((a, b) => b - a);
  const todas = orden.length === tb.filas.length;
  for (const i of orden) {
    if (todas && i === 0) tb.cuerpo.getRow(0).clear('Contents');
    else tb.t.rows.getItemAt(i).delete();
  }
}

/** Fila genérica: objeto {columna: valor} → array en el orden real; columnas ajenas → null. @param {Record<string,any>} obj @param {string[]} enc */
const objAFila = (obj, enc) => enc.map(h => (h in obj ? (obj[h] ?? '') : null));
const aObj = (/** @type {string[]} */ enc, /** @type {any[]} */ f) => Object.fromEntries(enc.map((h, i) => [h, f[i]]));

// ───────────────────────── Lectura ─────────────────────────

/** Lee todo el modelo. Nunca lanza por hojas ausentes: devuelve [] y un aviso.
 * @param {any} ctx */
export async function leerTodo(ctx) {
  /** @type {string[]} */ const avisos = [];
  const cargar = async (/** @type {NombreTabla} */ n) => {
    const tb = await cargarTabla(ctx, n);
    if (!tb) avisos.push(`Falta la tabla ${TABLAS[n].tabla} (hoja «${TABLAS[n].hoja}»).`);
    return tb;
  };
  const [tc, tl, tco, tcm, tcf] = [await cargar('citas'), await cargar('libros'), await cargar('colecciones'), await cargar('comentarios'), await cargar('config')];
  const citas = tc ? tc.filas.filter(f => !esVacia(f)).map(f => D.filaACita(f, tc.encabezados)) : [];
  const libros = tl ? tl.filas.filter(f => !esVacia(f)).map(f => D.filaALibro(f, tl.encabezados)) : [];
  const colecciones = tco ? tco.filas.filter(f => !esVacia(f)).map(f => D.filaAColeccion(f, tco.encabezados)) : [];
  const comentarios = tcm ? tcm.filas.filter(f => !esVacia(f)).map(f => D.filaAComentario(f, tcm.encabezados)) : [];
  const config = D.leerConfig(tcf ? tcf.filas.map(f => { const o = aObj(tcf.encabezados, f); return [o.clave, o.valor]; }) : []);
  return { citas: D.ordenarRecientes(citas), libros, colecciones, comentarios, config, avisos };
}

/** Búsqueda + filtros combinados (core/datos.filtrarCitas). @param {any} ctx @param {D.Filtros} filtros */
export async function buscarCitas(ctx, filtros) {
  const m = await leerTodo(ctx);
  return { citas: D.filtrarCitas(m.citas, filtros, m.config.usuario_local), avisos: m.avisos };
}

/** Atajo de un solo filtro: filtrarPor(ctx, 'author', 'Ludwig von Mises').
 * @param {any} ctx @param {'author'|'book'|'topic'|'collection'|'mine'|'soloGuardadas'} campo @param {any} valor @param {'likes'|'recent'} [orden] */
export const filtrarPor = (ctx, campo, valor, orden = 'likes') => buscarCitas(ctx, { [campo]: valor, sort: orden });

// ───────────────────────── Escritura ─────────────────────────

/** Crea una cita; si el libro o la colección no existen, los añade (como la web). Un único sync de escritura.
 * @param {any} ctx @param {Partial<D.Cita>} form @param {number} ahora */
export async function crearCita(ctx, form, ahora) {
  const tc = await tablaParaEscribir(ctx, 'citas');
  const tl = await cargarTabla(ctx, 'libros');
  const tco = await cargarTabla(ctx, 'colecciones');
  const tcf = await cargarTabla(ctx, 'config');
  const cfg = D.leerConfig(tcf ? tcf.filas.map(f => { const o = aObj(tcf.encabezados, f); return [o.clave, o.valor]; }) : []);
  const citas = tc.filas.map(f => D.filaACita(f, tc.encabezados));
  const r = D.crearCita(citas, form, { ahora, usuario: cfg.usuario_local });
  if (!r.ok) throw new ErrorReadCite(r.error);
  const libros = tl ? tl.filas.map(f => D.filaALibro(f, tl.encabezados)) : [];
  const colecciones = tco ? tco.filas.map(f => D.filaAColeccion(f, tco.encabezados)) : [];
  const ef = D.efectosAlCrear(r.cita, libros, colecciones, ahora);
  anadirFilas(tc, [D.citaAFila(r.cita, tc.encabezados).map(v => (v === undefined ? null : v))], r.cita.id);
  if (ef.libroNuevo && tl) anadirFilas(tl, [objAFila(ef.libroNuevo, tl.encabezados)], r.cita.id);
  if (ef.coleccionNueva && tco) anadirFilas(tco, [objAFila({ nombre: ef.coleccionNueva.nombre, descripcion: '', creada_fecha: D.msASerial(ahora) }, tco.encabezados)], r.cita.id);
  await ctx.sync();
  return { cita: r.cita, libroNuevo: ef.libroNuevo, coleccionNueva: ef.coleccionNueva };
}

/** Localiza una cita por id. @param {Tabla} tc @param {string} id */
function filaDeCita(tc, id) {
  const i = indices(tc, 'id', v => String(v) === id)[0];
  if (i === undefined) throw new ErrorReadCite(`Cita ${id}: no existe en tblCitas.`, id);
  return i;
}

/** Aplica una función de core/datos sobre una cita y reescribe solo su fila.
 * @param {any} ctx @param {string} id @param {(citas:D.Cita[])=>any} op */
async function modificarCita(ctx, id, op) {
  const tc = await tablaParaEscribir(ctx, 'citas');
  const i = filaDeCita(tc, id);
  const citas = tc.filas.map(f => D.filaACita(f, tc.encabezados));
  const r = op(citas);
  if (!r.ok) throw new ErrorReadCite(r.error, id);
  // solo se escriben las columnas cuyo valor cambia; el resto va como null (intacto)
  const nueva = D.citaAFila(r.cita, tc.encabezados);
  const vieja = D.citaAFila(citas[i], tc.encabezados);
  escribirFila(tc, i, nueva.map((v, j) => (v === undefined || v === vieja[j] ? null : v)), id);
  await ctx.sync();
  return r.cita;
}

/** @param {any} ctx @param {string} id @param {Partial<D.Cita>} cambios */
export const editarCita = (ctx, id, cambios) => modificarCita(ctx, id, citas => D.editarCita(citas, id, cambios));
/** @param {any} ctx @param {string} id @param {number} ahora */
export const guardarCita = (ctx, id, ahora) => modificarCita(ctx, id, citas => D.guardarCita(citas, id, ahora));
/** @param {any} ctx @param {string} id */
export const quitarGuardada = (ctx, id) => modificarCita(ctx, id, citas => D.quitarGuardada(citas, id));
/** @param {any} ctx @param {string} id */
export const alternarMeGusta = (ctx, id) => modificarCita(ctx, id, citas => D.alternarMeGusta(citas, id));

/** Borra una cita y sus comentarios. @param {any} ctx @param {string} id */
export async function borrarCita(ctx, id) {
  const tc = await tablaParaEscribir(ctx, 'citas');
  const i = filaDeCita(tc, id);
  const tcm = await cargarTabla(ctx, 'comentarios');
  borrarFilas(tc, [i]);
  let comentarios = 0;
  if (tcm && tcm.encabezados.includes('cita_id')) {
    const idx = indices(tcm, 'cita_id', v => String(v) === id);
    borrarFilas(tcm, idx);
    comentarios = idx.length;
  }
  await ctx.sync();
  return { id, comentariosBorrados: comentarios };
}

/** @param {any} ctx @param {string} citaId @param {string} texto @param {number} ahora */
export async function anadirComentario(ctx, citaId, texto, ahora) {
  const tcm = await tablaParaEscribir(ctx, 'comentarios');
  const tc = await tablaParaEscribir(ctx, 'citas');
  const m = await leerTodo(ctx);
  const r = D.anadirComentario([], tc.filas.map(f => D.filaACita(f, tc.encabezados)), { citaId, texto, usuario: m.config.usuario_local, ahora });
  if (!r.ok) throw new ErrorReadCite(r.error, citaId);
  anadirFilas(tcm, [objAFila({ cita_id: citaId, usuario: r.comentario.usuario, cuando: D.msASerial(ahora), texto: r.comentario.texto }, tcm.encabezados)], citaId);
  await ctx.sync();
  return r.comentario;
}

/** @param {any} ctx @param {{nombre:string, descripcion?:string}} datos @param {number} ahora */
export async function crearColeccion(ctx, datos, ahora) {
  const tco = await tablaParaEscribir(ctx, 'colecciones');
  const r = D.crearColeccion(tco.filas.map(f => D.filaAColeccion(f, tco.encabezados)), datos, ahora);
  if (!r.ok) throw new ErrorReadCite(r.error);
  anadirFilas(tco, [objAFila({ nombre: r.coleccion.nombre, descripcion: r.coleccion.descripcion, creada_fecha: D.msASerial(ahora) }, tco.encabezados)], r.coleccion.nombre);
  await ctx.sync();
  return r.coleccion;
}

/** Cambia estado / páginas / leídas de un libro (crea la fila si el libro solo existía en citas).
 * @param {any} ctx @param {string} titulo @param {{estado?:string, paginas?:number, leidas?:number}} cambios */
export async function actualizarLibro(ctx, titulo, cambios) {
  if (cambios.estado && !D.ESTADOS.includes(/** @type {any} */ (cambios.estado))) throw new ErrorReadCite(`Libro «${titulo}»: estado no válido (${cambios.estado}).`);
  const tl = await tablaParaEscribir(ctx, 'libros');
  /** @type {Record<string,any>} */ const c = { ...cambios };
  if ('leidas' in c) c.leidas = Math.max(0, parseInt(c.leidas) || 0);
  const i = indices(tl, 'titulo', v => String(v) === titulo)[0];
  if (i === undefined) anadirFilas(tl, [objAFila({ titulo, autor: '', paginas: 0, estado: 'quiero', leidas: 0, anio: '', ...c }, tl.encabezados)], titulo);
  else escribirFila(tl, i, tl.encabezados.map(h => (h in c ? c[h] : null)), titulo);
  await ctx.sync();
}

/** Guarda varias claves de tblConfig en un único sync. @param {any} ctx @param {Record<string, any>} pares */
export async function guardarConfig(ctx, pares) {
  const tcf = await tablaParaEscribir(ctx, 'config');
  /** @type {any[][]} */ const nuevas = [];
  for (const [clave, valor] of Object.entries(pares)) {
    const i = indices(tcf, 'clave', v => String(v) === clave)[0];
    if (i === undefined) nuevas.push(objAFila({ clave, valor }, tcf.encabezados));
    else escribirFila(tcf, i, tcf.encabezados.map(h => (h === 'valor' ? valor : null)), clave);
  }
  anadirFilas(tcf, nuevas, 'config');
  await ctx.sync();
}

/** Recrea las hojas/tablas que falten, con sus encabezados. No toca nada que exista:
 * si la hoja existe pero la tabla no y A1 no está vacía, se para con un error.
 * @param {any} ctx @returns {Promise<string[]>} lo que ha creado */
export async function repararEstructura(ctx) {
  /** @type {string[]} */ const hecho = [];
  for (const n of /** @type {NombreTabla[]} */ (Object.keys(TABLAS))) {
    const d = TABLAS[n];
    if (await cargarTabla(ctx, n)) continue;
    let hoja = ctx.workbook.worksheets.getItemOrNullObject(d.hoja);
    hoja.load('isNullObject');
    await ctx.sync();
    const rango = `A1:${letra(d.columnas.length - 1)}2`;
    if (hoja.isNullObject) { hoja = ctx.workbook.worksheets.add(d.hoja); hecho.push(`hoja ${d.hoja}`); }
    else {
      const r = hoja.getRange(rango).load('values');
      await ctx.sync();
      if (!r.values.every(esVacia)) throw new ErrorReadCite(`La hoja «${d.hoja}» existe y tiene datos en ${rango}; no creo ${d.tabla} encima. Muévelos o crea la tabla a mano.`);
    }
    hoja.getRange(`A1:${letra(d.columnas.length - 1)}1`).values = [d.columnas];
    const t = ctx.workbook.tables.add(`'${d.hoja}'!${rango}`, true);
    t.name = d.tabla;
    hecho.push(`tabla ${d.tabla}`);
    await ctx.sync();
  }
  return hecho;
}
