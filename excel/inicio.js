// @ts-check
// Genera la hoja «Inicio»: la portada de la web (cita más reciente en una tarjeta de cuaderno dentro del
// marco oscuro) y su reverso debajo. La crea y la rehace ReadCite (marca en Z1, columna oculta); si hay una
// hoja «Inicio» sin la marca, no se toca. Va la primera de las pestañas.
import { citaDestacada, planInicio, COLORES_PORTADA, FUENTES } from '../core/portada.js';
import { cargarTabla, ErrorReadCite } from './capa.js';
import { textoSeguro } from '../core/datos.js';

export const HOJA_INICIO = 'Inicio';
export const MARCA_INICIO = 'readcite:inicio';
const COLS = 'ABCDEFGHIJKLMN';
const H = { left: 'Left', center: 'Center', right: 'Right' }, V = { top: 'Top', center: 'Center', bottom: 'Bottom' };
const dir = (/** @type {number} */ f1, /** @type {number} */ c1, /** @type {number} */ f2, /** @type {number} */ c2) =>
  f1 === f2 && c1 === c2 ? `${COLS[c1]}${f1 + 1}` : `${COLS[c1]}${f1 + 1}:${COLS[c2]}${f2 + 1}`;

/** Dirección de la fila de una cita en tblCitas (para «Abrir ›»), o null. @param {any} ctx @param {string} id */
async function direccionCita(ctx, id) {
  const tc = await cargarTabla(ctx, 'citas');
  if (!tc) return null;
  const j = tc.encabezados.indexOf('id'); const i = tc.filas.findIndex(f => String(f[j]) === id);
  if (i < 0) return null;
  const r = tc.cuerpo.getRow(i).load('address');
  await ctx.sync();
  const [hoja, rango] = String(r.address).split('!');
  return `${hoja}!${rango.split(':')[0]}`;
}

/**
 * (Re)genera la hoja Inicio con la cita más reciente (o con `citaId` si se indica).
 * @param {any} ctx @param {{citas:any[], config:any}} modelo
 * @param {{medir:import('../core/maquetado.js').Medidor, citaId?:string, activar?:boolean}} op
 *   activar: false = actualización silenciosa (solo vuelve a la hoja Inicio si era la que estaba a la vista)
 */
export async function generarInicio(ctx, modelo, { medir, citaId, activar = true }) {
  const activa = ctx.workbook.worksheets.getActiveWorksheet ? ctx.workbook.worksheets.getActiveWorksheet().load('name') : null;
  const elegida = citaId ? modelo.citas.find(c => c.id === citaId) : null;
  if (citaId && !elegida) throw new ErrorReadCite(`Cita ${citaId}: no existe en tblCitas.`, citaId);
  const dest = elegida ? { cita: elegida, kicker: 'Cita elegida' } : citaDestacada(modelo.citas);
  const plan = planInicio(dest, medir);
  const enlaceCita = dest ? await direccionCita(ctx, dest.cita.id) : null;

  // ¿Existe ya? Solo se rehace si la creó ReadCite.
  const vieja = ctx.workbook.worksheets.getItemOrNullObject(HOJA_INICIO);
  vieja.load('isNullObject');
  await ctx.sync();
  if (!vieja.isNullObject) {
    const m = vieja.getRange('Z1').load('values');
    await ctx.sync();
    if (m.values[0][0] !== MARCA_INICIO) throw new ErrorReadCite(`Ya existe una hoja «${HOJA_INICIO}» que no creó ReadCite. Renómbrala o muévela para que pueda dibujar la portada.`);
    vieja.delete();
    await ctx.sync();
  }
  const ws = ctx.workbook.worksheets.add(HOJA_INICIO);
  ws.position = 0;
  ws.showGridlines = false;
  ws.tabColor = COLORES_PORTADA.verde;
  const ultima = plan.altosPt.length;

  plan.columnasPt.forEach((w, i) => { ws.getRange(`${COLS[i]}:${COLS[i]}`).format.columnWidth = w; });
  plan.altosPt.forEach((h, i) => { ws.getRange(`${i + 1}:${i + 1}`).format.rowHeight = h; });
  const todo = ws.getRange(`A1:N${ultima + 4}`);
  todo.format.fill.color = COLORES_PORTADA.fondo;
  todo.format.font.name = FUENTES.sans; todo.format.font.color = COLORES_PORTADA.tinta;
  ws.getRange('Z1').values = [[MARCA_INICIO]];
  ws.getRange('Z:Z').columnHidden = true;

  // Primero los rellenos (bloques sin texto), luego los textos encima.
  const orden = [...plan.bloques].sort((a, b) => Number(!!a.v) - Number(!!b.v));
  for (const b of orden) {
    const r = ws.getRange(dir(b.f1, b.c1, b.f2, b.c2)); const f = r.format; const e = b.e;
    if (e.fill) f.fill.color = e.fill;
    if (e.bordes?.bottom) { const bd = f.borders.getItem('EdgeBottom'); bd.style = 'Continuous'; bd.weight = e.bordes.bottom.weight; bd.color = e.bordes.bottom.color; }
    if (e.bordes?.right) { const bd = f.borders.getItem('EdgeRight'); bd.style = 'Continuous'; bd.weight = e.bordes.right.weight; bd.color = e.bordes.right.color; }
    if (!b.v) continue;
    const celda = ws.getRange(dir(b.f1, b.c1, b.f1, b.c1));
    celda.values = [[textoSeguro(b.v)]];
    const fc = celda.format;
    if (e.font) fc.font.name = e.font;
    if (e.size) fc.font.size = e.size;
    if (e.bold) fc.font.bold = true;
    if (e.italic) fc.font.italic = true;
    if (e.color) fc.font.color = e.color;
    if (e.h) fc.horizontalAlignment = H[e.h];
    if (e.v) fc.verticalAlignment = V[e.v];
    if (b.link?.fila !== undefined) celda.hyperlink = { documentReference: `'${HOJA_INICIO}'!A${b.link.fila + 1}`, textToDisplay: b.v, screenTip: b.v };
    else if (b.link?.cita && enlaceCita) celda.hyperlink = { documentReference: enlaceCita, textToDisplay: b.v, screenTip: 'Ir a la fila de esta cita en la hoja Citas' };
  }
  // Los enlaces cambian la fuente a azul subrayado: se devuelve el aspecto de botón.
  for (const b of plan.bloques.filter(x => x.link)) {
    const fc = ws.getRange(dir(b.f1, b.c1, b.f1, b.c1)).format.font;
    fc.color = b.e.color || COLORES_PORTADA.tinta; fc.underline = 'None'; fc.name = b.e.font || FUENTES.sans; fc.bold = !!b.e.bold;
  }
  if (activar || activa?.name === HOJA_INICIO) { ws.activate(); ws.getRange('A1').select(); }
  await ctx.sync();
  return { cita: dest?.cita.id ?? null, kicker: dest?.kicker ?? null };
}

/** ¿Hay una hoja Inicio creada por ReadCite? (para refrescarla sola tras crear/editar/borrar). @param {any} ctx */
export async function hayInicio(ctx) {
  const ws = ctx.workbook.worksheets.getItemOrNullObject(HOJA_INICIO);
  ws.load('isNullObject');
  await ctx.sync();
  if (ws.isNullObject) return false;
  const m = ws.getRange('Z1').load('values');
  await ctx.sync();
  return m.values[0][0] === MARCA_INICIO;
}
