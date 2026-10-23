// @ts-check
// Genera la hoja «Cuaderno» (A5, doble cara) a partir del plan de core/maquetado.js.
// - La hoja la crea y la regenera ReadCite: lleva una marca en L1 (columna oculta). Si existe una hoja
//   «Cuaderno» sin esa marca, no se toca y se avisa.
// - Columnas y filas en mm (puntos), márgenes de página a 0, A5 vertical, salto de página manual al inicio
//   de cada página y área de impresión A:J (todo el cuaderno o una sola letra).
import { estructuraCuaderno, NB_LETTERS } from '../core/cuaderno.js';
import { planCuaderno, FILAS_PAGINA, COLORES } from '../core/maquetado.js';
import { ErrorReadCite } from './capa.js';
import { textoSeguro } from '../core/datos.js';

export const HOJA = 'Cuaderno';
export const MARCA = 'readcite:cuaderno';
const LETRAS_COL = 'ABCDEFGHIJ';
const H = { left: 'Left', center: 'Center', right: 'Right' }, V = { top: 'Top', center: 'Center', bottom: 'Bottom' };
const PAGINAS_POR_LOTE = 6;

/** Opciones del cuaderno a partir de tblConfig. @param {any} config @param {number} ahora */
export function opcionesCuaderno(config, ahora) {
  const anio = parseInt(config.cuaderno_anio) || new Date(ahora).getFullYear();
  return { anio, porPagina: Number(config.cuaderno_citas_por_pagina) || 2, espejo: config.cuaderno_pestanas_espejo !== false, ahora };
}

/** Comprueba la hoja Cuaderno: null si no existe; lanza si existe y no la creó ReadCite. @param {any} ctx */
async function hojaPropia(ctx) {
  const ws = ctx.workbook.worksheets.getItemOrNullObject(HOJA);
  ws.load('isNullObject');
  await ctx.sync();
  if (ws.isNullObject) return null;
  const m = ws.getRange('L1').load('values');
  await ctx.sync();
  if (m.values[0][0] !== MARCA) throw new ErrorReadCite(`Ya existe una hoja «${HOJA}» que no creó ReadCite. Renómbrala o muévela para que pueda generar el cuaderno.`);
  return ws;
}

/**
 * (Re)genera la hoja Cuaderno.
 * @param {any} ctx  contexto de Excel.run
 * @param {{citas:any[], config:any}} modelo  resultado de leerTodo()
 * @param {{ahora:number, medir:import('../core/maquetado.js').Medidor, alProgresar?:(hecho:number,total:number)=>void}} op
 */
export async function generarCuaderno(ctx, modelo, { ahora, medir, alProgresar }) {
  const opciones = opcionesCuaderno(modelo.config, ahora);
  const nb = estructuraCuaderno(modelo.citas, opciones);
  const plan = planCuaderno(nb, medir);

  const vieja = await hojaPropia(ctx);
  if (vieja) { vieja.delete(); await ctx.sync(); }
  const ws = ctx.workbook.worksheets.add(HOJA);
  const ultima = plan.filas;

  // Rejilla en mm, fuente base y marca
  plan.columnasPt.forEach((w, i) => { ws.getRange(`${LETRAS_COL[i]}:${LETRAS_COL[i]}`).format.columnWidth = w; });
  ws.getRange(`1:${ultima}`).format.rowHeight = plan.altosPt[1];
  for (const f0 of plan.inicios) {
    ws.getRange(`${f0 + 1}:${f0 + 1}`).format.rowHeight = plan.altosPt[0];
    ws.getRange(`${f0 + FILAS_PAGINA}:${f0 + FILAS_PAGINA}`).format.rowHeight = plan.altosPt[FILAS_PAGINA - 1];
  }
  const todo = ws.getRange(`A1:J${ultima}`);
  todo.format.font.name = 'Georgia'; todo.format.font.size = 8; todo.format.font.color = COLORES.tinta;
  todo.format.verticalAlignment = 'Top'; todo.format.horizontalAlignment = 'Left'; // los números (año, páginas) también a la izquierda
  ws.getRange('L1').values = [[MARCA]];
  ws.getRange('L:L').columnHidden = true;
  ws.showGridlines = false;

  // Valores: una matriz por lote de páginas
  const porPagina = new Map();
  for (const b of plan.bloques) { const p = Math.floor(b.f1 / FILAS_PAGINA); if (!porPagina.has(p)) porPagina.set(p, []); porPagina.get(p).push(b); }
  await ctx.sync();

  for (let p0 = 0; p0 < plan.paginas; p0 += PAGINAS_POR_LOTE) {
    const p1 = Math.min(plan.paginas, p0 + PAGINAS_POR_LOTE);
    const fIni = p0 * FILAS_PAGINA, fFin = p1 * FILAS_PAGINA;
    const bloques = []; for (let p = p0; p < p1; p++) bloques.push(...(porPagina.get(p) || []));
    const m = Array.from({ length: fFin - fIni }, () => Array(10).fill(''));
    for (const b of bloques) { const [f, c] = ancla(b); m[f - fIni][c] = textoSeguro(b.v); }
    ws.getRange(`A${fIni + 1}:J${fFin}`).values = m;
    pintarLote(ws, bloques);
    await ctx.sync();
    if (alProgresar) alProgresar(p1, plan.paginas);
  }

  // Impresión: A5 vertical, márgenes 0, saltos manuales por página, área A:J
  if (ws.pageLayout) {
    const pl = ws.pageLayout;
    pl.paperSize = 'A5'; pl.orientation = 'Portrait';
    pl.leftMargin = 0; pl.rightMargin = 0; pl.topMargin = 0; pl.bottomMargin = 0; pl.headerMargin = 0; pl.footerMargin = 0;
    pl.centerHorizontally = true; pl.printGridlines = false; pl.zoom = { scale: 100 };
    pl.setPrintArea(`A1:J${ultima}`);
    ws.horizontalPageBreaks.removePageBreaks();
    for (const f0 of plan.inicios.slice(1)) ws.horizontalPageBreaks.add(`A${f0 + 1}`);
  }
  ws.activate();
  await ctx.sync();
  return { paginas: plan.paginas, hojas: plan.paginas / 2, citas: nb.total, anio: nb.anio, inicioLetra: plan.inicioLetra };
}

const dir = (/** @type {number} */ f1, /** @type {number} */ c1, /** @type {number} */ f2 = f1, /** @type {number} */ c2 = c1) =>
  f1 === f2 && c1 === c2 ? `${LETRAS_COL[c1]}${f1 + 1}` : `${LETRAS_COL[c1]}${f1 + 1}:${LETRAS_COL[c2]}${f2 + 1}`;
/** ¿El bloque necesita celdas combinadas? Solo si tiene texto y ocupa varias filas (Excel recorta el texto más alto
 * que su fila) o es texto multilínea; las pestañas y las etiquetas de una fila se colocan en su celda sin combinar.
 * @param {import('../core/maquetado.js').Bloque} b */
const combinar = b => !!b.v && !b.e.hueco && (b.f1 !== b.f2 || ((b.e.wrap || b.v.includes('\n')) && b.c1 !== b.c2));
/** Celda donde va el valor: la esquina si se combina; si no, la fila según la alineación vertical y la columna según la horizontal.
 * @param {import('../core/maquetado.js').Bloque} b @returns {[number, number]} */
export function ancla(b) {
  if (combinar(b)) return [b.f1, b.c1];
  const f = b.e.v === 'bottom' ? b.f2 : b.e.v === 'center' ? Math.floor((b.f1 + b.f2) / 2) : b.f1;
  return [f, b.e.h === 'right' ? b.c2 : b.c1];
}

/**
 * Formatea un lote de bloques con el mínimo de operaciones: combina solo lo imprescindible y aplica cada
 * formato una vez sobre todas las áreas que lo comparten (ws.getRanges, ExcelApi 1.9).
 * @param {any} ws @param {import('../core/maquetado.js').Bloque[]} bloques
 */
function pintarLote(ws, bloques) {
  /** @type {Map<string, string[]>} */ const grupos = new Map();
  const a = (/** @type {string} */ k, /** @type {string} */ d) => { if (!grupos.has(k)) grupos.set(k, []); /** @type {string[]} */ (grupos.get(k)).push(d); };
  for (const b of bloques) {
    const e = b.e, todo = dir(b.f1, b.c1, b.f2, b.c2);
    if (combinar(b)) ws.getRange(todo).merge(false);
    const [f, c] = ancla(b);
    const texto = combinar(b) ? todo : e.h === 'center' && b.c2 > b.c1 ? dir(f, b.c1, f, b.c2) : dir(f, c);
    if (b.v) a(JSON.stringify({ t: 1, size: e.size, bold: !!e.bold, italic: !!e.italic, color: e.color || '', wrap: !!e.wrap, v: e.v || '',
      h: e.h === 'center' && b.c2 > b.c1 && !combinar(b) ? 'CenterAcrossSelection' : e.h ? H[e.h] : '' }), texto);
    if (e.fill) a(JSON.stringify({ t: 2, fill: e.fill }), todo);
    if (e.reglaInf) a('{"t":3}', todo);
    if (e.hueco) a('{"t":4}', todo);
    if (b.link !== undefined && !e.hueco) ws.getRange(dir(f, c)).hyperlink = { documentReference: `'${HOJA}'!A${b.link + 1}`, textToDisplay: b.v, screenTip: 'Ir a esta letra' };
  }
  for (const [clave, dirs] of grupos) {
    const k = JSON.parse(clave);
    for (let i = 0; i < dirs.length; i += 40) { // direcciones por llamada (límite de longitud de getRanges)
      const fm = ws.getRanges(dirs.slice(i, i + 40).join(',')).format;
      if (k.t === 1) {
        if (k.size) fm.font.size = k.size;
        if (k.bold) fm.font.bold = true;
        if (k.italic) fm.font.italic = true;
        if (k.color) fm.font.color = k.color;
        if (k.wrap) fm.wrapText = true;
        if (k.v) fm.verticalAlignment = V[/** @type {'top'} */ (k.v)];
        if (k.h) fm.horizontalAlignment = k.h;
      } else if (k.t === 2) fm.fill.color = k.fill;
      else if (k.t === 3) { const bd = fm.borders.getItem('EdgeBottom'); bd.style = 'Continuous'; bd.weight = 'Hairline'; bd.color = COLORES.regla; }
      else for (const lado of ['EdgeTop', 'EdgeBottom']) { const bd = fm.borders.getItem(lado); bd.style = 'Continuous'; bd.weight = 'Thick'; bd.color = COLORES.blanco; }
    }
  }
}

/**
 * Cambia el área de impresión: todo el cuaderno o solo las hojas de una letra (que empieza en anverso
 * y ocupa un número par de páginas, así que se puede imprimir suelta a doble cara).
 * @param {any} ctx @param {{citas:any[], config:any}} modelo @param {string|null} letra @param {number} ahora
 */
export async function areaImpresion(ctx, modelo, letra, ahora) {
  const ws = await hojaPropia(ctx);
  if (!ws) throw new ErrorReadCite('Aún no hay cuaderno: pulsa «Generar cuaderno» primero.');
  const nb = estructuraCuaderno(modelo.citas, opcionesCuaderno(modelo.config, ahora));
  const total = nb.pages.length * FILAS_PAGINA;
  let rango = `A1:J${total}`;
  if (letra) {
    const i = NB_LETTERS.indexOf(letra); if (i < 0) throw new ErrorReadCite(`Letra no válida: ${letra}`);
    const ini = nb.start[letra] * FILAS_PAGINA;
    const fin = i < 25 ? nb.start[NB_LETTERS[i + 1]] * FILAS_PAGINA : total;
    rango = `A${ini + 1}:J${fin}`;
  }
  ws.pageLayout.setPrintArea(rango);
  await ctx.sync();
  return rango;
}
