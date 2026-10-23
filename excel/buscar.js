// @ts-check
// Hoja «Buscar»: la barra de búsqueda y filtros de la web dentro del libro, hecha con fórmulas dinámicas
// (FILTER, SORTBY, UNIQUE, VSTACK, HSTACK…): funciona sin el panel. Misma lógica que core/datos.filtrarCitas:
// texto «contiene» sin distinguir mayúsculas en cita/autor/libro/tema/concepto; filtros exactos combinados con Y;
// «Solo mis citas» = usuario de tblConfig; orden por likes (+1 si me gusta) o por fecha de creación.
// La crea y la rehace ReadCite (marca en Z1, columnas auxiliares ocultas); rehacerla = limpiar filtros.
import { ErrorReadCite } from './capa.js';

export const HOJA_BUSCAR = 'Buscar';
export const MARCA_BUSCAR = 'readcite:buscar';
const K = { fondo: '#e4e8ee', papel: '#ffffff', tinta: '#2a2d33', gris: '#5d6573', etiqueta: '#5d6573', borde: '#c3cad4', verde: '#3ecf74', amarillo: '#fff7d6' };

/** Celdas de entrada (las que el usuario cambia). */
export const ENTRADAS = { texto: 'B3', orden: 'B6', autor: 'C6', libro: 'D6', tema: 'E6', coleccion: 'F6', mias: 'G6' };
export const VALORES_INICIALES = { texto: '', orden: 'Más ♥', autor: 'Todos', libro: 'Todos', tema: 'Todos', coleccion: 'Todas', mias: 'No' };
export const ORDENES = ['Más ♥', 'Recientes'];

/** Lista de opciones de un campo, ordenada como la web (más frecuentes primero), con «Todos» delante. @param {string} campo @param {string} todos */
const lista = (campo, todos = 'Todos') =>
  `=IFERROR(LET(v,FILTER(tblCitas[${campo}],tblCitas[${campo}]<>""),u,UNIQUE(v),VSTACK("${todos}",SORTBY(u,COUNTIF(tblCitas[${campo}],u),-1))),"${todos}")`;

/** Fórmulas de la hoja (en inglés, como las guarda Excel). */
export const FORMULAS = {
  listaAutor: lista('autor'),
  listaLibro: lista('libro'),
  listaTema: lista('tema'),
  // colecciones: las de las citas en orden de aparición + las registradas sin citas (nombresColecciones)
  listaColeccion: '=IFERROR(LET(a,VSTACK(tblCitas[coleccion],tblColecciones[nombre]),u,UNIQUE(FILTER(a,a<>"")),VSTACK("Todas",u)),"Todas")',
  resultados: `=LET(
q,LOWER(TRIM(${ENTRADAS.texto})),au,${ENTRADAS.autor},li,${ENTRADAS.libro},te,${ENTRADAS.tema},co,${ENTRADAS.coleccion},
yo,IFERROR(INDEX(tblConfig[valor],MATCH("usuario_local",tblConfig[clave],0))&"","Tú"),
tx,tblCitas[texto],
ok,(tx<>"")*IF(q="",1,ISNUMBER(FIND(q,LOWER(tx&"|"&tblCitas[autor]&"|"&tblCitas[libro]&"|"&tblCitas[tema]&"|"&tblCitas[concepto]))))
*IF(OR(au="",au="Todos"),1,tblCitas[autor]=au)*IF(OR(li="",li="Todos"),1,tblCitas[libro]=li)
*IF(OR(te="",te="Todos"),1,tblCitas[tema]=te)*IF(OR(co="",co="Todas"),1,tblCitas[coleccion]=co)
*IF(${ENTRADAS.mias}="Sí",tblCitas[usuario]=IF(yo="","Tú",yo),1),
lk,IFERROR(--tblCitas[likes],0)+IF(tblCitas[me_gusta]=TRUE,1,0),
res,HSTACK(tx&"",tblCitas[autor]&"",tblCitas[libro]&"",tblCitas[tema]&"",tblCitas[coleccion]&"",lk,IF(tblCitas[guardada_fecha]<>"","✓","")),
clave,IF(${ENTRADAS.orden}="Recientes",IFERROR(--tblCitas[creada_fecha],0),lk),
IFERROR(SORTBY(FILTER(res,ok=1),FILTER(clave,ok=1),-1),"Ninguna cita coincide con la búsqueda."))`.replace(/\n/g, ''),
  cuenta: '=LET(n,IFERROR(ROWS(B11#)*(COLUMNS(B11#)>1),0),n&IF(n=1," cita"," citas"))',
};

/**
 * (Re)crea la hoja Buscar con los filtros a su valor inicial («Limpiar filtros»).
 * @param {any} ctx @param {{activar?:boolean}} [op]
 */
export async function generarBuscar(ctx, { activar = true } = {}) {
  const vieja = ctx.workbook.worksheets.getItemOrNullObject(HOJA_BUSCAR);
  vieja.load('isNullObject');
  await ctx.sync();
  if (!vieja.isNullObject) {
    const m = vieja.getRange('Z1').load('values');
    await ctx.sync();
    if (m.values[0][0] !== MARCA_BUSCAR) throw new ErrorReadCite(`Ya existe una hoja «${HOJA_BUSCAR}» que no creó ReadCite. Renómbrala o muévela para crear la búsqueda.`);
    vieja.delete();
    await ctx.sync();
  }
  const inicio = ctx.workbook.worksheets.getItemOrNullObject('Inicio'); inicio.load('isNullObject');
  await ctx.sync();
  const ws = ctx.workbook.worksheets.add(HOJA_BUSCAR);
  ws.position = inicio.isNullObject ? 0 : 1;
  ws.showGridlines = false;
  ws.tabColor = K.verde;

  // Rejilla y fondo
  const anchos = { A: 14, B: 300, C: 130, D: 150, E: 110, F: 140, G: 50, H: 70 };
  for (const [c, w] of Object.entries(anchos)) ws.getRange(`${c}:${c}`).format.columnWidth = w;
  const todo = ws.getRange('A1:I400');
  todo.format.fill.color = K.fondo; todo.format.font.name = 'Avenir Next'; todo.format.font.size = 11; todo.format.font.color = K.tinta;
  todo.format.verticalAlignment = 'Center';
  ws.getRange('B1:H7').format.fill.color = K.papel;

  // Barra: título, búsqueda, etiquetas y desplegables
  ws.getRange('B1').values = [['Buscar citas']];
  ws.getRange('B1').format.font.size = 18; ws.getRange('B1').format.font.bold = true;
  ws.getRange('1:1').format.rowHeight = 34;
  const etq = (/** @type {string} */ dir, /** @type {string} */ t) => { const r = ws.getRange(dir); r.values = [[t]]; r.format.font.size = 9; r.format.font.color = K.etiqueta; r.format.font.bold = true; r.format.verticalAlignment = 'Bottom'; };
  etq('B2', 'BUSCAR CITAS, AUTORES, LIBROS, TEMAS…');
  const caja = ws.getRange('B3:H3'); caja.merge(false);
  ws.getRange('3:3').format.rowHeight = 26;
  for (const [i, t] of ['ORDEN', 'AUTOR', 'LIBRO', 'TEMA', 'COLECCIÓN', 'SOLO MÍAS'].entries()) etq(`${'BCDEFG'[i]}5`, t);
  ws.getRange('G5:H5').merge(false); ws.getRange('G6:H6').merge(false);
  ws.getRange('6:6').format.rowHeight = 24;
  for (const dir of ['B3:H3', 'B6', 'C6', 'D6', 'E6', 'F6', 'G6:H6']) {
    const r = ws.getRange(dir); r.format.fill.color = K.amarillo;
    for (const lado of ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight']) { const b = r.format.borders.getItem(lado); b.style = 'Continuous'; b.color = K.borde; b.weight = 'Thin'; }
  }
  const e = ENTRADAS, v = VALORES_INICIALES;
  ws.getRange(e.texto).values = [[v.texto]];
  ws.getRange('B6:G6').values = [[v.orden, v.autor, v.libro, v.tema, v.coleccion, v.mias]];
  ws.getRange('B7').values = [['Para limpiar los filtros, elige «Todos» o pulsa «Hoja Buscar» en el panel ReadCite.']];
  ws.getRange('B7').format.font.size = 9; ws.getRange('B7').format.font.color = K.gris; ws.getRange('B7').format.font.italic = true;

  // Listas de opciones (columnas ocultas AA…AD) y validaciones
  ws.getRange('AA1').formulas = [[FORMULAS.listaAutor]];
  ws.getRange('AB1').formulas = [[FORMULAS.listaLibro]];
  ws.getRange('AC1').formulas = [[FORMULAS.listaTema]];
  ws.getRange('AD1').formulas = [[FORMULAS.listaColeccion]];
  ws.getRange('Z1').values = [[MARCA_BUSCAR]];
  ws.getRange('Z:AD').columnHidden = true;
  const desplegable = (/** @type {string} */ dir, /** @type {string} */ fuente) => {
    const dv = ws.getRange(dir).dataValidation; dv.clear();
    dv.rule = { list: { inCellDropDown: true, source: fuente } };
    dv.errorAlert = { showAlert: true, style: 'Stop', title: 'Valor no válido', message: 'Elige una opción de la lista.' };
  };
  desplegable(e.orden, ORDENES.join(','));
  desplegable(e.autor, '=$AA$1#'); desplegable(e.libro, '=$AB$1#'); desplegable(e.tema, '=$AC$1#'); desplegable(e.coleccion, '=$AD$1#');
  desplegable('G6', 'No,Sí');

  // Resultados
  ws.getRange('B9').formulas = [[FORMULAS.cuenta]];
  ws.getRange('B9').format.font.color = K.gris;
  const cab = ws.getRange('B10:H10');
  cab.values = [['Cita', 'Autor', 'Libro', 'Tema', 'Colección', '♥', 'Guardada']];
  cab.format.font.bold = true; cab.format.font.size = 9; cab.format.font.color = K.etiqueta; cab.format.fill.color = K.papel;
  const bd = cab.format.borders.getItem('EdgeBottom'); bd.style = 'Continuous'; bd.color = K.borde; bd.weight = 'Thin';
  const cuerpo = ws.getRange('B11:H400');
  cuerpo.format.verticalAlignment = 'Top'; cuerpo.format.wrapText = true; cuerpo.format.font.size = 10;
  // Excel no ajusta el alto de fila a los resultados de una fórmula: alto fijo de 3-4 líneas para ver la cita entera
  ws.getRange('11:400').format.rowHeight = 50;
  ws.getRange('B11:B400').format.font.name = 'Georgia'; ws.getRange('B11:B400').format.font.italic = true;
  ws.getRange('G11:H400').format.horizontalAlignment = 'Center';
  ws.getRange('B11').formulas = [[FORMULAS.resultados]];
  ws.freezePanes.freezeRows(10);

  if (activar) { ws.activate(); ws.getRange(e.texto).select(); }
  await ctx.sync();
  return { hoja: HOJA_BUSCAR };
}
