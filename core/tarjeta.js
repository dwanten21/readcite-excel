// @ts-check
// Tarjeta PNG de una cita (formato póster 2:3 o cuadrado), con el aspecto de las tarjetas de la web:
// papel de cuaderno (agujeros y doble margen rojo), tema, logo ReadCite, cita en cursiva sobre renglones
// azules —partida con guiones según nbHyph—, autor y libro. Puro: devuelve operaciones de dibujo.
import { nbHyph } from './silabeo.js';
import { partirLineas, recortar } from './maquetado.js';

/** Medidas como la web (RATIOS: póster k=1,5; cuadrada k=1) sobre 1200 px de ancho. */
export const FORMATOS = { poster: { label: 'Póster 2:3', k: 1.5 }, square: { label: 'Cuadrada', k: 1 } };
export const FUENTES_PNG = { serif: '"Source Serif 4", Georgia, serif', sans: 'Nunito, "Avenir Next", system-ui, sans-serif' };
const COL = { papel: '#ffffff', agujero: '#e4e8ee', agujeroBorde: '#c9d0da', margen: '#e8a49c', renglon: '#a9c3db', tinta: '#2a2d33', gris: '#5d6573', rojo: '#e8432f' };

/** @typedef {{t:'rect',x:number,y:number,w:number,h:number,color:string} | {t:'circulo',x:number,y:number,r:number,color:string,borde:string}
 *  | {t:'texto',x:number,y:number,texto:string,fuente:string,size:number,color:string,peso?:number,italic?:boolean,espaciado?:number,alinear?:'left'|'right'}} Op */

/**
 * @param {{text:string, author:string, book:string, topic?:string, ratio?:string}} q
 * @param {(t:string, f:{size:number,italic?:boolean,bold?:boolean,family?:string})=>number} medir  ancho en px
 * @returns {{W:number, H:number, ops:Op[], lineas:string[], tam:number}}
 */
export function planTarjeta(q, medir) {
  const W = 1200, H = Math.round(W * (FORMATOS[/** @type {'poster'} */ (q.ratio)] || FORMATOS.poster).k);
  /** @type {Op[]} */ const ops = [{ t: 'rect', x: 0, y: 0, w: W, h: H, color: COL.papel }];
  // agujeros del cuaderno y doble margen rojo
  for (let y = 75; y < H - 30; y += 152) ops.push({ t: 'circulo', x: 52, y, r: 17, color: COL.agujero, borde: COL.agujeroBorde });
  ops.push({ t: 'rect', x: 116, y: 0, w: 4, h: H, color: COL.margen }, { t: 'rect', x: 132, y: 0, w: 4, h: H, color: COL.margen });

  const x0 = 220, xF = W - 85, ancho = xF - x0;
  // cabecera: tema (versalitas espaciadas) y logo ReadCite
  ops.push({ t: 'texto', x: x0, y: 150, texto: (q.topic || 'Sin tema').toUpperCase(), fuente: FUENTES_PNG.sans, size: 34, color: COL.gris, espaciado: 6 });
  const wCite = medir('Cite', { size: 44, italic: true, bold: true, family: FUENTES_PNG.sans });
  ops.push({ t: 'texto', x: xF - wCite, y: 152, texto: 'Cite', fuente: FUENTES_PNG.sans, size: 44, color: COL.rojo, peso: 700, italic: true });
  ops.push({ t: 'texto', x: xF - wCite - 4, y: 152, texto: 'Read', fuente: FUENTES_PNG.sans, size: 44, color: COL.tinta, peso: 700, alinear: 'right' });

  // cita: el tamaño baja hasta que cabe; si aun así no cabe, se recorta con «…»
  const texto = '“' + nbHyph(q.text || 'Tu cita textual aparecerá aquí') + '”';
  const yTexto = 215, reservaAutor = 190, yMax = H - 90 - reservaAutor;
  let tam = 66, lineas = /** @type {string[]} */ ([]); const f = (/** @type {number} */ s) => ({ size: s, italic: true, family: FUENTES_PNG.serif });
  for (tam of [66, 60, 54, 48, 42, 36]) { lineas = partirLineas(texto, ancho, medir, f(tam)); if (yTexto + lineas.length * tam * 1.55 <= yMax) break; }
  const lh = tam * 1.55, caben = Math.max(1, Math.floor((yMax - yTexto) / lh));
  lineas = recortar(lineas, caben, ancho, medir, f(tam));
  lineas.forEach((l, i) => {
    const base = yTexto + (i + 1) * lh - tam * 0.45;
    ops.push({ t: 'texto', x: x0, y: base, texto: l, fuente: FUENTES_PNG.serif, size: tam, color: COL.tinta, italic: true });
    ops.push({ t: 'rect', x: x0, y: yTexto + (i + 1) * lh, w: ancho, h: 3, color: COL.renglon });
  });
  // autor y libro bajo la cita
  const yA = yTexto + lineas.length * lh + 75;
  ops.push({ t: 'texto', x: x0, y: yA, texto: q.author || 'Autor', fuente: FUENTES_PNG.sans, size: 46, color: COL.tinta, peso: 700 });
  ops.push({ t: 'texto', x: x0, y: yA + 62, texto: q.book || 'Libro', fuente: FUENTES_PNG.sans, size: 42, color: COL.gris });
  return { W, H, ops, lineas, tam };
}
