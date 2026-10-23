// @ts-check
// Portada («hero») de la web llevada a una hoja de Excel: cita destacada en una tarjeta de cuaderno
// dentro del marco oscuro, y su reverso debajo (sin giro 3D). Puro: la medición de texto se inyecta.
import { likesTotales } from './datos.js';
import { partirLineas } from './maquetado.js';

/** @typedef {import('./datos.js').Cita} Cita */
/** @typedef {import('./maquetado.js').Medidor} Medidor */

/**
 * Cita de la portada: la más reciente (por fecha de creación), sea de quien sea — decisión del usuario,
 * en lugar de la «cita destacada» de la web. Sin fechas, la de más likes. null si no hay citas.
 * @param {Cita[]} citas
 * @returns {{cita:Cita, kicker:string}|null}
 */
export function citaDestacada(citas) {
  const fechadas = citas.filter(c => c.createdAt != null).sort((a, b) => /** @type {number} */ (b.createdAt) - /** @type {number} */ (a.createdAt));
  if (fechadas.length) return { cita: fechadas[0], kicker: 'Cita más reciente' };
  const top = citas.slice().sort((a, b) => likesTotales(b) - likesTotales(a))[0];
  return top ? { cita: top, kicker: 'Cita más reciente' } : null;
}

export const COLORES_PORTADA = {
  fondo: '#e4e8ee', marco: '#3b3f47', marcoBorde: '#2c2f36', papel: '#ffffff', tinta: '#2a2d33', gris: '#5d6573', gris2: '#3a3f48',
  rojo: '#e8432f', margen: '#e8a49c', renglon: '#a9c3db', agujero: '#c9d0da', verde: '#3ecf74', verdeOsc: '#23964d', verdeTexto: '#0d3a20',
  botonSombra: '#b9c0ca', pastilla: '#d5dce5', amarillo: '#f5c518',
};
export const FUENTES = { mano: 'Bradley Hand', serif: 'Georgia', sans: 'Avenir Next' };

/** Columnas (pt): A margen | B marco | C agujeros | D doble margen rojo | E «Abrir» | F hueco | G «Ver reverso» | H resto del texto | I relleno | J-K-L decoración | M marco | N margen */
export const COLUMNAS_PT = [18, 16, 26, 6, 88, 12, 112, 420, 30, 12, 34, 12, 16, 18];
const C = { A: 0, B: 1, C: 2, D: 3, E: 4, F: 5, G: 6, H: 7, I: 8, J: 9, K: 10, L: 11, M: 12, N: 13 };
const ANCHO_TEXTO = COLUMNAS_PT.slice(C.E, C.H + 1).reduce((a, b) => a + b, 0) - 8;

/** @typedef {{bottom?:{color:string,weight:'Thin'|'Medium'|'Thick'}, right?:{color:string,weight:'Thin'|'Medium'|'Thick'}}} Bordes */
/** @typedef {{font?:string,size?:number,bold?:boolean,italic?:boolean,color?:string,fill?:string,h?:'left'|'center'|'right',v?:'top'|'center'|'bottom',bordes?:Bordes}} EstiloP */
/** @typedef {{f1:number,c1:number,f2:number,c2:number,v:string,e:EstiloP,link?:{fila?:number, cita?:true}}} BloqueP */

/**
 * Plan de la hoja Inicio.
 * @param {{cita:Cita, kicker:string}|null} dest @param {Medidor} medir
 * @returns {{columnasPt:number[], altosPt:number[], bloques:BloqueP[], frente:number, reverso:number}}
 */
export function planInicio(dest, medir) {
  /** @type {number[]} */ const altos = []; /** @type {BloqueP[]} */ const B = [];
  const fila = (/** @type {number} */ h) => { altos.push(h); return altos.length - 1; };
  const add = (/** @type {number} */ f1, /** @type {number} */ c1, /** @type {number} */ f2, /** @type {number} */ c2, /** @type {string} */ v, /** @type {EstiloP} */ e = {}, /** @type {any} */ link) => {
    /** @type {BloqueP} */ const b = { f1, c1, f2, c2, v, e }; if (link) b.link = link; B.push(b);
  };
  const K = COLORES_PORTADA;

  /** Dibuja un marco con su tarjeta; `contenido` recibe la fila actual y devuelve las filas de botones (para la decoración).
   * @param {(fila:(h:number)=>number)=>{botones:number|null, centro:number}} contenido */
  function tarjeta(contenido) {
    fila(14);                                     // margen superior
    const fMarco = fila(18);                      // marco superior
    const fIni = fila(22);                        // relleno superior de la tarjeta
    const { botones, centro } = contenido(fila);
    const fFin = fila(22);                        // relleno inferior
    const fMarcoInf = fila(18);
    // fondo del marco y tarjeta de papel
    add(fMarco, C.B, fMarco, C.M, '', { fill: K.marco });
    add(fMarcoInf, C.B, fMarcoInf, C.M, '', { fill: K.marco });
    add(fIni, C.B, fFin, C.B, '', { fill: K.marco });
    add(fIni, C.J, fFin, C.M, '', { fill: K.marco });
    add(fIni, C.C, fFin, C.I, '', { fill: K.papel });   // papel de toda la tarjeta (agujeros, márgenes y texto)
    add(fIni, C.C, fFin, C.C, '', { fill: K.papel, bordes: { right: { color: K.margen, weight: 'Thin' } } });
    add(fIni, C.D, fFin, C.D, '', { fill: K.papel, bordes: { right: { color: K.margen, weight: 'Thin' } } });
    add(fIni, C.I, fFin, C.I, '', { fill: K.papel });
    // agujeros del cuaderno, uno por fila de contenido
    for (let f = fIni + 1; f < fFin; f++) add(f, C.C, f, C.C, '●', { size: 9, color: K.agujero, h: 'center', v: 'center' });
    // decoración de la derecha: pastilla, aro amarillo y botón verde
    add(fIni, C.K, fIni, C.K, '▬', { fill: K.pastilla, color: K.tinta, size: 12, h: 'center', v: 'center', bordes: { bottom: { color: '#9aa2ae', weight: 'Thick' } } });
    add(centro, C.K, centro, C.K, '◯', { font: 'Arial', size: 30, bold: true, color: K.amarillo, h: 'center', v: 'center' });
    if (botones != null) add(botones, C.K, botones, C.K, '', { fill: K.verde, bordes: { bottom: { color: K.verdeOsc, weight: 'Thick' } } });
    return { fIni, fFin };
  }

  /** @param {number} f @param {string} txt @param {number} destinoFila */
  function botones(f, txt, destinoFila) {
    if (dest) add(f, C.E, f, C.E, 'Abrir ›', { font: FUENTES.sans, size: 12, bold: true, color: K.verdeTexto, fill: K.verde, h: 'center', v: 'center', bordes: { bottom: { color: K.verdeOsc, weight: 'Thick' } } }, { cita: true });
    add(f, C.G, f, C.G, txt, { font: FUENTES.sans, size: 12, bold: true, color: K.tinta, fill: '#f4f6f9', h: 'center', v: 'center', bordes: { bottom: { color: K.botonSombra, weight: 'Thick' } } }, { fila: destinoFila });
  }

  const q = dest?.cita;
  // ── Frente ──
  let filaBotonesFrente = 0;
  const frente = tarjeta(fila => {
    const fk = fila(32);
    add(fk, C.E, fk, C.H, dest ? `${dest.kicker} · ${q?.topic || 'Sin tema'}` : 'ReadCite', { font: FUENTES.mano, size: 20, color: K.rojo, v: 'bottom' });
    const texto = dest ? `“${q?.text}”` : 'Aún no hay citas.';
    let tam = 30, lineas = /** @type {string[]} */ ([]);
    for (tam of [30, 26, 22, 19, 16]) { lineas = partirLineas(texto, ANCHO_TEXTO * 0.95, medir, { size: tam, italic: true }); if (lineas.length <= 4) break; }
    const fl = lineas.map(l => { const f = fila(Math.ceil(tam * 1.5)); add(f, C.E, f, C.H, l, { font: FUENTES.serif, size: tam, italic: true, color: K.tinta, v: 'bottom', bordes: { bottom: { color: K.renglon, weight: 'Thin' } } }); return f; });
    const fa = fila(30);
    add(fa, C.E, fa, C.H, dest ? `${q?.author} — ${q?.book}` : 'Crea la primera desde el panel ReadCite → Nueva.', { font: FUENTES.sans, size: 13, color: K.gris, v: 'bottom' });
    fila(12);
    const fb = fila(32); filaBotonesFrente = fb;
    return { botones: fb, centro: fl[Math.floor((fl.length - 1) / 2)] };
  });

  if (!dest) {
    return { columnasPt: COLUMNAS_PT, altosPt: altos, bloques: B, frente: frente.fIni, reverso: frente.fIni };
  }

  // ── Reverso (debajo) ──
  let filaBotonesReverso = 0;
  const reverso = tarjeta(fila => {
    const fk = fila(28);
    add(fk, C.E, fk, C.H, 'LO QUE ENTENDÍ', { font: FUENTES.sans, size: 10, bold: true, color: K.rojo, v: 'bottom' });
    let tam = 20, lineas = /** @type {string[]} */ ([]);
    for (tam of [20, 18, 16, 14]) { lineas = partirLineas(q?.understood || '', ANCHO_TEXTO * 0.95, medir, { size: tam }); if (lineas.length <= 6) break; }
    const fl = lineas.map(l => { const f = fila(Math.ceil(tam * 1.45)); add(f, C.E, f, C.H, l, { font: FUENTES.sans, size: tam, color: K.tinta, v: 'bottom' }); return f; });
    if (q?.why) {
      const fw = fila(26); add(fw, C.E, fw, C.H, 'POR QUÉ ME LLAMÓ LA ATENCIÓN', { font: FUENTES.sans, size: 9, bold: true, color: K.gris, v: 'bottom' });
      for (const l of partirLineas(q.why, ANCHO_TEXTO * 0.95, medir, { size: 12 })) { const f = fila(20); add(f, C.E, f, C.H, l, { font: FUENTES.sans, size: 12, color: K.gris2, v: 'bottom' }); }
    }
    const locus = [q?.chapter ? `Cap. ${q.chapter}` : '', q?.page ? `p. ${q.page}` : ''].filter(Boolean).join(' · ') || 'Sin ubicación';
    const fa = fila(28); add(fa, C.E, fa, C.H, `${q?.author} — ${q?.book} · ${locus}`, { font: FUENTES.sans, size: 11, color: K.gris, v: 'bottom' });
    fila(12);
    const fb = fila(32); filaBotonesReverso = fb;
    return { botones: fb, centro: fl[Math.floor((fl.length - 1) / 2)] ?? fk };
  });

  botones(filaBotonesFrente, 'Ver reverso', reverso.fIni);
  botones(filaBotonesReverso, 'Ver frente', frente.fIni);
  return { columnasPt: COLUMNAS_PT, altosPt: altos, bloques: B, frente: frente.fIni, reverso: reverso.fIni };
}
