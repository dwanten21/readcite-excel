// @ts-check
// Maquetado del cuaderno A5 para una hoja de Excel: geometría en mm, partición de líneas con el silabeo
// de nbHyphWord y plan de bloques (celdas combinadas con texto y estilo) por página.
// Excel para Mac pinta U+00AD siempre como guion visible y no lo usa para cortar, así que aquí se cortan
// las líneas a mano en los mismos puntos que usa la web y solo se escribe "-" al final de la línea partida.
// Puro: la medición de texto se inyecta (`medir`), sin DOM ni Excel.
import { nbHyph } from './silabeo.js';
import { NB_LETTERS } from './cuaderno.js';

export const PT_POR_MM = 72 / 25.4;
/** Columnas A…J en mm. Anverso: [margen interior 16 = A+B+C][contenido D…H][margen 6 = I][pestañas J].
 *  Reverso (espejo): [pestañas A][margen 6 = B][contenido C…G][margen interior 16 = H+I+J]. Total 147 mm (< 148). */
export const COLUMNAS_MM = [9, 6, 1, 57, 22, 18, 18, 1, 6, 9];
export const FILA_MM = 2.3;               // filas de la rejilla (26 pestañas × 3 filas)
export const PAD_SUP_MM = 10, PAD_INF_MM = 3;
export const FILAS_RANURA = 3;            // filas por pestaña
export const FILAS_CUERPO = 26 * FILAS_RANURA; // 78
export const FILAS_PIE = 3;
export const FILAS_PAGINA = 1 + FILAS_CUERPO + FILAS_PIE + 1; // 83 → 199,3 mm por página
export const COLORES = { tinta: '#1a1a1a', gris: '#6b6b6b', gris2: '#8a8a8a', rojo: '#e8432f', regla: '#d9d6d0', caja: '#f1eee8',
  tabActivo: '#1a1a1a', tabCon: '#e4e0d9', tabVacio: '#f0eee9', tabTextoVacio: '#a3a3a3', blanco: '#ffffff' };
export const ALTO_LINEA = 1.35;           // alto de línea de Georgia en Excel para Mac (calibrado a ojo en la Fase 5: 1,2 recortaba)
export const ALTO_LINEA_TEXTO = 1.62;     // textos de varias líneas (cita, notas, APA) salen más espaciados: con 1,35 o 1,45 la
                                          // última línea quedaba medio cortada y no se veía el «…» (medido en Excel para Mac: ≈ 12 pt a 7,5 pt)
const SEGURIDAD = 0.86;                   // margen para diferencias de medición entre el panel y Excel (0,94 dejaba que Excel
                                          // para Mac volviera a partir líneas largas: palabras sueltas y cita cortada sin «…»)

/** @typedef {(texto:string, fuente:{size:number, italic?:boolean, bold?:boolean})=>number} Medidor  ancho en pt */
/** @typedef {{size?:number,bold?:boolean,italic?:boolean,color?:string,fill?:string,h?:'left'|'center'|'right',v?:'top'|'center'|'bottom',
 *   wrap?:boolean, reglaInf?:boolean, hueco?:boolean}} Estilo */
/** @typedef {{f1:number,c1:number,f2:number,c2:number,v:string,e:Estilo,link?:number}} Bloque  filas/columnas absolutas (0-based, inclusivas); link = fila destino */

const mm = (/** @type {number} */ x) => x * PT_POR_MM;
export const ANCHO_COLUMNAS_PT = COLUMNAS_MM.map(mm);
const filaPt = mm(FILA_MM);

/** Alto (en filas de rejilla) que ocupan `n` líneas de tamaño `size`. @param {number} n @param {number} size @param {number} [alto] */
export const filasPara = (n, size, alto = ALTO_LINEA) => Math.ceil((n * size * alto) / filaPt);
/** Líneas de tamaño `size` que caben en `filas` filas. @param {number} filas @param {number} size @param {number} [alto] */
export const lineasEn = (filas, size, alto = ALTO_LINEA) => Math.max(0, Math.floor((filas * filaPt) / (size * alto)));

/**
 * Parte un texto en líneas de ancho ≤ `ancho` (pt), de forma voraz como el navegador con `hyphens: manual`:
 * una palabra que no cabe se corta por el último guion blando (U+00AD) que deje caber "prefijo-";
 * si no hay corte posible, pasa entera a la línea siguiente. La salida no lleva U+00AD.
 * @param {string} texto  texto ya silabeado (con U+00AD donde se permite cortar)
 * @param {number} ancho @param {Medidor} medir @param {{size:number,italic?:boolean,bold?:boolean}} fuente
 * @param {number} [sangria] ancho (pt) de sangría de las líneas 2+ (referencia APA)
 */
export function partirLineas(texto, ancho, medir, fuente, sangria = 0) {
  const limpio = (/** @type {string} */ s) => s.replace(/­/g, '');
  const w = (/** @type {string} */ s) => medir(limpio(s), fuente);
  const palabras = (texto || '').split(/\s+/).filter(Boolean);
  /** @type {string[]} */ const lineas = [];
  let linea = '';
  const max = () => (lineas.length ? ancho - sangria : ancho);
  while (palabras.length) {
    const p = /** @type {string} */ (palabras[0]);
    const cand = linea ? linea + ' ' + p : p;
    if (w(cand) <= max()) { linea = cand; palabras.shift(); continue; }
    // ¿cortar la palabra en un guion blando?
    const cortes = []; for (let i = 0; i < p.length; i++) if (p[i] === '­') cortes.push(i);
    let hecho = false;
    for (let k = cortes.length - 1; k >= 0; k--) {
      const pre = p.slice(0, cortes[k]);
      if (w((linea ? linea + ' ' : '') + pre + '-') <= max()) {
        lineas.push(limpio((linea ? linea + ' ' : '') + pre) + '-');
        palabras[0] = p.slice(cortes[k] + 1); linea = ''; hecho = true; break;
      }
    }
    if (hecho) continue;
    if (linea) { lineas.push(limpio(linea)); linea = ''; continue; }
    lineas.push(limpio(p)); palabras.shift(); // palabra más ancha que la línea: se deja entera
  }
  if (linea) lineas.push(limpio(linea));
  return lineas;
}

/** Recorta a `n` líneas terminando en "…" si sobra texto. @param {string[]} lineas @param {number} n @param {number} ancho @param {Medidor} medir @param {any} fuente */
export function recortar(lineas, n, ancho, medir, fuente) {
  if (lineas.length <= n) return lineas;
  if (n <= 0) return [];
  const out = lineas.slice(0, n);
  let u = out[n - 1].replace(/-$/, '');
  while (u && medir(u + '…', fuente) > ancho) u = u.slice(0, -1);
  out[n - 1] = u.replace(/[\s,;:.]+$/, '') + '…';
  return out;
}

/** Une líneas para una celda; las líneas 2+ llevan `indent` espacios delante. @param {string[]} l @param {number} [indent] */
const celda = (l, indent = 0) => l.map((x, i) => (i && indent ? ' '.repeat(indent) + x : x)).join('\n');

/**
 * Plan completo de la hoja Cuaderno a partir de estructuraCuaderno().
 * @param {ReturnType<typeof import('./cuaderno.js').estructuraCuaderno>} nb
 * @param {Medidor} medir
 * @returns {{filas:number, columnasPt:number[], altosPt:number[], inicios:number[], inicioLetra:Record<string,number>, bloques:Bloque[], paginas:number}}
 */
export function planCuaderno(nb, medir) {
  /** @type {Bloque[]} */ const B = [];
  const altosPt = [];
  const inicios = nb.pages.map((_, i) => i * FILAS_PAGINA);
  /** @type {Record<string,number>} */ const inicioLetra = {};
  for (const l of NB_LETTERS) inicioLetra[l] = inicios[nb.start[l]];
  const add = (/** @type {number} */ f1, /** @type {number} */ c1, /** @type {number} */ f2, /** @type {number} */ c2, /** @type {string} */ v, /** @type {Estilo} */ e = {}, /** @type {number} */ link = -1) => {
    /** @type {Bloque} */ const b = { f1, c1, f2, c2, v, e }; if (link >= 0) b.link = link; B.push(b);
  };

  // Columnas lógicas del contenido (0 = principal, 1-2 = intermedias, 3 = última), mapeadas a A…J según la cara:
  // anverso D | E | F | G:H ; reverso C:D | E | F | G. Así la columna de 1 mm nunca queda sola.
  const LOG_ANVERSO = [[3, 3], [4, 4], [5, 5], [6, 7]], LOG_REVERSO = [[2, 3], [4, 4], [5, 5], [6, 6]];
  let L = LOG_ANVERSO;
  const addL = (/** @type {number} */ f1, /** @type {number} */ k1, /** @type {number} */ f2, /** @type {number} */ k2, /** @type {string} */ v, /** @type {Estilo} */ e = {}, /** @type {number} */ link = -1) =>
    add(f1, L[k1][0], f2, L[k2][1], v, e, link);

  nb.pages.forEach((p, i) => {
    const f0 = inicios[i];
    altosPt.push(mm(PAD_SUP_MM)); for (let k = 0; k < FILAS_CUERPO + FILAS_PIE; k++) altosPt.push(filaPt); altosPt.push(mm(PAD_INF_MM));
    const exterior = p.recto || !nb.espejo;          // anverso (o sin espejo): pestañas a la derecha
    L = exterior ? LOG_ANVERSO : LOG_REVERSO;
    const cT = p.tabsLeft ? 0 : 9;                    // columna de pestañas
    const cuerpo = f0 + 1;                            // primera fila del cuerpo
    const anchoPt = ANCHO_COLUMNAS_PT.slice(L[0][0], L[3][1] + 1).reduce((a, b) => a + b, 0) - 6; // menos relleno de celda
    const util = anchoPt * SEGURIDAD;

    // Pestañas A-Z (la letra activa en negro, con citas en gris, vacías en gris claro).
    NB_LETTERS.forEach((l, k) => {
      const activa = p.kind === 'letter' && p.letter === l; const con = nb.byL[l].length > 0;
      add(cuerpo + k * FILAS_RANURA, cT, cuerpo + k * FILAS_RANURA + FILAS_RANURA - 1, cT, l,
        { size: 6, bold: true, h: 'center', v: 'center', hueco: true, fill: activa ? COLORES.tabActivo : con ? COLORES.tabCon : COLORES.tabVacio,
          color: activa ? COLORES.blanco : con ? COLORES.tinta : COLORES.tabTextoVacio }, inicioLetra[l]);
    });

    // Pie: número de página por el lado exterior + texto.
    const fPie = cuerpo + FILAS_CUERPO, fPie2 = fPie + FILAS_PIE - 1;
    if (exterior) { addL(fPie, 0, fPie2, 2, p.foot, { size: 6.5, color: COLORES.gris, h: 'right', v: 'center' }); addL(fPie, 3, fPie2, 3, String(p.num), { size: 7, bold: true, h: 'right', v: 'center' }); }
    else { addL(fPie, 0, fPie2, 0, String(p.num), { size: 7, bold: true, h: 'left', v: 'center' }); addL(fPie, 1, fPie2, 3, p.foot, { size: 6.5, color: COLORES.gris, h: 'left', v: 'center' }); }

    if (p.kind === 'cover') portada(cuerpo);
    else if (p.kind === 'index') indice(cuerpo);
    else letra(p, cuerpo, util);
  });

  /** @param {number} f */
  function portada(f) {
    addL(f + 2, 0, f + 6, 3, 'READCITE', { size: 14, bold: true, color: COLORES.rojo, v: 'top' });
    addL(f + 40, 0, f + 45, 3, 'Cuaderno de lectura', { size: 24, bold: true, v: 'bottom' });
    addL(f + 46, 0, f + 55, 3, String(nb.anio), { size: 54, bold: true, color: COLORES.rojo, v: 'center' });
    addL(f + 56, 0, f + 59, 3, 'Citas textuales guardadas, ordenadas de la A a la Z por el apellido del autor.', { size: 9, italic: true, color: '#555555', wrap: true, v: 'top' });
    addL(f + 68, 0, f + 69, 0, 'Lector/a', { size: 8, v: 'bottom' }); addL(f + 68, 1, f + 69, 3, '', { reglaInf: true });
    addL(f + 72, 0, f + 73, 0, 'Comenzado', { size: 8, v: 'bottom' }); addL(f + 72, 1, f + 73, 1, '', { reglaInf: true });
    addL(f + 72, 2, f + 73, 2, 'Terminado', { size: 8, v: 'bottom', h: 'right' }); addL(f + 72, 3, f + 73, 3, '', { reglaInf: true });
  }

  /** @param {number} f */
  function indice(f) {
    addL(f, 0, f + 4, 3, 'Índice', { size: 18, bold: true, v: 'center' });
    nb.indice.forEach((r, k) => {
      const fr = f + 6 + k * 2;
      addL(fr, 0, fr + 1, 0, r.l, { size: 9, bold: true, v: 'center' }, inicioLetra[r.l]);
      addL(fr, 1, fr + 1, 2, r.count, { size: 7.5, color: COLORES.gris, v: 'center' });
      addL(fr, 3, fr + 1, 3, String(r.page), { size: 9, h: 'right', v: 'center' }, inicioLetra[r.l]);
    });
    addL(f + 60, 0, f + FILAS_CUERPO - 1, 3,
      'Imprime a doble cara, volteo por el borde largo. Cada letra empieza siempre en un anverso y ocupa hojas completas: puedes sacar, añadir o reordenar las hojas de una letra sin descuadrar las demás. Margen interior de 16 mm para encuadernar o perforar.',
      { size: 7, color: '#555555', wrap: true, v: 'bottom' });
  }

  /** @param {any} p @param {number} f @param {number} util */
  function letra(p, f, util) {
    addL(f, 0, f + 4, 0, p.letter, { size: 26, bold: true, v: 'bottom' });
    addL(f, 1, f + 4, 2, p.letterNote, { size: 7.5, color: COLORES.gris, v: 'bottom' });
    addL(f, 3, f + 4, 3, `LECTURA ${nb.anio}`, { size: 6, color: COLORES.gris, h: 'right', v: 'bottom' });
    const per = p.entries.length; const hueco = 2;
    const n = Math.floor((FILAS_CUERPO - 6 - hueco * (per - 1)) / per);
    p.entries.forEach((/** @type {any} */ e, k) => entrada(e, f + 6 + k * (n + hueco), n, util));
  }

  /**
   * Una cita (o plantilla vacía) en `n` filas desde `f`. Orden de lectura: autor → cita → lo que entendí →
   * datos de la obra → referencia APA. En una cita guardada los datos son una sola banda (tema, año original,
   * página): obra, año de edición, editorial y traductor ya están en la referencia APA. Las plantillas en blanco
   * llevan las dos bandas para rellenarlas a mano. Lo fijo (autor 2 + etiqueta de notas 2 + datos + etiqueta APA 2 + APA)
   * se descuenta y el resto se reparte entre cita y notas según lo que mide cada texto: si no caben enteros,
   * las notas conservan al menos el 40 % del espacio y ambos textos bajan de tamaño antes de recortarse con «…».
   * @param {any} e @param {number} f @param {number} n @param {number} util
   */
  function entrada(e, f, n, util) {
    const T = ALTO_LINEA_TEXTO;
    /** Filas que ocupa un texto ya partido (+1 de aire). @param {string[]} l @param {number} tam */
    const filasDe = (l, tam) => (l.length ? filasPara(l.length, tam, T) + 1 : 0);
    /** Encaja un texto en `filas`: prueba tamaños de mayor a menor y, si ni así cabe, recorta con «…».
     * @param {string} texto @param {number} filas @param {number[]} tams @param {boolean} cursiva */
    const encajar = (texto, filas, tams, cursiva) => {
      let l = /** @type {string[]} */ ([]), tam = tams[0];
      for (tam of tams) { l = partirLineas(texto, util, medir, { size: tam, italic: cursiva }); if (filasDe(l, tam) <= filas) return { l, tam }; }
      return { l: recortar(l, Math.max(1, lineasEn(filas - 1, tam, T)), util, medir, { size: tam, italic: cursiva }), tam };
    };
    const TAMS_CITA = [9, 8.5, 8, 7.5, 7], TAMS_NOTAS = [8, 7.5, 7];

    // 1) autor + código
    if (e.filled) addL(f, 0, f + 1, 2, e.author, { size: 9.5, bold: true, v: 'bottom' });
    else { addL(f, 0, f + 1, 0, 'Autor', { size: 7, color: COLORES.gris, v: 'bottom' }); addL(f, 1, f + 1, 2, '', { reglaInf: true }); }
    addL(f, 3, f + 1, 3, e.code, { size: 6.5, color: COLORES.gris2, h: 'right', v: 'bottom' });

    // 2) reparto del espacio
    const fuenteApa = { size: 7 };
    const apaTxt = e.filled ? `${e.apaPre}${e.apaTitle}${e.apaPost}` : '';
    const apaL = e.filled ? partirLineas(apaTxt, util, medir, fuenteApa, medir('    ', fuenteApa)) : [];
    const filasApa = e.filled ? Math.max(3, filasDe(apaL, 7)) : 3;
    const filasDatos = e.filled ? 4 : 8;
    const libre = n - 2 - 2 - filasDatos - 2 - filasApa;  // filas para cita + notas
    let filasCita, filasNotas;
    if (!e.filled) { filasCita = Math.ceil(libre * 0.45); filasNotas = libre - filasCita; }
    else {
      const quiere = (/** @type {string} */ t, /** @type {number} */ tam, /** @type {boolean} */ c) => filasDe(partirLineas(t, util, medir, { size: tam, italic: c }), tam);
      const nCita = quiere(e.text, 9, true);
      const nNotas = e.hasNotes ? quiere(e.notes, 8, false) : 6; // sin notas: 3 renglones para escribir a mano
      if (nCita + nNotas <= libre) { filasCita = nCita; filasNotas = libre - nCita; }
      else {
        const minNotas = Math.min(nNotas, Math.max(4, Math.round(libre * 0.4)));
        filasCita = Math.max(Math.min(3, libre), Math.min(nCita, libre - minNotas));
        filasNotas = libre - filasCita;
      }
    }

    // 3) cita
    let fc = f + 2;
    if (e.filled) {
      const c = encajar(e.text, filasCita, TAMS_CITA, true);
      const usadas = Math.min(filasCita, filasDe(c.l, c.tam));
      addL(fc, 0, fc + usadas - 1, 3, celda(c.l), { size: c.tam, italic: true, wrap: true, v: 'top' });
      filasNotas += filasCita - usadas; filasCita = usadas;   // lo que no use la cita, para las notas
    } else for (let r = 0; r < filasCita; r += 2) addL(fc + r, 0, Math.min(fc + r + 1, fc + filasCita - 1), 3, '', { reglaInf: true });
    fc += filasCita;

    // 4) lo que entendí (justo debajo de la cita)
    const etq = { size: 5.5, color: COLORES.gris2, v: /** @type {'bottom'} */ ('bottom') };
    addL(fc, 0, fc + 1, 3, 'LO QUE ENTENDÍ', { ...etq, color: COLORES.rojo, bold: true });
    fc += 2;
    if (filasNotas > 0) {
      if (e.filled && e.hasNotes) {
        const nt = encajar(e.notes, filasNotas, TAMS_NOTAS, false);
        addL(fc, 0, fc + filasNotas - 1, 3, celda(nt.l), { size: nt.tam, wrap: true, v: 'top' });
      } else for (let r = 0; r < filasNotas; r += 2) addL(fc + r, 0, Math.min(fc + r + 1, fc + filasNotas - 1), 3, '', { reglaInf: true });
    }
    fc += filasNotas;

    // 5) datos de la obra (dos bandas: etiqueta 2 filas + valor 2 filas)
    const val = { size: 7.5, v: /** @type {'top'} */ ('top'), reglaInf: !e.filled };
    /** @param {[string, any, number, number][]} campos */
    const banda = campos => { for (const [t, v, a, b] of campos) { addL(fc, a, fc + 1, b, t, etq); addL(fc + 2, a, fc + 3, b, e.filled ? String(v ?? '') : '', val); } fc += 4; };
    if (e.filled) banda([['TEMA · CONCEPTO', e.topic || '—', 0, 1], ['AÑO ORIG.', e.origYear, 2, 2], ['PÁG.', e.page, 3, 3]]);
    else {
      banda([['OBRA', '', 0, 0], ['AÑO ED.', '', 1, 1], ['AÑO ORIG.', '', 2, 2], ['PÁG.', '', 3, 3]]);
      banda([['EDITORIAL · TRADUCTOR', '', 0, 0], ['TEMA · CONCEPTO', '', 1, 3]]);
    }

    // 6) referencia APA 7 (caja de color, al pie de la ficha)
    addL(fc, 0, fc + 1, 3, 'REFERENCIA APA 7', { ...etq, fill: COLORES.caja });
    if (e.filled) addL(fc + 2, 0, fc + 1 + filasApa, 3, celda(recortar(apaL, Math.max(1, lineasEn(filasApa - 1, 7, T)), util, medir, fuenteApa), 4), { size: 7, wrap: true, v: 'top', fill: COLORES.caja });
    else for (let r = 0; r < filasApa; r++) addL(fc + 2 + r, 0, fc + 2 + r, 3, '', { fill: COLORES.caja, reglaInf: true });
  }

  return { filas: nb.pages.length * FILAS_PAGINA, columnasPt: ANCHO_COLUMNAS_PT, altosPt, inicios, inicioLetra, bloques: B, paginas: nb.pages.length };
}
