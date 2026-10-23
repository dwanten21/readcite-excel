// @ts-check
// Reto de lectura e intervalos Icaro. Reproduce las fórmulas de ReadCite.dc.html (renderVals → books / ch).
// Funciones puras: la hora actual siempre se pasa como argumento (`ahora`, en ms).

export const ICARO_INTERVALS=[1,3,7,14,30,90];
export const PAGINAS_DIA_DEF=20;
const DIA=86400000;

/** @typedef {{titulo:string, autor?:string, paginas:number, leidas:number, estado:string}} LibroReto */

/** Días completos transcurridos. @param {number} t @param {number} ahora */
export function daysSince(t,ahora){return Math.floor((ahora-t)/DIA)}

/** Páginas al día efectivas (0, vacío o no numérico → 20, como `perDay||20`). @param {unknown} v */
export function paginasDia(v){return parseInt(String(v))||PAGINAS_DIA_DEF}

/** Días para terminar un libro: ceil(max(0,págs−leídas)/págs_día). @param {number} paginas @param {number} leidas @param {number} perDay */
export function diasRestantes(paginas,leidas,perDay){return Math.ceil(Math.max(0,paginas-leidas)/perDay)}

/** Leídas acotadas a [0, páginas], como `Math.min(pages, read)` + `max(0, …)`. @param {number} paginas @param {unknown} leidas */
export function leidasAcotadas(paginas,leidas){return Math.min(paginas,Math.max(0,parseInt(String(leidas))||0))}

/** Porcentaje entero (0 si no hay páginas). @param {number} leidas @param {number} total */
export function porcentaje(leidas,total){return total?Math.round(leidas/total*100):0}

/** @param {number} days */
export function etiquetaDiasReto(days){return days>=30?Math.round(days/30*10)/10+' meses ('+days+' días)':days+' días'}
/** @param {number} days */
export function etiquetaDiasLibro(days){return days+(days===1?' día':' días')}
/** @param {number} n */
export function etiquetaLibros(n){return n+(n===1?' libro':' libros')}

/** Fila de libro con los datos del reto. @param {LibroReto} l @param {number} perDay */
export function datosLibro(l,perDay){
  const pages=l.paginas||0;const read=leidasAcotadas(pages,l.leidas);const daysLeft=diasRestantes(pages,read,perDay);
  return {...l,paginas:pages,leidas:read,enReto:l.estado==='quiero'&&pages>0,dias:daysLeft,diasLabel:etiquetaDiasLibro(daysLeft),
    pct:porcentaje(read,pages),repasoLabel:'día '+ICARO_INTERVALS.join(', ')};
}

/** Resumen del reto: solo cuentan los libros con estado 'quiero'.
 * @param {LibroReto[]} libros @param {unknown} perDayRaw */
export function resumenReto(libros,perDayRaw){
  const perDay=paginasDia(perDayRaw);
  const wish=libros.filter(b=>b.estado==='quiero').map(b=>datosLibro(b,perDay));
  const totalPages=wish.reduce((n,b)=>n+b.paginas,0);const readPages=wish.reduce((n,b)=>n+b.leidas,0);
  const days=diasRestantes(totalPages,readPages,perDay);const pct=porcentaje(readPages,totalPages);
  return {libros:wish,totalPages,readPages,perDay,pct,pctLabel:pct+'%',days,daysLabel:etiquetaDiasReto(days),booksLabel:etiquetaLibros(wish.length)};
}

/** @typedef {{id:string, book:string, user?:string, savedAt?:number|null, createdAt?:number|null}} CitaPlan */

/** Citas en el plan Icaro: de libros en 'quiero' y guardadas o propias.
 * @param {CitaPlan[]} citas @param {LibroReto[]} libros @param {string} usuarioLocal */
export function citasEnPlan(citas,libros,usuarioLocal){
  const wish=new Set(libros.filter(b=>b.estado==='quiero').map(b=>b.titulo));
  return citas.filter(x=>wish.has(x.book)&&(!!x.savedAt||x.user===usuarioLocal));
}

/** ¿Toca repaso hoy? Días desde creación (sin fecha → 1 día, como la web) ∈ intervalos Icaro.
 * @param {CitaPlan} x @param {number} ahora */
export function tocaRepaso(x,ahora){return ICARO_INTERVALS.includes(daysSince(x.createdAt||ahora-DIA*1,ahora))}

/** Sesiones de hoy y citas en el plan. @param {CitaPlan[]} citas @param {LibroReto[]} libros @param {string} usuarioLocal @param {number} ahora */
export function sesionesHoy(citas,libros,usuarioLocal,ahora){
  const plan=citasEnPlan(citas,libros,usuarioLocal);
  return {dueToday:plan.filter(x=>tocaRepaso(x,ahora)).length,quotesInPlan:plan.length,plan};
}
