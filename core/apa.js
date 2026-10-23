// @ts-check
// Referencia APA 7. Copia literal de ReadCite.dc.html (apaName, apa, apaText).
// Ojo: apaName conserva la partícula en el apellido ("von Mises, L."), a diferencia de nbSurname.
import { PARTICLES } from './nombres.js';

/** @typedef {{author?:string, year?:string|number, originalYear?:string|number, translator?:string, publisher?:string, book?:string}} DatosApa */

/** "Ludwig von Mises" → "von Mises, L." @param {string} author */
export function apaName(author){
  const w=(author||'').trim().split(/\s+/).filter(Boolean);if(!w.length)return '';
  if(w.length===1)return w[0];
  let i=w.findIndex((x,k)=>k>0&&PARTICLES.includes(x.toLowerCase()));if(i<0)i=w.length-1;
  const sur=w.slice(i).join(' ');const ini=w.slice(0,i).map(x=>x.replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ.-]/g,'')).filter(Boolean).map(x=>x.includes('-')?x.split('-').map(p=>p[0].toUpperCase()+'.').join('-'):x[0].toUpperCase()+'.').join(' ');
  return ini?`${sur}, ${ini}`:sur;
}

/** Partes de la referencia; `title` va en cursiva. @param {DatosApa} q @returns {{pre:string,title:string,post:string}} */
export function apa(q){
  const pre=`${apaName(q.author)||'Autor'} (${q.year||q.originalYear||'s. f.'}). `;
  let post='';if(q.translator)post+=` (${q.translator}, Trad.)`;post+='. ';
  if(q.publisher)post+=q.publisher+'.';
  if(q.originalYear&&q.year&&q.originalYear!==q.year)post+=` (Obra original publicada en ${q.originalYear})`;
  return {pre,title:q.book||'Título',post};
}

/** Referencia en texto plano. @param {DatosApa} q */
export function apaText(q){const a=apa(q);return a.pre+a.title+a.post}
