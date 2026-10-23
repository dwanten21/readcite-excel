// @ts-check
// Cuaderno A5: orden A-Z por apellido y estructura de páginas. Reproduce nbVals de ReadCite.dc.html sin la parte de pantalla.
// Reglas: portada + índice; por letra al menos 2 páginas y siempre en número par (cada letra empieza en anverso);
// recto = página impar; pestañas a la derecha en anverso y a la izquierda en reverso si `espejo`.
import { nbSurname, nbAuthor, nbInitial } from './nombres.js';
import { nbHyph } from './silabeo.js';
import { apa } from './apa.js';

export const NB_LETTERS='ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
export const NB_MONTHS=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

/** @typedef {import('./datos.js').Cita} Cita */

/** Comparador del cuaderno: apellido, luego libro, luego página (locale 'es'). @param {Cita} a @param {Cita} b */
export const compararCuaderno=(a,b)=>nbSurname(a.author).localeCompare(nbSurname(b.author),'es')||(a.book||'').localeCompare(b.book||'','es')||(+a.page||0)-(+b.page||0);

/** Año de guardado (sin fecha numérica → año actual). @param {number|null} v @param {number} anioActual */
const anioDe=(v,anioActual)=>typeof v==='number'?new Date(v).getFullYear():anioActual;

/** Años disponibles, del más reciente al más antiguo (incluye siempre el actual). @param {Cita[]} citas @param {number} ahora */
export function aniosCuaderno(citas,ahora){
  const now=new Date(ahora).getFullYear();
  return [...new Set([now,...citas.filter(x=>x.savedAt).map(x=>anioDe(x.savedAt,now))])].sort((a,b)=>b-a);
}

/** Citas guardadas de un año, agrupadas por letra y ordenadas. @param {Cita[]} citas @param {number} anio @param {number} ahora */
export function porLetra(citas,anio,ahora){
  const now=new Date(ahora).getFullYear();
  const list=citas.filter(x=>x.savedAt&&anioDe(x.savedAt,now)===anio);
  /** @type {Record<string,Cita[]>} */const byL={};NB_LETTERS.forEach(l=>byL[l]=[]);
  list.slice().sort(compararCuaderno).forEach(x=>byL[nbInitial(x.author)].push(x));
  return {byL,total:list.length};
}

/** Páginas que ocupa una letra. @param {number} n citas @param {number} per citas por página */
export function paginasLetra(n,per){let np=Math.max(2,Math.ceil(n/per));if(np%2)np++;return np}

/** Ficha de una cita en el cuaderno (todos los textos ya listos para imprimir).
 * @param {Cita} x @param {string} code @param {number} anio */
export function fichaCita(x,code,anio){
  const d=typeof x.savedAt==='number'?new Date(x.savedAt):null;const ap=apa(x);
  return {code,id:x.id,author:nbAuthor(x.author),text:'“'+nbHyph(x.text)+'”',book:x.book||'',edYear:x.year||'s. f.',origYear:x.originalYear||'—',
    pubTr:[x.publisher,x.translator?'trad. '+x.translator:''].filter(Boolean).join(' · ')||'—',apaPre:ap.pre,apaTitle:ap.title,apaPost:ap.post,
    page:x.page||'—',topic:[x.topic,x.concept].filter(Boolean).join(' · '),read:d?`${d.getDate()} ${NB_MONTHS[d.getMonth()]} ${d.getFullYear()}`:String(anio),
    notes:nbHyph(x.understood||''),hasNotes:!!(x.understood||'').trim()};
}

/** Estructura completa del cuaderno.
 * @param {Cita[]} citas todas las citas (se filtran las guardadas del año)
 * @param {{anio?:number, porPagina?:number, espejo?:boolean, ahora:number}} op */
export function estructuraCuaderno(citas,{anio,porPagina=2,espejo=true,ahora}){
  const per=Number(porPagina)||2;const year=anio||new Date(ahora).getFullYear();
  const {byL,total}=porLetra(citas,year,ahora);
  /** @type {any[]} */const raw=[{kind:'cover'},{kind:'index'}];/** @type {Record<string,number>} */const start={};
  NB_LETTERS.forEach(l=>{
    const items=byL[l];const np=paginasLetra(items.length,per);start[l]=raw.length;
    for(let pi=0;pi<np;pi++){
      const entries=[];
      for(let s=0;s<per;s++){const idx=pi*per+s,x=items[idx],code=`${l}-${String(idx+1).padStart(2,'0')}`;entries.push(x?{filled:true,...fichaCita(x,code,year)}:{filled:false,code})}
      raw.push({kind:'letter',letter:l,sheetInLetter:Math.floor(pi/2)+1,pi,np,entries});
    }
  });
  const pages=raw.map((r,i)=>{
    const num=i+1,recto=num%2===1,side=recto?'anverso':'reverso',sheet=Math.ceil(num/2);
    return {...r,num,recto,side,sheet,tabsLeft:espejo&&!recto,padL:recto||!espejo?16:6,padR:recto||!espejo?6:16,
      foot:r.kind==='letter'?`${r.letter} · hoja ${r.sheetInLetter} · ${side} — hoja ${sheet} del cuaderno`:`Cuaderno de lectura ${year} · ${side}`,
      letterNote:r.kind==='letter'?(r.pi===0?`${byL[r.letter].length} ${byL[r.letter].length===1?'cita':'citas'} · ${r.np} págs.`:'continuación'):''};
  });
  const indice=NB_LETTERS.map(l=>({l,page:start[l]+1,count:byL[l].length?byL[l].length+(byL[l].length===1?' cita':' citas'):'—'}));
  return {anio:year,porPagina:per,espejo,total,isEmpty:!total,pages,start,byL,indice,
    resumen:`${total} ${total===1?'cita guardada':'citas guardadas'} en ${year}, ordenadas de la A a la Z por apellido.`};
}
