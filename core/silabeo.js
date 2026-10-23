// @ts-check
// Silabeo español. Copia literal de ReadCite.dc.html.
// - nbHyphWord / nbHyph: guiones blandos (U+00AD) para el cuaderno. No parte palabras de menos de 5 letras
//   ni deja fragmentos de 1 letra al principio o al final ("poesía" → "po-e-sía", "aún" intacta).
// - syllables: sílabas completas para partir palabras en el PNG ("poesía" → po-e-sí-a).

export const NB_VOW='aeiouáéíóúü',NB_STRONG='aeoáéíóú';
export const NB_INSEP=['pr','br','tr','dr','cr','gr','fr','kr','pl','bl','cl','gl','fl','kl','ch','ll','rr'];

/** @param {string} w */
export function nbHyphWord(w){
  if(w.length<5)return w;const s=w.toLowerCase(),n=s.length;
  /** @param {number} i */
  const isV=i=>{const c=s[i];if(!NB_VOW.includes(c))return false;if(c==='u'&&i>0&&(s[i-1]==='q'||s[i-1]==='g')&&'eéií'.includes(s[i+1]||''))return false;return true};
  /** @type {number[][]} */
  const nuc=[];let i=0;
  while(i<n){if(!isV(i)){i++;continue}let st=i;while(i+1<n&&isV(i+1)){if(NB_STRONG.includes(s[i])&&NB_STRONG.includes(s[i+1])){nuc.push([st,i]);st=i+1}i++}nuc.push([st,i]);i++}
  /** @type {number[]} */
  const cuts=[];
  for(let k=0;k<nuc.length-1;k++){
    const a=nuc[k][1],b=nuc[k+1][0];const cons=s.slice(a+1,b);
    if(/[qg]u$/.test(cons)){const at2=b-2;if(at2>=2&&n-at2>=2)cuts.push(at2);continue}
    const m=cons.length;let at;
    if(m===0)at=b;else if(m===1)at=b-1;else if(m===2)at=NB_INSEP.includes(cons)?b-2:b-1;else if(m===3)at=NB_INSEP.includes(cons.slice(1))?b-2:b-1;else at=b-2;
    if(at>=2&&n-at>=2)cuts.push(at);
  }
  let out='',last=0;cuts.forEach(c=>{out+=w.slice(last,c)+'­';last=c});return out+w.slice(last);
}

/** Aplica nbHyphWord a cada palabra de un texto. @param {string} t */
export const nbHyph=t=>(t||'').replace(/[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+/g,nbHyphWord);

/** @param {string} w @returns {string[]} */
export function syllables(w){
  const V='aeiouáéíóúü';/** @param {string} c */const isV=c=>V.includes(c.toLowerCase());
  const clusters=['br','bl','cl','cr','dr','fl','fr','gl','gr','pl','pr','tr','ch','ll','rr'];
  /** @type {string[]} */
  const out=[];let cur='';
  for(let i=0;i<w.length;i++){
    cur+=w[i];if(!isV(w[i]))continue;
    let j=i+1,cons='';while(j<w.length&&!isV(w[j])){cons+=w[j];j++}
    if(j>=w.length){cur+=cons;i=w.length;break}
    if(cons.length===0){const a=w[i].toLowerCase(),b=w[j].toLowerCase();if('iuü'.includes(a)||'iuü'.includes(b))continue;out.push(cur);cur='';}
    else if(cons.length===1){out.push(cur);cur='';}
    else{const keep=clusters.includes(cons.slice(-2).toLowerCase())?2:1;const take=cons.length-keep;cur+=cons.slice(0,take);out.push(cur);cur='';i+=take;}
  }
  if(cur)out.push(cur);return out;
}

/** Para mostrar/depurar: guiones blandos → '-'. @param {string} t */
export const visibleHyph=t=>t.replace(/­/g,'-');
