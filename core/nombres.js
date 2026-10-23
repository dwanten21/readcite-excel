// @ts-check
// Nombres de autor: apellido con partículas, autor invertido e inicial del cuaderno.
// Copia literal de ReadCite.dc.html (PARTICLES, nbSurname, nbAuthor, nbInitial). No cambiar sin actualizar tests/paridad.test.js.

export const PARTICLES = ['von', 'de', 'del', 'van', 'der', 'la', 'di', 'da'];

/** Apellido para ordenar: salta las partículas ("Ludwig von Mises" → "Mises"). @param {string} author */
export function nbSurname(author){const w=(author||'').trim().split(/\s+/).filter(Boolean);if(w.length<2)return w[0]||'';let i=w.findIndex((x,k)=>k>0&&PARTICLES.includes(x.toLowerCase()));if(i<0)i=w.length-1;while(i<w.length-1&&PARTICLES.includes(w[i].toLowerCase()))i++;return w.slice(i).join(' ')}

/** Autor invertido para el cuaderno ("Mises, Ludwig von"). @param {string} author */
export function nbAuthor(author){const sur=nbSurname(author);const rest=(author||'').slice(0,(author||'').lastIndexOf(sur)).trim();return rest?`${sur}, ${rest}`:sur}

/** Letra del cuaderno: primera letra A-Z del apellido, sin tildes; 'A' si no hay. @param {string} author */
export function nbInitial(author){const c=nbSurname(author).normalize('NFD').replace(/[̀-ͯ]/g,'').toUpperCase().match(/[A-Z]/);return c?c[0]:'A'}
