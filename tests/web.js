// Datos y funciones ORIGINALES de la web de ReadCite para los tests.
// - Si existe ../ReadCite.dc.html (copia del autor), carga sus funciones en un sandbox para los tests de paridad.
// - Si no (p. ej. en el repositorio público), usa los datos de ejemplo de tests/datos/ejemplo-web.json y los
//   tests de paridad con la web se saltan (`omitirSinWeb`).
import { readFileSync, existsSync } from 'node:fs';
import vm from 'node:vm';

const RUTA = new URL('../ReadCite.dc.html', import.meta.url);
export const tieneWeb = existsSync(RUTA);
/** Opción para test(): salta los tests de paridad si no está la web original. */
export const omitirSinWeb = { skip: tieneWeb ? false : 'no está ReadCite.dc.html (repositorio público)' };

function cargarWeb() {
  const html = readFileSync(RUTA, 'utf8');
  const tramo = (desde, hasta) => {
    const a = html.indexOf(desde), b = html.indexOf(hasta, a);
    if (a < 0 || b < 0) throw new Error(`No encuentro el tramo «${desde}» en ReadCite.dc.html`);
    return html.slice(a, b);
  };
  // Comparador del cuaderno tal cual está en nbVals.
  const cmp = html.match(/list\.slice\(\)\.sort\((\(a,b\)=>nbSurname[^\n]*?)\)\.forEach\(x=>byL\[nbInitial/);
  if (!cmp) throw new Error('No encuentro el comparador del cuaderno en ReadCite.dc.html');
  const codigo = tramo('const KEY=', 'function load(){') + tramo('function syllables', 'function fillLine') +
    `;const nbCompare=${cmp[1]};` +
    ';({apaName,apa,apaText,nbHyphWord,nbHyph,nbSurname,nbAuthor,nbInitial,syllables,nbCompare,SEED,SEED_PAGES,SEED_LIB,ICARO_INTERVALS,PARTICLES})';
  return vm.runInNewContext(codigo, { Date, Math, JSON, String, Array, Object });
}

export const web = tieneWeb ? cargarWeb() : JSON.parse(readFileSync(new URL('./datos/ejemplo-web.json', import.meta.url), 'utf8'));
