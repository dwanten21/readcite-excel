import { test } from 'node:test';
import assert from 'node:assert/strict';
import { compararCuaderno, porLetra, paginasLetra, estructuraCuaderno, aniosCuaderno, NB_LETTERS } from '../core/cuaderno.js';
import { nbInitial } from '../core/nombres.js';
import { visibleHyph } from '../core/silabeo.js';
import { web, omitirSinWeb } from './web.js';

const AHORA = new Date(2026, 9, 15, 12).getTime();
const guardadas = () => [...web.SEED].map((q, i) => ({ ...q, savedAt: AHORA - i * 86400000, createdAt: null, liked: false }));

test('orden A-Z: apellido sin partículas, luego libro, luego página', () => {
  const ids = guardadas().sort(compararCuaderno).map(q => `${q.author} · ${q.book}`);
  assert.deepEqual(ids, [
    'Frédéric Bastiat · El Estado', 'Milton Friedman · Capitalismo y libertad', 'Friedrich A. Hayek · Camino de servidumbre',
    'Hans-Hermann Hoppe · Monarquía, democracia y orden natural', 'Carl Menger · Principios de economía política',
    'Ludwig von Mises · La acción humana', 'Ludwig von Mises · Liberalismo', 'Ludwig von Mises · Socialismo',
    'Gustave de Molinari · De la producción de seguridad', 'Murray N. Rothbard · Anatomía del Estado',
    'Murray N. Rothbard · El hombre, la economía y el Estado', 'Adam Smith · La riqueza de las naciones']);
  const mismo = [{ author: 'A B', book: 'L', page: '20' }, { author: 'A B', book: 'L', page: '3' }, { author: 'A B', book: 'L', page: '' }];
  assert.deepEqual(mismo.sort(compararCuaderno).map(q => q.page), ['', '3', '20']);
});

test('paridad con la web: el mismo comparador y la misma letra para cada cita', omitirSinWeb, () => {
  const extra = [{ author: 'José Ángel Úbeda', book: 'Ñ', page: '2' }, { author: 'Ana Álvarez', book: 'B', page: '1' }, { author: 'ana alvarez', book: 'A', page: '1' }, { author: '', book: 'Z' }];
  const todas = [...guardadas(), ...extra];
  assert.deepEqual(todas.slice().sort(compararCuaderno), todas.slice().sort(web.nbCompare));
  for (const q of todas) assert.equal(nbInitial(q.author), web.nbInitial(q.author));
});

test('páginas por letra: mínimo 2 y siempre par', () => {
  assert.deepEqual([0, 1, 2, 4, 5, 6, 9].map(n => paginasLetra(n, 2)), [2, 2, 2, 2, 4, 4, 6]);
  assert.deepEqual([0, 3, 4, 7].map(n => paginasLetra(n, 3)), [2, 2, 2, 4]);
  assert.deepEqual([1, 3].map(n => paginasLetra(n, 1)), [2, 4]);
});

test('estructura: portada, índice, cada letra empieza en anverso (página impar)', () => {
  const nb = estructuraCuaderno(guardadas(), { ahora: AHORA, porPagina: 1 });
  assert.equal(nb.pages[0].kind, 'cover');
  assert.equal(nb.pages[1].kind, 'index');
  for (const l of NB_LETTERS) assert.equal((nb.start[l] + 1) % 2, 1, `letra ${l} empieza en página ${nb.start[l] + 1}`);
  assert.equal(nb.pages.length % 2, 0); // total par → sin hojas a medias
  assert.equal(nb.byL.M.length, 5);
  assert.equal(nb.indice.find(i => i.l === 'M').count, '5 citas');
  assert.equal(nb.indice.find(i => i.l === 'Q').count, '—');
  // M con 5 citas a 1 por página → 6 páginas; códigos M-01…M-06
  const m = nb.pages.filter(p => p.letter === 'M');
  assert.equal(m.length, 6);
  assert.deepEqual(m.flatMap(p => p.entries.map(e => e.code)), ['M-01', 'M-02', 'M-03', 'M-04', 'M-05', 'M-06']);
  assert.equal(m[5].entries[0].filled, false);
  assert.equal(m[0].letterNote, '5 citas · 6 págs.');
  assert.equal(m[1].letterNote, 'continuación');
});

test('estructura: pestañas espejo, márgenes 16/6 mm y pie', () => {
  const nb = estructuraCuaderno(guardadas(), { ahora: AHORA });
  const [p3, p4] = [nb.pages[2], nb.pages[3]];
  assert.deepEqual([p3.recto, p3.tabsLeft, p3.padL, p3.padR], [true, false, 16, 6]);
  assert.deepEqual([p4.recto, p4.tabsLeft, p4.padL, p4.padR], [false, true, 6, 16]);
  assert.equal(p3.foot, 'A · hoja 1 · anverso — hoja 2 del cuaderno');
  assert.equal(nb.pages[0].foot, 'Cuaderno de lectura 2026 · anverso');
  const sinEspejo = estructuraCuaderno(guardadas(), { ahora: AHORA, espejo: false });
  assert.deepEqual([sinEspejo.pages[3].tabsLeft, sinEspejo.pages[3].padL], [false, 16]);
});

test('ficha: autor invertido, cita silabeada, APA y fecha de lectura', () => {
  const nb = estructuraCuaderno(guardadas(), { ahora: AHORA });
  const f = nb.pages.flatMap(p => p.entries || []).find(e => e.id === 's1');
  assert.equal(f.code, 'M-02'); // tras Menger
  assert.equal(f.author, 'Mises, Ludwig von');
  assert.equal(visibleHyph(f.text), '“La ac-ción hu-ma-na es con-duc-ta cons-cien-te y de-li-be-ra-da.”');
  assert.equal(f.apaPre, 'von Mises, L. (2011). ');
  assert.equal(f.read, '15 oct 2026');
  assert.equal(f.pubTr, 'Unión Editorial · trad. J. Reig Albiol');
});

test('filtro por año y años disponibles', () => {
  const citas = guardadas();
  citas[0].savedAt = new Date(2024, 2, 1).getTime();
  assert.deepEqual(aniosCuaderno(citas, AHORA), [2026, 2024]);
  assert.equal(porLetra(citas, 2024, AHORA).total, 1);
  assert.equal(porLetra(citas, 2026, AHORA).total, 11);
});

test('cuaderno vacío: 2 + 26×2 = 54 páginas, sin errores', () => {
  const nb = estructuraCuaderno([], { ahora: AHORA });
  assert.equal(nb.isEmpty, true);
  assert.equal(nb.pages.length, 54);
  assert.equal(nb.resumen, '0 citas guardadas en 2026, ordenadas de la A a la Z por apellido.');
});
