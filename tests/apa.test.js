import { test } from 'node:test';
import assert from 'node:assert/strict';
import { apaName, apa, apaText } from '../core/apa.js';
import { web, omitirSinWeb } from './web.js';

test('apaName: partículas, iniciales, guiones y casos límite', () => {
  assert.equal(apaName('Ludwig von Mises'), 'von Mises, L.');
  assert.equal(apaName('Gustave de Molinari'), 'de Molinari, G.');
  assert.equal(apaName('Friedrich A. Hayek'), 'Hayek, F. A.');
  assert.equal(apaName('Hans-Hermann Hoppe'), 'Hoppe, H.-H.');
  assert.equal(apaName('Platón'), 'Platón');
  assert.equal(apaName('  '), '');
  assert.equal(apaName(undefined), '');
});

test('apaText: traductor, año original distinto y valores por defecto', () => {
  assert.equal(apaText({ author: 'Friedrich A. Hayek', book: 'Camino de servidumbre', year: '2008', originalYear: '1944', publisher: 'Unión Editorial', translator: 'José Vergara' }),
    'Hayek, F. A. (2008). Camino de servidumbre (José Vergara, Trad.). Unión Editorial. (Obra original publicada en 1944)');
  assert.equal(apaText({ author: 'Ana López', book: 'X', year: '2020', originalYear: '2020' }), 'López, A. (2020). X. ');
  assert.equal(apaText({}), 'Autor (s. f.). Título. ');
  assert.equal(apaText({ author: 'Ana López', originalYear: '1990' }), 'López, A. (1990). Título. ');
});

test('apa: devuelve el título aparte (va en cursiva)', () => {
  const a = apa({ author: 'Adam Smith', book: 'La riqueza de las naciones', year: '2011' });
  assert.deepEqual(a, { pre: 'Smith, A. (2011). ', title: 'La riqueza de las naciones', post: '. ' });
});

test('paridad con la web: APA idéntica carácter a carácter (12 citas de ejemplo + casos límite)', omitirSinWeb, () => {
  const casos = [...web.SEED, {}, { author: 'Platón' }, { author: 'Hans-Hermann Hoppe', year: '2004', originalYear: '2001', translator: 'J. M.' },
    { author: 'María de la Cruz', book: 'Ñ', publisher: 'Ed.' }, { author: 'Jean-Baptiste Say', originalYear: '1803' }];
  for (const q of casos) {
    assert.equal(apaText(q), web.apaText(q), JSON.stringify(q));
    assert.equal(apaName(q.author), web.apaName(q.author));
  }
});
