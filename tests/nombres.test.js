import { test } from 'node:test';
import assert from 'node:assert/strict';
import { nbSurname, nbAuthor, nbInitial } from '../core/nombres.js';
import { web, omitirSinWeb } from './web.js';

test('nbSurname: salta partículas (a diferencia de apaName)', () => {
  assert.equal(nbSurname('Ludwig von Mises'), 'Mises');
  assert.equal(nbSurname('Gustave de Molinari'), 'Molinari');
  assert.equal(nbSurname('María de la Cruz'), 'Cruz');
  assert.equal(nbSurname('Murray N. Rothbard'), 'Rothbard');
  assert.equal(nbSurname('Platón'), 'Platón');
  assert.equal(nbSurname(''), '');
});

test('nbAuthor: autor invertido', () => {
  assert.equal(nbAuthor('Ludwig von Mises'), 'Mises, Ludwig von');
  assert.equal(nbAuthor('Adam Smith'), 'Smith, Adam');
  assert.equal(nbAuthor('Platón'), 'Platón');
});

test('nbInitial: letra sin tildes; A si no hay letra', () => {
  assert.equal(nbInitial('Ludwig von Mises'), 'M');
  assert.equal(nbInitial('José Ángel Úbeda'), 'U');
  assert.equal(nbInitial('Carl Ñandú'), 'N');
  assert.equal(nbInitial(''), 'A');
  assert.equal(nbInitial('123'), 'A');
});

test('paridad con la web: nbSurname, nbAuthor y nbInitial', omitirSinWeb, () => {
  const autores = [...new Set(web.SEED.map(q => q.author)), 'María de la Cruz', 'José Ángel Úbeda', 'Platón', '', 'Juan van der Berg', 'Ana Da Silva'];
  for (const a of autores) {
    assert.equal(nbSurname(a), web.nbSurname(a), a);
    assert.equal(nbAuthor(a), web.nbAuthor(a), a);
    assert.equal(nbInitial(a), web.nbInitial(a), a);
  }
});
