import { test } from 'node:test';
import assert from 'node:assert/strict';
import { nbHyphWord, nbHyph, syllables, visibleHyph } from '../core/silabeo.js';
import { web, omitirSinWeb } from './web.js';

const SH = '­';
const g = w => visibleHyph(nbHyphWord(w));

// Salida esperada de nbHyphWord (preservado tal cual). Donde difiere de la división silábica "de diccionario"
// es porque el algoritmo no deja fragmentos de 1 letra ni parte palabras de menos de 5 letras.
const CASOS = {
  'acción': 'ac-ción', 'constitución': 'cons-ti-tu-ción', 'nación': 'na-ción', 'guerra': 'gue-rra',
  'poesía': 'po-e-sía', 'cooperación': 'co-o-pe-ra-ción', 'institución': 'ins-ti-tu-ción',
  'abstracción': 'abs-trac-ción', 'exhibición': 'ex-hi-bi-ción', 'constelación': 'cons-te-la-ción',
  'reyerta': 're-yer-ta', 'caótico': 'ca-ó-ti-co', 'caída': 'ca-í-da', 'reúne': 're-ú-ne',
  'aún': 'aún', 'cooperar': 'co-o-pe-rar', 'leer': 'leer', 'creer': 'cre-er', 'aliado': 'alia-do',
};

test('nbHyphWord: las 19 palabras de prueba del brief (salida preservada)', () => {
  for (const [w, esperado] of Object.entries(CASOS)) assert.equal(g(w), esperado, w);
});

test('nbHyphWord: usa guiones blandos U+00AD, no guiones normales', () => {
  assert.equal(nbHyphWord('nación'), 'na' + SH + 'ción');
  assert.ok(!nbHyphWord('constitución').includes('-'));
});

test('nbHyphWord: palabras cortas y fragmentos de una letra', () => {
  for (const w of ['sol', 'casa', 'aún', 'leer']) assert.equal(nbHyphWord(w), w);
  assert.ok(!g('poesía').endsWith('-a'));
  assert.ok(!g('aliado').startsWith('a-'));
});

test('nbHyphWord: qu/gu + e/i no son diptongo, rr/ll/ch inseparables, conserva mayúsculas', () => {
  assert.equal(g('guerrero'), 'gue-rre-ro');
  assert.equal(g('quinientos'), 'qui-nien-tos');
  assert.equal(g('caballero'), 'ca-ba-lle-ro');
  assert.equal(g('muchacho'), 'mu-cha-cho');
  assert.equal(g('Constitución'), 'Cons-ti-tu-ción');
});

test('nbHyph: silabea solo palabras y respeta signos y espacios', () => {
  assert.equal(visibleHyph(nbHyph('La acción humana, «conducta»')), 'La ac-ción hu-ma-na, «con-duc-ta»');
  assert.equal(nbHyph(''), '');
  assert.equal(nbHyph(undefined), '');
});

test('syllables: sílabas completas (se usa en el PNG)', () => {
  assert.deepEqual(syllables('poesía'), ['po', 'e', 'sí', 'a']);
  assert.deepEqual(syllables('constitución'), ['cons', 'ti', 'tu', 'ción']);
  assert.deepEqual(syllables('aún'), ['a', 'ún']);
});

test('paridad con la web: nbHyph y syllables idénticos en todo el texto de las citas de ejemplo', omitirSinWeb, () => {
  const palabras = new Set();
  for (const q of web.SEED) for (const t of [q.text, q.understood, q.why]) (t.match(/[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+/g) || []).forEach(w => palabras.add(w));
  Object.keys(CASOS).forEach(w => palabras.add(w));
  assert.ok(palabras.size > 300);
  for (const w of palabras) {
    assert.equal(nbHyphWord(w), web.nbHyphWord(w), w);
    assert.deepEqual(syllables(w), [...web.syllables(w)], w);
  }
  for (const q of web.SEED) assert.equal(nbHyph(q.understood), web.nbHyph(q.understood));
});
