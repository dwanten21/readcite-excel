import { test } from 'node:test';
import assert from 'node:assert/strict';
import { limpiarTranscripcion, nombrePng, aplanarLuz } from '../core/ocr.js';
import { planTarjeta } from '../core/tarjeta.js';
import { nbHyphWord } from '../core/silabeo.js';

const medir = (t, f) => [...t].length * f.size * 0.5;

test('limpiarTranscripcion: une palabras partidas a final de línea y junta las líneas', () => {
  assert.equal(limpiarTranscripcion('La acción humana es con-\nducta consciente y\ndeliberada.'), 'La acción humana es conducta consciente y deliberada.');
  assert.equal(limpiarTranscripcion('cons-  \n  ciente'), 'consciente');
  assert.equal(limpiarTranscripcion('el anti-\nEstado'), 'el anti- Estado'); // mayúscula detrás: no se une
  assert.equal(limpiarTranscripcion('entre 1848-1849 y la guerra'), 'entre 1848-1849 y la guerra');
  assert.equal(limpiarTranscripcion('Dijo "basta" , y se fue .\r\n\r\nFin'), 'Dijo “basta”, y se fue. Fin');
  assert.equal(limpiarTranscripcion(''), '');
});

test('nombre del PNG como en la web', () => {
  assert.equal(nombrePng('Ludwig von Mises'), 'readcite-ludwig-von-mises.png');
  assert.equal(nombrePng('Frédéric Bastiat'), 'readcite-fr-d-ric-bastiat.png');
  assert.equal(nombrePng(''), 'readcite-cita.png');
});

const MISES = { text: 'La acción humana es conducta consciente y deliberada.', author: 'Ludwig von Mises', book: 'La acción humana', topic: 'Praxeología' };

test('tarjeta: póster 1200×1800 y cuadrada 1200×1200', () => {
  assert.deepEqual([planTarjeta({ ...MISES, ratio: 'poster' }, medir).W, planTarjeta({ ...MISES, ratio: 'poster' }, medir).H], [1200, 1800]);
  assert.equal(planTarjeta({ ...MISES, ratio: 'square' }, medir).H, 1200);
  assert.equal(planTarjeta({ ...MISES, ratio: 'raro' }, medir).H, 1800);
});

test('tarjeta: la cita se parte solo por los cortes de nbHyphWord y lleva guion visible', () => {
  const p = planTarjeta({ ...MISES, ratio: 'poster' }, (t, f) => [...t].length * f.size * 0.62);
  assert.ok(p.lineas.length >= 2);
  assert.ok(p.lineas.every(l => !l.includes('­')));
  const unido = p.lineas.map((l, i) => (l.endsWith('-') && i < p.lineas.length - 1 ? l.slice(0, -1) + '\u0000' : l + ' ')).join('').replace(/\u0000/g, '').trim();
  assert.equal(unido, '“' + MISES.text + '”');
  for (const [i, l] of p.lineas.entries()) {
    if (!l.endsWith('-') || i === p.lineas.length - 1) continue;
    const parte = l.slice(0, -1).split(' ').pop(), resto = p.lineas[i + 1].split(' ')[0].replace(/[.”]+$/, '');
    const cortes = nbHyphWord(parte + resto).split('­').map((_, k, a) => a.slice(0, k + 1).join(''));
    assert.ok(cortes.includes(parte), `${parte}|${resto}`);
  }
  // renglón azul bajo cada línea, tema en versalitas y logo
  assert.equal(p.ops.filter(o => o.t === 'rect' && o.color === '#a9c3db').length, p.lineas.length);
  assert.ok(p.ops.some(o => o.t === 'texto' && o.texto === 'PRAXEOLOGÍA'));
  assert.ok(p.ops.some(o => o.t === 'texto' && o.texto === 'Read') && p.ops.some(o => o.t === 'texto' && o.texto === 'Cite'));
});

test('tarjeta: cita larguísima en formato cuadrado baja de tamaño y no se sale', () => {
  const p = planTarjeta({ ...MISES, text: 'palabra '.repeat(200), ratio: 'square' }, medir);
  assert.ok(p.tam < 66);
  const textos = p.ops.filter(o => o.t === 'texto');
  assert.ok(textos.every(o => o.y < p.H));
  assert.ok(p.lineas.at(-1).endsWith('…'));
});

test('tarjeta sin datos: textos de ejemplo como la vista previa de la web', () => {
  const p = planTarjeta({ text: '', author: '', book: '' }, medir);
  const t = p.ops.filter(o => o.t === 'texto').map(o => o.texto);
  assert.ok(t.includes('SIN TEMA') && t.includes('Autor') && t.includes('Libro'));
  assert.ok(p.lineas.join(' ').includes('Tu cita textual'));
});

test('aplanarLuz: el papel en sombra queda blanco y la tinta sigue oscura', () => {
  // 128×1: mitad izquierda papel claro (200) con una letra (40), mitad derecha papel en sombra (100) con una letra (20)
  const ancho = 128, rgba = new Uint8ClampedArray(ancho * 4);
  for (let x = 0; x < ancho; x++) { const v = x < 64 ? (x === 10 ? 40 : 200) : (x === 110 ? 20 : 100); rgba.set([v, v, v, 255], 4 * x); }
  aplanarLuz(rgba, ancho, 1);
  assert.ok(rgba[4 * 0] > 245 && rgba[4 * 127] > 245, 'papel de ambas mitades ≈ blanco');
  assert.ok(rgba[4 * 10] < 70 && rgba[4 * 110] < 70, 'letras de ambas mitades oscuras');
  assert.equal(rgba[4 * 127 + 3], 255, 'no toca la opacidad');
});
