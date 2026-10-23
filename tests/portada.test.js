import { test } from 'node:test';
import assert from 'node:assert/strict';
import { citaDestacada, planInicio, COLUMNAS_PT } from '../core/portada.js';
import { filaACita } from '../core/datos.js';
import { web } from './web.js';

const medir = (t, f) => [...t].length * f.size * 0.5;
const AHORA = new Date(2026, 9, 15, 12).getTime();
const citas = () => [...web.SEED].map((q, i) => ({ ...filaACita([], []), ...q, createdAt: AHORA - (i + 1) * 86400000 }));

test('cita de la portada: la más reciente, sea de quien sea', () => {
  const c = citas();
  c[7].createdAt = AHORA; // s8 (Bastiat) pasa a ser la más reciente
  assert.deepEqual(citaDestacada(c), { cita: c[7], kicker: 'Cita más reciente' });
  c[2].createdAt = AHORA + 1000;
  assert.equal(citaDestacada(c).cita.id, 's3');
  // sin fechas: la de más likes
  const sinFecha = citas().map(x => ({ ...x, createdAt: null }));
  assert.equal(citaDestacada(sinFecha).cita.id, 's8');
  assert.equal(citaDestacada([]), null);
});

test('plan de la hoja Inicio: kicker, cita en renglones, autor y botones enlazados', () => {
  const c = citas(); c[7].createdAt = AHORA;
  const p = planInicio(citaDestacada(c), medir);
  const textos = p.bloques.map(b => b.v);
  assert.ok(textos.includes('Cita más reciente · Estado'));
  const renglones = p.bloques.filter(b => b.e.italic).map(b => b.v);
  assert.equal(renglones.join(' '), '“El Estado es la gran ficción a través de la cual todo el mundo se esfuerza en vivir a expensas de todo el mundo.”');
  assert.ok(renglones.length <= 4);
  assert.ok(textos.includes('Frédéric Bastiat — El Estado'));
  const abrir = p.bloques.filter(b => b.v === 'Abrir ›');
  assert.equal(abrir.length, 2);
  assert.ok(abrir.every(b => b.link?.cita));
  assert.equal(p.bloques.find(b => b.v === 'Ver reverso').link.fila, p.reverso);
  assert.equal(p.bloques.find(b => b.v === 'Ver frente').link.fila, p.frente);
  assert.ok(p.reverso > p.frente);
  // reverso: lo que entendí, por qué y ubicación
  assert.ok(textos.includes('LO QUE ENTENDÍ'));
  assert.ok(textos.includes('POR QUÉ ME LLAMÓ LA ATENCIÓN'));
  assert.ok(textos.includes('Frédéric Bastiat — El Estado · p. 9'));
  // geometría: todo dentro de la rejilla
  for (const b of p.bloques) {
    assert.ok(b.c1 >= 0 && b.c2 < COLUMNAS_PT.length && b.c1 <= b.c2);
    assert.ok(b.f1 >= 0 && b.f2 < p.altosPt.length && b.f1 <= b.f2);
  }
});

test('plan de la hoja Inicio sin citas: aviso y sin botón «Abrir»', () => {
  const p = planInicio(null, medir);
  assert.ok(p.bloques.some(b => b.v === 'Aún no hay citas.'));
  assert.ok(!p.bloques.some(b => b.v === 'Abrir ›' || b.v === 'Ver reverso'));
});

test('una cita muy larga reduce el tamaño de letra', () => {
  const c = citas(); const larga = { ...c[0], text: 'palabra '.repeat(80).trim(), createdAt: AHORA };
  const p = planInicio(citaDestacada([larga]), medir);
  const r = p.bloques.filter(b => b.e.italic);
  assert.ok(r[0].e.size < 30);
});
