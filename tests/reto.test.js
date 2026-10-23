import { test } from 'node:test';
import assert from 'node:assert/strict';
import { ICARO_INTERVALS, daysSince, paginasDia, diasRestantes, resumenReto, datosLibro, citasEnPlan, tocaRepaso, sesionesHoy } from '../core/reto.js';
import { web } from './web.js';

const DIA = 86400000;
const AHORA = new Date(2026, 9, 15, 12, 0).getTime();

test('intervalos Icaro idénticos a la web', () => {
  assert.deepEqual(ICARO_INTERVALS, [1, 3, 7, 14, 30, 90]);
  assert.deepEqual(ICARO_INTERVALS, [...web.ICARO_INTERVALS]);
});

test('días restantes, páginas/día y etiquetas', () => {
  assert.equal(diasRestantes(320, 100, 20), 11);
  assert.equal(diasRestantes(320, 400, 20), 0);
  assert.equal(paginasDia(''), 20);
  assert.equal(paginasDia(0), 20);
  assert.equal(paginasDia('35'), 35);
  const l = datosLibro({ titulo: 'X', paginas: 100, leidas: 500, estado: 'quiero' }, 20);
  assert.equal(l.leidas, 100);
  assert.equal(l.diasLabel, '0 días');
  assert.equal(datosLibro({ titulo: 'Y', paginas: 21, leidas: 1, estado: 'quiero' }, 20).diasLabel, '1 día');
  assert.equal(datosLibro({ titulo: 'Z', paginas: 0, leidas: 0, estado: 'quiero' }, 20).enReto, false);
});

test('resumen del reto: solo libros en "quiero"', () => {
  const libros = [
    { titulo: 'A', paginas: 816, leidas: 16, estado: 'quiero' },
    { titulo: 'B', paginas: 600, leidas: 0, estado: 'quiero' },
    { titulo: 'C', paginas: 1040, leidas: 900, estado: 'leyendo' },
  ];
  const r = resumenReto(libros, 20);
  assert.equal(r.totalPages, 1416);
  assert.equal(r.readPages, 16);
  assert.equal(r.pctLabel, '1%');
  assert.equal(r.days, 70);
  assert.equal(r.daysLabel, '2.3 meses (70 días)');
  assert.equal(r.booksLabel, '2 libros');
  assert.equal(resumenReto([{ titulo: 'A', paginas: 100, leidas: 50, estado: 'quiero' }], 25).daysLabel, '2 días');
  const vacio = resumenReto([], '');
  assert.deepEqual([vacio.totalPages, vacio.pct, vacio.days, vacio.daysLabel, vacio.booksLabel], [0, 0, 0, '0 días', '0 libros']);
});

test('reto con los datos de ejemplo de la web (estado + páginas)', () => {
  const libros = Object.entries(web.SEED_LIB).map(([titulo, estado]) => ({ titulo, estado, paginas: web.SEED_PAGES[titulo] || 0, leidas: 0 }));
  const r = resumenReto(libros, 20);
  // quiero: La riqueza (816) + Monarquía (420) + El hombre (1200) + Socialismo (600)
  assert.equal(r.totalPages, 3036);
  assert.equal(r.days, 152);
  assert.equal(r.daysLabel, '5.1 meses (152 días)');
});

test('Icaro: repaso cuando los días desde la creación coinciden con un intervalo', () => {
  for (const d of ICARO_INTERVALS) assert.equal(tocaRepaso({ id: 'x', book: 'A', createdAt: AHORA - d * DIA }, AHORA), true, `día ${d}`);
  for (const d of [0, 2, 5, 15, 91]) assert.equal(tocaRepaso({ id: 'x', book: 'A', createdAt: AHORA - d * DIA }, AHORA), false, `día ${d}`);
  assert.equal(tocaRepaso({ id: 'x', book: 'A', createdAt: null }, AHORA), true); // sin fecha → 1 día, como la web
  assert.equal(daysSince(AHORA - 3 * DIA - 1000, AHORA), 3);
});

test('Icaro: plan = citas de libros en "quiero" guardadas o propias', () => {
  const libros = [{ titulo: 'A', paginas: 100, leidas: 0, estado: 'quiero' }, { titulo: 'B', paginas: 100, leidas: 0, estado: 'leyendo' }];
  const citas = [
    { id: '1', book: 'A', user: 'Ana', savedAt: AHORA, createdAt: AHORA - 7 * DIA },
    { id: '2', book: 'A', user: 'Tú', savedAt: null, createdAt: AHORA - 2 * DIA },
    { id: '3', book: 'A', user: 'Otro', savedAt: null, createdAt: AHORA - 1 * DIA },
    { id: '4', book: 'B', user: 'Tú', savedAt: AHORA, createdAt: AHORA - 1 * DIA },
  ];
  assert.deepEqual(citasEnPlan(citas, libros, 'Tú').map(c => c.id), ['1', '2']);
  const s = sesionesHoy(citas, libros, 'Tú', AHORA);
  assert.equal(s.quotesInPlan, 2);
  assert.equal(s.dueToday, 1);
});
