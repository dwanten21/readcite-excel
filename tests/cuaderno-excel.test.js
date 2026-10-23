import { test } from 'node:test';
import assert from 'node:assert/strict';
import { generarCuaderno, areaImpresion, HOJA, MARCA } from '../excel/cuaderno.js';
import { FILAS_PAGINA } from '../core/maquetado.js';
import { CONFIG_DEF } from '../core/datos.js';
import { web } from './web.js';

const medir = (t, f) => [...t].length * f.size * 0.5;
const AHORA = new Date(2026, 9, 15, 12).getTime();
const citas = () => [...web.SEED].map((q, i) => ({ ...q, savedAt: AHORA - i * 86400000, createdAt: null, liked: false }));

/** Libro falso: registra cada operación de Office.js sobre la hoja Cuaderno. */
function libro({ existente = null } = {}) {
  const log = []; let syncs = 0; let hoja = existente; // existente: {marca}
  const grabadora = (ruta) => new Proxy(function () {}, {
    get(_, k) {
      if (k === 'then') return undefined;
      if (k === 'isNullObject') return false;
      return grabadora(`${ruta}.${String(k)}`);
    },
    set(_, k, v) { log.push({ op: 'set', ruta: `${ruta}.${String(k)}`, v }); return true; },
    apply(_, __, args) {
      log.push({ op: 'call', ruta, args });
      const m = ruta.match(/getRange$/); if (m) return grabadora(`${ruta}(${args[0]})`);
      if (/getItem$/.test(ruta)) return grabadora(`${ruta}(${args[0]})`);
      return grabadora(`${ruta}()`);
    },
  });
  const ctx = {
    workbook: {
      worksheets: {
        getItemOrNullObject(n) {
          assert.equal(n, HOJA);
          const h = hoja;
          return { isNullObject: !h, load() { return this; }, getRange: () => ({ load() { return this; }, values: [[h?.marca ?? '']] }),
            delete() { log.push({ op: 'delete' }); hoja = null; }, pageLayout: { setPrintArea: a => log.push({ op: 'area', a }) } };
        },
        add(n) { log.push({ op: 'add', n }); hoja = { marca: MARCA }; return grabadora('ws'); },
      },
    },
    async sync() { syncs++; },
  };
  return { ctx, log, get syncs() { return syncs; } };
}

const modelo = (c = citas(), config = {}) => ({ citas: c, config: { ...CONFIG_DEF, ...config } });

test('genera la hoja: rejilla en mm, valores por lotes, combinaciones, enlaces, A5 y saltos por página', async () => {
  const L = libro();
  const r = await generarCuaderno(L.ctx, modelo(), { ahora: AHORA, medir });
  assert.equal(r.paginas % 2, 0);
  assert.equal(r.citas, 12);
  const sets = L.log.filter(x => x.op === 'set');
  assert.equal(sets.filter(x => x.ruta.endsWith('format.columnWidth')).length, 10);
  assert.ok(sets.some(x => x.ruta === 'ws.getRange(L1).values' && x.v[0][0] === MARCA));
  assert.ok(sets.some(x => x.ruta.endsWith('pageLayout.paperSize') && x.v === 'A5'));
  for (const m of ['leftMargin', 'rightMargin', 'topMargin', 'bottomMargin']) assert.ok(sets.some(x => x.ruta.endsWith(`pageLayout.${m}`) && x.v === 0), m);
  const saltos = L.log.filter(x => x.op === 'call' && x.ruta.endsWith('horizontalPageBreaks.add'));
  assert.equal(saltos.length, r.paginas - 1);
  assert.deepEqual(saltos.map(s => s.args[0]).slice(0, 2), [`A${FILAS_PAGINA + 1}`, `A${2 * FILAS_PAGINA + 1}`]);
  assert.ok(L.log.some(x => x.op === 'call' && x.ruta.endsWith('pageLayout.setPrintArea') && x.args[0] === `A1:J${r.paginas * FILAS_PAGINA}`));
  // valores escritos en bloques de páginas (nunca celda a celda)
  const valores = sets.filter(x => /getRange\(A\d+:J\d+\)\.values$/.test(x.ruta));
  assert.equal(valores.length, Math.ceil(r.paginas / 6));
  const todos = valores.flatMap(x => x.v.flat());
  assert.ok(todos.some(v => typeof v === 'string' && v.startsWith('“La ac')));
  assert.ok(!todos.some(v => typeof v === 'string' && v.includes('­')));
  // enlaces solo en el índice (26 letras × 2: letra y número de página), a la fila de inicio de cada letra
  const enlaces = sets.filter(x => x.ruta.endsWith('.hyperlink'));
  assert.equal(enlaces.length, 52);
  assert.ok(enlaces.every(x => /^'Cuaderno'!A\d+$/.test(x.v.documentReference) && x.v.textToDisplay));
  // formatos agrupados: pocas llamadas a getRanges, y las pestañas no se combinan
  const areas = L.log.filter(x => x.op === 'call' && x.ruta.endsWith('getRanges'));
  assert.ok(areas.length < 40 * Math.ceil(r.paginas / 6), `${areas.length} getRanges`);
  assert.ok(areas.every(x => x.args[0].split(',').length <= 40));
  const merges = L.log.filter(x => x.op === 'call' && x.ruta.endsWith('.merge'));
  assert.ok(merges.length < 35 * r.paginas, `${merges.length} combinaciones`);
  assert.ok(!merges.some(x => /getRange\((A|J)\d+:\1\d+\)\.merge$/.test(x.ruta)), 'las pestañas (columna A o J) no se combinan');
  assert.ok(L.syncs <= 3 + Math.ceil(r.paginas / 6) + 1, `${L.syncs} sync`);
});

test('regenerar: borra solo una hoja Cuaderno creada por ReadCite; si no lleva la marca, para', async () => {
  const propio = libro({ existente: { marca: MARCA } });
  await generarCuaderno(propio.ctx, modelo(), { ahora: AHORA, medir });
  assert.deepEqual(propio.log.filter(x => x.op === 'delete' || x.op === 'add').map(x => x.op), ['delete', 'add']);
  const ajeno = libro({ existente: { marca: 'mis notas' } });
  await assert.rejects(generarCuaderno(ajeno.ctx, modelo(), { ahora: AHORA, medir }), /Ya existe una hoja «Cuaderno» que no creó ReadCite/);
  assert.equal(ajeno.log.filter(x => x.op === 'delete' || x.op === 'add').length, 0);
});

test('cuaderno vacío (tabla de citas vacía): 54 páginas de plantillas', async () => {
  const L = libro();
  const r = await generarCuaderno(L.ctx, modelo([]), { ahora: AHORA, medir });
  assert.deepEqual([r.paginas, r.citas], [54, 0]);
});

test('opciones desde Config: año, citas por página y espejo', async () => {
  const L = libro();
  const r = await generarCuaderno(L.ctx, modelo(citas(), { cuaderno_anio: '2025', cuaderno_citas_por_pagina: 1 }), { ahora: AHORA, medir });
  assert.deepEqual([r.anio, r.citas, r.paginas], [2025, 0, 54]);
});

test('área de impresión por letra: empieza en anverso y ocupa páginas pares', async () => {
  const L = libro({ existente: { marca: MARCA } });
  const r = await areaImpresion(L.ctx, modelo(), 'M', AHORA);
  const [, a, b] = r.match(/^A(\d+):J(\d+)$/).map(Number);
  assert.equal((a - 1) % FILAS_PAGINA, 0);
  assert.equal(((a - 1) / FILAS_PAGINA) % 2, 0);
  assert.equal(((b - a + 1) / FILAS_PAGINA) % 2, 0);
  assert.match(await areaImpresion(L.ctx, modelo(), null, AHORA), /^A1:J\d+$/);
  await assert.rejects(areaImpresion(libro().ctx, modelo(), 'M', AHORA), /Aún no hay cuaderno/);
});
