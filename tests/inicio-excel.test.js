import { test } from 'node:test';
import assert from 'node:assert/strict';
import { generarInicio, HOJA_INICIO, MARCA_INICIO } from '../excel/inicio.js';
import { filaACita, CONFIG_DEF } from '../core/datos.js';
import { web } from './web.js';

const medir = (t, f) => [...t].length * f.size * 0.5;
const AHORA = new Date(2026, 9, 15, 12).getTime();

/** Libro falso: tblCitas con ids y una hoja Inicio opcional; registra las operaciones de la hoja nueva. */
function libro({ existente = null, ids = ['s1', 's8'] } = {}) {
  const log = []; let hoja = existente;
  const grab = (ruta) => new Proxy(function () {}, {
    get(_, k) { if (k === 'then') return undefined; return grab(`${ruta}.${String(k)}`); },
    set(_, k, v) { log.push({ op: 'set', ruta: `${ruta}.${String(k)}`, v }); return true; },
    apply(_, __, a) { log.push({ op: 'call', ruta, a }); return grab(`${ruta}(${a[0] ?? ''})`); },
  });
  const tabla = {
    isNullObject: false, load() { return this; },
    getHeaderRowRange: () => ({ load() { return this; }, values: [['id', 'texto']] }),
    getDataBodyRange: () => ({ load() { return this; }, values: ids.map(i => [i, 't']), formulas: ids.map(i => [i, 't']),
      getRow: i => ({ load() { return this; }, address: `Citas!A${i + 2}:X${i + 2}` }) }),
  };
  const ctx = {
    workbook: {
      tables: { getItemOrNullObject: n => (n === 'tblCitas' ? tabla : { isNullObject: true, load() { return this; } }) },
      worksheets: {
        getItemOrNullObject: () => ({ isNullObject: !hoja, load() { return this; }, getRange: () => ({ load() { return this; }, values: [[hoja?.marca ?? '']] }), delete() { log.push({ op: 'delete' }); } }),
        add(n) { log.push({ op: 'add', n }); return grab('ws'); },
      },
    },
    async sync() {},
  };
  return { ctx, log };
}
const modelo = () => ({ config: CONFIG_DEF, citas: [...web.SEED].map((q, i) => ({ ...filaACita([], []), ...q, createdAt: AHORA - i * 86400000 })) });

test('crea la hoja Inicio primera, con la cita más reciente y «Abrir» enlazado a su fila de Citas', async () => {
  const L = libro();
  const r = await generarInicio(L.ctx, modelo(), { medir });
  assert.deepEqual(r, { cita: 's1', kicker: 'Cita más reciente' });
  const sets = L.log.filter(x => x.op === 'set');
  assert.ok(sets.some(x => x.ruta === 'ws.position' && x.v === 0));
  assert.ok(sets.some(x => x.ruta === 'ws.getRange(Z1).values' && x.v[0][0] === MARCA_INICIO));
  assert.ok(sets.some(x => x.ruta.endsWith('.values') && x.v[0][0] === 'Cita más reciente · Praxeología'));
  const links = sets.filter(x => x.ruta.endsWith('.hyperlink')).map(x => x.v);
  assert.ok(links.some(l => l.documentReference === 'Citas!A2' && l.textToDisplay === 'Abrir ›'));
  assert.ok(links.some(l => l.textToDisplay === 'Ver reverso' && /^'Inicio'!A\d+$/.test(l.documentReference)));
  assert.ok(links.some(l => l.textToDisplay === 'Ver frente'));
  assert.equal(L.log.filter(x => x.op === 'call' && x.ruta.endsWith('.merge')).length, 0);
});

test('cita elegida desde el panel y errores', async () => {
  const r = await generarInicio(libro().ctx, modelo(), { medir, citaId: 's8' });
  assert.deepEqual(r, { cita: 's8', kicker: 'Cita elegida' });
  await assert.rejects(generarInicio(libro().ctx, modelo(), { medir, citaId: 'zz' }), { message: 'Cita zz: no existe en tblCitas.' });
});

test('rehace solo su propia hoja Inicio; si no lleva la marca, para', async () => {
  const propio = libro({ existente: { marca: MARCA_INICIO } });
  await generarInicio(propio.ctx, modelo(), { medir });
  assert.deepEqual(propio.log.filter(x => x.op === 'delete' || x.op === 'add').map(x => x.op), ['delete', 'add']);
  const ajeno = libro({ existente: { marca: 'otra cosa' } });
  await assert.rejects(generarInicio(ajeno.ctx, modelo(), { medir }), new RegExp(`Ya existe una hoja «${HOJA_INICIO}» que no creó ReadCite`));
  assert.equal(ajeno.log.filter(x => x.op === 'delete' || x.op === 'add').length, 0);
});

test('sin citas: portada con aviso', async () => {
  const L = libro({ ids: [] });
  const r = await generarInicio(L.ctx, { config: CONFIG_DEF, citas: [] }, { medir });
  assert.deepEqual(r, { cita: null, kicker: null });
  assert.ok(L.log.some(x => x.op === 'set' && x.ruta.endsWith('.values') && x.v[0][0] === 'Aún no hay citas.'));
});
