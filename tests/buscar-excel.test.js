import { test } from 'node:test';
import assert from 'node:assert/strict';
import { generarBuscar, FORMULAS, ENTRADAS, HOJA_BUSCAR, MARCA_BUSCAR } from '../excel/buscar.js';
import { TABLAS } from '../excel/esquema.js';

function libro({ existente = null, inicio = true } = {}) {
  const log = [];
  const grab = (ruta) => new Proxy(function () {}, {
    get(_, k) { if (k === 'then') return undefined; return grab(`${ruta}.${String(k)}`); },
    set(_, k, v) { log.push({ op: 'set', ruta: `${ruta}.${String(k)}`, v }); return true; },
    apply(_, __, a) { log.push({ op: 'call', ruta, a }); return grab(`${ruta}(${a[0] ?? ''})`); },
  });
  const ctx = {
    workbook: { worksheets: {
      getItemOrNullObject: n => ({ isNullObject: n === HOJA_BUSCAR ? !existente : !inicio, load() { return this; },
        getRange: () => ({ load() { return this; }, values: [[existente?.marca ?? '']] }), delete() { log.push({ op: 'delete' }); } }),
      add(n) { log.push({ op: 'add', n }); return grab('ws'); },
    } },
    async sync() {},
  };
  return { ctx, log };
}

test('fórmulas: solo usan columnas reales de las tablas', () => {
  const cols = new Set([...TABLAS.citas.columnas, ...TABLAS.colecciones.columnas, ...TABLAS.config.columnas]);
  const todas = Object.values(FORMULAS).join(' ');
  for (const [, campo] of todas.matchAll(/tbl\w+\[(\w+)\]/g)) assert.ok(cols.has(campo), `columna desconocida: ${campo}`);
  // todas las entradas participan en el filtro
  for (const dir of Object.values(ENTRADAS)) assert.ok(FORMULAS.resultados.includes(dir), dir);
  // mismas reglas que la web: contiene sin mayúsculas en 5 campos, likes +1 si me gusta, orden descendente
  assert.match(FORMULAS.resultados, /LOWER\(tx&"\|"&tblCitas\[autor\]&"\|"&tblCitas\[libro\]&"\|"&tblCitas\[tema\]&"\|"&tblCitas\[concepto\]\)/);
  assert.match(FORMULAS.resultados, /IF\(tblCitas\[me_gusta\]=TRUE,1,0\)/);
  assert.match(FORMULAS.resultados, /tblCitas\[coleccion\]&""/); // celdas vacías como texto vacío, no 0
  assert.match(FORMULAS.resultados, /SORTBY\(FILTER\(res,ok=1\),FILTER\(clave,ok=1\),-1\)/);
  assert.ok(!FORMULAS.resultados.includes('\n'));
  // listas ordenadas por frecuencia, con «Todos/Todas» delante
  assert.match(FORMULAS.listaAutor, /VSTACK\("Todos",SORTBY\(u,COUNTIF\(tblCitas\[autor\],u\),-1\)\)/);
  assert.match(FORMULAS.listaColeccion, /VSTACK\("Todas",u\)/);
});

test('genera la hoja: barra con desplegables, resultados dinámicos y marca', async () => {
  const L = libro();
  await generarBuscar(L.ctx);
  const sets = L.log.filter(x => x.op === 'set');
  assert.ok(sets.some(x => x.ruta === 'ws.position' && x.v === 1)); // detrás de Inicio
  assert.ok(sets.some(x => x.ruta === 'ws.getRange(Z1).values' && x.v[0][0] === MARCA_BUSCAR));
  assert.ok(sets.some(x => x.ruta === 'ws.getRange(B11).formulas' && x.v[0][0] === FORMULAS.resultados));
  assert.ok(sets.some(x => x.ruta === 'ws.getRange(B6:G6).values' && x.v[0].join('|') === 'Más ♥|Todos|Todos|Todos|Todas|No'));
  const reglas = sets.filter(x => x.ruta.endsWith('dataValidation.rule')).map(x => x.v.list.source);
  assert.deepEqual(reglas, ['Más ♥,Recientes', '=$AA$1#', '=$AB$1#', '=$AC$1#', '=$AD$1#', 'No,Sí']);
  assert.ok(sets.some(x => x.ruta === 'ws.getRange(Z:AD).columnHidden' && x.v === true));
});

test('sin hoja Inicio va la primera; rehace solo su propia hoja', async () => {
  const L = libro({ inicio: false });
  await generarBuscar(L.ctx);
  assert.ok(L.log.some(x => x.op === 'set' && x.ruta === 'ws.position' && x.v === 0));
  const propia = libro({ existente: { marca: MARCA_BUSCAR } });
  await generarBuscar(propia.ctx);
  assert.deepEqual(propia.log.filter(x => x.op === 'delete' || x.op === 'add').map(x => x.op), ['delete', 'add']);
  const ajena = libro({ existente: { marca: 'otra' } });
  await assert.rejects(generarBuscar(ajena.ctx), /Ya existe una hoja «Buscar» que no creó ReadCite/);
});
