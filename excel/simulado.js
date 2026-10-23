// Excel simulado (tests y modo demostración del panel fuera de Excel): el subconjunto de Office.js que usa excel/capa.js, con la semántica que importa:
// las operaciones se encolan y solo se aplican en context.sync(); `null` en values deja la celda intacta;
// las filas nuevas heredan las columnas calculadas; las tablas exigen al menos una fila de datos.
import { TABLAS } from './esquema.js';

export function libroSimulado() {
  /** tablas: nombre → {hoja, encabezados, filas, formulas} */
  const tablas = {};
  const hojas = new Map(); // nombre → {celdas: Map('A1' → valor)}
  const cola = [];
  const stats = { syncs: 0, escrituras: 0, porSync: [] };

  const nueva = (hoja, encabezados) => ({ hoja, encabezados: [...encabezados], filas: [encabezados.map(() => '')], formulas: [encabezados.map(() => '')] });
  // Estructura exacta creada en la Fase 3: 5 tablas con su fila en blanco + tblConfig con valores.
  for (const d of Object.values(TABLAS)) { hojas.set(d.hoja, { celdas: new Map() }); tablas[d.tabla] = nueva(d.hoja, d.columnas); }
  const cfg = [['idioma', 'es'], ['traduccion_auto', true], ['usuario_local', 'Tú'], ['paginas_dia', 20], ['cuaderno_citas_por_pagina', 2], ['cuaderno_pestanas_espejo', true], ['cuaderno_anio', '']];
  tablas.tblConfig.filas = cfg.map(f => [...f]); tablas.tblConfig.formulas = cfg.map(f => f.map(v => (typeof v === 'string' ? v : String(v))));
  hojas.set('Reto', { celdas: new Map() });

  const valorFormula = (f, enc) => (typeof f === 'string' && f.startsWith('=') ? `«${f}»` : f);
  // Como Excel: un texto con apóstrofo delante se guarda como texto sin el apóstrofo; un texto que empieza por «=» se guarda como fórmula.
  const guardar = x => (typeof x === 'string' && x.startsWith("'") ? { v: x.slice(1), f: x.slice(1) } : typeof x === 'string' && x.startsWith('=') ? { v: `«${x}»`, f: x } : { v: x, f: x === '' ? '' : String(x) });
  const encolar = fn => { cola.push(fn); stats.escrituras++; };

  function filaRango(tb, i) {
    return {
      set values(v) { encolar(() => { v[0].forEach((x, j) => { if (x !== null) { const g = guardar(x); tb.filas[i][j] = g.v; tb.formulas[i][j] = g.f; } }); }); },
      clear() { encolar(() => { tb.filas[i] = tb.filas[i].map((x, j) => (tb.formulas[i][j]?.startsWith?.('=') ? x : '')); tb.formulas[i] = tb.formulas[i].map(f => (f?.startsWith?.('=') ? f : '')); }); },
    };
  }

  function tabla(nombre) {
    const tb = tablas[nombre];
    const t = { isNullObject: !tb, load() { return t; } };
    if (!tb) return t;
    Object.defineProperty(t, 'name', { set(n) { tablas[n] = tablas[nombre]; delete tablas[nombre]; } });
    t.getHeaderRowRange = () => ({ values: [[...tb.encabezados]], load() { return this; } });
    t.getDataBodyRange = () => {
      const b = { values: tb.filas.map(f => [...f]), formulas: tb.formulas.map(f => [...f]), rowCount: tb.filas.length, load() { return b; }, getRow: i => filaRango(tb, i) };
      return b;
    };
    t.rows = {
      add(_i, valores) {
        encolar(() => valores.forEach(v => {
          const calc = tb.formulas[0];
          tb.filas.push(v.map((x, j) => (x === null ? (calc[j]?.startsWith?.('=') ? valorFormula(calc[j]) : '') : guardar(x).v)));
          tb.formulas.push(v.map((x, j) => (x === null ? (calc[j]?.startsWith?.('=') ? calc[j] : '') : guardar(x).f)));
        }));
      },
      getItemAt(i) { return { delete() { encolar(() => { if (tb.filas.length === 1) throw new Error('Excel: una tabla necesita al menos una fila'); tb.filas.splice(i, 1); tb.formulas.splice(i, 1); }); } }; },
    };
    return t;
  }

  const ctx = {
    workbook: {
      tables: {
        getItemOrNullObject: tabla,
        add(direccion, _conEncabezados) {
          const m = direccion.match(/^'(.+)'!A1:([A-Z]+)2$/);
          const hoja = hojas.get(m[1]);
          const nombre = 'Tabla' + Object.keys(tablas).length;
          encolar(() => { // los encabezados se leen al aplicar (la escritura previa ya está aplicada)
            const enc = [...hoja.celdas.entries()].filter(([k]) => /^[A-Z]+1$/.test(k)).map(([, v]) => v);
            tablas[nombre] = nueva(m[1], enc);
          });
          let destino = nombre;
          return { set name(n) { encolar(() => { tablas[n] = tablas[destino]; delete tablas[destino]; destino = n; }); } };
        },
      },
      worksheets: {
        getItemOrNullObject(n) {
          const h = hojas.get(n);
          return { isNullObject: !h, load() { return this; }, getRange: r => rangoHoja(n, r) };
        },
        add(n) { hojas.set(n, { celdas: new Map() }); return { getRange: r => rangoHoja(n, r) }; },
      },
    },
    async sync() {
      stats.syncs++;
      stats.porSync.push(cola.length);
      while (cola.length) cola.shift()();
    },
  };

  function rangoHoja(n, r) {
    const [a, b] = r.split(':');
    const col = s => s.match(/[A-Z]+/)[0], fil = s => +s.match(/\d+/)[0];
    const cols = []; for (let c = col(a).charCodeAt(0); c <= col(b || a).charCodeAt(0); c++) cols.push(String.fromCharCode(c));
    const filas = []; for (let f = fil(a); f <= fil(b || a); f++) filas.push(f);
    const celdas = hojas.get(n).celdas;
    // Celdas ocupadas por tablas existentes en esa hoja
    const deTabla = (c, f) => { for (const tb of Object.values(tablas)) if (tb.hoja === n) { const j = c.charCodeAt(0) - 65; if (j < tb.encabezados.length) { if (f === 1) return tb.encabezados[j]; if (tb.filas[f - 2]) return tb.filas[f - 2][j]; } } return undefined; };
    return {
      get values() { return filas.map(f => cols.map(c => celdas.get(c + f) ?? deTabla(c, f) ?? '')); },
      set values(v) { encolar(() => v.forEach((fila, i) => fila.forEach((x, j) => celdas.set(cols[j] + filas[i], x)))); },
      load() { return this; },
    };
  }

  return { ctx, tablas, hojas, stats };
}
