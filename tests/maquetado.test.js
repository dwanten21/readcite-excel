import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as M from '../core/maquetado.js';
import { estructuraCuaderno, NB_LETTERS } from '../core/cuaderno.js';
import { nbHyph, nbHyphWord } from '../core/silabeo.js';
import { web } from './web.js';

// Medidor de prueba: cada carácter mide 0,5 × tamaño (cursiva y negrita un 5 % más).
const medir = (t, f) => [...t].length * f.size * 0.5 * (f.italic || f.bold ? 1.05 : 1);
const AHORA = new Date(2026, 9, 15, 12).getTime();
const guardadas = () => [...web.SEED].map((q, i) => ({ ...q, savedAt: AHORA - i * 86400000, createdAt: null, liked: false }));

test('geometría: 147 mm de ancho y cada página por debajo de 210 mm', () => {
  assert.equal(M.COLUMNAS_MM.reduce((a, b) => a + b, 0), 147);
  // anverso: margen interior A+B+C = 16 mm, exterior I = 6 mm; reverso al revés
  assert.equal(M.COLUMNAS_MM[0] + M.COLUMNAS_MM[1] + M.COLUMNAS_MM[2], 16);
  assert.equal(M.COLUMNAS_MM[7] + M.COLUMNAS_MM[8] + M.COLUMNAS_MM[9], 16);
  const altoMm = M.PAD_SUP_MM + (M.FILAS_CUERPO + M.FILAS_PIE) * M.FILA_MM + M.PAD_INF_MM;
  assert.ok(altoMm < 205, `página de ${altoMm} mm`);
});

test('partirLineas: corta solo en los puntos de nbHyphWord y sin U+00AD en la salida', () => {
  const texto = nbHyph('La acción humana es conducta consciente y deliberada de la constitución');
  const f = { size: 10 };
  for (const ancho of [40, 60, 90, 130]) {
    const l = M.partirLineas(texto, ancho, medir, f);
    for (const x of l) {
      assert.ok(!x.includes('­'), x);
      assert.ok(medir(x, f) <= ancho || !x.includes(' '), `«${x}» mide más de ${ancho}`);
    }
    // Reconstrucción: quitando los guiones de corte se recupera el texto original
    const unido = l.map((x, i) => (x.endsWith('-') && i < l.length - 1 ? x.slice(0, -1) + '\u0000' : x + ' ')).join('').replace(/\u0000/g, '').trim();
    assert.equal(unido, texto.replace(/­/g, ''));
    // cada corte con guion coincide con un corte permitido por nbHyphWord
    l.forEach((x, i) => {
      if (!x.endsWith('-') || i === l.length - 1) return;
      const parte = x.slice(0, -1).split(' ').pop(); const resto = l[i + 1].split(' ')[0];
      const palabra = parte + resto.replace(/-$/, '');
      const permitido = nbHyphWord(palabra).split('­');
      const acumulados = permitido.map((_, k) => permitido.slice(0, k + 1).join(''));
      assert.ok(acumulados.includes(parte), `corte «${parte}|${resto}» no permitido en ${palabra}`);
    });
  }
});

test('partirLineas: palabras sin corte pasan enteras; sangría APA; texto vacío', () => {
  assert.deepEqual(M.partirLineas('aaa bbb', 3 * 5, medir, { size: 10 }), ['aaa', 'bbb']);
  assert.deepEqual(M.partirLineas('supercalifragilístico', 20, medir, { size: 10 }), ['supercalifragilístico']);
  const l = M.partirLineas('uno dos tres cuatro cinco', 50, medir, { size: 10 }, 20);
  assert.ok(l.slice(1).every(x => medir(x, { size: 10 }) <= 30));
  assert.deepEqual(M.partirLineas('', 50, medir, { size: 10 }), []);
});

test('recortar: añade «…» y respeta el ancho', () => {
  const l = M.recortar(['uno dos', 'tres cua-', 'tro cinco'], 2, 30, medir, { size: 10 });
  assert.equal(l.length, 2);
  assert.ok(l[1].endsWith('…') && !l[1].includes('-…'));
  assert.deepEqual(M.recortar(['a', 'b'], 5, 30, medir, { size: 10 }), ['a', 'b']);
});

/** Comprueba invariantes del plan. */
function invariantes(nb, plan) {
  assert.equal(plan.altosPt.length, plan.filas);
  assert.equal(plan.paginas, nb.pages.length);
  const altoPagMm = plan.altosPt.slice(0, M.FILAS_PAGINA).reduce((a, b) => a + b, 0) / M.PT_POR_MM;
  assert.ok(altoPagMm < 205, `página de ${altoPagMm} mm`);
  // ningún bloque se solapa con otro ni sale de su página
  const ocupado = new Map();
  for (const b of plan.bloques) {
    const pag = Math.floor(b.f1 / M.FILAS_PAGINA);
    assert.equal(Math.floor(b.f2 / M.FILAS_PAGINA), pag, `bloque «${b.v}» cruza de página`);
    assert.ok(b.c1 >= 0 && b.c2 <= 9 && b.c1 <= b.c2 && b.f1 <= b.f2);
    for (let f = b.f1; f <= b.f2; f++) for (let c = b.c1; c <= b.c2; c++) {
      const k = f + ':' + c; assert.ok(!ocupado.has(k), `solape en fila ${f} col ${c}: «${ocupado.get(k)}» / «${b.v}»`); ocupado.set(k, b.v);
    }
    assert.ok(!b.v.includes('­'), 'sin guiones blandos en Excel');
    if (b.v) assert.ok(M.COLUMNAS_MM.slice(b.c1, b.c2 + 1).reduce((x, y) => x + y, 0) >= 5, `texto «${b.v}» en columnas de menos de 5 mm`);
  }
  // cada letra empieza en una página impar (anverso) y las pestañas enlazan a su inicio
  for (const l of NB_LETTERS) assert.equal(plan.inicioLetra[l] / M.FILAS_PAGINA % 2, 0, `letra ${l}`);
  const tabs = plan.bloques.filter(b => b.e.hueco);
  assert.equal(tabs.length, 26 * nb.pages.length);
  for (const t of tabs) assert.equal(t.link, plan.inicioLetra[t.v]);
}

test('plan del cuaderno con las 12 citas de ejemplo (2 por página)', () => {
  const nb = estructuraCuaderno(guardadas(), { ahora: AHORA, porPagina: 2 });
  const plan = M.planCuaderno(nb, medir);
  invariantes(nb, plan);
  // pestañas: en anverso a la derecha (J), en reverso a la izquierda (A), como la web
  const tabsPag = i => plan.bloques.filter(b => b.e.hueco && Math.floor(b.f1 / M.FILAS_PAGINA) === i);
  assert.ok(tabsPag(2).every(b => b.c1 === 9));
  assert.ok(tabsPag(3).every(b => b.c1 === 0));
  // la pestaña activa de la página de la M es negra
  const pM = plan.inicioLetra.M / M.FILAS_PAGINA;
  assert.equal(tabsPag(pM).find(b => b.v === 'M').e.fill, M.COLORES.tabActivo);
  // la cita de Mises está, silabeada a mano, y su APA también
  const cita = plan.bloques.find(b => b.v.startsWith('“La ac'));
  assert.ok(cita, 'falta la cita de La acción humana');
  assert.ok(plan.bloques.some(b => b.v.startsWith('von Mises, L. (2011). La acción humana')));
});

test('plan: 1 y 3 citas por página, sin espejo, y cuaderno vacío', () => {
  for (const porPagina of [1, 3]) invariantes(...(nb => [nb, M.planCuaderno(nb, medir)])(estructuraCuaderno(guardadas(), { ahora: AHORA, porPagina })));
  const sinEspejo = estructuraCuaderno(guardadas(), { ahora: AHORA, espejo: false });
  const p = M.planCuaderno(sinEspejo, medir);
  invariantes(sinEspejo, p);
  assert.ok(p.bloques.filter(b => b.e.hueco).every(b => b.c1 === 9));
  const vacio = estructuraCuaderno([], { ahora: AHORA });
  const pv = M.planCuaderno(vacio, medir);
  invariantes(vacio, pv);
  assert.equal(pv.paginas, 54);
});

test('plan: una cita larguísima se recorta con «…» sin salirse de su entrada', () => {
  const larga = { ...guardadas()[0], text: 'palabra '.repeat(400).trim(), understood: 'nota '.repeat(300) };
  const nb = estructuraCuaderno([larga], { ahora: AHORA, porPagina: 3 });
  const plan = M.planCuaderno(nb, medir);
  invariantes(nb, plan);
  const cita = plan.bloques.find(b => b.v.startsWith('“palabra'));
  assert.ok(cita.v.endsWith('…'));
  assert.ok(cita.e.size >= 7);
});

test('plan: las notas («lo que entendí») van bajo la cita y tienen espacio aunque la cita sea larga', () => {
  const larga = { ...guardadas()[0], text: 'imprenta '.repeat(120).trim(), understood: 'Los comienzos del comercio de libros se dieron con impresores y libreros. '.repeat(3).trim() };
  const nb = estructuraCuaderno([larga], { ahora: AHORA, porPagina: 2 });
  const plan = M.planCuaderno(nb, medir);
  invariantes(nb, plan);
  const cita = plan.bloques.find(b => b.v.startsWith('“imprenta'));
  const etiqueta = plan.bloques.find(b => b.v === 'LO QUE ENTENDÍ' && b.f1 > cita.f1);
  const notas = plan.bloques.find(b => b.v.startsWith('Los comienzos'));
  const obra = plan.bloques.find(b => b.v === 'TEMA · CONCEPTO' && b.f1 > cita.f1);
  assert.ok(cita.f2 < etiqueta.f1 && etiqueta.f2 < notas.f1 && notas.f2 < obra.f1, 'orden: cita → lo que entendí → datos');
  assert.ok(notas.v.split('\n').length >= 3, `notas con ${notas.v.split('\n').length} líneas`);
  assert.ok(notas.f2 - notas.f1 + 1 >= 6, 'al menos el 40 % del espacio para las notas');
});

test('plan: cita y notas cortas no dejan huecos; sin notas quedan renglones para escribir', () => {
  const corta = { ...guardadas()[0], text: 'Breve.', understood: 'Una idea.' };
  const sinNotas = { ...guardadas()[1], understood: '' };
  const nb = estructuraCuaderno([corta, sinNotas], { ahora: AHORA, porPagina: 2 });
  const plan = M.planCuaderno(nb, medir);
  invariantes(nb, plan);
  const cita = plan.bloques.find(b => b.v.startsWith('“Breve'));
  assert.ok(cita.f2 - cita.f1 + 1 <= 4, 'la cita corta ocupa solo lo que necesita');
});
