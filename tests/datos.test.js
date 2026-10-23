import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as D from '../core/datos.js';
import { web } from './web.js';

const AHORA = new Date(2026, 9, 15, 10, 30).getTime();
const ENC = D.COLUMNAS_CITAS.map(([, col]) => col);
// Citas de ejemplo de la web convertidas al modelo del núcleo (como si vinieran de tblCitas).
const seed = () => [...web.SEED].map((q, i) => ({ ...D.filaACita([], []), ...q, user: q.user, createdAt: AHORA - (i + 1) * 3600000, savedAt: null, liked: false }));
const FORM = { text: 'Una cita', author: 'Ana López', book: 'Libro nuevo', understood: 'Algo' };

test('esquema de tblCitas: encabezados exactos del brief + me_gusta y posicion_imagen', () => {
  assert.deepEqual(ENC, ['id', 'texto', 'autor', 'libro', 'tema', 'concepto', 'capitulo', 'pagina', 'anio_edicion', 'anio_original',
    'editorial', 'traductor', 'coleccion', 'entendido', 'por_que', 'likes', 'me_gusta', 'guardada_fecha', 'creada_fecha', 'usuario',
    'ratio_tarjeta', 'posicion_imagen', 'imagen_path', 'mask_json']);
});

test('fechas: serial de Excel ↔ ms (ida y vuelta, vacío)', () => {
  assert.equal(D.msASerial(new Date(1900, 0, 1).getTime()), 2);
  assert.equal(D.msASerial(new Date(2026, 9, 15).getTime()), 46310);
  assert.equal(D.serialAMs(46310.5), new Date(2026, 9, 15, 12).getTime());
  assert.equal(D.serialAMs(D.msASerial(AHORA)), AHORA);
  assert.equal(D.serialAMs(''), null);
  assert.equal(D.msASerial(null), '');
});

test('fila ↔ cita: ida y vuelta conserva todo, en cualquier orden de columnas', () => {
  const c = { ...D.filaACita([], []), id: 'u1', text: 'T', author: 'A B', book: 'L', likes: 3, liked: true, savedAt: AHORA, createdAt: AHORA - 1000, mask: [[0.1, 0.5], null] };
  const enc = [...ENC].reverse();
  assert.deepEqual(D.filaACita(D.citaAFila(c, enc), enc), c);
  const conExtra = [...ENC, 'mi_columna'];
  assert.equal(D.citaAFila(c, conExtra).at(-1), undefined); // columnas ajenas no se tocan
  assert.equal(D.filaACita(['x'], ['id']).mask, null);
  assert.equal(D.filaACita(['{roto'], ['mask_json']).mask, null);
  assert.equal(D.filaACita(['VERDADERO'], ['me_gusta']).liked, true);
});

test('crearCita: obligatorios, id, valores por defecto y no muta la entrada', () => {
  const citas = seed();
  const mal = D.crearCita(citas, { text: 'x' }, { ahora: AHORA, usuario: 'Tú' });
  assert.deepEqual(mal, { ok: false, error: 'Completa autor, libro, lo que entendiste' });
  const r = D.crearCita(citas, FORM, { ahora: AHORA, usuario: 'Tú' });
  assert.ok(r.ok);
  assert.equal(r.cita.id, 'u' + AHORA);
  assert.equal(r.cita.topic, 'Sin tema');
  assert.equal(r.cita.likes, 0);
  assert.equal(r.cita.ratio, 'poster');
  assert.equal(r.citas[0], r.cita); // la nueva va primero, como en la web
  assert.equal(citas.length, 12);
  const r2 = D.crearCita(r.citas, FORM, { ahora: AHORA, usuario: 'Tú' });
  assert.equal(r2.cita.id, 'u' + (AHORA + 1)); // nunca repite id
});

test('editarCita y borrarCita: errores en español con el id', () => {
  const citas = seed();
  assert.deepEqual(D.editarCita(citas, 'nope', { text: 'x' }), { ok: false, error: 'Cita nope: no existe' });
  assert.deepEqual(D.editarCita(citas, 's1', { author: ' ' }), { ok: false, error: 'Cita s1: falta autor' });
  const e = D.editarCita(citas, 's1', { text: 'Nuevo', id: 'hack', topic: '' });
  assert.equal(e.cita.text, 'Nuevo');
  assert.equal(e.cita.id, 's1');
  assert.equal(e.cita.topic, 'Sin tema');
  assert.equal(citas[0].text, web.SEED[0].text);
  const com = [{ citaId: 's1', usuario: 'Tú', cuando: AHORA, texto: 'a' }, { citaId: 's2', usuario: 'Tú', cuando: AHORA, texto: 'b' }];
  const b = D.borrarCita(citas, com, 's1');
  assert.equal(b.citas.length, 11);
  assert.deepEqual(b.comentarios.map(c => c.citaId), ['s2']);
  assert.equal(D.borrarCita(citas, com, 'zz').error, 'Cita zz: no existe');
});

test('guardar, quitar guardada y me gusta', () => {
  let citas = seed();
  citas = D.guardarCita(citas, 's3', AHORA).citas;
  assert.equal(citas.find(c => c.id === 's3').savedAt, AHORA);
  assert.deepEqual(D.filtrarCitas(citas, { soloGuardadas: true }, 'Tú').map(c => c.id), ['s3']);
  citas = D.quitarGuardada(citas, 's3').citas;
  assert.equal(D.filtrarCitas(citas, { soloGuardadas: true }, 'Tú').length, 0);
  assert.equal(D.guardarCita(citas, 'zz', AHORA).error, 'Cita zz: no existe');
  citas = D.alternarMeGusta(citas, 's12').citas;
  assert.equal(D.likesTotales(citas.find(c => c.id === 's12')), 75);
});

test('búsqueda: contiene, sin mayúsculas, en cita/autor/libro/tema/concepto', () => {
  const citas = seed();
  assert.deepEqual(D.filtrarCitas(citas, { query: '  ESTADO ', sort: 'recent' }, 'Tú').map(c => c.id), ['s4', 's5', 's6', 's7', 's8', 's12']);
  assert.deepEqual(D.filtrarCitas(citas, { query: 'praxeología' }, 'Tú').map(c => c.id), ['s1']);
  assert.equal(D.filtrarCitas(citas, { query: 'zzz' }, 'Tú').length, 0);
});

test('filtros combinados con Y, "mías" y orden por likes', () => {
  const citas = seed();
  assert.deepEqual(D.filtrarCitas(citas, { topic: 'Estado' }, 'Tú').map(c => c.id), ['s8', 's4', 's6', 's12']);
  assert.deepEqual(D.filtrarCitas(citas, { topic: 'Estado', author: 'Murray N. Rothbard' }, 'Tú').map(c => c.id), ['s4']);
  assert.deepEqual(D.filtrarCitas(citas, { collection: 'Orden espontáneo', query: 'mercado' }, 'Tú').map(c => c.id), ['s9', 's5']);
  assert.deepEqual(D.filtrarCitas(citas, { mine: true }, 'Carlos').map(c => c.id), ['s4', 's6', 's10', 's12']);
  assert.equal(D.filtrosActivos({ topic: 'Estado', mine: true, query: 'x' }), 2);
});

test('paridad con la web: búsqueda + filtros + orden por likes', () => {
  const citas = seed();
  const webFiltrar = (all, q, F) => all.filter(x => { // copia del filtro de renderVals
    const norm = s => (s || '').toLowerCase();
    if (q && ![x.text, x.author, x.book, x.topic, x.concept].some(v => norm(v).includes(q))) return false;
    if (F.author && x.author !== F.author) return false; if (F.book && x.book !== F.book) return false; if (F.topic && x.topic !== F.topic) return false; if (F.collection && x.collection !== F.collection) return false;
    return true;
  }).slice().sort((a, b) => (b.likes || 0) - (a.likes || 0));
  const casos = [{}, { topic: 'Estado' }, { author: 'Ludwig von Mises' }, { collection: 'Orden espontáneo' }, { book: 'El Estado' }];
  for (const q of ['', 'estado', 'la', 'mises', 'valor'])
    for (const F of casos) assert.deepEqual(D.filtrarCitas(citas, { ...F, query: q }, 'Tú').map(c => c.id), webFiltrar([...web.SEED], q, F).map(c => c.id), `${q} ${JSON.stringify(F)}`);
});

test('contar y opciones de filtros', () => {
  const citas = seed();
  assert.deepEqual(D.contar(citas, 'topic')[0], ['Estado', 4]);
  const op = D.opcionesFiltros(citas, [{ nombre: 'Vacía', descripcion: '', creada: null }]);
  assert.equal(op.autores[0], 'Ludwig von Mises');
  assert.deepEqual(op.colecciones, ['Mises sobre propiedad y acción', 'Ideas que cambiaron mi forma de pensar', 'El Estado bajo la lupa', 'Orden espontáneo', 'Vacía']);
});

test('colecciones: crear, duplicados, resumen y extracto', () => {
  const r = D.crearColeccion([], { nombre: ' Clásicos ', descripcion: 'Viejos' }, AHORA);
  assert.deepEqual(r.coleccion, { nombre: 'Clásicos', descripcion: 'Viejos', creada: AHORA });
  assert.equal(D.crearColeccion(r.colecciones, { nombre: 'Clásicos' }, AHORA).error, 'Ya existe la colección «Clásicos»');
  assert.equal(D.crearColeccion([], { nombre: '' }, AHORA).error, 'La colección necesita un nombre');
  const res = D.resumenColecciones(seed(), r.colecciones);
  const estado = res.find(c => c.nombre === 'El Estado bajo la lupa');
  assert.equal(estado.count, 4);
  assert.equal(res.at(-1).nombre, 'Clásicos');
  assert.equal(res.at(-1).count, 0);
  assert.equal(D.extracto('abcdef', 3), '“abc…”');
  assert.equal(D.extracto('abc', 3), '“abc”');
});

test('efectos al crear: libro nuevo en "leyendo" y colección nueva', () => {
  const cita = D.crearCita([], { ...FORM, collection: 'Nueva' }, { ahora: AHORA, usuario: 'Tú' }).cita;
  const e = D.efectosAlCrear(cita, [], [], AHORA);
  assert.equal(e.libroNuevo.estado, 'leyendo');
  assert.equal(e.coleccionNueva.nombre, 'Nueva');
  const e2 = D.efectosAlCrear(cita, [{ titulo: 'Libro nuevo' }], [{ nombre: 'Nueva' }], AHORA);
  assert.deepEqual([e2.libroNuevo, e2.coleccionNueva], [null, null]);
});

test('comentarios', () => {
  const citas = seed();
  const r = D.anadirComentario([], citas, { citaId: 's8', texto: ' Bien ', usuario: 'Tú', ahora: AHORA });
  assert.deepEqual(r.comentario, { citaId: 's8', usuario: 'Tú', cuando: AHORA, texto: 'Bien' });
  assert.equal(D.anadirComentario([], citas, { citaId: 's8', texto: ' ', usuario: 'Tú', ahora: AHORA }).error, 'Cita s8: el comentario está vacío');
  assert.equal(D.anadirComentario([], citas, { citaId: 'x', texto: 'a', usuario: 'Tú', ahora: AHORA }).error, 'Cita x: no existe');
  assert.equal(D.comentariosDe(r.comentarios, 's8').length, 1);
});

test('biblioteca y configuración', () => {
  const lib = D.biblioteca(seed(), [{ titulo: 'Liberalismo', autor: 'Ludwig von Mises', paginas: 224, estado: 'leidos', leidas: 224, anio: '1927' }]);
  assert.equal(lib.length, 12);
  assert.equal(lib[0].citas, 1);
  assert.equal(D.librosPorEstado(lib, 'quiero').length, 11);
  assert.equal(D.filaALibro(['X', '', '100', 'raro', '-5'], D.COLUMNAS_LIBROS).estado, 'quiero');
  assert.equal(D.filaALibro(['X', '', '100', 'raro', '-5'], D.COLUMNAS_LIBROS).leidas, 0);
  const cfg = D.leerConfig([['paginas_dia', '35'], ['traduccion_auto', 'FALSO'], ['usuario_local', 'Dani'], ['', 'x']]);
  assert.deepEqual([cfg.paginas_dia, cfg.traduccion_auto, cfg.usuario_local, cfg.idioma], [35, false, 'Dani', 'es']);
});

test('tabla vacía: todo funciona sin citas', () => {
  assert.deepEqual(D.filtrarCitas([], { query: 'x', topic: 'y' }, 'Tú'), []);
  assert.deepEqual(D.opcionesFiltros([], []), { autores: [], libros: [], temas: [], colecciones: [] });
  assert.deepEqual(D.biblioteca([], []), []);
  assert.deepEqual(D.leerConfig([]), D.CONFIG_DEF);
});

test('textoSeguro: antepone apóstrofo a lo que Excel tomaría por fórmula', () => {
  assert.equal(D.textoSeguro('=SUMA(A1)'), "'=SUMA(A1)");
  assert.equal(D.textoSeguro('+34 600'), "'+34 600");
  assert.equal(D.textoSeguro('-nota'), "'-nota");
  assert.equal(D.textoSeguro('@usuario'), "'@usuario");
  assert.equal(D.textoSeguro('La acción humana'), 'La acción humana');
  assert.equal(D.textoSeguro('— Mises'), '— Mises');
  assert.equal(D.textoSeguro(42), 42);
  assert.equal(D.textoSeguro(true), true);
});
