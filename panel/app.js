// @ts-check
// Panel lateral de ReadCite. Toda la lógica vive en core/ y excel/capa.js; aquí solo se pinta y se reacciona.
// Dentro de Excel usa Excel.run; fuera (navegador) arranca en modo demostración con un libro simulado.
import * as X from '../excel/capa.js';
import * as D from '../core/datos.js';
import { apa, apaText } from '../core/apa.js';
import { resumenReto, datosLibro, sesionesHoy, ICARO_INTERVALS } from '../core/reto.js';
import { libroSimulado } from '../excel/simulado.js';
import { generarCuaderno, areaImpresion, opcionesCuaderno } from '../excel/cuaderno.js';
import { generarInicio, hayInicio } from '../excel/inicio.js';
import { generarBuscar } from '../excel/buscar.js';
import { guardarPng, dibujarTarjeta } from './png.js';
import { limpiarTranscripcion } from '../core/ocr.js';
import { imagenLegible } from './imagen.js';
import { FORMATOS } from '../core/tarjeta.js';
import { estructuraCuaderno, NB_LETTERS } from '../core/cuaderno.js';

/** Medidor de texto con la misma fuente que usa la hoja (Georgia), en puntos. */
const lienzo = document.createElement('canvas').getContext('2d');
/** @type {import('../core/maquetado.js').Medidor} */
const medir = (t, f) => { lienzo.font = `${f.italic ? 'italic ' : ''}${f.bold ? 'bold ' : ''}${f.size}pt Georgia`; return lienzo.measureText(t).width * 0.75; };

/* global Office, Excel */
const VERSION = '1.0.0';
let enExcel = false;
/** @type {any} */ let sim = null;
/** @param {(ctx:any)=>Promise<any>} fn */
const ejecutar = fn => (enExcel ? Excel.run(ctx => fn(ctx)) : fn(sim.ctx));

const FILTROS0 = { query: '', author: '', book: '', topic: '', collection: '', mine: false, soloGuardadas: false, sort: /** @type {'likes'|'recent'} */ ('likes') };
const st = {
  vista: 'citas',
  /** @type {any} */ modelo: null,
  filtros: { ...FILTROS0 },
  reverso: new Set(), comentarios: new Set(),
  /** @type {string|null} */ confirmarBorrar: null,
  /** @type {string|null} */ editId: null,
  /** @type {Record<string,any>} */ form: {},
  libTab: 'leyendo', confirmarReparar: false, ocrEstado: '', ocrNota: '',
  /** @type {string} */ progresoCuaderno: '', letraImpresion: '',
};

// ───────── utilidades ─────────
const $ = (/** @type {string} */ s) => /** @type {HTMLElement} */ (document.querySelector(s));
const esc = (/** @type {any} */ s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c] || c);
const opciones = (/** @type {string[]} */ lista, /** @type {string} */ sel, vacio = 'Todos') => `<option value="">${vacio}</option>` + lista.map(v => `<option${v === sel ? ' selected' : ''}>${esc(v)}</option>`).join('');
const COLORES = ['#3b3f47', '#e0533a', '#4f86d9', '#2fae63', '#d9a400', '#8a6cd1', '#2b9aa0'];
const hash = (/** @type {string} */ s) => { let x = 0; for (const c of s || '') x = (x * 31 + c.charCodeAt(0)) >>> 0; return x; };
const fecha = (/** @type {number|null} */ ms) => (ms ? new Date(ms).toLocaleString('es', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '');
const apaHtml = (/** @type {any} */ q) => { const a = apa(q); return `${esc(a.pre)}<i>${esc(a.title)}</i>${esc(a.post)}`; };

let tToast = 0;
/** @param {string} msg @param {boolean} [mal] */
function toast(msg, mal = false) {
  const t = $('#toast'); t.textContent = msg; t.className = 'toast' + (mal ? ' mal' : ''); t.hidden = false;
  clearTimeout(tToast); tToast = window.setTimeout(() => { t.hidden = true; }, mal ? 6000 : 2200);
}
/** @param {string} texto */
async function copiar(texto) {
  try { await navigator.clipboard.writeText(texto); }
  catch { const a = document.createElement('textarea'); a.value = texto; document.body.append(a); a.select(); document.execCommand('copy'); a.remove(); }
  toast('Copiado al portapapeles');
}

async function recargar() { st.modelo = await ejecutar(ctx => X.leerTodo(ctx)); pintar(); }

/** Tras crear, editar o borrar: si existe la hoja Inicio, se redibuja sola con la cita más reciente (sin cambiar de hoja). */
async function refrescarInicio() {
  if (!enExcel) return;
  try { await Excel.run(async ctx => { if (await hayInicio(ctx)) await generarInicio(ctx, st.modelo, { medir, activar: false }); }); }
  catch (e) { toast('No pude actualizar la hoja Inicio: ' + (e instanceof Error ? e.message : e), true); }
}
/** Excel de escritorio sin servidor local (Mac/Windows) no deja descargar desde el panel: el PNG se inserta como
 * imagen en la hoja activa (clic derecho → «Guardar como imagen»). En Excel para la web se descarga normal. */
function insertarPng() {
  // @ts-ignore
  const plataforma = enExcel && Office.context ? String(Office.context.platform) : '';
  if (!enExcel || plataforma === 'OfficeOnline') return null;
  return async (/** @type {string} */ base64) => {
    await Excel.run(async ctx => {
      const img = ctx.workbook.worksheets.getActiveWorksheet().shapes.addImage(base64);
      img.left = 20; img.top = 20; img.width = 300; img.name = 'ReadCite PNG';
      await ctx.sync();
    });
    return 'Imagen insertada en la hoja: clic derecho sobre ella → «Guardar como imagen…».';
  };
}

/** @param {string} [citaId] */
async function pintarInicio(citaId) {
  try {
    const r = await ejecutar(ctx => generarInicio(ctx, st.modelo, { medir, citaId }));
    toast(r.cita ? `Hoja Inicio: ${r.kicker.toLowerCase()} (${r.cita}).` : 'Hoja Inicio creada (aún no hay citas).');
  } catch (e) { toast(e instanceof Error ? e.message : String(e), true); }
}
/** Ejecuta una operación de la capa Excel, avisa y recarga. @param {(ctx:any)=>Promise<any>} fn @param {string} [ok] */
async function hacer(fn, ok) {
  try { await ejecutar(fn); if (ok) toast(ok); await recargar(); return true; }
  catch (e) { toast(e instanceof Error ? e.message : String(e), true); return false; }
}

// ───────── pintado ─────────
function pintar() {
  document.querySelectorAll('.pestanas button').forEach(b => b.setAttribute('aria-selected', String(/** @type {HTMLElement} */ (b).dataset.vista === st.vista)));
  $('#avisos').innerHTML = (st.modelo?.avisos || []).map((/** @type {string} */ a) => `<div class="aviso">${esc(a)} Ve a ⚙ y usa «Reparar estructura».</div>`).join('');
  const v = { citas: vistaCitas, nueva: vistaFormulario, biblioteca: vistaBiblioteca, colecciones: vistaColecciones, cuaderno: vistaCuaderno, ajustes: vistaAjustes }[st.vista];
  $('#vista').innerHTML = v ? v() : '';
  if (st.vista === 'citas') pintarLista();
  if (st.vista === 'nueva') { pintarApaForm(); pintarVistaPrevia(); }
}

function vistaCitas() {
  const F = st.filtros; const op = D.opcionesFiltros(st.modelo.citas, st.modelo.colecciones); const n = D.filtrosActivos(F) + (F.soloGuardadas ? 1 : 0);
  return `<div class="caja">
    <input id="q" type="search" placeholder="Busca en citas, autores, libros, temas…" value="${esc(F.query)}" aria-label="Buscar">
    <details${n ? ' open' : ''}><summary class="contador">Filtros ${n ? `(${n})` : ''}</summary>
      <div class="rejilla2" style="margin-top:6px">
        <label>Autor<select data-filtro="author">${opciones(op.autores, F.author)}</select></label>
        <label>Libro<select data-filtro="book">${opciones(op.libros, F.book)}</select></label>
        <label>Tema<select data-filtro="topic">${opciones(op.temas, F.topic)}</select></label>
        <label>Colección<select data-filtro="collection">${opciones(op.colecciones, F.collection, 'Todas')}</select></label>
      </div>
      <div class="acciones" style="margin-top:6px">
        <label class="fila"><input type="checkbox" data-filtro="mine"${F.mine ? ' checked' : ''}>Solo mías</label>
        <label class="fila"><input type="checkbox" data-filtro="soloGuardadas"${F.soloGuardadas ? ' checked' : ''}>Guardadas</label>
        <select data-filtro="sort" style="width:auto" aria-label="Orden"><option value="likes"${F.sort === 'likes' ? ' selected' : ''}>Más likes</option><option value="recent"${F.sort === 'recent' ? ' selected' : ''}>Recientes</option></select>
        <button class="mini" data-acc="limpiar">Limpiar</button>
      </div>
    </details>
    <div class="acciones"><button class="mini" data-acc="pintar-inicio"${enExcel ? '' : ' disabled'} title="Dibuja la portada en la hoja Inicio con la cita más reciente">⌂ Actualizar hoja Inicio</button>
      <button class="mini" data-acc="hoja-buscar"${enExcel ? '' : ' disabled'} title="Crea la hoja Buscar (o limpia sus filtros)">🔎 Hoja Buscar</button></div>
  </div>
  <div class="contador" id="cuenta"></div>
  <div id="lista" style="display:flex;flex-direction:column;gap:12px"></div>`;
}

function pintarLista() {
  const lista = D.filtrarCitas(st.modelo.citas, st.filtros, st.modelo.config.usuario_local);
  $('#cuenta').textContent = `${lista.length} ${lista.length === 1 ? 'cita' : 'citas'}`;
  $('#lista').innerHTML = lista.length ? lista.map(tarjeta).join('')
    : `<p class="vacio">${st.modelo.citas.length ? 'Ninguna cita coincide con la búsqueda.' : 'Aún no hay citas. Crea la primera en «Nueva».'}</p>`;
}

/** @param {D.Cita} c */
function tarjeta(c) {
  const rev = st.reverso.has(c.id);
  const coms = D.comentariosDe(st.modelo.comentarios, c.id);
  const locus = [c.chapter ? `Cap. ${c.chapter}` : '', c.page ? `p. ${c.page}` : ''].filter(Boolean).join(' · ') || 'Sin ubicación';
  const cara = rev
    ? `<div class="reverso">
        <div class="etiqueta">Lo que entendí</div><p>${esc(c.understood)}</p>
        ${c.why ? `<div class="etiqueta">Por qué me llamó la atención</div><p>${esc(c.why)}</p>` : ''}
        <div class="locus">${esc(locus)}${c.collection ? ` · Colección: ${esc(c.collection)}` : ''}</div>
        <div class="apa" style="margin-top:6px">${apaHtml(c)}</div>
      </div>`
    : `<div class="tema">${esc(c.topic)}</div>
       <blockquote class="cita">“${esc(c.text)}”</blockquote>
       <div class="autor">— ${esc(c.author)}</div><div class="libro">${esc(c.book)}</div>`;
  const borrar = st.confirmarBorrar === c.id
    ? `<span class="error">¿Borrar la cita y sus comentarios?</span><button class="peligro mini" data-acc="borrar-si" data-id="${esc(c.id)}">Sí, borrar</button><button class="mini" data-acc="borrar-no">No</button>`
    : `<button data-acc="editar" data-id="${esc(c.id)}">Editar</button><button data-acc="borrar" data-id="${esc(c.id)}" aria-label="Borrar">🗑</button>`;
  return `<article class="tarjeta" aria-label="Cita de ${esc(c.author)}">
    <span class="cinta${c.savedAt ? ' guardada' : ''}" title="${c.savedAt ? 'Guardada' : 'No guardada'}"></span>
    ${cara}
    <div class="pie-tarjeta">
      <button data-acc="girar" data-id="${esc(c.id)}">${rev ? 'Ver frente' : 'Ver reverso'}</button>
      <button data-acc="like" data-id="${esc(c.id)}" class="likes${c.liked ? ' mio' : ''}" aria-pressed="${c.liked}">♥ ${D.likesTotales(c)}</button>
      <button data-acc="${c.savedAt ? 'quitar' : 'guardar'}" data-id="${esc(c.id)}">${c.savedAt ? 'Guardada ✓' : 'Guardar'}</button>
      <button data-acc="copiar-apa" data-id="${esc(c.id)}">APA</button>
      <button data-acc="png" data-id="${esc(c.id)}" title="Descargar la tarjeta en PNG (${FORMATOS[c.ratio === 'square' ? 'square' : 'poster'].label})">PNG</button>
      <button data-acc="destacar" data-id="${esc(c.id)}"${enExcel ? '' : ' disabled'} title="Poner esta cita en la hoja Inicio">⌂</button>
      <button data-acc="comentarios" data-id="${esc(c.id)}">💬 ${coms.length}</button>
      ${borrar}
    </div>
    ${st.comentarios.has(c.id) ? `<div class="comentarios">
      ${coms.length ? coms.map((/** @type {D.Comentario} */ m) => `<div class="comentario"><b>${esc(m.usuario)}</b><time>${esc(fecha(m.cuando))}</time><div>${esc(m.texto)}</div></div>`).join('') : '<div class="contador">Sin comentarios todavía.</div>'}
      <textarea id="com-${esc(c.id)}" placeholder="Añade tu comentario…" style="min-height:44px"></textarea>
      <div><button class="mini primario" data-acc="comentar" data-id="${esc(c.id)}">Comentar</button></div>
    </div>` : ''}
  </article>`;
}

const CAMPOS = /** @type {const} */ ([
  ['text', 'Cita textual *', 'area'], ['author', 'Autor *'], ['book', 'Libro *'], ['understood', 'Lo que entendí *', 'area'], ['why', 'Por qué me llamó la atención', 'area'],
  ['topic', 'Tema'], ['concept', 'Concepto'], ['chapter', 'Capítulo'], ['page', 'Página'], ['year', 'Año de la edición'], ['originalYear', 'Año original'],
  ['publisher', 'Editorial'], ['translator', 'Traductor'], ['collection', 'Colección'],
]);

function vistaFormulario() {
  const f = st.form; const editando = !!st.editId;
  const campo = (/** @type {readonly string[]} */ [k, label, tipo]) => tipo === 'area'
    ? `<label style="grid-column:1/-1">${label}<textarea data-campo="${k}">${esc(f[k])}</textarea></label>`
    : `<label${['author', 'book', 'collection'].includes(k) ? ' style="grid-column:1/-1"' : ''}>${label}<input data-campo="${k}" value="${esc(f[k])}"${k === 'collection' ? ' list="lista-colecciones"' : ''}></label>`;
  return `<form class="caja" id="formulario" novalidate>
    <h2>${editando ? 'Editar cita' : 'Nueva cita'}</h2>
    <div class="caja" style="background:var(--suave);gap:4px">
      <div class="acciones"><span style="flex:1">Transcribir desde una foto de la página</span>
        <label class="boton-archivo${st.ocrEstado ? ' ocupado' : ''}"><span id="ocr-estado">${esc(st.ocrEstado || 'Subir foto')}</span><input id="foto" type="file" accept="image/*,.heic,.heif" hidden${st.ocrEstado ? ' disabled' : ''}></label></div>
      <div class="contador">${esc(st.ocrNota || 'La foto no se guarda: se lee aquí mismo, en tu Mac. El texto extraído queda en el campo de abajo para que lo corrijas.')}</div>
    </div>
    <div class="rejilla2">${CAMPOS.map(campo).join('')}</div>
    <datalist id="lista-colecciones">${D.nombresColecciones(st.modelo.citas, st.modelo.colecciones).map(n => `<option value="${esc(n)}">`).join('')}</datalist>
    <details><summary class="contador">Tarjeta e imagen</summary><div class="rejilla2" style="margin-top:6px">
      <label>Posición del texto<select data-campo="imagePos">${D.POSICIONES.map(p => `<option value="${p}"${(f.imagePos || 'center') === p ? ' selected' : ''}>${{ top: 'Arriba', center: 'Centro', bottom: 'Abajo' }[p]}</option>`).join('')}</select></label>
      <label style="grid-column:1/-1">Ruta de la imagen<input data-campo="image" value="${esc(f.image)}" placeholder="p. ej. imagenes/mises.jpg"></label>
    </div></details>
    <div><div class="contador">Referencia APA 7</div><div class="apa" id="apa-form"></div></div>
    <div style="display:flex;flex-direction:column;gap:6px">
      <div class="contador" style="letter-spacing:.12em;text-transform:uppercase">Vista previa</div>
      <div class="segmentos" role="group" aria-label="Formato de tarjeta">${Object.entries(FORMATOS).map(([k, v]) => `<button type="button" data-acc="formato" data-ratio="${k}" aria-pressed="${(f.ratio === 'square' ? 'square' : 'poster') === k}">${v.label}</button>`).join('')}</div>
      <div id="vista-previa" class="vista-previa"></div>
      <div><button type="button" data-acc="png-form">Descargar PNG</button></div>
    </div>
    <div class="acciones">
      <button type="submit" class="primario">${editando ? 'Guardar cambios' : 'Guardar cita'}</button>
      ${editando ? '<button type="button" data-acc="cancelar-edicion">Cancelar</button>' : '<button type="button" data-acc="vaciar-form">Vaciar</button>'}
    </div>
  </form>`;
}
function pintarApaForm() { const el = $('#apa-form'); if (el) el.innerHTML = apaHtml(st.form); }

let tPrevia = 0, nPrevia = 0;
/** Vista previa de la tarjeta PNG (se redibuja al escribir, con un pequeño retardo). */
function pintarVistaPrevia() {
  clearTimeout(tPrevia);
  tPrevia = window.setTimeout(async () => {
    const yo = ++nPrevia; const f = st.form;
    const c = await dibujarTarjeta({ text: f.text, author: f.author, book: f.book, topic: f.topic, ratio: f.ratio });
    const caja = $('#vista-previa'); if (!caja || yo !== nPrevia) return;
    c.style.width = '100%'; c.style.height = 'auto'; caja.replaceChildren(c);
  }, 200);
}

/* global Tesseract */
/** OCR en el propio panel con Tesseract.js (se descarga la primera vez desde jsDelivr; la foto no sale del Mac). */
async function cargarOcr() {
  // @ts-ignore
  if (window.Tesseract) return window.Tesseract;
  await new Promise((res, rej) => {
    const sc = document.createElement('script'); sc.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
    // Sello de integridad (SRI): si alguien altera el archivo en el CDN, el navegador se niega a ejecutarlo.
    sc.integrity = 'sha256-qOKZGNCYsrBuEBK9rv+0rsBEXF1WVHCQI+C9H0QqgOg='; sc.crossOrigin = 'anonymous'; sc.referrerPolicy = 'no-referrer';
    sc.onload = res; sc.onerror = () => rej(new Error('No se pudo cargar el lector de fotos (¿sin conexión a internet?).')); document.head.append(sc); });
  // @ts-ignore
  return window.Tesseract;
}
/** @param {File} archivo */
async function transcribir(archivo) {
  const marca = (/** @type {string} */ t) => { st.ocrEstado = t; const el = document.getElementById('ocr-estado'); if (el) el.textContent = t; };
  st.ocrNota = ''; marca('Preparando…'); pintar();
  try {
    marca('Abriendo foto…');
    const imagen = await imagenLegible(archivo);
    const T = await cargarOcr();
    const r = await T.recognize(imagen, 'spa', { logger: (/** @type {any} */ m) => {
      if (m.status === 'recognizing text') marca(`Leyendo… ${Math.round(m.progress * 100)} %`);
      else if (/load|initializ/.test(m.status)) marca('Cargando español…');
    } });
    const texto = limpiarTranscripcion(r.data.text);
    if (!texto) throw new Error('vacío');
    st.form.text = texto;
    st.ocrNota = 'Texto transcrito. Revísalo y recorta lo que no sea la cita.';
  } catch (e) {
    st.ocrNota = e instanceof Error && e.message.startsWith('No se pudo cargar') ? e.message
      : e instanceof Error && e.message === 'formato' ? 'No pude abrir ese archivo de imagen. Prueba con una foto JPG, PNG o HEIC.'
      : 'No se pudo leer la foto. Prueba con más luz o escribe la cita a mano.';
  }
  st.ocrEstado = ''; pintar();
}

function vistaBiblioteca() {
  const cfg = st.modelo.config; const perDay = cfg.paginas_dia;
  const libros = D.biblioteca(st.modelo.citas, st.modelo.libros).map(l => ({ ...datosLibro(l, perDay), citas: l.citas }));
  const tabs = D.ESTADOS.map(k => `<button data-acc="libtab" data-tab="${k}" aria-selected="${st.libTab === k}">${D.ETIQUETAS_ESTADO[k]} <span class="contador">${libros.filter(b => b.estado === k).length}</span></button>`).join('');
  let reto = '';
  if (st.libTab === 'quiero' && libros.some(b => b.estado === 'quiero')) {
    const r = resumenReto(libros, perDay); // incluye libros que solo aparecen en citas (estado 'quiero'), como la web
    const s = sesionesHoy(st.modelo.citas, libros, cfg.usuario_local, Date.now());
    reto = `<section class="caja reto" aria-label="Reto de lectura">
      <h3 class="rojo">Reto de lectura</h3>
      <div><span class="cifra">${r.totalPages}</span> páginas por leer · ${esc(r.booksLabel)}</div>
      <div>A tu ritmo terminas en <b>${esc(r.daysLabel)}</b>.</div>
      <label>Páginas al día: <b>${r.perDay}</b><input type="range" min="5" max="120" step="5" value="${r.perDay}" data-acc-cambio="paginas-dia"></label>
      <div class="acciones contador"><span>${r.readPages} leídas</span><span style="margin-left:auto">${esc(r.pctLabel)}</span></div>
      <div class="barra"><i style="width:${r.pct}%"></i></div>
      <h3 style="margin-top:6px">Memorizar · Método Ícaro</h3>
      <div class="intervalos">${ICARO_INTERVALS.map(d => `<span>Día ${d}</span>`).join('')}</div>
      <div class="contador">Sesiones de hoy: <b>${s.dueToday}</b> · Citas en el plan: <b>${s.quotesInPlan}</b></div>
    </section>`;
  }
  const filas = libros.filter(b => b.estado === st.libTab).map(b => {
    const q = st.modelo.citas.find((/** @type {D.Cita} */ c) => c.book === b.titulo) || { author: b.autor, book: b.titulo, originalYear: b.anio };
    return `<div class="libro-fila">
      <div class="lomo" style="background:${COLORES[hash(b.titulo) % COLORES.length]}">${esc(b.titulo[0] || '?')}</div>
      <div style="min-width:0;display:flex;flex-direction:column;gap:4px">
        <button data-acc="ver-libro" data-libro="${esc(b.titulo)}" style="all:unset;cursor:pointer;font:600 15px var(--serif)">${esc(b.titulo)}</button>
        <div class="contador">${esc(b.autor || q.author || '')} · ${b.citas} ${b.citas === 1 ? 'cita' : 'citas'} · ${b.paginas ? b.paginas + ' págs.' : 'páginas sin dato'}</div>
        <div class="rejilla2">
          <label>Estado<select data-libro="${esc(b.titulo)}" data-autor="${esc(b.autor || q.author || '')}" data-campo-libro="estado">${D.ESTADOS.map(k => `<option value="${k}"${b.estado === k ? ' selected' : ''}>${D.ETIQUETAS_ESTADO[k]}</option>`).join('')}</select></label>
          <label>Páginas<input type="number" min="0" value="${b.paginas}" data-libro="${esc(b.titulo)}" data-autor="${esc(b.autor || q.author || '')}" data-campo-libro="paginas"></label>
        </div>
        ${b.enReto ? `<div class="contador">≈ ${esc(b.diasLabel)} a tu ritmo · Repasos: ${esc(b.repasoLabel)}</div>
          <label class="fila">Leídas<input type="number" min="0" max="${b.paginas}" value="${b.leidas}" style="width:80px" data-libro="${esc(b.titulo)}" data-autor="${esc(b.autor || q.author || '')}" data-campo-libro="leidas"></label>
          <div class="barra"><i style="width:${b.pct}%"></i></div>` : ''}
        <div class="apa">${apaHtml(q)}</div>
      </div>
    </div>`;
  }).join('');
  return `<h2>Biblioteca</h2><nav class="pestanas" style="padding:0">${tabs}</nav>${reto}${filas || '<p class="vacio">No hay libros en este estante.</p>'}`;
}

function vistaColecciones() {
  const cols = D.resumenColecciones(st.modelo.citas, st.modelo.colecciones);
  return `<h2>Colecciones</h2>
  <form class="caja" id="form-coleccion">
    <label>Nombre<input name="nombre" required></label>
    <label>Descripción<input name="descripcion"></label>
    <div><button class="primario" type="submit">Crear colección</button></div>
  </form>
  ${cols.map(c => `<section class="caja">
    <div class="acciones"><h3>${esc(c.nombre)}</h3><span class="contador">${c.count} ${c.count === 1 ? 'cita' : 'citas'}</span>
      <button class="mini" style="margin-left:auto" data-acc="ver-coleccion" data-nombre="${esc(c.nombre)}"${c.count ? '' : ' disabled'}>Ver citas</button></div>
    ${c.descripcion ? `<div class="contador">${esc(c.descripcion)}</div>` : ''}
    ${c.muestras.map(m => `<div class="cita" style="font-size:13px;margin:0">${esc(m)}</div>`).join('')}
  </section>`).join('') || '<p class="vacio">Aún no hay colecciones.</p>'}`;
}

function vistaCuaderno() {
  const nb = estructuraCuaderno(st.modelo.citas, opcionesCuaderno(st.modelo.config, Date.now()));
  const ocupado = !!st.progresoCuaderno;
  return `<h2>Cuaderno A5</h2>
  <section class="caja">
    <div>${esc(nb.resumen)}</div>
    <div class="contador">${nb.pages.length} páginas · ${nb.pages.length / 2} hojas A5 a doble cara · ${nb.porPagina} ${nb.porPagina === 1 ? 'cita' : 'citas'} por página${nb.espejo ? ' · pestañas en espejo' : ''}. Año, citas por página y espejo se cambian en ⚙.</div>
    ${nb.isEmpty ? '<div class="contador">Sin citas guardadas ese año: el cuaderno sale con plantillas en blanco para escribir a mano. Pulsa «Guardar» en una cita para que aparezca en su letra.</div>' : ''}
    <div class="acciones">
      <button class="primario" data-acc="generar-cuaderno"${ocupado || !enExcel ? ' disabled' : ''}>${ocupado ? esc(st.progresoCuaderno) : 'Generar cuaderno'}</button>
    </div>
    <div class="contador">Crea o rehace la hoja «Cuaderno». Esa hoja la escribe ReadCite: lo que edites a mano en ella se pierde al regenerarla.</div>
  </section>
  <section class="caja">
    <h3>Imprimir</h3>
    <label>Qué imprimir<select id="letra-impresion"><option value="">Todo el cuaderno</option>${NB_LETTERS.map(l => `<option value="${l}"${st.letraImpresion === l ? ' selected' : ''}>Solo la letra ${l} (${nb.byL[l].length} ${nb.byL[l].length === 1 ? 'cita' : 'citas'})</option>`).join('')}</select></label>
    <div><button data-acc="area-impresion"${enExcel ? '' : ' disabled'}>Preparar área de impresión</button></div>
    <ol class="contador" style="margin:0;padding-left:18px">
      <li>Pulsa «Generar cuaderno» después de guardar o quitar citas.</li>
      <li>Elige qué imprimir y pulsa «Preparar área de impresión».</li>
      <li>⌘P → Tamaño de papel <b>A5</b>, <b>Doble cara</b>: encuadernación por el <b>borde largo</b>, escala 100 %.</li>
      <li>Cada letra empieza en anverso: puedes imprimir solo una letra sin descuadrar las demás.</li>
    </ol>
  </section>
  ${enExcel ? '' : '<p class="contador">El cuaderno se genera dentro de Excel; en el modo demostración solo ves el resumen.</p>'}`;
}

function vistaAjustes() {
  const c = st.modelo.config;
  return `<h2>Ajustes</h2>
  <form class="caja" id="form-ajustes">
    <label>Tu nombre (para «Solo mías» y los comentarios)<input name="usuario_local" value="${esc(c.usuario_local)}"></label>
    <label>Páginas al día del reto<input name="paginas_dia" type="number" min="1" value="${c.paginas_dia}"></label>
    <div class="rejilla2">
      <label>Citas por página (cuaderno)<select name="cuaderno_citas_por_pagina">${[1, 2, 3].map(n => `<option${+c.cuaderno_citas_por_pagina === n ? ' selected' : ''}>${n}</option>`).join('')}</select></label>
      <label>Año del cuaderno<input name="cuaderno_anio" value="${esc(c.cuaderno_anio)}" placeholder="actual"></label>
    </div>
    <label class="fila"><input type="checkbox" name="cuaderno_pestanas_espejo"${c.cuaderno_pestanas_espejo ? ' checked' : ''}>Pestañas en espejo (doble cara)</label>
    <div><button class="primario" type="submit">Guardar ajustes</button></div>
  </form>
  <section class="caja">
    <h3>Estructura del libro</h3>
    <div class="contador">Si borraste una hoja (Citas, Libros, Colecciones, Comentarios o Config), esto la vuelve a crear vacía. No toca nada que exista.</div>
    <div class="acciones">${st.confirmarReparar
      ? '<span class="error">¿Crear las hojas y tablas que falten?</span><button class="primario mini" data-acc="reparar-si">Sí</button><button class="mini" data-acc="reparar-no">No</button>'
      : '<button data-acc="reparar">Reparar estructura</button>'}</div>
  </section>
  <section class="caja">
    <h3>Acerca de</h3>
    <div class="contador">ReadCite para Excel ${VERSION} · © 2026 Walter Daniel Castillo Talancha · licencia <a href="https://creativecommons.org/licenses/by/4.0/deed.es" target="_blank" rel="noopener noreferrer">CC BY 4.0</a> (puedes copiarlo y modificarlo citando al autor e indicando los cambios).</div>
    <div class="contador">Versión oficial: id. <code>b790dff5…f36686</code> · firma <code>SHA256:26a3bP3v…6q0</code>. Si tu copia viene de otra fuente, compruébala con <code>scripts/verificar_version.sh</code>.</div>
  </section>
  <p class="contador">${enExcel ? 'Conectado a Excel.' : 'Modo demostración: los cambios viven solo en esta pestaña.'}</p>`;
}

// ───────── eventos ─────────
document.addEventListener('click', async e => {
  const b = /** @type {HTMLElement|null} */ (/** @type {HTMLElement} */ (e.target).closest('[data-vista],[data-acc]'));
  if (!b) return;
  if (b.dataset.vista) {
    if (b.dataset.vista === 'nueva' && st.vista !== 'nueva' && !st.editId) st.form = { ...st.form };
    st.vista = b.dataset.vista; pintar(); return;
  }
  const id = b.dataset.id || '';
  const ahora = Date.now();
  switch (b.dataset.acc) {
    case 'girar': st.reverso.has(id) ? st.reverso.delete(id) : st.reverso.add(id); pintarLista(); break;
    case 'like': await hacer(ctx => X.alternarMeGusta(ctx, id)); break;
    case 'guardar': await hacer(ctx => X.guardarCita(ctx, id, ahora), 'Guardada'); break;
    case 'quitar': await hacer(ctx => X.quitarGuardada(ctx, id), 'Quitada de guardadas'); break;
    case 'png': { const c = st.modelo.citas.find((/** @type {D.Cita} */ x) => x.id === id); try { toast(await guardarPng(c, insertarPng())); } catch (err) { toast(err instanceof Error ? err.message : String(err), true); } break; }
    case 'png-form': {
      const f = st.form; try { toast(await guardarPng({ text: f.text || 'Tu cita textual', author: f.author || 'Autor', book: f.book || '', topic: f.topic || 'Sin tema', ratio: f.ratio }, insertarPng())); }
      catch (err) { toast(err instanceof Error ? err.message : String(err), true); } break;
    }
    case 'formato': st.form.ratio = b.dataset.ratio; document.querySelectorAll('[data-acc=formato]').forEach(x => x.setAttribute('aria-pressed', String(/** @type {HTMLElement} */ (x).dataset.ratio === st.form.ratio))); pintarVistaPrevia(); break;
    case 'copiar-apa': { const c = st.modelo.citas.find((/** @type {D.Cita} */ x) => x.id === id); await copiar(`“${c.text}”\n— ${c.author}, ${c.book}${c.page ? ', p. ' + c.page : ''}\n\n${apaText(c)}`); break; }
    case 'comentarios': st.comentarios.has(id) ? st.comentarios.delete(id) : st.comentarios.add(id); pintarLista(); break;
    case 'comentar': { const t = /** @type {HTMLTextAreaElement} */ ($('#com-' + CSS.escape(id))).value; await hacer(ctx => X.anadirComentario(ctx, id, t, ahora), 'Comentario añadido'); break; }
    case 'editar': { const c = st.modelo.citas.find((/** @type {D.Cita} */ x) => x.id === id); st.editId = id; st.form = { ...c }; st.vista = 'nueva'; pintar(); break; }
    case 'borrar': st.confirmarBorrar = id; pintarLista(); break;
    case 'borrar-no': st.confirmarBorrar = null; pintarLista(); break;
    case 'borrar-si': st.confirmarBorrar = null; if (await hacer(ctx => X.borrarCita(ctx, id), 'Cita borrada')) await refrescarInicio(); break;
    case 'pintar-inicio': await pintarInicio(); break;
    case 'hoja-buscar': try { await ejecutar(ctx => generarBuscar(ctx)); toast('Hoja Buscar lista: escribe arriba o elige en los desplegables.'); } catch (err) { toast(err instanceof Error ? err.message : String(err), true); } break;
    case 'destacar': await pintarInicio(id); break;
    case 'cancelar-edicion': st.editId = null; st.form = {}; st.vista = 'citas'; pintar(); break;
    case 'vaciar-form': st.form = {}; pintar(); break;
    case 'limpiar': st.filtros = { ...FILTROS0, sort: st.filtros.sort }; pintar(); break;
    case 'libtab': st.libTab = b.dataset.tab || 'leyendo'; pintar(); break;
    case 'ver-libro': st.filtros = { ...FILTROS0, book: b.dataset.libro || '' }; st.vista = 'citas'; pintar(); break;
    case 'ver-coleccion': st.filtros = { ...FILTROS0, collection: b.dataset.nombre || '' }; st.vista = 'citas'; pintar(); break;
    case 'generar-cuaderno': {
      st.progresoCuaderno = 'Preparando…'; pintar();
      try {
        const r = await ejecutar(ctx => generarCuaderno(ctx, st.modelo, { ahora, medir, alProgresar: (h, t) => { st.progresoCuaderno = `Generando… ${h}/${t} págs.`; pintar(); } }));
        toast(`Cuaderno ${r.anio}: ${r.paginas} páginas (${r.hojas} hojas), ${r.citas} citas.`);
      } catch (err) { toast(err instanceof Error ? err.message : String(err), true); }
      st.progresoCuaderno = ''; pintar(); break;
    }
    case 'area-impresion': {
      const letra = /** @type {HTMLSelectElement} */ ($('#letra-impresion')).value || null; st.letraImpresion = letra || '';
      try { const r = await ejecutar(ctx => areaImpresion(ctx, st.modelo, letra, ahora)); toast(`Área de impresión: ${letra ? 'letra ' + letra : 'todo el cuaderno'} (${r}). Ahora ⌘P.`); }
      catch (err) { toast(err instanceof Error ? err.message : String(err), true); }
      break;
    }
    case 'reparar': st.confirmarReparar = true; pintar(); break;
    case 'reparar-no': st.confirmarReparar = false; pintar(); break;
    case 'reparar-si': {
      st.confirmarReparar = false;
      try { const hecho = await ejecutar(ctx => X.repararEstructura(ctx)); toast(hecho.length ? 'Creado: ' + hecho.join(', ') : 'La estructura está completa; no hacía falta nada.'); await recargar(); }
      catch (err) { toast(err instanceof Error ? err.message : String(err), true); }
      break;
    }
  }
});

let tBusca = 0;
document.addEventListener('input', e => {
  const el = /** @type {HTMLInputElement} */ (e.target);
  if (el.id === 'q') { st.filtros.query = el.value; clearTimeout(tBusca); tBusca = window.setTimeout(pintarLista, 120); }
  else if (el.dataset.campo) { st.form[el.dataset.campo] = el.value; pintarApaForm(); pintarVistaPrevia(); }
});

document.addEventListener('change', async e => {
  const el = /** @type {HTMLInputElement} */ (e.target);
  if (el.dataset.filtro) {
    const k = el.dataset.filtro; /** @type {any} */ (st.filtros)[k] = el.type === 'checkbox' ? el.checked : el.value;
    pintar(); return;
  }
  if (el.dataset.campoLibro) {
    const titulo = el.dataset.libro || ''; const campo = el.dataset.campoLibro;
    const valor = campo === 'estado' ? el.value : Math.max(0, parseInt(el.value) || 0);
    await hacer(ctx => X.actualizarLibro(ctx, titulo, { [campo]: valor, ...(st.modelo.libros.some((/** @type {any} */ l) => l.titulo === titulo) ? {} : { autor: el.dataset.autor }) }));
    return;
  }
  if (el.id === 'foto' && el.files && el.files[0]) { const f = el.files[0]; el.value = ''; await transcribir(f); return; }
  if (el.dataset.accCambio === 'paginas-dia') await hacer(ctx => X.guardarConfig(ctx, { paginas_dia: parseInt(el.value) || 20 }));
});

document.addEventListener('submit', async e => {
  e.preventDefault();
  const f = /** @type {HTMLFormElement} */ (e.target);
  if (f.id === 'formulario') {
    const datos = { ...st.form };
    const ok = st.editId
      ? await hacer(ctx => X.editarCita(ctx, /** @type {string} */ (st.editId), datos), 'Cambios guardados')
      : await hacer(ctx => X.crearCita(ctx, datos, Date.now()), 'Cita guardada en tblCitas');
    if (ok) { st.editId = null; st.form = {}; st.filtros = { ...FILTROS0, sort: 'recent' }; st.vista = 'citas'; pintar(); await refrescarInicio(); }
  } else if (f.id === 'form-coleccion') {
    const d = new FormData(f);
    await hacer(ctx => X.crearColeccion(ctx, { nombre: String(d.get('nombre')), descripcion: String(d.get('descripcion') || '') }, Date.now()), 'Colección creada');
  } else if (f.id === 'form-ajustes') {
    const d = new FormData(f);
    const cambios = [['usuario_local', String(d.get('usuario_local') || '').trim() || 'Tú'], ['paginas_dia', parseInt(String(d.get('paginas_dia'))) || 20],
      ['cuaderno_citas_por_pagina', parseInt(String(d.get('cuaderno_citas_por_pagina'))) || 2], ['cuaderno_anio', String(d.get('cuaderno_anio') || '').trim()],
      ['cuaderno_pestanas_espejo', d.get('cuaderno_pestanas_espejo') === 'on']];
    await hacer(ctx => X.guardarConfig(ctx, Object.fromEntries(cambios)), 'Ajustes guardados');
  }
});

// ───────── arranque ─────────
const DEMO = [
  { text: 'El Estado es la gran ficción a través de la cual todo el mundo se esfuerza en vivir a expensas de todo el mundo.', author: 'Frédéric Bastiat', book: 'El Estado', year: '2004', originalYear: '1848', publisher: 'Alianza Editorial', page: '9', topic: 'Estado', concept: 'Expolio legal', understood: 'Cada grupo pide privilegios creyendo que los paga «otro». Como todos lo hacen, todos acaban pagando más de lo que reciben.', why: 'Una sola frase que explica la lógica de los grupos de presión.', collection: 'El Estado bajo la lupa' },
  { text: 'La acción humana es conducta consciente y deliberada.', author: 'Ludwig von Mises', book: 'La acción humana', year: '2011', originalYear: '1949', publisher: 'Unión Editorial', translator: 'J. Reig Albiol', chapter: 'I', page: '15', topic: 'Praxeología', concept: 'Acción humana', understood: 'Toda acción parte de un fin elegido y de medios escasos.', collection: 'Mises sobre propiedad y acción' },
  { text: 'Cuanto más planifica el Estado, más difícil se hace la planificación para el individuo.', author: 'Friedrich A. Hayek', book: 'Camino de servidumbre', year: '2008', originalYear: '1944', publisher: 'Unión Editorial', translator: 'José Vergara', chapter: 'VI', page: '112', topic: 'Planificación', concept: 'Orden espontáneo', understood: 'La planificación central no sustituye a la individual: la anula.', collection: 'Orden espontáneo' },
];
async function arrancarDemo() {
  sim = libroSimulado();
  const t0 = Date.now() - 3 * 86400000;
  for (const [i, q] of DEMO.entries()) await X.crearCita(sim.ctx, q, t0 + i * 3600000);
  const m = await X.leerTodo(sim.ctx);
  await X.guardarCita(sim.ctx, m.citas.find((/** @type {D.Cita} */ c) => c.author === 'Frédéric Bastiat').id, Date.now());
  await X.actualizarLibro(sim.ctx, 'Camino de servidumbre', { estado: 'quiero', paginas: 320, leidas: 40 });
  await X.actualizarLibro(sim.ctx, 'El Estado', { estado: 'quiero', paginas: 64 });
  const modo = $('#modo'); modo.textContent = 'Demo'; modo.title = 'Fuera de Excel: datos de ejemplo en memoria'; modo.hidden = false;
}

async function arrancar(/** @type {boolean} */ excel) {
  enExcel = excel;
  try { if (!enExcel) await arrancarDemo(); await recargar(); }
  catch (e) { $('#vista').innerHTML = `<p class="error">No pude leer el libro: ${esc(e instanceof Error ? e.message : e)}</p>`; }
}
// @ts-ignore
if (typeof Office !== 'undefined' && Office.onReady) Office.onReady(info => arrancar(info.host === Office.HostType.Excel));
else arrancar(false);
