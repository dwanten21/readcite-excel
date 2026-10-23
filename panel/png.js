// @ts-check
// Dibuja la tarjeta PNG (core/tarjeta.js) en un canvas y la guarda: primero en ~/Downloads/ReadCite a través
// del servidor local (los paneles de Excel no permiten descargas normales); si no hay servidor, descarga normal.
import { planTarjeta, FUENTES_PNG } from '../core/tarjeta.js';
import { nombrePng } from '../core/ocr.js';

const lienzoMedida = /** @type {CanvasRenderingContext2D} */ (document.createElement('canvas').getContext('2d'));
/** @param {{size:number,italic?:boolean,bold?:boolean,family?:string}} f */
const fuenteCss = f => `${f.italic ? 'italic ' : ''}${f.bold ? '700 ' : ''}${f.size}px ${f.family || FUENTES_PNG.serif}`;
/** @type {(t:string, f:{size:number,italic?:boolean,bold?:boolean,family?:string})=>number} */
const medirPx = (t, f) => { lienzoMedida.font = fuenteCss(f); return lienzoMedida.measureText(t).width; };

/** Espera a las fuentes de la tarjeta (Source Serif 4 y Nunito, cargadas por index.html). */
async function fuentesListas() {
  try {
    await Promise.all(['italic 400 66px "Source Serif 4"', '700 46px Nunito', '400 34px Nunito', 'italic 700 44px Nunito'].map(f => document.fonts.load(f)));
    await document.fonts.ready;
  } catch { /* si no cargan, se dibuja con Georgia / la sans del sistema */ }
}

/** Dibuja la tarjeta de una cita. @param {{text:string,author:string,book:string,topic?:string,ratio?:string}} q */
export async function dibujarTarjeta(q) {
  await fuentesListas();
  const plan = planTarjeta(q, medirPx);
  const c = document.createElement('canvas'); c.width = plan.W; c.height = plan.H;
  const g = /** @type {CanvasRenderingContext2D} */ (c.getContext('2d'));
  g.textBaseline = 'alphabetic';
  for (const o of plan.ops) {
    if (o.t === 'rect') { g.fillStyle = o.color; g.fillRect(o.x, o.y, o.w, o.h); }
    else if (o.t === 'circulo') { g.beginPath(); g.arc(o.x, o.y, o.r, 0, Math.PI * 2); g.fillStyle = o.color; g.fill(); g.lineWidth = 2; g.strokeStyle = o.borde; g.stroke(); }
    else {
      g.font = `${o.italic ? 'italic ' : ''}${o.peso || 400} ${o.size}px ${o.fuente}`; g.fillStyle = o.color;
      g.textAlign = o.alinear === 'right' ? 'right' : 'left';
      if (o.espaciado) { let x = o.x; for (const ch of o.texto) { g.fillText(ch, x, o.y); x += g.measureText(ch).width + o.espaciado; } }
      else g.fillText(o.texto, o.x, o.y);
    }
  }
  return c;
}

/** Guarda el PNG de una cita. Devuelve dónde quedó.
 * @param {{text:string,author:string,book:string,topic?:string,ratio?:string}} q
 * @param {((base64:string)=>Promise<string>)|null} [insertar] si no hay servidor local y no se puede descargar
 *   (Excel de escritorio), inserta la imagen en el libro y devuelve el mensaje para el usuario */
export async function guardarPng(q, insertar = null) {
  const c = await dibujarTarjeta(q);
  const blob = await new Promise(res => c.toBlob(res, 'image/png'));
  if (!blob) throw new Error('No se pudo generar el PNG');
  const nombre = nombrePng(q.author);
  /** @type {Response|null} */ let r = null;
  try { r = await fetch(`/guardar-png?nombre=${encodeURIComponent(nombre)}`, { method: 'POST', headers: { 'Content-Type': 'image/png' }, body: blob }); }
  catch { r = null; } // sin servidor: descarga normal
  if (r && (r.headers.get('Content-Type') || '').startsWith('application/json')) {
    const j = await r.json();
    if (r.ok && j.ruta) return `Guardado en ${j.ruta}`;
    if (r.status !== 404) throw new Error(`No se pudo guardar el PNG: ${j.error || r.status}`);
  }
  if (insertar) return insertar(c.toDataURL('image/png').split(',')[1]);
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nombre; document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  return `Descargado: ${nombre}`;
}
