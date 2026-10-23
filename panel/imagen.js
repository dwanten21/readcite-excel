// Prepara la foto para el OCR: Tesseract solo entiende JPG/PNG/BMP, así que cualquier foto (incluida la HEIC del
// iPhone) se dibuja en un lienzo y se entrega como PNG, ya girada según su orientación, a un tamaño razonable y con
// la iluminación aplanada (las zonas en sombra de la página se leen peor).

import { aplanarLuz } from '../core/ocr.js';

const LADO_MAX = 3000;

/** @param {File} f */
export const esHeic = f => /hei[cf]/i.test(f.type) || /\.(heic|heif)$/i.test(f.name);

/** Decodifica con el propio navegador (Safari/Excel para Mac leen HEIC; Chromium/Edge no).
 * @param {Blob} blob @returns {Promise<ImageBitmap | HTMLImageElement | null>} */
async function decodificar(blob) {
  try { return await createImageBitmap(blob, { imageOrientation: 'from-image' }); } catch { /* sigue con <img> */ }
  const url = URL.createObjectURL(blob);
  try {
    return await new Promise(res => { const im = new Image(); im.onload = () => res(im); im.onerror = () => res(null); im.src = url; });
  } finally { URL.revokeObjectURL(url); }
}

/**
 * @param {File} archivo
 * @returns {Promise<Blob>} PNG listo para el OCR
 */
export async function imagenLegible(archivo) {
  let img = await decodificar(archivo);
  if (!img && esHeic(archivo)) {
    // Excel en Windows/web no abre HEIC: se convierte aquí mismo con heic-to (libheif en WebAssembly, sin eval,
    // incluido en el complemento y cubierto por la firma oficial). Excel para Mac no llega a cargarlo.
    const { heicTo } = await import('./vendor/heic-to/heic-to.js');
    img = await decodificar(await heicTo({ blob: archivo, type: 'image/jpeg', quality: 0.95 }));
  }
  if (!img) throw new Error('formato');
  const ancho = img.width, alto = img.height;
  const k = Math.min(1, LADO_MAX / Math.max(ancho, alto));
  const lienzo = document.createElement('canvas');
  lienzo.width = Math.round(ancho * k); lienzo.height = Math.round(alto * k);
  const ctx = /** @type {CanvasRenderingContext2D} */ (lienzo.getContext('2d'));
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, lienzo.width, lienzo.height);
  ctx.drawImage(img, 0, 0, lienzo.width, lienzo.height);
  if ('close' in img) img.close();
  const datos = ctx.getImageData(0, 0, lienzo.width, lienzo.height);
  aplanarLuz(datos.data, lienzo.width, lienzo.height);
  ctx.putImageData(datos, 0, 0);
  return new Promise((res, rej) => lienzo.toBlob(b => (b ? res(b) : rej(new Error('formato'))), 'image/png'));
}
