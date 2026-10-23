// @ts-check
// Limpieza del texto que devuelve el OCR de una foto de página (equivale a lo que la web pedía a la IA:
// «une las palabras partidas por guion de fin de línea» y devuelve el texto seguido).

const LETRA = 'A-Za-zÁÉÍÓÚÜÑáéíóúüñ';

/**
 * - «pala-\nbra» → «palabra» (guion de fin de línea seguido de minúscula; también guion blando o «‐»).
 * - Saltos de línea → espacios; espacios repetidos → uno; comillas rectas dobles → tipográficas.
 * - Los guiones de verdad («anti-Estado», «1848-1849») y los de una línea a otra con mayúscula se conservan.
 * @param {string} texto
 */
export function limpiarTranscripcion(texto) {
  return (texto || '')
    .replace(/\r/g, '')
    .replace(new RegExp(`([${LETRA}])[-‐­]\\s*\\n\\s*([a-záéíóúüñ])`, 'g'), '$1$2')
    .replace(/\s*\n\s*/g, ' ')
    .replace(/[ \t ]{2,}/g, ' ')
    .replace(/"([^"]+)"/g, '“$1”')
    .replace(/\s+([,.;:!?»”)])/g, '$1')
    .trim();
}

/** Nombre del PNG, como la web: «readcite-» + autor en minúsculas con lo que no sea a-z0-9 cambiado por «-». @param {string} autor */
export const nombrePng = autor => ('readcite-' + (autor || 'cita')).toLowerCase().replace(/[^a-z0-9]+/g, '-') + '.png';

/**
 * Aplana la iluminación de una foto de página antes del OCR: pasa a gris y divide cada píxel por el «blanco del
 * papel» de su zona (máximo por bloques de `lado` px, interpolado). Así una esquina en sombra queda tan clara como
 * el resto y las letras conservan el contraste. Modifica `rgba` (datos de un ImageData) en su sitio.
 * @param {Uint8ClampedArray} rgba @param {number} ancho @param {number} alto @param {number} [lado]
 */
export function aplanarLuz(rgba, ancho, alto, lado = 64) {
  const n = ancho * alto, gris = new Float32Array(n);
  for (let i = 0; i < n; i++) gris[i] = 0.299 * rgba[4 * i] + 0.587 * rgba[4 * i + 1] + 0.114 * rgba[4 * i + 2];
  const bx = Math.ceil(ancho / lado), by = Math.ceil(alto / lado), papel = new Float32Array(bx * by);
  for (let y = 0; y < alto; y++) {
    const fila = Math.floor(y / lado) * bx;
    for (let x = 0; x < ancho; x++) { const b = fila + Math.floor(x / lado); if (gris[y * ancho + x] > papel[b]) papel[b] = gris[y * ancho + x]; }
  }
  for (let y = 0; y < alto; y++) {
    const fy = Math.min(by - 1, Math.max(0, y / lado - 0.5)), y0 = Math.floor(fy), y1 = Math.min(by - 1, y0 + 1), ay = fy - y0;
    for (let x = 0; x < ancho; x++) {
      const fx = Math.min(bx - 1, Math.max(0, x / lado - 0.5)), x0 = Math.floor(fx), x1 = Math.min(bx - 1, x0 + 1), ax = fx - x0;
      const blanco = (papel[y0 * bx + x0] * (1 - ax) + papel[y0 * bx + x1] * ax) * (1 - ay)
                   + (papel[y1 * bx + x0] * (1 - ax) + papel[y1 * bx + x1] * ax) * ay;
      const i = y * ancho + x, v = Math.min(255, (255 * gris[i]) / Math.max(1, blanco));
      rgba[4 * i] = rgba[4 * i + 1] = rgba[4 * i + 2] = v;
    }
  }
}
