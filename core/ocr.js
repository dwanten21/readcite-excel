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
