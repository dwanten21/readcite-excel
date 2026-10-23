// @ts-check
// Nombres exactos de hojas, tablas y columnas de ReadCiteExcel.xlsx (creados en la Fase 3).
import { COLUMNAS_CITAS, COLUMNAS_LIBROS, COLUMNAS_COLECCIONES, COLUMNAS_COMENTARIOS, COLUMNAS_CONFIG } from '../core/datos.js';

export const TABLAS = {
  citas:       { tabla: 'tblCitas',       hoja: 'Citas',       columnas: COLUMNAS_CITAS.map(([, c]) => c), obligatorias: ['id', 'texto', 'autor', 'libro', 'entendido'], clave: 'id' },
  libros:      { tabla: 'tblLibros',      hoja: 'Libros',      columnas: COLUMNAS_LIBROS,      obligatorias: ['titulo', 'estado'], clave: 'titulo' },
  colecciones: { tabla: 'tblColecciones', hoja: 'Colecciones', columnas: COLUMNAS_COLECCIONES, obligatorias: ['nombre'], clave: 'nombre' },
  comentarios: { tabla: 'tblComentarios', hoja: 'Comentarios', columnas: COLUMNAS_COMENTARIOS, obligatorias: ['cita_id', 'texto'], clave: null },
  config:      { tabla: 'tblConfig',      hoja: 'Config',      columnas: COLUMNAS_CONFIG,      obligatorias: ['clave', 'valor'], clave: 'clave' },
};

/** Letra de columna (0 → A, 23 → X, 26 → AA). @param {number} i */
export function letra(i) { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
