# Instalar ReadCite en Excel

Solo necesitas el archivo **`manifest.xml`** de la versión oficial. Descárgalo de
https://raw.githubusercontent.com/dwanten21/readcite-excel/main/manifest.xml (clic derecho → *Guardar como…*,
nombre `manifest.xml`). Antes, si quieres, comprueba que es el oficial: [es-oficial.md](es-oficial.md).

## Excel para la web (cualquier computadora)

1. Abre tu libro en excel.office.com.
2. **Inicio → Complementos → Más complementos → Mis complementos → Cargar mi complemento**.
3. Elige `manifest.xml` → **Cargar**. Aparece el botón **ReadCite** en la pestaña Inicio.

## Excel para Windows (Microsoft 365)

1. **Inicio → Complementos → Más complementos → Mis complementos → Cargar mi complemento** y elige `manifest.xml`.
2. Si tu Excel no muestra «Cargar mi complemento», usa una carpeta compartida de confianza:
   crea una carpeta, compártela (*Propiedades → Compartir*), copia ahí `manifest.xml`, y en Excel ve a
   **Archivo → Opciones → Centro de confianza → Configuración → Catálogos de complementos de confianza**,
   añade la ruta `\\NOMBRE-PC\carpeta`, marca *Mostrar en el menú* y reinicia Excel. Luego
   **Complementos → Carpeta compartida → ReadCite**.

## Excel para Mac

1. En Finder, **Ir → Ir a la carpeta…** y pega:
   `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
   (si la carpeta `wef` no existe, créala).
2. Copia ahí `manifest.xml`.
3. Cierra Excel con **⌘Q**, ábrelo y busca **Inicio → ReadCite**
   (o **Insertar → Complementos → Mis complementos → Complementos de desarrollador**).

## Primer uso

- Abre ReadCite en un libro vacío: verás avisos de que faltan tablas. Pulsa **⚙ → Reparar estructura → Sí**
  y ReadCite crea sus hojas (Citas, Libros, Colecciones, Comentarios, Config). No toca nada que ya exista.
- Las hojas **Inicio**, **Buscar** y **Cuaderno** se crean desde el panel cuando las pidas.
- **PNG:** en Excel de escritorio la tarjeta se inserta como imagen en la hoja (clic derecho → *Guardar como
  imagen*); en Excel para la web se descarga.
- **Transcribir foto:** la primera vez descarga el lector de texto (≈15 MB). La foto no sale de tu equipo.
