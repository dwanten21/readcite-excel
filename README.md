# ReadCite para Excel

Complemento de Excel para guardar **citas textuales de libros**: autor, obra, lo que entendiste y por qué te
llamó la atención; referencia **APA 7** automática, colecciones, biblioteca de lecturas con **reto de lectura**
y repaso **Ícaro**, búsqueda con filtros, transcripción desde una **foto** de la página, tarjetas **PNG**
(póster 2:3 o cuadrada) y un **cuaderno A5 imprimible** ordenado de la A a la Z, con silabeo español.

Funciona en Excel para **Windows, Mac y la web** (Microsoft 365). Tus citas se guardan en tu propio archivo
`.xlsx`; el complemento no envía tus datos a ningún servidor.

## 👉 Cómo instalarlo

**Paso 1 — Descarga ReadCite (un clic):**

[![Descargar ReadCite](https://img.shields.io/badge/⬇%20%20Descargar%20ReadCite-23964d?style=for-the-badge)](https://github.com/dwanten21/readcite-excel/releases/latest/download/ReadCite.xml)

Se guarda un archivo pequeño, `ReadCite.xml`, en tu carpeta **Descargas**.

**Paso 2 — Ábrelo en Excel:** sigue los pasos para tu Excel (navegador, Windows o Mac) en la página de instalación:

### ➡️ [dwanten21.github.io/readcite-excel](https://dwanten21.github.io/readcite-excel/)

> Excel no permite instalar complementos con un solo clic salvo desde la tienda de Microsoft, por eso hay
> que cargar el archivo una vez. Después, ReadCite queda en la pestaña **Inicio** de Excel.

## Más información

- Guía detallada de instalación: [docs/instalar.md](docs/instalar.md)
- **¿Es oficial mi copia?** [docs/es-oficial.md](docs/es-oficial.md) — descarga ReadCite solo desde aquí.
- **Imprimir el cuaderno A5:** [docs/imprimir-cuaderno.md](docs/imprimir-cuaderno.md)
- **Licencia:** [CC BY 4.0](LICENSE.md) — puedes copiarlo y modificarlo citando al autor e indicando los cambios.

## Estructura

| Carpeta | Qué hay |
|---|---|
| `panel/` | Panel lateral (HTML, CSS y JS sin compilar) |
| `core/` | Lógica pura: APA 7, silabeo, nombres, datos, reto, cuaderno, maquetado, tarjetas |
| `excel/` | Capa Office.js: tablas, hojas Inicio, Buscar y Cuaderno |
| `tests/` | Tests (`node --test tests/`, sin dependencias) |
| `firma/` | Lista de huellas SHA-256 firmada de la versión oficial y clave pública |
| `scripts/` | Firmar/verificar versiones y generar el manifiesto |

## Desarrollo

```bash
node --test tests/                       # tests
python3 servidor_https.py                # panel en https://localhost:3000 (manifest.local.xml)
sh scripts/verificar_version.sh          # comprobar que la copia es la oficial
```

© 2026 Walter Daniel Castillo Talancha · CC BY 4.0
