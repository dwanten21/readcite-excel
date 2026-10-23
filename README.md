# ReadCite para Excel

Complemento de Excel para guardar **citas textuales de libros**: autor, obra, lo que entendiste y por qué te
llamó la atención; referencia **APA 7** automática, colecciones, biblioteca de lecturas con **reto de lectura**
y repaso **Ícaro**, búsqueda con filtros, transcripción desde una **foto** de la página, tarjetas **PNG**
(póster 2:3 o cuadrada) y un **cuaderno A5 imprimible** ordenado de la A a la Z, con silabeo español.

Funciona en Excel para **Windows, Mac y la web** (Microsoft 365). Tus citas se guardan en tu propio archivo
`.xlsx`; el complemento no envía tus datos a ningún servidor.

- **Instalar:** [docs/instalar.md](docs/instalar.md) — manifiesto: [manifest.xml](https://raw.githubusercontent.com/dwanten21/readcite-excel/main/manifest.xml)
- **¿Es oficial mi copia?** [docs/es-oficial.md](docs/es-oficial.md)
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
