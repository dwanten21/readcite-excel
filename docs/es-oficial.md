# ¿Es oficial mi copia de ReadCite?

ReadCite es libre (licencia [CC BY 4.0](../LICENSE.md)): cualquiera puede copiarlo, modificarlo y volver a
publicarlo. Eso es bueno, pero también significa que alguien podría repartir una copia **alterada** con
código dañino. Esta guía te permite comprobar en un minuto que tienes la versión **oficial e intacta**.

## Datos de la versión oficial

| | |
|---|---|
| Autor | Walter Daniel Castillo Talancha |
| Id. del complemento (en `manifest.xml`, etiqueta `<Id>`) | `b790dff5-9208-4d2b-a842-bd4a62f36686` |
| Repositorio oficial | https://github.com/dwanten21/readcite-excel |
| Dirección oficial del panel | https://dwanten21.github.io/readcite-excel/panel/index.html |
| Manifiesto oficial | https://github.com/dwanten21/readcite-excel/blob/main/manifest.xml |
| Huella de la clave de firma | `SHA256:26a3bP3v/11QPGHSRojptkKjrKrn1GSAKz7YFnbk6q0` |

Si alguien te pasa un `manifest.xml` cuya dirección no es la oficial, **no es la versión oficial**: puede ser una
adaptación legítima (la licencia lo permite, y debe decir que está modificada) o algo peor. Úsala solo si confías
en quien te la da.

## Comprobar una copia descargada (macOS o Linux)

Desde la carpeta del complemento:

```bash
sh scripts/verificar_version.sh
```

Debe terminar con **«Copia oficial e intacta.»**. El script comprueba tres cosas:

1. que la lista de archivos (`firma/SHA256SUMS`) está **firmada con la clave del autor**;
2. que **ningún archivo** ha cambiado (huella SHA-256 de cada uno);
3. que **no se ha añadido** ningún archivo de código.

Comprueba además que la huella de `firma/readcite_firma.pub` coincide con la de la tabla de arriba:

```bash
ssh-keygen -lf firma/readcite_firma.pub
```

(Si alguien cambió también la clave pública, esta huella no coincidirá con la publicada aquí y en la web oficial.)

## Qué protege y qué no

- ✅ Detecta cualquier cambio, archivo añadido o lista de huellas falsificada.
- ✅ La versión oficial bloquea fórmulas maliciosas en tus citas, solo ejecuta su propio código (política de
  seguridad CSP) y comprueba el sello de integridad del OCR que descarga.
- ❌ No impide que alguien publique una copia modificada: la licencia lo permite. Lo que impide es que te la
  hagan pasar por la oficial sin que puedas notarlo.

## Para el autor: publicar una versión

```bash
sh scripts/firmar_version.sh
```

Firma con la clave privada `~/.readcite/firma/readcite_firma`. **Esa clave no se publica nunca** y conviene
guardar una copia de seguridad cifrada: si se pierde, las versiones nuevas no se podrán firmar con la misma
identidad; si alguien la roba, podría firmar versiones falsas (habría que crear una clave nueva y avisar).
