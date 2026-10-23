#!/bin/sh
# Comprueba que una copia de ReadCite es la versión oficial firmada por su autor y que ningún archivo
# ha sido alterado. Funciona en macOS y Linux sin instalar nada (usa ssh-keygen y shasum/sha256sum).
# Uso (desde la carpeta del complemento):  sh scripts/verificar_version.sh
set -u
cd "$(dirname "$0")/.."
ok=1

# 1) La lista de huellas está firmada por la clave oficial
if ssh-keygen -Y verify -f firma/firmantes_autorizados -I readcite-oficial -n readcite-version \
     -s firma/SHA256SUMS.sig < firma/SHA256SUMS > /dev/null 2>&1; then
  echo "✅ Firma válida: la lista de archivos es la que publicó el autor."
else
  echo "❌ FIRMA NO VÁLIDA: esta copia no es la versión oficial (o la lista fue manipulada)."; ok=0
fi

# 2) Cada archivo coincide con su huella
if command -v shasum > /dev/null; then H="shasum -a 256 -c --quiet"; else H="sha256sum -c --quiet"; fi
if $H firma/SHA256SUMS; then
  echo "✅ Todos los archivos coinciden con la versión firmada."
else
  echo "❌ HAY ARCHIVOS ALTERADOS (listados arriba): no uses esta copia."; ok=0
fi

# 3) Archivos de código añadidos que no estaban en la versión firmada
extra=$(find panel core excel -type f ! -name '.DS_Store' | LC_ALL=C sort | while IFS= read -r f; do
  grep -q "  $f\$" firma/SHA256SUMS || echo "$f"; done)
if [ -n "$extra" ]; then echo "❌ ARCHIVOS AÑADIDOS que no son de la versión oficial:"; echo "$extra"; ok=0; fi

[ $ok = 1 ] && echo "Copia oficial e intacta." || { echo "Copia NO oficial o modificada."; exit 1; }
