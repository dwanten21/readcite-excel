#!/bin/sh
# Firma una versión oficial de ReadCite: calcula la huella SHA-256 de cada archivo que se publica y firma
# esa lista con la clave privada del autor (~/.readcite/firma/readcite_firma, que nunca sale de su Mac).
# Uso (desde la carpeta del proyecto):  sh scripts/firmar_version.sh
set -eu
cd "$(dirname "$0")/.."
CLAVE="${READCITE_CLAVE:-$HOME/.readcite/firma/readcite_firma}"
[ -f "$CLAVE" ] || { echo "No encuentro la clave de firma en $CLAVE" >&2; exit 1; }

# Archivos que forman la versión publicada (panel, lógica, capa Excel, manifiesto y licencia)
find panel core excel -type f ! -name '.DS_Store' | LC_ALL=C sort > firma/.lista
printf '%s\n' manifest.xml LICENSE.md >> firma/.lista
: > firma/SHA256SUMS
while IFS= read -r f; do shasum -a 256 "$f" >> firma/SHA256SUMS; done < firma/.lista
rm firma/.lista

rm -f firma/SHA256SUMS.sig
ssh-keygen -q -Y sign -f "$CLAVE" -n readcite-version firma/SHA256SUMS
echo "Firmados $(wc -l < firma/SHA256SUMS | tr -d ' ') archivos → firma/SHA256SUMS y firma/SHA256SUMS.sig"
sh scripts/verificar_version.sh
