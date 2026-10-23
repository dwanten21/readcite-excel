#!/bin/sh
# Genera manifest.xml (el que se reparte) a partir de manifest.local.xml cambiando la dirección del panel.
# Uso:  sh scripts/generar_manifiesto.sh https://USUARIO.github.io/readcite-excel
set -eu
cd "$(dirname "$0")/.."
BASE="${1:?Indica la dirección pública, p. ej. https://usuario.github.io/readcite-excel}"
BASE="${BASE%/}"
case "$BASE" in https://*) ;; *) echo "La dirección debe empezar por https://" >&2; exit 1;; esac
ORIGEN=$(printf '%s' "$BASE" | sed -E 's#^(https://[^/]+).*#\1#')
sed -e "s#<AppDomain>https://localhost:3000</AppDomain>#<AppDomain>$ORIGEN</AppDomain>#" \
    -e "s#https://localhost:3000#$BASE#g" \
    -e 's#BASE = https://localhost:3000 (servidor local HTTPS)#Versión publicada#' manifest.local.xml > manifest.xml
echo "manifest.xml → $BASE/panel/index.html"
