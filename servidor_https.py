"""Servidor local HTTPS del panel de ReadCite para Excel: https://localhost:3000/panel/index.html

Solo sirve las carpetas panel/, core/ y excel/ (nunca el .xlsx ni las copias de seguridad)
y solo escucha en 127.0.0.1. Certificado: ~/.readcite/certs/localhost.{crt,key}.

Uso:  python3 servidor_https.py        (Ctrl+C para parar)
"""
import http.server
import json
import os
import re
import ssl
import sys
from urllib.parse import parse_qs, urlparse

RAIZ = os.path.dirname(os.path.abspath(__file__))
CERTS = os.path.expanduser("~/.readcite/certs")
PUERTO = 3000
PERMITIDAS = ("/panel/", "/core/", "/excel/")
DESCARGAS = os.path.expanduser("~/Downloads/ReadCite")
NOMBRE_PNG = re.compile(r"^readcite-[a-z0-9-]{1,80}\.png$")
MAX_PNG = 15 * 1024 * 1024
ORIGEN = f"https://localhost:{PUERTO}"


class Manejador(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map, ".js": "text/javascript", ".mjs": "text/javascript"}

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=RAIZ, **kwargs)

    def send_head(self):
        ruta = self.path.split("?", 1)[0]
        if not ruta.startswith(PERMITIDAS) or ".." in ruta or ruta.endswith("/"):
            self.send_error(404, "No disponible")
            return None
        return super().send_head()

    def do_POST(self):
        """POST /guardar-png?nombre=readcite-autor.png  (cuerpo = bytes PNG) → ~/Downloads/ReadCite/.
        Solo desde el propio panel (Origin https://localhost:3000), nombre fijo, PNG real y ≤ 15 MB; nunca sobrescribe."""
        url = urlparse(self.path)
        if url.path != "/guardar-png":
            return self._json(404, {"error": "No disponible"})
        if self.headers.get("Origin") != ORIGEN:
            return self._json(403, {"error": "Origen no permitido"})
        nombre = (parse_qs(url.query).get("nombre") or [""])[0]
        if not NOMBRE_PNG.match(nombre):
            return self._json(400, {"error": "Nombre de archivo no válido"})
        largo = int(self.headers.get("Content-Length") or 0)
        if not 0 < largo <= MAX_PNG:
            return self._json(413, {"error": "Imagen vacía o demasiado grande"})
        datos = self.rfile.read(largo)
        if not datos.startswith(b"\x89PNG\r\n\x1a\n"):
            return self._json(400, {"error": "No es un PNG"})
        os.makedirs(DESCARGAS, exist_ok=True)
        base, ext = os.path.splitext(nombre)
        ruta, n = os.path.join(DESCARGAS, nombre), 2
        while os.path.exists(ruta):
            ruta, n = os.path.join(DESCARGAS, f"{base}-{n}{ext}"), n + 1
        with open(ruta, "xb") as f:
            f.write(datos)
        return self._json(200, {"ruta": ruta.replace(os.path.expanduser("~"), "~")})

    def _json(self, codigo, obj):
        cuerpo = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.end_headers()
        self.wfile.write(cuerpo)

    def end_headers(self):
        # Sin caché: Excel siempre carga la última versión del panel.
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, formato, *args):
        sys.stderr.write("[readcite] %s\n" % (formato % args))


def main():
    crt, key = os.path.join(CERTS, "localhost.crt"), os.path.join(CERTS, "localhost.key")
    if not (os.path.exists(crt) and os.path.exists(key)):
        sys.exit(f"Falta el certificado en {CERTS}. Vuelve a generarlo (ver docs/instalacion-mac.md).")
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(crt, key)
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", PUERTO), Manejador)
    srv.socket = ctx.wrap_socket(srv.socket, server_side=True)
    print(f"ReadCite sirviendo en https://localhost:{PUERTO}/panel/index.html  (Ctrl+C para parar)", flush=True)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
