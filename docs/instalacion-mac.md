# Instalar el panel de ReadCite en Excel para Mac (modo local)

El panel se sirve desde tu propio Mac en `https://localhost:3000`. Nada sale de tu equipo:
tus citas viven solo en `ReadCiteExcel.xlsx`.

## Piezas

| Pieza | Dónde |
|---|---|
| Certificado del servidor | `~/.readcite/certs/localhost.crt` y `.key` (válido hasta nov. 2027) |
| CA local (solo para `localhost` / `127.0.0.1`) | `~/.readcite/certs/readcite-ca.crt` — su clave privada se borró al crearla |
| Servidor | `servidor_https.py` (solo sirve `panel/`, `core/` y `excel/`; escucha en 127.0.0.1) |
| Manifiesto instalado | `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/readcite-manifest.xml` |

## Primera vez

1. Confía en la CA local (una sola vez; te pedirá la contraseña de tu Mac):

   ```bash
   security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db ~/.readcite/certs/readcite-ca.crt
   ```

   La CA tiene *Name Constraints*: aunque alguien la usara, solo vale para `localhost`.

2. El servidor arranca solo: lo lanza macOS al iniciar sesión (LaunchAgent
   `~/Library/LaunchAgents/com.readcite.servidor.plist`) y lo reinicia si se cae.
   Registro en `~/.readcite/servidor.log`.

3. Cierra Excel por completo (⌘Q) y ábrelo con `ReadCiteExcel.xlsx`.
4. Pestaña **Inicio → ReadCite** (a la derecha). Si no aparece:
   **Insertar → Complementos → Mis complementos**, apartado *Complementos de desarrollador* → ReadCite.

## Uso diario

Nada: abre Excel y pulsa **Inicio → ReadCite**. El servidor ya está en marcha.

Comprobar el servidor / reiniciarlo a mano:

```bash
launchctl print gui/$(id -u)/com.readcite.servidor | grep -E "state|pid"
launchctl kickstart -k gui/$(id -u)/com.readcite.servidor
```

## Problemas frecuentes

- **«Error del complemento» o panel en blanco**: el servidor no responde. Reinícialo con el
  `launchctl kickstart` de arriba y pulsa **Reintentar** en el panel. Comprueba en Safari que
  `https://localhost:3000/panel/index.html` abre sin aviso de seguridad (si avisa, falta el paso 1).
  Si moviste la carpeta `ReadCiteCarpeta`, el LaunchAgent apunta a la ruta vieja: pide a Claude Code que lo actualice.
- **No aparece el botón**: cierra Excel con ⌘Q (no basta con cerrar la ventana) y vuelve a abrirlo.
- **Cambios del panel que no se ven**: menú **ⓘ** del panel → **Borrar caché web**, luego cierra y abre el panel.

## Desinstalar

```bash
launchctl bootout gui/$(id -u) ~/Library/LaunchAgents/com.readcite.servidor.plist
rm ~/Library/LaunchAgents/com.readcite.servidor.plist
rm ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/readcite-manifest.xml
security delete-certificate -c "ReadCite Local Dev CA (solo localhost)" ~/Library/Keychains/login.keychain-db
rm -r ~/.readcite
```

## Renovar el certificado (cuando caduque)

Pídeselo a Claude Code ("regenera el certificado local de ReadCite"): crea una CA nueva limitada a
localhost, firma `localhost.crt`, borra la clave de la CA y repite el paso 1 de arriba con la nueva CA.
