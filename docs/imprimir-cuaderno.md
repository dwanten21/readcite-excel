# Imprimir el cuaderno A5 de ReadCite

El cuaderno vive en la hoja **Cuaderno** de `ReadCiteExcel.xlsx`. Reúne las citas **guardadas**
de un año, ordenadas de la A a la Z por el apellido del autor (sin partículas: «von Mises» va en la M).

## Qué sale

- Página 1: portada (año, lector/a, fechas). Página 2: índice con la página de cada letra (con enlace).
- Cada letra empieza **siempre en un anverso** (página impar) y ocupa un número **par** de páginas:
  hojas completas que puedes sacar, añadir o reordenar sin descuadrar las demás.
- Sin citas, una letra tiene igualmente 2 páginas de plantillas para escribir a mano.
- Pestañas A–Z en el borde exterior (derecha en anverso, izquierda en reverso), la letra actual en negro.
- Margen interior 16 mm (para perforar o encuadernar) y exterior 6 mm.
- El texto va silabeado con las mismas reglas que la web: el guion solo aparece donde se parte una palabra.

## Paso a paso

1. **Guarda las citas** que quieras en el cuaderno: en el panel, botón **Guardar** de cada cita.
2. Si quieres otro año, otras citas por página (1, 2 o 3) o quitar el espejo de pestañas: panel → **⚙**.
3. Panel → **Cuaderno** → **Generar cuaderno** (tarda menos de un minuto; rehace la hoja entera).
4. En **Qué imprimir** elige *Todo el cuaderno* o *Solo la letra X* y pulsa **Preparar área de impresión**.
5. **⌘P** (Archivo → Imprimir) y comprueba:
   - **Tamaño de papel: A5** (148 × 210 mm). Si tu impresora no tiene A5, ver «Sin papel A5» abajo.
   - **Escala: 100 %** (no «Ajustar a página»: deshace los saltos de página).
   - **Doble cara / Dúplex: encuadernación por el borde largo** (en macOS: «Dúplex» → «Encuadernación
     borde largo»; en algunos controladores está en el apartado *Diseño* o *Acabado*).
   - Vista previa: la primera página es la portada y el total coincide con el panel (p. ej. 54 páginas).
6. Imprime. Con doble cara, la página 1 y la 2 quedan en la misma hoja (portada / índice), y así sucesivamente.

## Imprimir solo una letra

Como cada letra empieza en anverso y tiene páginas pares, puedes reimprimir solo la letra que ha crecido
(paso 4 con *Solo la letra X*) y sustituir sus hojas en la carpeta. Vuelve a *Todo el cuaderno* después.

## Problemas frecuentes

- **Sale una hoja en blanco entre medias o las letras empiezan en reverso**: la impresora está escalando.
  Pon escala 100 % y desactiva «Ajustar a página».
- **Las pestañas salen cortadas en el borde**: tu impresora no imprime hasta el borde del papel (márgenes
  físicos de 3–5 mm). El contenido no se ve afectado; si te molesta, desactiva el espejo en ⚙ para que
  las pestañas vayan siempre a la derecha.
- **Sin papel A5**: imprime en A4 con «2 páginas por hoja» (multipágina) y corta por la mitad; o usa la
  función *Folleto* del controlador. La maqueta está pensada para A5 a tamaño real.
- **Cambié algo a mano en la hoja Cuaderno y desapareció**: esa hoja la rehace ReadCite cada vez que
  pulsas *Generar cuaderno*. Edita las citas en la hoja **Citas** (o desde el panel), no en el cuaderno.
- **La referencia APA sale sin cursiva en el título**: Excel no permite cursiva en parte de una celda desde
  el complemento; la referencia se imprime entera en redonda.
