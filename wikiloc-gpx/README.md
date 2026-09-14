# wikiloc_gpx.py

Recorre una carpeta con ficheros `.gpx`, calcula las métricas de cada ruta y genera
una tabla en varios formatos. Desde la página HTML puedes marcar rutas y descargar
un `.bat` que las copia o las mueve, agrupadas en carpetas por país.

Es un único fichero de Python, sin dependencias obligatorias, y funciona sin conexión.

---

## Requisitos

- **Python 3.7 o superior.** Comprueba con `python --version`.
- **Opcional:** `openpyxl`, solo si quieres salida en Excel real.
  Sin él, el formato `xlsx` genera un CSV en su lugar y te avisa.
  ```bat
  pip install openpyxl
  ```
- **Opcional:** `reverse_geocoder`, para que el país se detecte con precisión.
  Sin él se usa una tabla interna aproximada y los países aparecen marcados
  con «(aprox.)».
  ```bat
  pip install reverse_geocoder
  ```

---

## Sintaxis

```
python wikiloc_gpx.py [directorio] [-o SALIDA] [-f FORMATOS] [-r]
                      [--umbral METROS] [--enlace ruta|autor] [--solo-wikiloc]
```

### La regla que más problemas da

**Los formatos van separados por comas y SIN espacios.** `cmd` corta los argumentos
por los espacios, así que `-f all, movil` se interpreta como `-f "all,"` más un
argumento suelto `movil` y el script falla con
`error: unrecognized arguments: movil`.

| Mal                    | Bien                                        |
| ---------------------- | ------------------------------------------- |
| `-f html, movil`       | `-f html,movil`  o  `-f "html, movil"`      |
| `-f all, movil`        | `-f all`  (all ya incluye movil)            |

Lo mismo con las rutas: si llevan espacios, van entre comillas.
`"C:\Mis rutas\gpx"`, no `C:\Mis rutas\gpx`.

---

## Parámetros

| Parámetro | Qué hace | Por defecto |
| --- | --- | --- |
| `directorio` | Carpeta donde están los `.gpx`. | La carpeta actual |
| `-o`, `--salida` | Nombre base de los ficheros generados, **sin extensión**. Si es solo un nombre, se crean dentro de la carpeta explorada. Si incluye ruta (`D:\informes\rutas`), se usa esa ruta. | El nombre de la carpeta explorada |
| `-f`, `--formato` | Formatos separados por comas: `txt`, `csv`, `xlsx`, `html`, `movil`, o `all` para todos. | `txt` |
| `-r`, `--recursivo` | Busca también en las subcarpetas. | No |
| `--umbral` | Metros de umbral para el desnivel acumulado. Sube el valor si el GPS te infla los desniveles. | `4` |
| `--enlace` | Si la columna Enlace muestra el de la ruta (`ruta`) o el del autor (`autor`). | `ruta` |
| `--solo-wikiloc` | Descarta los `.gpx` que no vengan de Wikiloc en lugar de incluirlos etiquetados. | No |

---

## Ejemplos listos para copiar

Lo más habitual, todo de una vez y mirando también en subcarpetas:

```bat
python C:\input\wikiloc_gpx.py C:\input -r -o rutas -f all --umbral 5
```

Genera dentro de `C:\input`: `rutas.txt`, `rutas.csv`, `rutas.xlsx`, `rutas.html`
y `rutas_movil.html`.

Solo la tabla de escritorio:

```bat
python C:\input\wikiloc_gpx.py C:\input -r -o rutas -f html
```

Solo la versión de móvil:

```bat
python C:\input\wikiloc_gpx.py C:\input -r -o rutas -f movil
```

Dejar los informes en otro disco, sin ensuciar la carpeta de los GPX:

```bat
python C:\input\wikiloc_gpx.py C:\input -r -o "D:\informes\rutas 2026" -f html,movil
```

Sin `-o`, los ficheros toman el nombre de la carpeta explorada
(`C:\input` produce `input.html`, `input_movil.html`…):

```bat
python C:\input\wikiloc_gpx.py C:\input -f all
```

Descartando lo que no sea de Wikiloc:

```bat
python C:\input\wikiloc_gpx.py C:\input -r -o rutas -f all --solo-wikiloc
```

---

## Qué contiene cada formato

| Formato | Fichero | Para qué |
| --- | --- | --- |
| `txt` | `rutas.txt` | Texto alineado en columnas, para mirar de un vistazo |
| `csv` | `rutas.csv` | Con BOM y `;`, lo abre Excel en español con coma decimal |
| `xlsx` | `rutas.xlsx` | Excel real, con números, autofiltro y enlaces pinchables |
| `html` | `rutas.html` | La versión completa: tabla, filtros, miniaturas y generación del `.bat` |
| `movil` | `rutas_movil.html` | Versión ligera para el teléfono, solo rutas de Wikiloc |

### Columnas

`Nombre del archivo`, `Wikiloc` (Sí/No), `Origen` (Wikiloc, Garmin Connect,
StravaGPX…), `Duplicado`, `Autor`, `Enlace`, `País`, `Tipo` (Circular o Lineal),
`Inicio (mapa)`, `Dist (km)`, `Alt mín (m)`, `Alt máx (m)`, `Desnivel + (m)`,
`Inicio (UTC)`, `Fin (UTC)`, `Duración (h)`, `En movim. (h)`, `Parado (h)`.

Las horas están en UTC, que es como vienen en los GPX. El tiempo en movimiento
separa la marcha real de las paradas usando un umbral de 1 km/h.

---

## La página de escritorio (`rutas.html`)

Se abre con doble clic. No necesita conexión ni servidor.

**Miniaturas.** Cada fila lleva el trazado dibujado. Al pasar el ratón por encima
se amplía en una lupa flotante, con el punto de inicio en verde, el final en rojo
y el perfil de altitud debajo.

**Filtros de la cabecera.** Texto libre (nombre, autor, país, origen), país,
Wikiloc sí/no, circular o lineal, duplicados, y rangos mínimo y máximo de
kilómetros y de desnivel. El botón «Limpiar filtros» los deja todos a cero.

**Generar el `.bat`.**

1. Marca las casillas de las rutas. La casilla de la cabecera marca y desmarca
   solo lo que esté visible con el filtro puesto.
2. Escribe la carpeta raíz de destino.
3. Elige **Copiar** o **Mover**.
4. Pulsa **Descargar .bat por paises** (o **Descargar .bat** para volcarlo todo
   a una sola carpeta, sin clasificar).

También tienes **Previsualizar**, que enseña el contenido del `.bat` antes de
bajarlo, y **Copiar comandos**, para pegarlos directamente en una ventana de CMD.

### Cómo se reparten las carpetas

Dentro de la raíz se crea una subcarpeta por país, y tres especiales. El orden de
prioridad es:

1. **`Sospechosos`** — rutas con el mismo inicio, final y distancia que otra,
   pero con nombre de fichero distinto. Probables duplicados, para revisar a mano.
2. **`No_Wikiloc`** — ficheros que no vienen de Wikiloc, aunque se les haya
   detectado el país.
3. **`<País>`** — `Espana`, `Peru`, `Estados Unidos`… sin tildes, a propósito,
   para evitar problemas de codificación en la consola.
4. **`desconocido`** — si no se ha podido determinar el país.

**Si un fichero ya existe en destino con el mismo nombre, NO se sobrescribe:**
se omite y el `.bat` lo dice por pantalla. Así los duplicados con el mismo nombre
que estaban en subcarpetas distintas no se pisan entre sí.

### Ejecutar el `.bat`

Con doble clic usa la carpeta raíz que escribiste en la página. También puedes
pasarle otra raíz como primer parámetro, sin volver a generar nada:

```bat
cd %USERPROFILE%\Downloads
rutas_copiar_paises.bat "E:\Backup GPX\2026"
```

Las comillas son obligatorias si la ruta lleva espacios.

---

## La página de móvil (`rutas_movil.html`)

Un solo fichero de unos 16 KB, solo con rutas de Wikiloc, en fichas en vez de tabla:
trazado en miniatura, nombre, país, circular o lineal, y los datos en rejilla
(km, desnivel, duración en h:mm, altura mínima y máxima, y la distancia hasta ti).
Abajo, dos botones grandes: **Cómo llegar al inicio** y **Ver en Wikiloc**.

Arriba, selector de país y de orden (nombre, más cercanas, más largas, más desnivel)
y el botón **Usar mi posición**.

### Sobre la ubicación

El enlace **Cómo llegar al inicio** no lleva punto de partida, así que Google Maps
traza la ruta desde donde estés en ese momento. Funciona siempre, dé o no la página
permiso de ubicación.

**Ordenar por cercanía es otra cosa:** eso sí necesita que el navegador le dé la
posición a la página, y los navegadores solo lo hacen en contextos que consideran
seguros. Abriendo el fichero con doble clic desde el propio teléfono es probable que
lo bloquee. Tienes dos salidas:

- El botón **A mano**: escribe `40.4168, -3.7038` o **pega un enlace de Google Maps**
  (el de compartir ubicación) y él extrae las coordenadas.
- Subir el fichero a cualquier alojamiento con `https`. Servirlo desde tu PC no vale:
  `http://192.168.x.x` tampoco cuenta como contexto seguro.

Las distancias son en línea recta, no por carretera.

---

## Errores frecuentes

| Mensaje | Causa |
| --- | --- |
| `error: unrecognized arguments: movil` | Un espacio detrás de la coma en `-f`. Usa `-f html,movil`. |
| `ERROR: 'X' no es un directorio` | El primer argumento debe ser una carpeta, no un `.gpx`. |
| `ERROR: formato(s) no válido(s)` | Solo valen `txt`, `csv`, `xlsx`, `html`, `movil` y `all`. |
| `'openpyxl' no está instalado` | Aviso, no error: genera CSV en su lugar. |
| `Error en fichero.gpx: XML no válido` | Ese `.gpx` está corrupto o incompleto; los demás se procesan igual. |

---

## Limitaciones conocidas

- **Tiempo en movimiento.** Usa un umbral fijo de 1 km/h. En progresión muy lenta
  (una trepada, nieve profunda) contará como parada lo que fue esfuerzo.
- **Duplicados.** La comparación es geométrica: inicio, final y distancia. Dos
  subidas distintas al mismo pico desde el mismo aparcamiento, con la misma
  distancia, pueden salir marcadas como sospechosas. Por eso se desvían a una
  carpeta en vez de borrarse.
- **País aproximado.** Sin `reverse_geocoder` se usa el punto de referencia más
  cercano de una tabla interna, sin límite de distancia: un track en medio del mar
  se asignará al país más próximo.
- **Nombres repetidos.** Dos rutas distintas con el mismo nombre de fichero y del
  mismo país acabarán en la misma carpeta, y la segunda se omitirá por existir ya.
