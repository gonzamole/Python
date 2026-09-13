# wikiloc-gpx

Extrae datos y métricas de los ficheros **GPX descargados de Wikiloc** que haya
en una carpeta y los vuelca en una tabla, en varios formatos a la vez:
texto alineado, CSV, Excel (`.xlsx`) y HTML.

Para cada track obtiene el nombre del archivo, el autor, el enlace de Wikiloc,
el país y las principales métricas de la ruta (distancia, altitudes, desnivel
acumulado, horas de inicio/fin y duración).

## Índice

- [Características](#características)
- [Requisitos e instalación](#requisitos-e-instalación)
- [Uso](#uso)
- [Dónde se guardan los resultados](#dónde-se-guardan-los-resultados)
- [Opciones](#opciones)
- [Columnas de salida](#columnas-de-salida)
- [Formatos de salida](#formatos-de-salida)
- [Copiar o mover ficheros desde el HTML](#copiar-o-mover-ficheros-desde-el-html)
- [Cómo se detecta el país](#cómo-se-detecta-el-país)
- [Ejemplo de salida](#ejemplo-de-salida)
- [Limitaciones y notas](#limitaciones-y-notas)
- [Licencia](#licencia)

## Características

- Recorre un directorio (opcionalmente también sus subcarpetas) buscando `.gpx`.
- Reconoce automáticamente los ficheros de Wikiloc e ignora el resto.
- Extrae **autor** y **enlace** (el de la ruta `.../hiking-trails/...` o, si lo
  prefieres, el del perfil del autor `.../user.do?id=...`).
- Calcula por track: **distancia**, **altitud mínima y máxima**,
  **desnivel positivo acumulado**, **hora de inicio y fin** y **duración**.
- Determina el **país** sin conexión (offline).
- Genera la salida en **txt, csv, xlsx y html**, o en los que elijas.
- En el HTML puedes **marcar tracks y copiarlos o moverlos** a otra carpeta
  (genera un `.bat` de Windows con los comandos).
- Sin dependencias obligatorias: funciona solo con la biblioteca estándar de
  Python. Las bibliotecas opcionales mejoran el resultado (Excel real y país
  más preciso).

## Requisitos e instalación

- **Python 3.8 o superior.**
- Bibliotecas **opcionales**:
  - [`openpyxl`](https://pypi.org/project/openpyxl/) — para generar Excel
    `.xlsx` real. Si no está instalada, al pedir `xlsx` el programa crea un
    `.csv` en su lugar y avisa por pantalla.
  - [`reverse_geocoder`](https://pypi.org/project/reverse_geocoder/) — para
    detectar el país con precisión. Funciona offline una vez instalada. Sin
    ella se usa una tabla interna aproximada y los países se marcan con
    `(aprox.)`.

```bash
git clone https://github.com/<tu-usuario>/wikiloc-gpx.git
cd wikiloc-gpx

# Opcionales (recomendado)
pip install openpyxl reverse_geocoder
```

No hay nada que compilar: es un único script, `wikiloc_gpx.py`.

## Uso

```bash
python wikiloc_gpx.py [DIRECTORIO] [-o BASE] [-f FORMATOS] [-r]
                      [--enlace {ruta,autor}] [--umbral METROS]
```

Ejemplos:

```bash
# Todos los formatos, incluyendo subcarpetas
python wikiloc_gpx.py "C:\rutas\wikiloc" -o rutas -f all -r

# Explorar una subcarpeta y dejar los resultados EN ella (nombre = la carpeta)
python wikiloc_gpx.py C:\gpx\montenegro -f all

# Solo Excel y HTML
python wikiloc_gpx.py ~/gpx -o rutas -f xlsx,html

# Solo texto alineado de la carpeta actual
python wikiloc_gpx.py

# Ajustar el umbral del desnivel a 5 m
python wikiloc_gpx.py ~/gpx -o rutas -f xlsx --umbral 5
```

En Windows, si `python` no funciona pero tienes Python instalado, prueba con
`py` en lugar de `python`.

## Dónde se guardan los resultados

El primer argumento indica **qué carpeta explorar**; `-o` decide **cómo se
llaman y dónde van** los ficheros generados:

- **Sin `-o`** → se crean **dentro de la carpeta explorada**, con el nombre de
  esa carpeta.
- **`-o nombre`** (solo un nombre) → se crean dentro de la carpeta explorada con
  ese nombre.
- **`-o C:\ruta\nombre`** (con ruta) → se crean en la ruta que indiques.

Así puedes tener el script en una carpeta fija y explorar otras. Por ejemplo,
con `wikiloc_gpx.py` en `C:\gpx`:

```bash
cd C:\gpx
python wikiloc_gpx.py montenegro -f all
```

explora `C:\gpx\montenegro` y deja ahí mismo `montenegro.txt`, `montenegro.csv`,
`montenegro.xlsx` y `montenegro.html`. Lo mismo con ruta absoluta desde
cualquier sitio:

```bash
python C:\gpx\wikiloc_gpx.py C:\gpx\montenegro -f all
```

## Opciones

| Opción | Descripción | Por defecto |
| --- | --- | --- |
| `DIRECTORIO` | Carpeta donde buscar los `.gpx`. | carpeta actual |
| `-o`, `--salida` | Nombre base de salida **sin extensión**. Si es solo un nombre, los ficheros se crean **dentro de la carpeta explorada**; si incluye una ruta, se usa esa ruta. Vacío = nombre de la carpeta explorada. | (nombre de la carpeta) |
| `-f`, `--formato` | Formatos separados por coma: `txt`, `csv`, `xlsx`, `html`, o `all`. | `txt` |
| `-r`, `--recursivo` | Buscar también en subcarpetas. | desactivado |
| `--enlace` | Qué enlace usar en la columna *Enlace wikiloc*: `ruta` (página del track) o `autor` (perfil del autor). | `ruta` |
| `--umbral` | Umbral en metros para el desnivel acumulado (ver más abajo). | `4` |

## Columnas de salida

| Columna | Contenido |
| --- | --- |
| Nombre del archivo | Nombre del `.gpx`. |
| Autor | Autor del track (de `<author>` o `<copyright>`); `(desconocido)` si no consta. |
| Enlace wikiloc | Enlace de la ruta o del autor, según `--enlace`. |
| País | País de la ruta (ver [detección](#cómo-se-detecta-el-país)). |
| Dist (km) | Distancia horizontal (2D) sumando la separación entre puntos. |
| Alt mín (m) | Altitud mínima del track. |
| Alt máx (m) | Altitud máxima del track. |
| Desnivel + (m) | Desnivel positivo acumulado (con filtro de `--umbral`). |
| Inicio (UTC) | Hora del primer punto, en UTC. |
| Fin (UTC) | Hora del último punto, en UTC. |
| Duración (h) | Diferencia Fin − Inicio, en horas decimales. |

## Formatos de salida

- **txt** — Texto alineado en columnas. Se lee bien con una fuente
  monoespaciada (Consolas, Courier, etc.).
- **csv** — Con BOM y separador `;`, y coma decimal, para que **Excel en
  español** lo abra correctamente y con los acentos bien.
- **xlsx** — Excel real: cabecera fija, filtros automáticos, enlaces clicables
  y los números y fechas como **valores reales** (puedes ordenarlos y sumarlos).
  Requiere `openpyxl`.
- **html** — Página web con la tabla; al pulsar una cabecera **ordena** por esa
  columna (orden numérico en las columnas de números). Además incluye casillas
  para **seleccionar tracks y copiarlos o moverlos** a otra carpeta (ver la
  sección siguiente).

## Copiar o mover ficheros desde el HTML

El HTML permite quedarte con los tracks que te interesen y llevar sus `.gpx` a
otra carpeta. Como una página abierta desde un archivo local no puede escribir
en el disco por seguridad del navegador, el HTML **genera un `.bat` de Windows**
(o te copia los comandos al portapapeles) y eres tú quien hace la copia o el
movimiento al ejecutarlo.

Pasos:

1. Abre el `.html` en el navegador.
2. Marca las casillas de los tracks que quieras (o usa la casilla de la cabecera
   para marcar/desmarcar todos).
3. Escribe la **carpeta destino** (p. ej. `C:\gpx\seleccion`).
4. Elige **Copiar** o **Mover**.
5. Pulsa una opción:
   - **Previsualizar** — muestra los comandos exactos antes de ejecutar nada.
   - **Copiar comandos** — los copia para pegarlos en una ventana CMD.
   - **Descargar .bat** — baja un fichero que ejecutas con doble clic.

El `.bat` crea la carpeta destino si no existe y respeta rutas con espacios.

Notas:

- Solo **Windows** (usa `copy`/`move`). Para macOS/Linux haría falta generar un
  `.sh` con `cp`/`mv`.
- Al descargar un `.bat`, Windows puede mostrar un aviso de SmartScreen; si te
  fías del archivo: «Más información → Ejecutar de todas formas».
- Las rutas de origen son las que tenían los `.gpx` al generar el HTML. Si
  después mueves o renombras los originales, vuelve a generar la tabla.

## Cómo se detecta el país

1. Si el nombre o la descripción del track menciona explícitamente un país
   (p. ej. «España», «Portugal», «Macedonia»…), se usa ese.
2. Si no, se hace **geocodificación inversa offline** con las coordenadas del
   primer y último punto:
   - Con `reverse_geocoder` instalado → resultado preciso.
   - Sin él → tabla interna aproximada; el país aparece con el sufijo
     `(aprox.)` para avisarte de que conviene revisarlo, sobre todo en zonas
     de frontera.

## Ejemplo de salida

```text
Nombre del archivo           | Autor         | Enlace wikiloc                                     | País                | Dist (km) | Alt mín (m) | Alt máx (m) | Desnivel + (m) |     Inicio (UTC) |        Fin (UTC) | Duración (h)
-----------------------------+---------------+---------------------------------------------------+---------------------+-----------+-------------+-------------+----------------+------------------+------------------+-------------
vejce-kobilica-treskavec.gpx | (desconocido) | https://www.wikiloc.com/hiking-trails/vejce-...    | Macedonia del Norte |     15,81 |        1126 |        2511 |           1746 | 2025-11-02 06:55 | 2025-11-02 16:22 |         9,44
```

## Limitaciones y notas

- **Desnivel acumulado.** Se calcula solo el positivo (D+) y depende del
  `--umbral`. El valor por defecto (4 m) filtra el ruido del GPS y da una cifra
  realista; bajarlo a 0 dispara el resultado, subirlo lo hace más conservador.
  Ajústalo comparando con una ruta que ya conozcas.
- **Horas en UTC.** Wikiloc guarda las marcas de tiempo en UTC, así que
  *Inicio* y *Fin* se muestran en UTC (en España, 1–2 h menos que la hora
  local del reloj). La **duración sí es exacta**, porque es una resta.
- **Distancia 2D.** Es la distancia horizontal, como la de la mayoría de
  visores; la distancia 3D (con la pendiente) sería algo mayor.
- Los ficheros que no son de Wikiloc y los `.gpx` ilegibles se ignoran y se
  informa de ellos en pantalla.

## Licencia

Publicado bajo licencia **MIT**. Añade un fichero `LICENSE` con el texto de la
licencia si aún no lo tienes.
