# gpx_a_google_maps

Convierte un track GPX (por ejemplo, uno descargado de **Wikiloc**) en enlaces de ruta de **Google Maps** que puedes abrir en el móvil y seguir con navegación paso a paso.

Google Maps no permite importar un GPX para navegarlo. Este script extrae del track los puntos clave, los ordena y genera los enlaces `maps/dir` con sus paradas intermedias. Si la ruta tiene más paradas de las que Google admite en un enlace, la divide en tramos encadenados.

## Características

- Usa los **waypoints** del GPX como paradas y los ordena según el sentido de avance del track, aunque en el archivo vengan desordenados.
- Si el GPX no tiene waypoints, **simplifica el track** y se queda con los puntos que mejor conservan su forma (curvas, desvíos, cruces), no con puntos a intervalos fijos.
- **Fusiona puntos casi coincidentes**, para no generar paradas inútiles a pocos metros unas de otras.
- **Reparte los tramos de forma equilibrada**: el final de un tramo es el inicio del siguiente y nunca queda un tramo residual de unos pocos metros.
- Genera un **HTML** con un enlace por tramo, pensado para abrirlo en el móvil.
- Solo necesita **Python 3** con la biblioteca estándar. No hay dependencias que instalar.

## Requisitos

Python 3.6 o superior.

## Uso

```bash
python gpx_a_google_maps.py ruta.gpx
```

En Windows:

```bat
python c:\ruta\gpx_a_google_maps.py c:\ruta\mi_track.gpx
```

El script muestra los enlaces por pantalla y guarda un archivo `ruta_google_maps.html` junto al GPX. Pásate ese HTML al móvil y abre los tramos en orden: se abrirán en la app de Google Maps listos para iniciar la navegación.

### Opciones

| Opción | Valor por defecto | Descripción |
|---|---|---|
| `--modo` | `driving` | Modo de desplazamiento: `driving`, `walking`, `bicycling` o `transit`. |
| `--usar` | `auto` | `auto` usa los waypoints si hay al menos dos; `wpt` fuerza el uso de waypoints; `track` los ignora y simplifica el track. |
| `--puntos` | `25` | Número de puntos que se extraen del track cuando se usa `--usar track` o el GPX no tiene waypoints. |
| `--max-paradas` | `9` | Paradas intermedias por enlace. Google admite 9; si abres los enlaces desde el navegador del móvil en lugar de la app, puede que solo acepte 3. |
| `--umbral` | según el modo | Distancia en metros por debajo de la cual dos puntos se consideran el mismo: 500 en coche, 250 en bici y 100 a pie. |
| `--salida` | `<gpx>_google_maps.html` | Ruta del archivo HTML de salida. |

### Ejemplos

```bash
# Ruta en coche usando los waypoints del GPX
python gpx_a_google_maps.py espinar.gpx

# Ruta a pie ciñéndose más al trazado original
python gpx_a_google_maps.py senda.gpx --modo walking --usar track --puntos 40

# Enlaces para abrir desde el navegador del móvil
python gpx_a_google_maps.py ruta.gpx --max-paradas 3 --salida enlaces.html
```

### Salida de ejemplo

```
Puntos usados: waypoints (10)
Longitud aproximada: 133.6 km  |  Tramos: 1

Tramo 1:
https://www.google.com/maps/dir/?api=1&origin=40.725906,-4.268220&destination=40.587697,-4.399618&travelmode=driving&waypoints=...

Enlaces guardados en: espinar_google_maps.html
```

## Cómo funciona

1. **Lectura del GPX.** Lee los waypoints (`<wpt>`) y los puntos de track (`<trkpt>`). Si no hay track, usa los puntos de ruta (`<rtept>`).
2. **Selección de puntos.**
   - **Con waypoints y track:** asigna a cada waypoint el punto del track más cercano y los ordena por esa posición. Añade el inicio y el final del track si no coinciden con ningún waypoint.
   - **Solo con waypoints:** los usa en el orden del archivo.
   - **Solo con track:** aplica una variante de Douglas-Peucker con número de puntos fijo. Parte del inicio y el final y va añadiendo siempre el punto que más se aleja de la línea simplificada.
3. **Limpieza.** Fusiona los puntos consecutivos que estén a menos del umbral. Si uno de ellos tiene nombre propio, conserva ese.
4. **División en tramos.** Calcula cuántos enlaces hacen falta y reparte los puntos de forma equilibrada entre ellos.
5. **Generación de enlaces** con la [API de URLs de Google Maps](https://developers.google.com/maps/documentation/urls/get-started), que no necesita clave de API.

## Limitaciones

- **Google recalcula el camino entre parada y parada** con su propia cartografía. En carretera el resultado se ajusta bien al track, pero en sendas o pistas que Google no conoce te llevará por otro sitio. Cuantos más puntos uses, más se ceñirá al trazado.
- **El orden de los waypoints se deduce por proximidad al track.** Si el recorrido pasa varias veces muy cerca de un mismo waypoint, conviene revisar el orden en la salida.
- **La navegación depende de la conexión de datos**, igual que cualquier ruta de Google Maps.
- Para senderismo de montaña son más adecuadas aplicaciones que navegan el GPX directamente, como Wikiloc, OsmAnd u Organic Maps.
