#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gpx_a_google_maps.py
Convierte un track GPX (por ejemplo, descargado de Wikiloc) en uno o varios
enlaces de ruta de Google Maps que se pueden abrir y navegar desde el móvil.

Lógica:
  - Si el GPX tiene waypoints (<wpt>) y track (<trkpt>), usa los waypoints
    como paradas, ordenados según el sentido de avance del track, y añade
    el inicio y el final del track si no coinciden con ningún waypoint.
  - Si solo tiene waypoints, los usa en el orden del archivo.
  - Si solo tiene track, selecciona automáticamente los puntos que mejor
    conservan la forma del recorrido (curvas, desvíos).
  - Como Google Maps limita las paradas por enlace, divide la ruta en
    tramos encadenados (el final de un tramo es el inicio del siguiente).

Uso:
  python gpx_a_google_maps.py ruta.gpx
  python gpx_a_google_maps.py ruta.gpx --modo walking --puntos 30
  python gpx_a_google_maps.py ruta.gpx --usar track --salida enlaces.html

Solo usa la biblioteca estándar de Python 3 (no hay que instalar nada).
"""

import argparse
import heapq
import html
import math
import os
import sys
import xml.etree.ElementTree as ET
from urllib.parse import quote

RADIO_TIERRA = 6371000.0  # metros


# ---------------------------------------------------------------- lectura GPX

def _etiqueta(elem):
    """Nombre de la etiqueta sin espacio de nombres."""
    return elem.tag.rsplit('}', 1)[-1]


def _hijo(elem, nombre):
    for h in elem:
        if _etiqueta(h) == nombre:
            return (h.text or '').strip()
    return ''


def leer_gpx(ruta):
    """Devuelve (waypoints, track). Cada punto es un dict lat, lon, nombre."""
    raiz = ET.parse(ruta).getroot()
    wpts, trk, rte = [], [], []
    for e in raiz.iter():
        t = _etiqueta(e)
        if t not in ('wpt', 'trkpt', 'rtept'):
            continue
        p = {'lat': float(e.get('lat')), 'lon': float(e.get('lon')),
             'nombre': _hijo(e, 'name')}
        {'wpt': wpts, 'trkpt': trk, 'rtept': rte}[t].append(p)
    return wpts, (trk or rte)


# ------------------------------------------------------------------ geometría

def distancia(a, b):
    """Distancia haversine en metros."""
    la1, la2 = math.radians(a['lat']), math.radians(b['lat'])
    dla = la2 - la1
    dlo = math.radians(b['lon'] - a['lon'])
    h = math.sin(dla / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin(dlo / 2) ** 2
    return 2 * RADIO_TIERRA * math.asin(math.sqrt(h))


def _xy(p, lat0):
    """Proyección plana local (suficiente para distancias cortas)."""
    return (math.radians(p['lon']) * math.cos(math.radians(lat0)) * RADIO_TIERRA,
            math.radians(p['lat']) * RADIO_TIERRA)


def _dist_segmento(p, a, b):
    (px, py), (ax, ay), (bx, by) = p, a, b
    dx, dy = bx - ax, by - ay
    if dx == 0 and dy == 0:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def simplificar_track(track, n_puntos):
    """
    Elige n_puntos del track que mejor conservan su forma (variante de
    Douglas-Peucker con número de puntos fijo): va añadiendo siempre el punto
    que más se aleja de la línea simplificada.
    """
    if len(track) <= n_puntos:
        return list(track)
    lat0 = sum(p['lat'] for p in track) / len(track)
    xy = [_xy(p, lat0) for p in track]

    def peor(i, j):
        mejor_d, mejor_k = -1.0, None
        for k in range(i + 1, j):
            d = _dist_segmento(xy[k], xy[i], xy[j])
            if d > mejor_d:
                mejor_d, mejor_k = d, k
        return mejor_d, mejor_k

    elegidos = {0, len(track) - 1}
    cola = []
    d, k = peor(0, len(track) - 1)
    if k is not None:
        heapq.heappush(cola, (-d, 0, len(track) - 1, k))
    while cola and len(elegidos) < n_puntos:
        _, i, j, k = heapq.heappop(cola)
        elegidos.add(k)
        for a, b in ((i, k), (k, j)):
            d, kk = peor(a, b)
            if kk is not None:
                heapq.heappush(cola, (-d, a, b, kk))
    return [track[i] for i in sorted(elegidos)]


def ordenar_waypoints_por_track(wpts, track, umbral=75.0):
    """
    Ordena los waypoints según el punto del track más cercano a cada uno
    y añade inicio/fin del track si no hay waypoint a menos de 'umbral' m.
    """
    con_indice = []
    for w in wpts:
        idx = min(range(len(track)), key=lambda i: distancia(w, track[i]))
        con_indice.append((idx, w))
    con_indice.sort(key=lambda x: x[0])
    ordenados = [w for _, w in con_indice]

    inicio, fin = dict(track[0]), dict(track[-1])
    inicio['nombre'] = inicio['nombre'] or 'Inicio del track'
    fin['nombre'] = fin['nombre'] or 'Final del track'
    if not ordenados or distancia(ordenados[0], inicio) > umbral:
        ordenados.insert(0, inicio)
    if distancia(ordenados[-1], fin) > umbral:
        ordenados.append(fin)
    return ordenados


# ---------------------------------------------------------- enlaces de Maps

def _coord(p):
    return '%.6f,%.6f' % (p['lat'], p['lon'])


def fusionar_cercanos(puntos, umbral):
    """
    Quita puntos consecutivos separados por menos de 'umbral' metros
    (conserva el que tenga nombre propio o, si no, el primero). Siempre
    mantiene el último punto como destino.
    """
    if len(puntos) < 2:
        return list(puntos)
    res = [puntos[0]]
    for p in puntos[1:]:
        if distancia(res[-1], p) < umbral:
            genericos = ('', 'Waypoint', 'Inicio del track', 'Final del track')
            if res[-1]['nombre'] in genericos and p['nombre'] not in genericos:
                res[-1] = p
            continue
        res.append(p)
    if len(res) == 1:
        res.append(puntos[-1])
    return res


def dividir_en_tramos(puntos, max_paradas):
    """
    Reparte los puntos en tramos equilibrados: cada tramo tiene origen,
    hasta max_paradas intermedias y destino, y el destino de un tramo es el
    origen del siguiente. Evita que el último tramo quede con 1 o 2 puntos.
    """
    segmentos = len(puntos) - 1               # saltos entre puntos
    por_tramo = max_paradas + 1               # saltos máximos por enlace
    n_tramos = max(1, math.ceil(segmentos / por_tramo))
    base, resto = divmod(segmentos, n_tramos)
    tramos, i = [], 0
    for k in range(n_tramos):
        saltos = base + (1 if k < resto else 0)
        tramos.append(puntos[i:i + saltos + 1])
        i += saltos
    return tramos


def url_google_maps(tramo, modo):
    url = ('https://www.google.com/maps/dir/?api=1'
           '&origin=' + _coord(tramo[0]) +
           '&destination=' + _coord(tramo[-1]) +
           '&travelmode=' + modo)
    intermedios = tramo[1:-1]
    if intermedios:
        url += '&waypoints=' + quote('|'.join(_coord(p) for p in intermedios))
    return url


def escribir_html(ruta_html, titulo, tramos, urls, km):
    filas = []
    for n, (t, u) in enumerate(zip(tramos, urls), 1):
        nombres = ' → '.join(html.escape(p['nombre'] or _coord(p)) for p in t)
        filas.append('<li><a href="%s" target="_blank">Tramo %d (%d puntos)</a>'
                     '<br><small>%s</small></li>' % (html.escape(u), n, len(t), nombres))
    with open(ruta_html, 'w', encoding='utf-8') as f:
        f.write('<!doctype html><html lang="es"><head><meta charset="utf-8">'
                '<meta name="viewport" content="width=device-width, initial-scale=1">'
                '<title>%s</title><style>body{font-family:sans-serif;max-width:700px;'
                'margin:2em auto;padding:0 1em;line-height:1.5}li{margin:1em 0}'
                'a{font-size:1.15em}</style></head><body><h1>%s</h1>'
                '<p>Longitud aproximada del track: %.1f km. Abre los tramos en orden.</p>'
                '<ol>%s</ol></body></html>'
                % (html.escape(titulo), html.escape(titulo), km, ''.join(filas)))


# ---------------------------------------------------------------- principal

def main():
    ap = argparse.ArgumentParser(description='Convierte un GPX en rutas de Google Maps.')
    ap.add_argument('gpx', help='Archivo GPX de entrada')
    ap.add_argument('--modo', default='driving',
                    choices=['walking', 'bicycling', 'driving', 'transit'],
                    help='Modo de desplazamiento (por defecto: driving)')
    ap.add_argument('--usar', default='auto', choices=['auto', 'wpt', 'track'],
                    help='auto: waypoints si existen; wpt: solo waypoints; '
                         'track: ignora waypoints y simplifica el track')
    ap.add_argument('--puntos', type=int, default=25,
                    help='Nº de puntos a extraer del track cuando se usa el track (por defecto 25)')
    ap.add_argument('--max-paradas', type=int, default=9,
                    help='Paradas intermedias por enlace (Google admite 9; '
                         'en navegador móvil a veces solo 3)')
    ap.add_argument('--umbral', type=float,
                    help='Distancia en metros por debajo de la cual dos puntos se consideran '
                         'el mismo (por defecto: 500 en coche, 250 en bici, 100 a pie)')
    ap.add_argument('--salida', help='Archivo HTML de salida (por defecto: <gpx>_google_maps.html)')
    args = ap.parse_args()

    umbral = args.umbral or {'driving': 500, 'transit': 500,
                             'bicycling': 250, 'walking': 100}[args.modo]

    wpts, track = leer_gpx(args.gpx)
    if not wpts and not track:
        sys.exit('El GPX no contiene waypoints ni puntos de track.')

    usar_wpt = args.usar == 'wpt' or (args.usar == 'auto' and len(wpts) >= 2)
    if usar_wpt and not wpts:
        sys.exit('Has pedido usar waypoints, pero el GPX no tiene ninguno.')
    if not usar_wpt and not track:
        sys.exit('El GPX no tiene track; prueba con --usar wpt.')

    if usar_wpt:
        puntos = ordenar_waypoints_por_track(wpts, track, umbral) if track else list(wpts)
        origen = 'waypoints (%d)' % len(wpts)
    else:
        puntos = simplificar_track(track, max(2, args.puntos))
        origen = 'track simplificado (%d de %d puntos)' % (len(puntos), len(track))

    base = track if track else puntos
    km = sum(distancia(base[i], base[i + 1]) for i in range(len(base) - 1)) / 1000

    puntos = fusionar_cercanos(puntos, umbral)
    tramos = dividir_en_tramos(puntos, args.max_paradas)
    urls = [url_google_maps(t, args.modo) for t in tramos]

    salida = args.salida or os.path.splitext(args.gpx)[0] + '_google_maps.html'
    titulo = os.path.splitext(os.path.basename(args.gpx))[0]
    escribir_html(salida, titulo, tramos, urls, km)

    print('Puntos usados: %s' % origen)
    print('Longitud aproximada: %.1f km  |  Tramos: %d\n' % (km, len(tramos)))
    for n, u in enumerate(urls, 1):
        print('Tramo %d:\n%s\n' % (n, u))
    print('Enlaces guardados en: %s' % salida)


if __name__ == '__main__':
    main()
