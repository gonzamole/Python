#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wikiloc_gpx.py
==============
Recorre un directorio buscando ficheros .gpx (de Wikiloc o no) y genera
una tabla con una fila por track y estas columnas:

    Nombre del archivo | Wikiloc | Origen | Duplicado | Autor | Enlace |
    País | Tipo | Inicio (mapa) | Dist (km) | Alt mín (m) | Alt máx (m) |
    Desnivel + (m) | Inicio (UTC) | Fin (UTC) | Duración (h) |
    En movim. (h) | Parado (h)

  · «Wikiloc» vale Sí/No según el fichero proceda o no de Wikiloc (se detecta
    por el atributo creator o por un enlace a wikiloc.com). Con --solo-wikiloc
    se descartan los que no lo son.
  · «Origen» es el programa o sitio que generó el GPX, y «Enlace» usa el de
    Wikiloc o, si no lo hay, cualquier otro enlace que traiga el fichero.
  · «Duplicado» marca las rutas con el mismo inicio, fin y distancia: la
    primera es el original y las demás salen como «Sospechoso» (otro nombre)
    o «Mismo nombre».
  · «Tipo» distingue rutas circulares de lineales, e «Inicio (mapa)» enlaza
    con Google Maps para llegar al punto de partida.
  · «En movim.» y «Parado» separan el tiempo real de marcha de las paradas.

Las métricas se calculan a partir de los trackpoints (<ele> y <time>):
    - Dist (km):      suma de distancias horizontales (haversine) entre puntos.
    - Alt mín/máx:    mínimo y máximo de <ele>.
    - Desnivel + (m): desnivel positivo acumulado con filtro de umbral
                      (--umbral, por defecto 4 m) para descartar el ruido del GPS.
    - Inicio/Fin:     primer y último <time> del track, en UTC (así vienen).
    - Duración (h):   diferencia Fin-Inicio en horas decimales.

Formatos de salida (uno o varios a la vez con -f):
    txt   -> texto alineado en columnas (monoespaciado)
    csv   -> CSV con BOM y ';' (Excel en español lo abre bien, coma decimal)
    xlsx  -> Excel real (números reales; requiere 'openpyxl'; si falta, cae a CSV)
    html  -> página web con tabla ordenable, filtros en la cabecera (texto, país,
             origen, tipo, duplicados, rangos de distancia y desnivel) y una
             miniatura del trazado en cada fila que se amplía con el perfil de
             altitud al pasar el ratón por encima.
             Permite marcar filas y generar un .bat que copia o mueve esos ficheros
             a una carpeta, o que los agrupa en subcarpetas por país (botón
             «Descargar .bat por paises»). Los que no son de Wikiloc van a
             «No_Wikiloc» y los posibles duplicados con otro nombre a
             «Sospechosos». Si el fichero ya existe en destino, se omite.
             La carpeta raíz se puede pasar como primer parámetro al .bat.
    movil -> versión ligera para el teléfono, en <salida>_movil.html: solo rutas
             de Wikiloc, en fichas con duración, país, distancia, alturas,
             desnivel, circular/lineal, enlace a Wikiloc y enlace a Maps para
             llegar al inicio. Filtra por país y ordena por cercanía usando el
             GPS del dispositivo (o unas coordenadas escritas a mano).

País (en este orden): mención explícita en el nombre/desc; si no, geocodificación
inversa OFFLINE por coordenadas (reverse_geocoder si está instalado; si no, tabla
embebida aproximada, marcada con "(aprox.)").

Uso:
    python wikiloc_gpx.py [DIRECTORIO] [-o BASE] [-f FORMATOS] [-r]
                          [--enlace {ruta,autor}] [--umbral METROS]

Ejemplos:
    python wikiloc_gpx.py "C:\\rutas\\wikiloc" -o rutas -f all -r
    python wikiloc_gpx.py ~/gpx -o rutas -f xlsx,html --umbral 5
"""

import argparse
import csv as _csv
import html as _html
import math
import os
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

# --------------------------------------------------------------------------- #
#  Dependencias opcionales
# --------------------------------------------------------------------------- #
try:
    import reverse_geocoder as _rg          # pip install reverse_geocoder
    _RG_DISPONIBLE = True
except Exception:
    _RG_DISPONIBLE = False

try:
    import openpyxl                          # pip install openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    _XLSX_DISPONIBLE = True
except Exception:
    _XLSX_DISPONIBLE = False

# --------------------------------------------------------------------------- #
#  Definición de columnas: (clave, cabecera, tipo, decimales)
#  tipo: txt | url | num | dt
# --------------------------------------------------------------------------- #
COLS = [
    ("archivo",      "Nombre del archivo", "txt", None),
    ("ruta_rel",     "Carpeta origen",     "txt", None),
    ("wikiloc",      "Wikiloc",            "txt", None),
    ("origen",       "Origen",             "txt", None),
    ("dup",          "Duplicado",          "txt", None),
    ("carpeta",      "Carpeta destino",    "txt", None),
    ("autor",        "Autor",              "txt", None),
    ("enlace",       "Enlace",             "url", None),
    ("pais",         "País",               "txt", None),
    ("tipo",         "Tipo",               "txt", None),
    ("coord_inicio", "Inicio (mapa)",      "geo", None),
    ("dist_km",      "Dist (km)",          "num", 2),
    ("alt_min",      "Alt mín (m)",        "num", 0),
    ("alt_max",      "Alt máx (m)",        "num", 0),
    ("desnivel_pos", "Desnivel + (m)",     "num", 0),
    ("inicio",       "Inicio (UTC)",       "dt",  None),
    ("fin",          "Fin (UTC)",          "dt",  None),
    ("dur_h",        "Duración (h)",       "num", 2),
    ("t_mov_h",      "En movim. (h)",      "num", 2),
    ("t_par_h",      "Parado (h)",         "num", 2),
]
CABECERA = [c[1] for c in COLS]

# --------------------------------------------------------------------------- #
#  Nombres de país en español por código ISO-3166 alfa-2
# --------------------------------------------------------------------------- #
COUNTRY_ES = {
    "AD": "Andorra", "AL": "Albania", "AT": "Austria", "BA": "Bosnia y Herzegovina",
    "BE": "Bélgica", "BG": "Bulgaria", "BY": "Bielorrusia", "CH": "Suiza",
    "CY": "Chipre", "CZ": "Chequia", "DE": "Alemania", "DK": "Dinamarca",
    "EE": "Estonia", "ES": "España", "FI": "Finlandia", "FR": "Francia",
    "GB": "Reino Unido", "GR": "Grecia", "HR": "Croacia", "HU": "Hungría",
    "IE": "Irlanda", "IS": "Islandia", "IT": "Italia", "LI": "Liechtenstein",
    "LT": "Lituania", "LU": "Luxemburgo", "LV": "Letonia", "MC": "Mónaco",
    "MD": "Moldavia", "ME": "Montenegro", "MK": "Macedonia del Norte",
    "MT": "Malta", "NL": "Países Bajos", "NO": "Noruega", "PL": "Polonia",
    "PT": "Portugal", "RO": "Rumanía", "RS": "Serbia", "RU": "Rusia",
    "SE": "Suecia", "SI": "Eslovenia", "SK": "Eslovaquia", "SM": "San Marino",
    "TR": "Turquía", "UA": "Ucrania", "XK": "Kosovo",
    "US": "Estados Unidos", "CA": "Canadá", "MX": "México", "AR": "Argentina",
    "BR": "Brasil", "CL": "Chile", "CO": "Colombia", "PE": "Perú",
    "MA": "Marruecos", "DZ": "Argelia", "TN": "Túnez", "EG": "Egipto",
    "ZA": "Sudáfrica", "CN": "China", "JP": "Japón", "IN": "India",
    "AU": "Australia", "NZ": "Nueva Zelanda", "TH": "Tailandia", "NP": "Nepal",
}

# --------------------------------------------------------------------------- #
#  Detección de país por el TEXTO (nombre/descripción del track)
# --------------------------------------------------------------------------- #
NAME_HINTS = {
    "españa": "ES", "spain": "ES", "espagne": "ES", "portugal": "PT",
    "france": "FR", "francia": "FR", "italia": "IT", "italy": "IT",
    "italie": "IT", "andorra": "AD", "marruecos": "MA", "morocco": "MA",
    "maroc": "MA", "suiza": "CH", "switzerland": "CH", "schweiz": "CH",
    "austria": "AT", "österreich": "AT", "alemania": "DE", "germany": "DE",
    "deutschland": "DE", "eslovenia": "SI", "slovenia": "SI", "croacia": "HR",
    "croatia": "HR", "hrvatska": "HR", "grecia": "GR", "greece": "GR",
    "macedonia": "MK", "makedonija": "MK", "kosovo": "XK", "albania": "AL",
    "shqipëri": "AL", "shqiperi": "AL", "montenegro": "ME", "crna gora": "ME",
    "serbia": "RS", "srbija": "RS", "bosnia": "BA", "noruega": "NO",
    "norway": "NO", "norge": "NO", "islandia": "IS", "iceland": "IS",
    "nepal": "NP", "himalaya": "NP",
}

# --------------------------------------------------------------------------- #
#  Tabla de referencia embebida (fallback sin dependencias)
# --------------------------------------------------------------------------- #
REF_POINTS = [
    (40.4168, -3.7038, "ES"), (41.3874, 2.1686, "ES"), (37.3891, -5.9845, "ES"),
    (43.2630, -2.9350, "ES"), (39.4699, -0.3763, "ES"), (43.3619, -8.4115, "ES"),
    (42.6000, -6.8000, "ES"), (40.2440, -5.2797, "ES"), (37.1773, -3.5986, "ES"),
    (28.4636, -16.2518, "ES"),
    (38.7223, -9.1393, "PT"), (41.1579, -8.6291, "PT"), (37.0194, -7.9304, "PT"),
    (32.6669, -16.9241, "PT"), (42.5063, 1.5218, "AD"),
    (48.8566, 2.3522, "FR"), (43.6047, 1.4442, "FR"), (45.7640, 4.8357, "FR"),
    (43.7102, 7.2620, "FR"), (43.2965, 5.3698, "FR"), (48.5734, 7.7521, "FR"),
    (50.8503, 4.3517, "BE"), (49.6116, 6.1319, "LU"), (52.3676, 4.9041, "NL"),
    (46.9480, 7.4474, "CH"), (46.2044, 6.1432, "CH"), (46.0037, 8.9511, "CH"),
    (47.1410, 9.5209, "LI"), (43.7384, 7.4246, "MC"),
    (41.9028, 12.4964, "IT"), (45.4642, 9.1900, "IT"), (45.0703, 7.6869, "IT"),
    (43.7696, 11.2558, "IT"), (40.8518, 14.2681, "IT"), (38.1157, 13.3615, "IT"),
    (43.9160, 12.4467, "SM"),
    (48.2082, 16.3738, "AT"), (47.2692, 11.4041, "AT"),
    (52.5200, 13.4050, "DE"), (48.1351, 11.5820, "DE"), (50.9375, 6.9603, "DE"),
    (50.0755, 14.4378, "CZ"), (48.1486, 17.1077, "SK"), (47.4979, 19.0402, "HU"),
    (52.2297, 21.0122, "PL"), (50.0647, 19.9450, "PL"),
    (54.6872, 25.2797, "LT"), (56.9496, 24.1052, "LV"), (59.4370, 24.7536, "EE"),
    (53.9006, 27.5590, "BY"), (50.4501, 30.5234, "UA"), (47.0105, 28.8638, "MD"),
    (44.4268, 26.1025, "RO"), (42.6977, 23.3219, "BG"),
    (46.0569, 14.5058, "SI"), (45.8150, 15.9819, "HR"), (43.5081, 16.4402, "HR"),
    (43.8563, 18.4131, "BA"), (42.4304, 19.2594, "ME"),
    (44.7866, 20.4489, "RS"), (43.1367, 20.5122, "RS"),
    (42.6629, 21.1655, "XK"), (42.2139, 20.7397, "XK"),
    (41.9981, 21.4254, "MK"), (42.0100, 20.9700, "MK"),
    (41.3275, 19.8187, "AL"), (42.0693, 19.5033, "AL"),
    (37.9838, 23.7275, "GR"), (40.6401, 22.9444, "GR"),
    (55.6761, 12.5683, "DK"), (59.9139, 10.7522, "NO"), (60.3913, 5.3221, "NO"),
    (59.3293, 18.0686, "SE"), (60.1699, 24.9384, "FI"), (64.1466, -21.9426, "IS"),
    (51.5074, -0.1278, "GB"), (55.9533, -3.1883, "GB"), (53.4808, -2.2426, "GB"),
    (53.3498, -6.2603, "IE"), (55.7558, 37.6173, "RU"),
    (35.1856, 33.3823, "CY"), (35.8989, 14.5146, "MT"),
    (39.9334, 32.8597, "TR"), (41.0082, 28.9784, "TR"),
    (34.0209, -6.8416, "MA"), (31.6295, -7.9811, "MA"), (35.7595, -5.8340, "MA"),
    (36.7538, 3.0588, "DZ"), (36.8065, 10.1815, "TN"), (30.0444, 31.2357, "EG"),
    (38.9072, -77.0369, "US"), (40.7128, -74.0060, "US"), (34.0522, -118.2437, "US"),
    (45.4215, -75.6972, "CA"), (43.6532, -79.3832, "CA"),
    (19.4326, -99.1332, "MX"), (-34.6037, -58.3816, "AR"), (-15.7939, -47.8828, "BR"),
    (-33.4489, -70.6693, "CL"), (4.7110, -74.0721, "CO"), (-12.0464, -77.0428, "PE"),
    (-25.7479, 28.2293, "ZA"), (39.9042, 116.4074, "CN"), (35.6762, 139.6503, "JP"),
    (28.6139, 77.2090, "IN"), (27.7172, 85.3240, "NP"), (13.7563, 100.5018, "TH"),
    (-35.2809, 149.1300, "AU"), (-33.8688, 151.2093, "AU"), (-41.2865, 174.7762, "NZ"),
]


# --------------------------------------------------------------------------- #
#  Utilidades XML / tiempo
# --------------------------------------------------------------------------- #
def _local(tag):
    return tag.rsplit("}", 1)[-1].lower() if "}" in tag else tag.lower()


def _find_first(root, nombre_local):
    for el in root.iter():
        if _local(el.tag) == nombre_local:
            return el
    return None


def _find_in(parent, nombre_local):
    for el in parent.iter():
        if el is not parent and _local(el.tag) == nombre_local:
            return el
    return None


def parse_time(s):
    """Convierte una marca de tiempo ISO-8601 GPX en datetime (o None)."""
    if not s:
        return None
    s = s.strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except Exception:
        pass
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z",
                "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None


def _a_utc_naive(dt):
    """datetime -> UTC sin tzinfo (para mostrar y para Excel)."""
    if dt is None:
        return None
    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt


# --------------------------------------------------------------------------- #
#  Cálculos geográficos
# --------------------------------------------------------------------------- #
def _haversine(lat1, lon1, lat2, lon2):
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlmb / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def _cc_por_tabla(lat, lon):
    mejor_cc, mejor_d = None, float("inf")
    for rlat, rlon, cc in REF_POINTS:
        d = _haversine(lat, lon, rlat, rlon)
        if d < mejor_d:
            mejor_cc, mejor_d = cc, d
    return mejor_cc


def cc_por_coordenadas(coords):
    if not coords:
        return None, False
    if _RG_DISPONIBLE:
        try:
            res = _rg.search(coords, verbose=False)
            ccs = [r.get("cc") for r in res if r.get("cc")]
            if ccs:
                return max(set(ccs), key=ccs.count), False
        except Exception:
            pass
    ccs = [c for c in (_cc_por_tabla(la, lo) for la, lo in coords) if c]
    if ccs:
        return max(set(ccs), key=ccs.count), True
    return None, True


def cc_por_nombre(texto):
    if not texto:
        return None
    t = " " + texto.lower() + " "
    for clave, cc in NAME_HINTS.items():
        if clave in t:
            return cc
    return None


def distancia_km(puntos):
    coords = [(p["lat"], p["lon"]) for p in puntos
              if p["lat"] is not None and p["lon"] is not None]
    if not coords:
        return None
    d = 0.0
    for (la1, lo1), (la2, lo2) in zip(coords, coords[1:]):
        d += _haversine(la1, lo1, la2, lo2)
    return d


def desnivel_positivo(eles, umbral):
    """Desnivel positivo acumulado con filtro de histéresis (umbral en metros)."""
    eles = [e for e in eles if e is not None]
    if len(eles) < 2:
        return None
    ganancia, ref = 0.0, eles[0]
    for e in eles[1:]:
        d = e - ref
        if d >= umbral:
            ganancia += d
            ref = e
        elif d <= -umbral:
            ref = e
    return ganancia


def tiempos_movimiento(pts, v_min=1.0, hueco_max_s=3600):
    """(horas en movimiento, horas parado) a partir de los tiempos de cada punto.

    Un tramo cuenta como movimiento si la velocidad media supera v_min km/h.
    Los saltos de más de hueco_max_s (grabación pausada) se ignoran salvo que
    el desplazamiento sea pequeño, en cuyo caso se cuentan como parada.
    """
    mov = par = 0.0
    prev = None
    hay = False
    for p in pts:
        if p["time"] is None or p["lat"] is None or p["lon"] is None:
            continue
        if prev is not None:
            dt = (_a_utc_naive(p["time"]) - _a_utc_naive(prev["time"])).total_seconds()
            d = _haversine(prev["lat"], prev["lon"], p["lat"], p["lon"])
            if dt > 0:
                hay = True
                if dt > hueco_max_s:
                    if d < 1.0:
                        par += dt
                elif d / (dt / 3600.0) >= v_min:
                    mov += dt
                else:
                    par += dt
        prev = p
    if not hay:
        return None, None
    return mov / 3600.0, par / 3600.0


def tipo_ruta(pts, dist_km):
    """'Circular' si el final vuelve cerca del inicio, 'Lineal' en caso contrario."""
    coords = [(p["lat"], p["lon"]) for p in pts
              if p["lat"] is not None and p["lon"] is not None]
    if len(coords) < 3 or not dist_km:
        return ""
    cierre = _haversine(coords[0][0], coords[0][1], coords[-1][0], coords[-1][1])
    return "Circular" if cierre <= max(0.2, dist_km * 0.08) else "Lineal"


def _muestrear(lista, n):
    """Reduce una lista a como mucho n elementos conservando el primero y el último."""
    if len(lista) <= n:
        return list(lista)
    paso = (len(lista) - 1) / float(n - 1)
    out = [lista[int(round(i * paso))] for i in range(n)]
    out[-1] = lista[-1]
    return out


def traza_normalizada(pts, n=140, lado=100.0):
    """Polilínea del trazado escalada a una caja de `lado` unidades.

    Devuelve {'w', 'h', 'p'} con 'p' en formato 'x,y x,y ...' listo para un
    <polyline> de SVG (eje Y ya invertido). None si no hay puntos suficientes.
    """
    coords = [(p["lat"], p["lon"]) for p in pts
              if p["lat"] is not None and p["lon"] is not None]
    if len(coords) < 2:
        return None
    coords = _muestrear(coords, n)
    lat_m = sum(c[0] for c in coords) / len(coords)
    k = math.cos(math.radians(lat_m)) or 1e-6
    xs = [c[1] * k for c in coords]
    ys = [c[0] for c in coords]
    dx, dy = max(xs) - min(xs), max(ys) - min(ys)
    if dx <= 0 and dy <= 0:
        return None
    if dx >= dy:
        w, h = lado, (lado * dy / dx if dx > 0 else 0.0)
    else:
        w, h = (lado * dx / dy if dy > 0 else 0.0), lado
    w, h = max(w, 3.0), max(h, 3.0)
    x0, y0 = min(xs), min(ys)
    pares = []
    for x, y in zip(xs, ys):
        px = (x - x0) / dx * w if dx > 0 else w / 2.0
        py = h - (y - y0) / dy * h if dy > 0 else h / 2.0
        pares.append(f"{px:.1f},{py:.1f}")
    return {"w": round(w, 1), "h": round(h, 1), "p": " ".join(pares)}


def perfil_normalizado(pts, n=140, ancho=100.0, alto=40.0):
    """Perfil de altitud escalado a una caja ancho x alto, en formato polyline."""
    eles = [p["ele"] for p in pts if p["ele"] is not None]
    if len(eles) < 2:
        return ""
    eles = _muestrear(eles, n)
    lo, hi = min(eles), max(eles)
    rango = (hi - lo) or 1.0
    paso = ancho / (len(eles) - 1)
    return " ".join(f"{i * paso:.1f},{alto - (e - lo) / rango * alto:.1f}"
                    for i, e in enumerate(eles))


def marcar_duplicados(regs, tol_km=0.1):
    """Agrupa rutas con el mismo inicio, final y distancia, y las etiqueta.

    Dentro de cada grupo, el primero es el original; los demás son
    «Mismo nombre» (el .bat los omitirá al existir ya en destino) o
    «Sospechoso» si el nombre difiere pero el trazado coincide.
    """
    for r in regs:
        r["dup"] = ""
        r["dup_cod"] = ""
    grupos = {}
    for r in regs:
        if r.get("lat0") is None or r.get("dist_km") is None:
            continue
        firma = (round(r["lat0"], 3), round(r["lon0"], 3),
                 round(r["latn"], 3), round(r["lonn"], 3),
                 round(r["dist_km"] / tol_km))
        grupos.setdefault(firma, []).append(r)
    n = 0
    for lista in grupos.values():
        if len(lista) < 2:
            continue
        n += 1
        lista.sort(key=lambda r: (not r["es_wikiloc"], r["archivo"].lower()))
        canon = lista[0]
        canon["dup"], canon["dup_cod"] = f"Original #{n}", "orig"
        for r in lista[1:]:
            if r["archivo"].lower() == canon["archivo"].lower():
                r["dup"], r["dup_cod"] = f"Mismo nombre #{n}", "rep"
            else:
                r["dup"], r["dup_cod"] = f"Sospechoso #{n}", "sosp"
    return n


# --------------------------------------------------------------------------- #
#  Análisis de un fichero GPX
# --------------------------------------------------------------------------- #
def extraer_puntos(root):
    """Lista ordenada de puntos {lat, lon, ele, time} a partir de trkpt/rtept/wpt."""
    def leer(el):
        try:
            lat = float(el.get("lat"))
            lon = float(el.get("lon"))
        except (TypeError, ValueError):
            return None
        ele, t = None, None
        for ch in el:
            lc = _local(ch.tag)
            if lc == "ele":
                try:
                    ele = float(ch.text)
                except (TypeError, ValueError):
                    pass
            elif lc == "time":
                t = parse_time(ch.text)
        return {"lat": lat, "lon": lon, "ele": ele, "time": t}

    for etiqueta in ("trkpt", "rtept", "wpt"):
        pts = []
        for el in root.iter():
            if _local(el.tag) == etiqueta:
                p = leer(el)
                if p:
                    pts.append(p)
        if pts:
            return pts
    return []


def analizar_gpx(ruta, umbral):
    try:
        root = ET.parse(ruta).getroot()
    except ET.ParseError:
        return {"error": "XML no válido"}
    except Exception as e:
        return {"error": str(e)}

    creator = (root.get("creator") or "").lower()
    hrefs = [(el.get("href") or "") for el in root.iter() if _local(el.tag) == "link"]
    es_wikiloc = ("wikiloc" in creator) or any("wikiloc.com" in h.lower() for h in hrefs)

    # Autor
    autor = ""
    autor_el = _find_first(root, "author")
    if autor_el is not None:
        n = _find_in(autor_el, "name")
        if n is not None and (n.text or "").strip():
            autor = n.text.strip()
    if not autor:
        cop = _find_first(root, "copyright")
        if cop is not None and (cop.get("author") or "").strip():
            autor = cop.get("author").strip()

    # Enlaces
    enlace_ruta, enlace_autor, enlace_ext = "", "", ""
    for h in hrefs:
        hl = h.lower()
        if "wikiloc.com" not in hl:
            if not enlace_ext and hl.startswith(("http://", "https://")):
                enlace_ext = h
            continue
        if "user.do" in hl or "/user" in hl:
            enlace_autor = enlace_autor or h
        elif not enlace_ruta:
            enlace_ruta = h
    if not enlace_ruta:
        for h in hrefs:
            if "wikiloc.com" in h.lower() and "user.do" not in h.lower():
                enlace_ruta = h
                break

    # Sitio de origen: el creator del GPX, o el dominio del enlace externo
    origen = (root.get("creator") or "").strip()
    origen = re.split(r"\s+[-–]\s+|\s+https?://", origen)[0].strip()
    if not origen and enlace_ext:
        m = re.match(r"https?://(?:www\.)?([^/]+)", enlace_ext, re.I)
        origen = m.group(1) if m else ""
    if es_wikiloc and not origen:
        origen = "Wikiloc"

    # Nombre y descripción
    #  Ojo: el primer <name> del fichero suele ser el del autor, dentro de
    #  <metadata><author>. El nombre bueno es el hijo directo de <trk>/<rte>.
    nombre_track = ""
    for etiqueta in ("trk", "rte", "metadata"):
        cont = _find_first(root, etiqueta)
        if cont is None:
            continue
        for ch in cont:
            if _local(ch.tag) == "name" and (ch.text or "").strip():
                nombre_track = ch.text.strip()
                break
        if nombre_track:
            break
    if not nombre_track:
        for el in root.iter():
            if _local(el.tag) == "name" and (el.text or "").strip():
                nombre_track = el.text.strip()
                break
    d = _find_first(root, "desc")
    desc = (d.text or "").strip() if d is not None else ""

    # Puntos y métricas
    pts = extraer_puntos(root)
    eles = [p["ele"] for p in pts if p["ele"] is not None]
    tiempos = [p["time"] for p in pts if p["time"] is not None]

    dist_km = distancia_km(pts)
    alt_min = min(eles) if eles else None
    alt_max = max(eles) if eles else None
    desn_pos = desnivel_positivo([p["ele"] for p in pts], umbral)
    inicio = _a_utc_naive(tiempos[0]) if tiempos else None
    fin = _a_utc_naive(tiempos[-1]) if tiempos else None
    dur_h = None
    if inicio and fin:
        dur_h = (fin - inicio).total_seconds() / 3600.0
    t_mov_h, t_par_h = tiempos_movimiento(pts)
    tipo = tipo_ruta(pts, dist_km)
    traza = traza_normalizada(pts)
    perfil = perfil_normalizado(pts)

    # Punto de inicio y de fin (para el enlace a Maps y para los duplicados)
    con_coord = [p for p in pts if p["lat"] is not None and p["lon"] is not None]
    lat0 = lon0 = latn = lonn = None
    coord_inicio = ""
    if con_coord:
        lat0, lon0 = con_coord[0]["lat"], con_coord[0]["lon"]
        latn, lonn = con_coord[-1]["lat"], con_coord[-1]["lon"]
        coord_inicio = f"{lat0:.5f}, {lon0:.5f}"

    # País
    coords_clave = []
    if pts:
        coords_clave.append((pts[0]["lat"], pts[0]["lon"]))
        if len(pts) > 1:
            coords_clave.append((pts[-1]["lat"], pts[-1]["lon"]))
    aprox = False
    cc = cc_por_nombre(nombre_track + " " + desc)
    if not cc:
        cc, aprox = cc_por_coordenadas(coords_clave)
    pais = COUNTRY_ES.get(cc, cc or "Desconocido")
    if aprox and cc:
        pais += " (aprox.)"

    return {
        "archivo": os.path.basename(ruta),
        "nombre": nombre_track,
        "es_wikiloc": es_wikiloc,
        "wikiloc": "Sí" if es_wikiloc else "No",
        "origen": origen or "(desconocido)",
        "autor": autor or "(desconocido)",
        "enlace_ruta": enlace_ruta,
        "enlace_autor": enlace_autor,
        "enlace_ext": enlace_ext,
        "pais": pais,
        "tipo": tipo,
        "coord_inicio": coord_inicio,
        "lat0": lat0, "lon0": lon0, "latn": latn, "lonn": lonn,
        "traza": traza,
        "perfil": perfil,
        "dist_km": dist_km,
        "alt_min": alt_min,
        "alt_max": alt_max,
        "desnivel_pos": desn_pos,
        "inicio": inicio,
        "fin": fin,
        "dur_h": dur_h,
        "t_mov_h": t_mov_h,
        "t_par_h": t_par_h,
    }


# --------------------------------------------------------------------------- #
#  Formateo de valores para mostrar (texto)
# --------------------------------------------------------------------------- #
def _fmt_num(x, dec):
    if x is None:
        return ""
    if dec == 0:
        return f"{round(x):d}"
    return f"{x:.{dec}f}".replace(".", ",")   # coma decimal (español)


def _fmt_dt(dt):
    return dt.strftime("%Y-%m-%d %H:%M") if dt is not None else ""


def disp(rec, col):
    """Cadena visible de una celda (para txt, csv y HTML)."""
    key, _cab, tipo, dec = col
    v = rec.get(key)
    if tipo == "dt":
        return _fmt_dt(v)
    if tipo == "num":
        return _fmt_num(v, dec)
    return "" if v is None else str(v)


# --------------------------------------------------------------------------- #
#  Escritores de salida
# --------------------------------------------------------------------------- #
def escribir_txt(regs, ruta):
    filas = [[disp(r, c) for c in COLS] for r in regs]
    anchos = [max([len(CABECERA[i])] + [len(f[i]) for f in filas]) if filas
              else len(CABECERA[i]) for i in range(len(COLS))]
    der = {"num", "dt"}  # columnas alineadas a la derecha

    def just(txt, i):
        return txt.rjust(anchos[i]) if COLS[i][2] in der else txt.ljust(anchos[i])

    def linea(vals):
        return " | ".join(just(vals[i], i) for i in range(len(COLS)))

    with open(ruta, "w", encoding="utf-8") as fh:
        fh.write(linea(CABECERA) + "\n")
        fh.write("-+-".join("-" * a for a in anchos) + "\n")
        for f in filas:
            fh.write(linea(f) + "\n")


def escribir_csv(regs, ruta):
    with open(ruta, "w", encoding="utf-8-sig", newline="") as fh:
        w = _csv.writer(fh, delimiter=";")
        w.writerow(CABECERA)
        for r in regs:
            w.writerow([disp(r, c) for c in COLS])


def carpeta_pais(pais):
    """Nombre de carpeta seguro (ASCII) a partir del valor de la columna País.

    'España' -> 'Espana' | 'Perú (aprox.)' -> 'Peru' | vacío/None -> 'desconocido'.
    Se quitan las tildes a propósito: el .bat se genera en UTF-8 sin BOM y los
    nombres acentuados son la primera fuente de problemas en consolas con la
    página de códigos heredada.
    """
    p = (pais or "").replace("(aprox.)", "").strip()
    if not p or p.lower().startswith("desconocido"):
        return "desconocido"
    p = unicodedata.normalize("NFKD", p)
    p = "".join(ch for ch in p if not unicodedata.combining(ch))
    p = re.sub(r'[<>:"/\\|?*]', "", p).strip(" .")
    return p or "desconocido"


def carpeta_destino(r):
    """Subcarpeta a la que va un fichero al agrupar por países.

    Prioridad: Sospechosos (posible duplicado con otro nombre) > No_Wikiloc >
    carpeta del país. Los duplicados con el MISMO nombre no se desvían: van a
    su carpeta normal y el .bat los omite porque el fichero ya existe allí.
    """
    if (r.get("dup_cod") or "") == "sosp":
        return "Sospechosos"
    if not r.get("es_wikiloc"):
        return "No_Wikiloc"
    return carpeta_pais(r.get("pais"))


def _fmt_horas(h):
    """1.48 -> '1:29'. Cadena vacía si no hay dato."""
    if h is None:
        return ""
    total = int(round(h * 60))
    return f"{total // 60}:{total % 60:02d}"


def svg_mini(traza, clase="mini"):
    """<svg> con el trazado de la ruta, o cadena vacía si no hay puntos."""
    if not traza:
        return ""
    w, h, p = traza["w"], traza["h"], traza["p"]
    prim = p.split(" ")[0].split(",")
    ult = p.split(" ")[-1].split(",")
    rad = max(w, h) / 28.0
    return (f'<svg class="{clase}" viewBox="{-rad:.1f} {-rad:.1f} '
            f'{w + 2 * rad:.1f} {h + 2 * rad:.1f}" preserveAspectRatio="xMidYMid meet">'
            f'<polyline points="{p}"/>'
            f'<circle class="ini" cx="{prim[0]}" cy="{prim[1]}" r="{rad:.1f}"/>'
            f'<circle class="fin" cx="{ult[0]}" cy="{ult[1]}" r="{rad:.1f}"/></svg>')


def escribir_html(regs, ruta):
    def td(rec, col):
        key, _c, tipo, _d = col
        texto = disp(rec, col)
        if tipo == "url" and texto.startswith("http"):
            return (f'<td class="u"><a href="{_html.escape(texto, quote=True)}" '
                    f'target="_blank" rel="noopener">{_html.escape(texto)}</a></td>')
        if tipo == "geo" and rec.get("lat0") is not None:
            url = ("https://www.google.com/maps/dir/?api=1&destination="
                   f'{rec["lat0"]:.6f},{rec["lon0"]:.6f}')
            return (f'<td class="g"><a href="{_html.escape(url, quote=True)}" '
                    f'target="_blank" rel="noopener" '
                    f'title="Como llegar al inicio con Google Maps">'
                    f'{_html.escape(texto)} &#128205;</a></td>')
        if key == "wikiloc":
            clase = ' class="wk-no"' if texto == "No" else ' class="wk-si"'
        elif key == "dup":
            cod = rec.get("dup_cod") or ""
            clase = f' class="dp dp-{cod}"' if cod else ' class="dp"'
        elif key == "carpeta":
            if texto == "Sospechosos":
                clase = ' class="cp cp-sosp"'
            elif texto == "No_Wikiloc":
                clase = ' class="cp cp-nowk"'
            else:
                clase = ' class="cp"'
        else:
            clase = ' class="n"' if tipo in ("num", "dt") else ""
        return f"<td{clase}>{_html.escape(texto)}</td>"

    def td_traza(r):
        tz = r.get("traza")
        if not tz:
            return '<td class="tz"></td>'
        pf = r.get("perfil") or ""
        return (f'<td class="tz" data-w="{tz["w"]}" data-h="{tz["h"]}" '
                f'data-p="{tz["p"]}" data-pf="{pf}" '
                f'data-nom="{_html.escape(r.get("archivo", ""), quote=True)}">'
                f'{svg_mini(tz)}</td>')

    def fila(r):
        busca = " ".join(str(r.get(k) or "") for k in
                         ("archivo", "autor", "pais", "origen", "tipo")).lower()
        attrs = (f'data-busca="{_html.escape(busca, quote=True)}" '
                 f'data-pais="{_html.escape(r.get("pais") or "", quote=True)}" '
                 f'data-wk="{"1" if r.get("es_wikiloc") else "0"}" '
                 f'data-tipo="{_html.escape(r.get("tipo") or "", quote=True)}" '
                 f'data-dup="{r.get("dup_cod") or ""}" '
                 f'data-carp="{_html.escape(r.get("carpeta") or "", quote=True)}" '
                 f'data-dist="{"" if r.get("dist_km") is None else round(r["dist_km"], 2)}" '
                 f'data-desn="{"" if r.get("desnivel_pos") is None else round(r["desnivel_pos"])}"')
        clase = "" if r.get("es_wikiloc") else "nowk"
        if (r.get("dup_cod") or "") == "sosp":
            clase = (clase + " sosp").strip()
        clase_tr = f' class="{clase}"' if clase else ""
        cb = (f'<td class="chk"><input type="checkbox" class="chk-row" '
              f'data-path="{_html.escape(r.get("ruta_abs", ""), quote=True)}" '
              f'data-nom="{_html.escape(r.get("archivo", ""), quote=True)}" '
              f'data-carpeta="{_html.escape(r.get("carpeta") or carpeta_destino(r), quote=True)}" '
              f'onchange="actualizar()"></td>')
        return (f"<tr{clase_tr} {attrs}>" + cb + td_traza(r)
                + "".join(td(r, c) for c in COLS) + "</tr>")

    trs = [fila(r) for r in regs]
    ths = ('<th class="chk"><input type="checkbox" id="all" '
           'onclick="toggleTodos(this)" title="Marcar / desmarcar lo visible"></th>'
           '<th class="tz">Traza</th>')
    ths += "".join(
        f'<th onclick="ordenar({i + 2})">{_html.escape(c[1])} '
        f'<span class="a">&#8597;</span></th>' for i, c in enumerate(COLS)
    )
    paises = sorted({r.get("pais") or "" for r in regs} - {""})
    carpetas = sorted({r.get("carpeta") or "" for r in regs} - {""})
    opt_carp = "".join(f'<option value="{_html.escape(c, quote=True)}">'
                       f'{_html.escape(c)}</option>' for c in carpetas)
    opt_pais = "".join(f'<option value="{_html.escape(p, quote=True)}">'
                       f'{_html.escape(p)}</option>' for p in paises)
    n_wk = sum(1 for r in regs if r.get("es_wikiloc"))
    n_sosp = sum(1 for r in regs if (r.get("dup_cod") or "") == "sosp")
    n_rep = sum(1 for r in regs if (r.get("dup_cod") or "") == "rep")

    doc = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mis rutas GPX</title>
<style>
  :root {{ color-scheme: light dark; }}
  body {{ font-family: system-ui, "Segoe UI", Roboto, Arial, sans-serif;
         margin: 1.5rem; line-height: 1.4; }}
  h1 {{ font-size: 1.35rem; margin: 0 0 .25rem; }}
  p.info {{ color:#666; margin:0 0 1rem; font-size:.9rem; }}
  table {{ border-collapse: collapse; width: 100%; font-size:.9rem; }}
  th, td {{ border:1px solid #d0d0d0; padding:.4rem .5rem; text-align:left;
           vertical-align:middle; }}
  th {{ background:#f3f4f6; cursor:pointer; user-select:none;
       position:sticky; top:0; white-space:nowrap; z-index:2; }}
  th .a {{ color:#aaa; font-size:.8em; }}
  td.n {{ text-align:right; white-space:nowrap; font-variant-numeric:tabular-nums; }}
  td.g {{ white-space:nowrap; font-variant-numeric:tabular-nums; }}
  tbody tr:nth-child(even) {{ background:rgba(0,0,0,.03); }}
  tbody tr.nowk td {{ background:rgba(217,119,6,.10); }}
  tbody tr.sosp td {{ background:rgba(220,38,38,.10); }}
  td.wk-si {{ text-align:center; white-space:nowrap; }}
  td.wk-no {{ text-align:center; white-space:nowrap; color:#b45309; font-weight:600; }}
  td.dp {{ white-space:nowrap; font-size:.85em; }}
  td.dp-sosp {{ color:#b91c1c; font-weight:600; }}
  td.dp-rep {{ color:#666; }}
  td.cp {{ white-space:nowrap; font-size:.85em; font-family:Consolas,monospace; }}
  td.cp-sosp {{ color:#b91c1c; font-weight:600; }}
  td.cp-nowk {{ color:#b45309; font-weight:600; }}
  td.u {{ word-break:break-all; max-width:300px; }}
  a {{ color:#1a56db; text-decoration:none; }} a:hover {{ text-decoration:underline; }}
  th.chk {{ cursor:default; width:2.2rem; text-align:center; }}
  td.chk {{ text-align:center; }}
  th.tz {{ cursor:default; width:7rem; }}
  td.tz {{ width:7rem; padding:.2rem; }}
  svg.mini {{ width:6.5rem; height:3.2rem; display:block; }}
  svg polyline {{ fill:none; stroke:#1a56db; stroke-width:1.6;
                 stroke-linejoin:round; stroke-linecap:round;
                 vector-effect:non-scaling-stroke; }}
  svg circle.ini {{ fill:#16a34a; }}
  svg circle.fin {{ fill:#dc2626; }}
  #lupa {{ display:none; position:fixed; z-index:50; pointer-events:none;
          padding:.5rem; border:1px solid #999; border-radius:10px;
          background:#fff; box-shadow:0 8px 26px rgba(0,0,0,.28); }}
  #lupa .ttl {{ font-size:.8rem; color:#555; margin-bottom:.3rem;
               max-width:19rem; overflow:hidden; text-overflow:ellipsis;
               white-space:nowrap; }}
  #lupa svg.big {{ width:19rem; height:13rem; }}
  #lupa svg.bigpf {{ width:19rem; height:4.5rem; margin-top:.3rem; }}
  #lupa svg.bigpf polyline {{ stroke:#7c3aed; }}
  #lupa .pie {{ font-size:.72rem; color:#777; margin-top:.25rem; }}
  .tb {{ display:flex; flex-wrap:wrap; gap:.5rem .8rem; align-items:center;
        margin:0 0 .6rem; padding:.6rem .7rem; border:1px solid #d0d0d0;
        border-radius:8px; background:rgba(0,0,0,.02); font-size:.9rem; }}
  .tb input[type=text], .tb input[type=number], .tb select {{ padding:.25rem .4rem; }}
  .tb input[type=number] {{ width:5rem; }}
  .tb button {{ padding:.3rem .7rem; cursor:pointer; }}
  #cuenta {{ color:#666; }}
  .prev {{ display:none; white-space:pre; overflow:auto; max-height:260px;
          margin:0 0 1rem; padding:.6rem .7rem; border:1px solid #d0d0d0;
          border-radius:8px; background:#0d1117; color:#e6edf3;
          font-family:Consolas,"Courier New",monospace; font-size:.82rem; }}
  @media (prefers-color-scheme: dark) {{
    th {{ background:#222; }} th, td {{ border-color:#444; }}
    p.info {{ color:#999; }} a {{ color:#7aa7ff; }}
    tbody tr.nowk td {{ background:rgba(217,119,6,.18); }}
    tbody tr.sosp td {{ background:rgba(220,38,38,.18); }}
    td.wk-no {{ color:#fbbf24; }}
    td.dp-sosp, td.cp-sosp {{ color:#fca5a5; }}
    td.cp-nowk {{ color:#fbbf24; }}
    svg polyline {{ stroke:#7aa7ff; }}
    #lupa {{ background:#15181d; border-color:#555; }}
    #lupa .ttl, #lupa .pie {{ color:#aaa; }}
    #lupa svg.bigpf polyline {{ stroke:#c4b5fd; }}
  }}
</style></head><body>
<h1>Mis rutas GPX</h1>
<p class="info">{len(regs)} fichero(s): {n_wk} de Wikiloc, {len(regs) - n_wk} de otras
 fuentes. Duplicados detectados: {n_sosp} sospechoso(s) y {n_rep} con el mismo nombre.
 Pasa el raton por encima de la miniatura para verla ampliada con el perfil de
 altitud. Pulsa una cabecera para ordenar. Horas en UTC.</p>
<div class="tb">
  <input type="text" id="fTxt" size="22" placeholder="Buscar nombre / autor / pais"
         oninput="filtrar()">
  <select id="fPais" onchange="filtrar()"><option value="">Todos los paises</option>
    {opt_pais}</select>
  <select id="fWk" onchange="filtrar()"><option value="">Wikiloc y otros</option>
    <option value="1">Solo Wikiloc</option><option value="0">Solo NO Wikiloc</option></select>
  <select id="fTipo" onchange="filtrar()"><option value="">Circulares y lineales</option>
    <option value="Circular">Solo circulares</option>
    <option value="Lineal">Solo lineales</option></select>
  <select id="fDup" onchange="filtrar()"><option value="">Con duplicados</option>
    <option value="no">Ocultar duplicados</option>
    <option value="sosp">Solo sospechosos</option>
    <option value="si">Solo duplicados</option></select>
  <select id="fCarp" onchange="filtrar()"><option value="">Todas las carpetas</option>
    {opt_carp}</select>
  <label>km <input type="number" id="fD1" step="1" placeholder="min" oninput="filtrar()">
    <input type="number" id="fD2" step="1" placeholder="max" oninput="filtrar()"></label>
  <label>D+ <input type="number" id="fH1" step="50" placeholder="min" oninput="filtrar()">
    <input type="number" id="fH2" step="50" placeholder="max" oninput="filtrar()"></label>
  <button onclick="limpiar()">Limpiar filtros</button>
</div>
<div class="tb">
  <label>Carpeta destino / raiz:
    <input type="text" id="dest" size="34" placeholder="C:\\rutas\\seleccion"></label>
  <label><input type="radio" name="op" id="copiar" checked> Copiar</label>
  <label><input type="radio" name="op" id="mover"> Mover</label>
  <button onclick="descargar()">Descargar .bat</button>
  <button onclick="descargarPaises()">Descargar .bat por paises</button>
  <button onclick="copiar()">Copiar comandos</button>
  <button onclick="previsualizar()">Previsualizar</button>
  <span id="cuenta">0 marcados</span>
</div>
<pre id="prev" class="prev"></pre>
<div id="lupa"></div>
<table id="t"><thead><tr>{ths}</tr></thead><tbody>
{os.linesep.join(trs)}
</tbody></table>
<script>
function F(id){{ return document.getElementById(id); }}
function num(s){{ s=(s||'').trim(); if(s==='') return null;
  var n=Number(s.replace(/\\s/g,'').replace(',','.')); return isNaN(n)?null:n; }}
function filas(){{ return Array.prototype.slice.call(
  document.querySelectorAll('#t tbody tr')); }}
function visible(tr){{ return tr.style.display!=='none'; }}

function ordenar(col){{
  var tb=document.querySelector('#t tbody');
  var fs=filas();
  var asc=tb.getAttribute('data-col')!=col||tb.getAttribute('data-dir')!='asc';
  fs.sort(function(a,b){{
    var x=a.cells[col].innerText, y=b.cells[col].innerText;
    var nx=num(x), ny=num(y), cmp;
    if(nx!==null&&ny!==null) cmp=nx-ny;
    else if(x===''&&y!=='') cmp=1;
    else if(y===''&&x!=='') cmp=-1;
    else cmp=(x.toLowerCase()<y.toLowerCase()?-1:x.toLowerCase()>y.toLowerCase()?1:0);
    return cmp*(asc?1:-1);
  }});
  fs.forEach(function(f){{ tb.appendChild(f); }});
  tb.setAttribute('data-col',col); tb.setAttribute('data-dir',asc?'asc':'desc');
}}

function filtrar(){{
  var txt=(F('fTxt').value||'').toLowerCase().trim();
  var pais=F('fPais').value, wk=F('fWk').value, tipo=F('fTipo').value;
  var dup=F('fDup').value, carp=F('fCarp').value;
  var d1=num(F('fD1').value), d2=num(F('fD2').value);
  var h1=num(F('fH1').value), h2=num(F('fH2').value);
  var vis=0;
  filas().forEach(function(tr){{
    var d=tr.dataset, ok=true;
    if(txt && d.busca.indexOf(txt)<0) ok=false;
    if(ok&&pais && d.pais!==pais) ok=false;
    if(ok&&wk && d.wk!==wk) ok=false;
    if(ok&&tipo && d.tipo!==tipo) ok=false;
    if(ok&&dup==='no' && (d.dup==='sosp'||d.dup==='rep')) ok=false;
    if(ok&&dup==='sosp' && d.dup!=='sosp') ok=false;
    if(ok&&dup==='si' && !d.dup) ok=false;
    if(ok&&carp && d.carp!==carp) ok=false;
    var dk=d.dist===''?null:+d.dist, dh=d.desn===''?null:+d.desn;
    if(ok&&d1!==null && (dk===null||dk<d1)) ok=false;
    if(ok&&d2!==null && (dk===null||dk>d2)) ok=false;
    if(ok&&h1!==null && (dh===null||dh<h1)) ok=false;
    if(ok&&h2!==null && (dh===null||dh>h2)) ok=false;
    tr.style.display=ok?'':'none';
    if(ok) vis++;
  }});
  actualizar(vis);
}}
function limpiar(){{
  ['fTxt','fD1','fD2','fH1','fH2'].forEach(function(i){{ F(i).value=''; }});
  ['fPais','fWk','fTipo','fDup','fCarp'].forEach(function(i){{ F(i).value=''; }});
  filtrar();
}}
function actualizar(vis){{
  if(vis===undefined) vis=filas().filter(visible).length;
  var marc=filas().filter(function(tr){{ return tr.cells[0].firstChild.checked; }});
  var ocultos=marc.filter(function(tr){{ return !visible(tr); }}).length;
  var t=marc.length+' marcados de '+vis+' visibles';
  if(ocultos) t+=' ('+ocultos+' marcados estan ocultos por el filtro)';
  F('cuenta').textContent=t;
}}
function toggleTodos(cb){{
  filas().forEach(function(tr){{
    if(visible(tr)) tr.cells[0].firstChild.checked=cb.checked;
  }});
  actualizar();
}}

/* ---------- lupa del trazado ---------- */
function svgGrande(td){{
  var w=+td.dataset.w, h=+td.dataset.h, p=td.dataset.p, pf=td.dataset.pf;
  var pt=p.split(' '), a=pt[0].split(','), z=pt[pt.length-1].split(',');
  var r=Math.max(w,h)/45;
  var s='<div class="ttl">'+td.dataset.nom+'</div>';
  s+='<svg class="big" viewBox="'+(-r*2)+' '+(-r*2)+' '+(w+r*4)+' '+(h+r*4)+
     '" preserveAspectRatio="xMidYMid meet"><polyline points="'+p+'"/>'+
     '<circle class="ini" cx="'+a[0]+'" cy="'+a[1]+'" r="'+r+'"/>'+
     '<circle class="fin" cx="'+z[0]+'" cy="'+z[1]+'" r="'+r+'"/></svg>';
  if(pf) s+='<svg class="bigpf" viewBox="-1 -1 102 42" '+
            'preserveAspectRatio="none"><polyline points="'+pf+'"/></svg>';
  s+='<div class="pie">Verde: inicio &nbsp;|&nbsp; Rojo: final'+
     (pf?' &nbsp;|&nbsp; Abajo, perfil de altitud':'')+'</div>';
  return s;
}}
function colocarLupa(e){{
  var l=F('lupa'), m=14;
  var x=e.clientX+m, y=e.clientY+m;
  if(x+l.offsetWidth>window.innerWidth-8) x=e.clientX-l.offsetWidth-m;
  if(y+l.offsetHeight>window.innerHeight-8) y=e.clientY-l.offsetHeight-m;
  l.style.left=Math.max(4,x)+'px'; l.style.top=Math.max(4,y)+'px';
}}
document.addEventListener('mouseover',function(e){{
  var td=e.target.closest?e.target.closest('td.tz'):null;
  if(!td||!td.dataset.p) return;
  var l=F('lupa'); l.innerHTML=svgGrande(td); l.style.display='block';
  colocarLupa(e);
}});
document.addEventListener('mousemove',function(e){{
  if(F('lupa').style.display==='block'){{
    var td=e.target.closest?e.target.closest('td.tz'):null;
    if(td&&td.dataset.p) colocarLupa(e); else F('lupa').style.display='none';
  }}
}});

/* ---------- generacion del .bat ---------- */
function comandos(porPaises){{
  var dest=F('dest').value.trim();
  var mover=F('mover').checked;
  var cbs=Array.prototype.slice.call(document.querySelectorAll('.chk-row:checked'));
  if(!dest){{ alert('Escribe la carpeta de destino.'); return null; }}
  if(cbs.length===0){{ alert('No has marcado ningun archivo.'); return null; }}
  var op=mover?'move':'copy';
  var L=['@echo off','chcp 65001 >nul','setlocal',
         'rem Carpeta raiz: primer parametro del .bat; si no, el valor de abajo.',
         'rem Si el archivo ya existe en destino NO se sobrescribe: se omite.',
         'set "RAIZ=%~1"',
         'if "%RAIZ%"=="" set "RAIZ='+dest+'"',
         'if not exist "%RAIZ%" mkdir "%RAIZ%"',''];
  var carpetas={{}};
  cbs.forEach(function(cb){{
    carpetas[porPaises?(cb.dataset.carpeta||'desconocido'):'']=1;
  }});
  Object.keys(carpetas).sort().forEach(function(p){{
    if(p) L.push('if not exist "%RAIZ%\\\\'+p+'" mkdir "%RAIZ%\\\\'+p+'"');
  }});
  L.push('');
  cbs.forEach(function(cb){{
    var sub=porPaises?('\\\\'+(cb.dataset.carpeta||'desconocido')):'';
    var destino='%RAIZ%'+sub+'\\\\'+cb.dataset.nom;
    L.push('if exist "'+destino+'" echo   [=] ya estaba, omitido: '+cb.dataset.nom);
    L.push('if not exist "'+destino+'" '+op+' /Y "'+cb.dataset.path+
           '" "%RAIZ%'+sub+'\\\\" >nul');
  }});
  L.push('','echo.','echo Procesados '+cbs.length+' archivo(s).','pause');
  return L.join('\\r\\n');
}}
function descargarBat(texto,nombre){{
  var blob=new Blob([texto],{{type:'text/plain;charset=utf-8'}});
  var a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=nombre;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}}
function descargar(){{
  var t=comandos(false); if(t===null) return;
  descargarBat(t, F('mover').checked?'rutas_mover.bat':'rutas_copiar.bat');
}}
function descargarPaises(){{
  var t=comandos(true); if(t===null) return;
  var p=F('prev'); p.textContent=t; p.style.display='block';
  descargarBat(t, F('mover').checked?'rutas_mover_paises.bat':'rutas_copiar_paises.bat');
}}
function copiar(){{
  var t=comandos(false); if(t===null) return;
  var ta=document.createElement('textarea'); ta.value=t;
  document.body.appendChild(ta); ta.select();
  try{{ document.execCommand('copy');
        alert('Comandos copiados. Pegalos en una ventana CMD y pulsa Enter.'); }}
  catch(e){{ alert('No se pudo copiar automaticamente.'); }}
  document.body.removeChild(ta);
}}
function previsualizar(){{
  var t=comandos(false); if(t===null) return;
  var p=F('prev'); p.textContent=t; p.style.display='block';
}}
actualizar();
</script>
</body></html>"""
    with open(ruta, "w", encoding="utf-8") as fh:
        fh.write(doc)


def escribir_movil(regs, ruta):
    """Página ligera para el móvil: solo rutas de Wikiloc, en fichas.

    Filtra por país y ordena por cercanía usando el GPS del dispositivo (o unas
    coordenadas pegadas a mano). El enlace «Cómo llegar» no lleva origen, así
    que Google Maps traza la ruta desde la posición actual del teléfono.
    """
    wk = [r for r in regs if r.get("es_wikiloc")]

    def ficha(r):
        titulo = r.get("nombre") or os.path.splitext(r.get("archivo", ""))[0]
        pais = r.get("pais") or "País desconocido"
        tipo = r.get("tipo") or ""
        sub = " · ".join(x for x in (pais, tipo) if x)
        cod = r.get("dup_cod") or ""
        chapa = ""
        if cod:
            texto_chapa = {"sosp": "Posible duplicado",
                           "rep": "Repetido",
                           "orig": "Tiene copias"}.get(cod, "")
            num = (r.get("dup") or "").split("#")[-1]
            chapa = (f'<span class="bg bg-{cod}">{texto_chapa}'
                     f'{" #" + num if num else ""}</span>')
        dat = [
            (_fmt_num(r.get("dist_km"), 2), "km"),
            (_fmt_num(r.get("desnivel_pos"), 0), "desnivel +"),
            (_fmt_horas(r.get("dur_h")), "duración"),
            (_fmt_num(r.get("alt_min"), 0), "alt. mín"),
            (_fmt_num(r.get("alt_max"), 0), "alt. máx"),
        ]
        celdas = "".join(
            f'<li><b>{_html.escape(v) if v else "—"}</b><span>{e}</span></li>'
            for v, e in dat)
        celdas += '<li class="cerca"><b>—</b><span>desde ti</span></li>'

        botones = ""
        if r.get("lat0") is not None:
            u = ("https://www.google.com/maps/dir/?api=1&destination="
                 f'{r["lat0"]:.6f},{r["lon0"]:.6f}')
            botones += (f'<a class="b b1" href="{_html.escape(u, quote=True)}" '
                        f'target="_blank" rel="noopener">Cómo llegar al inicio</a>')
        enl = r.get("enlace") or ""
        if enl.startswith("http"):
            botones += (f'<a class="b b2" href="{_html.escape(enl, quote=True)}" '
                        f'target="_blank" rel="noopener">Ver en Wikiloc</a>')

        lat = "" if r.get("lat0") is None else f'{r["lat0"]:.6f}'
        lon = "" if r.get("lon0") is None else f'{r["lon0"]:.6f}'
        return (
            f'<article class="c{" dup-" + cod if cod else ""}" '
            f'data-pais="{_html.escape(pais, quote=True)}" '
            f'data-lat="{lat}" data-lon="{lon}" data-dup="{cod}" '
            f'data-dist="{"" if r.get("dist_km") is None else round(r["dist_km"], 2)}" '
            f'data-desn="{"" if r.get("desnivel_pos") is None else round(r["desnivel_pos"])}" '
            f'data-nom="{_html.escape(titulo.lower(), quote=True)}">'
            f'<div class="hd">{svg_mini(r.get("traza"), "tz")}'
            f'<div class="tt"><h2>{_html.escape(titulo)}</h2>'
            f'<p>{_html.escape(sub)}</p>{chapa}</div></div>'
            f'<ul class="ds">{celdas}</ul>'
            f'<div class="bt">{botones}</div></article>'
        )

    paises = sorted({r.get("pais") or "" for r in wk} - {""})
    opt = "".join(f'<option value="{_html.escape(p, quote=True)}">{_html.escape(p)}</option>'
                  for p in paises)
    fichas = os.linesep.join(ficha(r) for r in wk) or \
        '<p class="vacio">No hay rutas de Wikiloc en esta carpeta.</p>'

    doc = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Mis rutas</title>
<style>
  :root {{ color-scheme: light dark; --bd:#d6d8dc; --sub:#6b7280; --bg:#fff;
          --az:#1a56db; }}
  * {{ box-sizing:border-box; -webkit-tap-highlight-color:transparent; }}
  body {{ font-family:system-ui,"Segoe UI",Roboto,Arial,sans-serif; margin:0;
         padding:0 .7rem 2.5rem; line-height:1.35; font-size:16px; }}
  h1 {{ font-size:1.15rem; margin:.8rem 0 .1rem; }}
  .sub {{ color:var(--sub); font-size:.8rem; margin:0 0 .6rem; }}
  .barra {{ position:sticky; top:0; z-index:5; background:var(--bg);
           padding:.5rem 0 .6rem; border-bottom:1px solid var(--bd); }}
  .barra .fila {{ display:flex; gap:.4rem; margin-bottom:.4rem; }}
  select, input, button {{ font:inherit; padding:.55rem .5rem; border-radius:9px;
          border:1px solid var(--bd); background:transparent; color:inherit;
          min-height:2.6rem; }}
  select {{ flex:1; }}
  button {{ cursor:pointer; }}
  button.gps {{ flex:1; background:var(--az); color:#fff; border-color:var(--az);
               font-weight:600; }}
  button.lim {{ width:6.5rem; }}
  #manual {{ display:none; gap:.4rem; }}
  #manual input {{ flex:1; min-width:0; }}
  #estado {{ font-size:.78rem; color:var(--sub); min-height:1.1rem; }}
  #estado.err {{ color:#b91c1c; }}
  .c {{ border:1px solid var(--bd); border-radius:14px; padding:.7rem .8rem;
       margin:.7rem 0; }}
  .hd {{ display:flex; gap:.7rem; align-items:center; }}
  .hd .tt {{ min-width:0; }}
  .c h2 {{ font-size:1rem; margin:0; overflow:hidden; text-overflow:ellipsis;
          white-space:nowrap; }}
  .c .hd p {{ margin:.1rem 0 0; font-size:.78rem; color:var(--sub); }}
  .bg {{ display:inline-block; margin-top:.25rem; padding:.1rem .45rem;
        border-radius:99px; font-size:.68rem; font-weight:700;
        letter-spacing:.02em; }}
  .bg-sosp {{ background:rgba(220,38,38,.14); color:#b91c1c; }}
  .bg-rep {{ background:rgba(0,0,0,.08); color:var(--sub); }}
  .bg-orig {{ background:rgba(22,163,74,.14); color:#15803d; }}
  .c.dup-sosp {{ border-color:#dc2626; }}
  svg.tz {{ width:4.4rem; height:2.9rem; flex:0 0 auto; }}
  svg polyline {{ fill:none; stroke:var(--az); stroke-width:1.7;
                 stroke-linejoin:round; stroke-linecap:round;
                 vector-effect:non-scaling-stroke; }}
  svg circle.ini {{ fill:#16a34a; }} svg circle.fin {{ fill:#dc2626; }}
  ul.ds {{ list-style:none; display:grid; grid-template-columns:repeat(3,1fr);
          gap:.45rem .3rem; margin:.7rem 0 .1rem; padding:0; }}
  ul.ds li {{ text-align:center; }}
  ul.ds b {{ display:block; font-size:1rem; font-variant-numeric:tabular-nums; }}
  ul.ds span {{ font-size:.7rem; color:var(--sub); }}
  li.cerca b {{ color:var(--az); }}
  .bt {{ display:flex; gap:.45rem; margin-top:.7rem; }}
  a.b {{ flex:1; text-align:center; padding:.6rem .4rem; border-radius:10px;
        text-decoration:none; font-size:.85rem; font-weight:600; }}
  a.b1 {{ background:var(--az); color:#fff; }}
  a.b2 {{ border:1px solid var(--az); color:var(--az); }}
  .vacio {{ color:var(--sub); }}
  @media (prefers-color-scheme: dark) {{
    :root {{ --bd:#3a3d42; --sub:#9aa0a6; --bg:#101214; --az:#7aa7ff; }}
    body {{ background:var(--bg); color:#e8eaed; }}
    a.b1 {{ color:#0b1220; }}
    .bg-sosp {{ background:rgba(220,38,38,.22); color:#fca5a5; }}
    .bg-orig {{ background:rgba(22,163,74,.22); color:#86efac; }}
    .bg-rep {{ background:rgba(255,255,255,.10); }}
  }}
</style></head><body>
<h1>Mis rutas</h1>
<p class="sub" id="cuenta">{len(wk)} rutas de Wikiloc</p>
<div class="barra">
  <div class="fila">
    <select id="fPais" onchange="pintar()">
      <option value="">Todos los países</option>{opt}</select>
    <select id="fOrd" onchange="pintar()">
      <option value="nom">Por nombre</option>
      <option value="cerca">Más cercanas</option>
      <option value="dist">Más largas</option>
      <option value="desn">Más desnivel</option>
    </select>
  </div>
  <div class="fila">
    <select id="fDup" onchange="pintar()">
      <option value="">Todas, con duplicados</option>
      <option value="no">Ocultar duplicados</option>
      <option value="sosp">Solo posibles duplicados</option>
    </select>
  </div>
  <div class="fila">
    <button class="gps" onclick="ubicar()">Usar mi posición</button>
    <button class="lim" onclick="manual()">A mano</button>
  </div>
  <div class="fila" id="manual">
    <input id="coord" inputmode="decimal" placeholder="40.4168, -3.7038 o enlace de Maps">
    <button onclick="usarManual()">Ir</button>
  </div>
  <div id="estado"></div>
</div>
<div id="lista">
{fichas}
</div>
<script>
var MI=null;
function F(i){{ return document.getElementById(i); }}
function fichas(){{ return Array.prototype.slice.call(
  document.querySelectorAll('#lista .c')); }}
function estado(t,err){{ var e=F('estado'); e.textContent=t||'';
  e.className=err?'err':''; }}

function hav(la1,lo1,la2,lo2){{
  var R=6371, t=Math.PI/180;
  var dla=(la2-la1)*t, dlo=(lo2-lo1)*t;
  var a=Math.sin(dla/2)*Math.sin(dla/2)+
        Math.cos(la1*t)*Math.cos(la2*t)*Math.sin(dlo/2)*Math.sin(dlo/2);
  return 2*R*Math.asin(Math.sqrt(a));
}}
function fmt(n){{
  if(n===null) return '—';
  if(n<10) return n.toFixed(1).replace('.',',')+' km';
  return Math.round(n)+' km';
}}
function fijar(la,lo,txt){{
  MI={{la:la,lo:lo}};
  fichas().forEach(function(c){{
    var b=c.querySelector('li.cerca b');
    if(c.dataset.lat===''){{ b.textContent='—'; c.dataset.km=''; return; }}
    var d=hav(la,lo,+c.dataset.lat,+c.dataset.lon);
    c.dataset.km=d; b.textContent=fmt(d);
  }});
  F('fOrd').value='cerca';
  estado('Distancias en línea recta desde '+txt+'.');
  pintar();
}}
function ubicar(){{
  if(!navigator.geolocation){{
    estado('Este navegador no puede darme la posición.',true); return; }}
  estado('Buscando tu posición…');
  navigator.geolocation.getCurrentPosition(
    function(p){{ fijar(p.coords.latitude,p.coords.longitude,'tu posición'); }},
    function(e){{
      estado('No he podido leer el GPS ('+e.message+'). Si has abierto el fichero '+
             'con doble clic, el navegador bloquea la ubicación: pulsa «A mano» '+
             'y pega unas coordenadas.',true);
      F('manual').style.display='flex';
    }},
    {{enableHighAccuracy:true, timeout:12000, maximumAge:120000}});
}}
function manual(){{
  var m=F('manual');
  m.style.display = m.style.display==='flex' ? 'none' : 'flex';
  if(m.style.display==='flex') F('coord').focus();
}}
function usarManual(){{
  var v=F('coord').value.trim();
  var m=v.match(/(-?\\d{{1,3}}[.,]\\d+)[^\\d-]+(-?\\d{{1,3}}[.,]\\d+)/);
  if(!m){{ estado('No reconozco esas coordenadas. Prueba con 40.4168, -3.7038',true);
           return; }}
  fijar(parseFloat(m[1].replace(',','.')), parseFloat(m[2].replace(',','.')),
        'el punto indicado');
}}
function pintar(){{
  var pais=F('fPais').value, ord=F('fOrd').value, dup=F('fDup').value;
  var cs=fichas(), vis=0;
  cs.forEach(function(c){{
    var ok=!pais||c.dataset.pais===pais;
    if(ok&&dup==='no' && (c.dataset.dup==='sosp'||c.dataset.dup==='rep')) ok=false;
    if(ok&&dup==='sosp' && c.dataset.dup!=='sosp') ok=false;
    c.style.display=ok?'':'none';
    if(ok) vis++;
  }});
  if(ord==='cerca'&&!MI){{ estado('Pulsa «Usar mi posición» para ordenar por '+
                                 'cercanía.',true); }}
  var val=function(c){{
    if(ord==='cerca') return c.dataset.km===''||c.dataset.km===undefined
                             ? Infinity : +c.dataset.km;
    if(ord==='dist') return c.dataset.dist===''?-1:-c.dataset.dist;
    if(ord==='desn') return c.dataset.desn===''?-1:-c.dataset.desn;
    return null;
  }};
  cs.sort(function(a,b){{
    var x=val(a), y=val(b);
    if(x===null) return a.dataset.nom<b.dataset.nom?-1:
                        a.dataset.nom>b.dataset.nom?1:0;
    return x-y;
  }});
  var L=F('lista');
  cs.forEach(function(c){{ L.appendChild(c); }});
  F('cuenta').textContent=vis+' ruta'+(vis===1?'':'s')+
    (pais?' en '+pais:' de Wikiloc')+
    (dup==='sosp'?' · posibles duplicados':dup==='no'?' · sin duplicados':'');
}}
pintar();
</script>
</body></html>"""
    with open(ruta, "w", encoding="utf-8") as fh:
        fh.write(doc)


def escribir_xlsx(regs, ruta):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wikiloc"

    cab_fill = PatternFill("solid", fgColor="1A56DB")
    cab_font = Font(bold=True, color="FFFFFF")
    ws.append(CABECERA)
    for c in ws[1]:
        c.fill = cab_fill
        c.font = cab_font
        c.alignment = Alignment(vertical="center")

    for r in regs:
        valores = []
        for key, _cab, tipo, dec in COLS:
            v = r.get(key)
            if tipo == "num":
                valores.append(None if v is None else
                               (round(v) if dec == 0 else round(v, dec)))
            elif tipo == "dt":
                valores.append(v)  # ya es datetime naive (UTC) o None
            else:
                valores.append("" if v is None else v)
        ws.append(valores)
        fila = ws.max_row
        for i, (_key, _cab, tipo, dec) in enumerate(COLS, start=1):
            celda = ws.cell(row=fila, column=i)
            if tipo == "num" and celda.value is not None:
                celda.number_format = "0" if dec == 0 else "0.00"
            elif tipo == "dt" and celda.value is not None:
                celda.number_format = "yyyy-mm-dd hh:mm"
            elif tipo == "url" and str(celda.value).startswith("http"):
                celda.hyperlink = celda.value
                celda.font = Font(color="1A56DB", underline="single")
            elif tipo == "geo" and r.get("lat0") is not None:
                celda.hyperlink = ("https://www.google.com/maps/dir/?api=1&destination="
                                   f'{r["lat0"]:.6f},{r["lon0"]:.6f}')
                celda.font = Font(color="1A56DB", underline="single")

    # Anchos por contenido (topes por columna)
    topes = {"archivo": 42, "ruta_rel": 22, "wikiloc": 9, "origen": 18, "dup": 16,
             "carpeta": 18, "autor": 26, "enlace": 52, "pais": 24, "tipo": 10,
             "coord_inicio": 20}
    for i, (key, cab, _t, _d) in enumerate(COLS, start=1):
        ancho = max([len(cab)] + [len(disp(r, COLS[i - 1])) for r in regs] + [6]) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(ancho, topes.get(key, 16))

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{ws.max_row}"
    wb.save(ruta)


# --------------------------------------------------------------------------- #
#  Recorrido y orquestación
# --------------------------------------------------------------------------- #
def recolectar_gpx(directorio, recursivo):
    if recursivo:
        for base, _, files in os.walk(directorio):
            for f in sorted(files):
                if f.lower().endswith(".gpx"):
                    yield os.path.join(base, f)
    else:
        for f in sorted(os.listdir(directorio)):
            r = os.path.join(directorio, f)
            if os.path.isfile(r) and f.lower().endswith(".gpx"):
                yield r


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Extrae datos y métricas de ficheros GPX de Wikiloc.")
    ap.add_argument("directorio", nargs="?", default=".",
                    help="Directorio con los .gpx (por defecto: el actual).")
    ap.add_argument("-o", "--salida", default="",
                    help="Nombre base de salida SIN extensión. Si es solo un nombre, "
                         "los ficheros se crean DENTRO de la carpeta explorada. "
                         "Si incluye una ruta (p. ej. C:\\otra\\salida\\rutas), se usa "
                         "esa ruta. Vacío = nombre de la carpeta explorada.")
    ap.add_argument("-f", "--formato", default="txt",
                    help="Formatos separados por coma: txt,csv,xlsx,html,movil "
                         "o 'all'. 'movil' genera <salida>_movil.html.")
    ap.add_argument("-r", "--recursivo", action="store_true",
                    help="Buscar también en subdirectorios.")
    ap.add_argument("--enlace", choices=["ruta", "autor"], default="ruta",
                    help="Enlace de la columna 'Enlace wikiloc'.")
    ap.add_argument("--umbral", type=float, default=4.0,
                    help="Umbral en metros para el desnivel acumulado (def.: 4).")
    ap.add_argument("--solo-wikiloc", action="store_true",
                    help="Descartar los .gpx que no sean de Wikiloc en lugar de "
                         "incluirlos con la etiqueta Wikiloc = No.")
    args = ap.parse_args(argv)

    if not os.path.isdir(args.directorio):
        print(f"ERROR: '{args.directorio}' no es un directorio.", file=sys.stderr)
        return 2

    formatos = ["txt", "csv", "xlsx", "html", "movil"] if args.formato.lower() == "all" \
        else [x.strip().lower() for x in args.formato.split(",") if x.strip()]
    validos = {"txt", "csv", "xlsx", "html", "movil"}
    desconocidos = [x for x in formatos if x not in validos]
    if desconocidos:
        print(f"ERROR: formato(s) no válido(s): {', '.join(desconocidos)}.",
              file=sys.stderr)
        return 2

    # Dónde y con qué nombre se escribe la salida:
    #  - Si --salida trae una ruta (o es absoluta) -> se respeta tal cual.
    #  - Si es solo un nombre -> se crea DENTRO de la carpeta explorada.
    #  - Si está vacío -> se usa el nombre de la carpeta explorada.
    salida = args.salida.strip()
    if salida and (os.path.dirname(salida) or os.path.isabs(salida)):
        base = os.path.splitext(salida)[0]
    else:
        nombre = os.path.splitext(salida)[0] if salida else ""
        if not nombre:
            nombre = os.path.basename(os.path.normpath(
                os.path.abspath(args.directorio))) or "wikiloc_tracks"
        base = os.path.join(args.directorio, nombre)
    motor = "reverse_geocoder (preciso)" if _RG_DISPONIBLE \
        else "tabla embebida (aproximada; instala 'reverse_geocoder' para más precisión)"
    print(f"Motor de país: {motor}", file=sys.stderr)

    raiz_explorada = os.path.abspath(args.directorio)
    regs, n_total, n_ignorados, n_errores = [], 0, 0, 0
    for ruta in recolectar_gpx(args.directorio, args.recursivo):
        n_total += 1
        r = analizar_gpx(ruta, args.umbral)
        if "error" in r:
            n_errores += 1
            print(f"  · Error en {os.path.basename(ruta)}: {r['error']}",
                  file=sys.stderr)
            continue
        if not r["es_wikiloc"] and args.solo_wikiloc:
            n_ignorados += 1
            print(f"  · Ignorado (no es de Wikiloc): {os.path.basename(ruta)}",
                  file=sys.stderr)
            continue
        enlace = r["enlace_autor"] if args.enlace == "autor" else r["enlace_ruta"]
        if not enlace:
            enlace = (r["enlace_ruta"] or r["enlace_autor"] or r["enlace_ext"]
                      or "(sin enlace)")
        r["enlace"] = enlace
        r["ruta_abs"] = os.path.abspath(ruta)
        rel = os.path.relpath(os.path.dirname(r["ruta_abs"]), raiz_explorada)
        r["ruta_rel"] = "(raíz)" if rel == "." else rel
        regs.append(r)

    n_grupos = marcar_duplicados(regs)
    for r in regs:
        r["carpeta"] = carpeta_destino(r)

    generados = []
    for fmt in formatos:
        destino = f"{base}_movil.html" if fmt == "movil" else f"{base}.{fmt}"
        if fmt == "txt":
            escribir_txt(regs, destino)
        elif fmt == "csv":
            escribir_csv(regs, destino)
        elif fmt == "html":
            escribir_html(regs, destino)
        elif fmt == "movil":
            escribir_movil(regs, destino)
        elif fmt == "xlsx":
            if _XLSX_DISPONIBLE:
                escribir_xlsx(regs, destino)
            else:
                destino = f"{base}.csv"
                escribir_csv(regs, destino)
                print("  · 'openpyxl' no está instalado; genero CSV en su lugar "
                      "(instálalo con: pip install openpyxl).", file=sys.stderr)
        generados.append(destino)

    n_wk = sum(1 for r in regs if r.get("es_wikiloc"))
    n_sosp = sum(1 for r in regs if r.get("dup_cod") == "sosp")
    n_rep = sum(1 for r in regs if r.get("dup_cod") == "rep")
    print(f"\nProcesados {n_total} .gpx | Wikiloc: {n_wk} | "
          f"No Wikiloc: {len(regs) - n_wk} | Descartados: {n_ignorados} | "
          f"Errores: {n_errores}", file=sys.stderr)
    if n_grupos:
        print(f"Duplicados: {n_grupos} grupo(s) | {n_sosp} sospechoso(s) con otro "
              f"nombre | {n_rep} con el mismo nombre", file=sys.stderr)
    print("Generado(s): " + ", ".join(dict.fromkeys(generados)), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
