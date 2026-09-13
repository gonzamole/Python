#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wikiloc_gpx.py
==============
Recorre un directorio buscando ficheros .gpx descargados de Wikiloc y genera
una tabla con una fila por track y estas columnas:

    Nombre del archivo | Autor | Enlace wikiloc | País |
    Dist (km) | Alt mín (m) | Alt máx (m) | Desnivel + (m) |
    Inicio (UTC) | Fin (UTC) | Duración (h)

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
    html  -> página web con tabla ordenable (orden numérico en columnas de números)

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
import sys
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
    ("autor",        "Autor",              "txt", None),
    ("enlace",       "Enlace wikiloc",     "url", None),
    ("pais",         "País",               "txt", None),
    ("dist_km",      "Dist (km)",          "num", 2),
    ("alt_min",      "Alt mín (m)",        "num", 0),
    ("alt_max",      "Alt máx (m)",        "num", 0),
    ("desnivel_pos", "Desnivel + (m)",     "num", 0),
    ("inicio",       "Inicio (UTC)",       "dt",  None),
    ("fin",          "Fin (UTC)",          "dt",  None),
    ("dur_h",        "Duración (h)",       "num", 2),
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
    if not es_wikiloc:
        return None

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
    enlace_ruta, enlace_autor = "", ""
    for h in hrefs:
        hl = h.lower()
        if "wikiloc.com" not in hl:
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

    # Nombre y descripción
    nombre_track = ""
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
        "autor": autor or "(desconocido)",
        "enlace_ruta": enlace_ruta,
        "enlace_autor": enlace_autor,
        "pais": pais,
        "dist_km": dist_km,
        "alt_min": alt_min,
        "alt_max": alt_max,
        "desnivel_pos": desn_pos,
        "inicio": inicio,
        "fin": fin,
        "dur_h": dur_h,
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


def escribir_html(regs, ruta):
    def td(rec, col):
        key, _c, tipo, _d = col
        texto = disp(rec, col)
        if tipo == "url" and texto.startswith("http"):
            return (f'<td><a href="{_html.escape(texto)}" target="_blank" '
                    f'rel="noopener">{_html.escape(texto)}</a></td>')
        clase = ' class="n"' if tipo in ("num", "dt") else ""
        return f"<td{clase}>{_html.escape(texto)}</td>"

    def fila(r):
        cb = (f'<td class="chk"><input type="checkbox" class="chk-row" '
              f'data-path="{_html.escape(r.get("ruta_abs", ""), quote=True)}" '
              f'onchange="actualizar()"></td>')
        return "<tr>" + cb + "".join(td(r, c) for c in COLS) + "</tr>"

    trs = [fila(r) for r in regs]
    ths = ('<th class="chk"><input type="checkbox" id="all" '
           'onclick="toggleTodos(this)" title="Marcar todo"></th>')
    ths += "".join(
        f'<th onclick="ordenar({i + 1})">{_html.escape(c[1])} '
        f'<span class="a">&#8597;</span></th>' for i, c in enumerate(COLS)
    )
    doc = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tracks de Wikiloc</title>
<style>
  :root {{ color-scheme: light dark; }}
  body {{ font-family: system-ui, "Segoe UI", Roboto, Arial, sans-serif;
         margin: 1.5rem; line-height: 1.4; }}
  h1 {{ font-size: 1.35rem; margin: 0 0 .25rem; }}
  p.info {{ color:#666; margin:0 0 1rem; font-size:.9rem; }}
  table {{ border-collapse: collapse; width: 100%; font-size:.9rem; }}
  th, td {{ border:1px solid #d0d0d0; padding:.45rem .55rem; text-align:left;
           vertical-align:top; }}
  th {{ background:#f3f4f6; cursor:pointer; user-select:none;
       position:sticky; top:0; white-space:nowrap; }}
  th .a {{ color:#aaa; font-size:.8em; }}
  td.n {{ text-align:right; white-space:nowrap; font-variant-numeric:tabular-nums; }}
  tbody tr:nth-child(even) {{ background:rgba(0,0,0,.03); }}
  td:nth-child(4) {{ word-break:break-all; max-width:360px; }}
  a {{ color:#1a56db; text-decoration:none; }} a:hover {{ text-decoration:underline; }}
  th.chk {{ cursor:default; width:2.2rem; text-align:center; }}
  td.chk {{ text-align:center; }}
  .tb {{ display:flex; flex-wrap:wrap; gap:.5rem .8rem; align-items:center;
        margin:0 0 .8rem; padding:.6rem .7rem; border:1px solid #d0d0d0;
        border-radius:8px; background:rgba(0,0,0,.02); font-size:.9rem; }}
  .tb input[type=text] {{ padding:.25rem .4rem; }}
  .tb button {{ padding:.3rem .7rem; cursor:pointer; }}
  #cuenta {{ color:#666; }}
  .prev {{ display:none; white-space:pre; overflow:auto; max-height:260px;
          margin:0 0 1rem; padding:.6rem .7rem; border:1px solid #d0d0d0;
          border-radius:8px; background:#0d1117; color:#e6edf3;
          font-family:Consolas,"Courier New",monospace; font-size:.82rem; }}
  @media (prefers-color-scheme: dark) {{
    th {{ background:#222; }} th, td {{ border-color:#444; }}
    p.info {{ color:#999; }} a {{ color:#7aa7ff; }}
  }}
</style></head><body>
<h1>Tracks de Wikiloc</h1>
<p class="info">{len(regs)} track(s). Pulsa una cabecera para ordenar. Marca casillas,
 escribe la carpeta destino y genera un .bat para copiar o mover esos archivos.
 Horas en UTC. &laquo;(aprox.)&raquo; = pais estimado sin <code>reverse_geocoder</code>.</p>
<div class="tb">
  <label>Carpeta destino:
    <input type="text" id="dest" size="42" placeholder="C:\\rutas\\seleccion"></label>
  <label><input type="radio" name="op" id="copiar" checked> Copiar</label>
  <label><input type="radio" name="op" id="mover"> Mover</label>
  <button onclick="descargar()">Descargar .bat</button>
  <button onclick="copiar()">Copiar comandos</button>
  <button onclick="previsualizar()">Previsualizar</button>
  <span id="cuenta">0 marcados</span>
</div>
<pre id="prev" class="prev"></pre>
<table id="t"><thead><tr>{ths}</tr></thead><tbody>
{os.linesep.join(trs)}
</tbody></table>
<script>
function num(s){{ s=s.trim(); if(s==='') return null;
  var n=Number(s.replace(/\\s/g,'').replace(',','.')); return isNaN(n)?null:n; }}
function ordenar(col){{
  var tb=document.querySelector('#t tbody');
  var filas=Array.from(tb.rows);
  var asc=tb.getAttribute('data-col')!=col||tb.getAttribute('data-dir')!='asc';
  filas.sort(function(a,b){{
    var x=a.cells[col].innerText, y=b.cells[col].innerText;
    var nx=num(x), ny=num(y), cmp;
    if(nx!==null&&ny!==null) cmp=nx-ny;
    else cmp=(x.toLowerCase()<y.toLowerCase()?-1:x.toLowerCase()>y.toLowerCase()?1:0);
    return cmp*(asc?1:-1);
  }});
  filas.forEach(function(f){{ tb.appendChild(f); }});
  tb.setAttribute('data-col',col); tb.setAttribute('data-dir',asc?'asc':'desc');
}}
function actualizar(){{
  var n=document.querySelectorAll('.chk-row:checked').length;
  document.getElementById('cuenta').textContent=n+' marcados';
}}
function toggleTodos(cb){{
  var todos=document.querySelectorAll('.chk-row');
  for(var i=0;i<todos.length;i++) todos[i].checked=cb.checked;
  actualizar();
}}
function comandos(){{
  var dest=document.getElementById('dest').value.trim();
  var mover=document.getElementById('mover').checked;
  var cbs=Array.prototype.slice.call(document.querySelectorAll('.chk-row:checked'));
  if(!dest){{ alert('Escribe la carpeta de destino.'); return null; }}
  if(cbs.length===0){{ alert('No has marcado ningun archivo.'); return null; }}
  var op=mover?'move':'copy';
  var L=['@echo off','chcp 65001 >nul','set "DEST='+dest+'"',
         'if not exist "%DEST%" mkdir "%DEST%"'];
  cbs.forEach(function(cb){{
    L.push(op+' /Y "'+cb.dataset.path+'" "%DEST%\\\\"');
  }});
  L.push('echo.','echo Terminado: '+cbs.length+' archivo(s).','pause');
  return L.join('\\r\\n');
}}
function descargar(){{
  var t=comandos(); if(t===null) return;
  var mover=document.getElementById('mover').checked;
  var blob=new Blob([t],{{type:'text/plain;charset=utf-8'}});
  var a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=mover?'wikiloc_mover.bat':'wikiloc_copiar.bat';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}}
function copiar(){{
  var t=comandos(); if(t===null) return;
  var ta=document.createElement('textarea'); ta.value=t;
  document.body.appendChild(ta); ta.select();
  try{{ document.execCommand('copy');
        alert('Comandos copiados. Pegalos en una ventana CMD y pulsa Enter.'); }}
  catch(e){{ alert('No se pudo copiar automaticamente.'); }}
  document.body.removeChild(ta);
}}
function previsualizar(){{
  var t=comandos(); if(t===null) return;
  var p=document.getElementById('prev'); p.textContent=t; p.style.display='block';
}}
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

    # Anchos por contenido (topes por columna)
    topes = {"archivo": 42, "autor": 26, "enlace": 64, "pais": 24}
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
                    help="Formatos separados por coma: txt,csv,xlsx,html  o  'all'.")
    ap.add_argument("-r", "--recursivo", action="store_true",
                    help="Buscar también en subdirectorios.")
    ap.add_argument("--enlace", choices=["ruta", "autor"], default="ruta",
                    help="Enlace de la columna 'Enlace wikiloc'.")
    ap.add_argument("--umbral", type=float, default=4.0,
                    help="Umbral en metros para el desnivel acumulado (def.: 4).")
    args = ap.parse_args(argv)

    if not os.path.isdir(args.directorio):
        print(f"ERROR: '{args.directorio}' no es un directorio.", file=sys.stderr)
        return 2

    formatos = ["txt", "csv", "xlsx", "html"] if args.formato.lower() == "all" \
        else [x.strip().lower() for x in args.formato.split(",") if x.strip()]
    validos = {"txt", "csv", "xlsx", "html"}
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

    regs, n_total, n_ignorados, n_errores = [], 0, 0, 0
    for ruta in recolectar_gpx(args.directorio, args.recursivo):
        n_total += 1
        r = analizar_gpx(ruta, args.umbral)
        if r is None:
            n_ignorados += 1
            print(f"  · Ignorado (no es de Wikiloc): {os.path.basename(ruta)}",
                  file=sys.stderr)
            continue
        if "error" in r:
            n_errores += 1
            print(f"  · Error en {os.path.basename(ruta)}: {r['error']}",
                  file=sys.stderr)
            continue
        enlace = r["enlace_autor"] if args.enlace == "autor" else r["enlace_ruta"]
        if not enlace:
            enlace = r["enlace_ruta"] or r["enlace_autor"] or "(sin enlace)"
        r["enlace"] = enlace
        r["ruta_abs"] = os.path.abspath(ruta)
        regs.append(r)

    generados = []
    for fmt in formatos:
        destino = f"{base}.{fmt}"
        if fmt == "txt":
            escribir_txt(regs, destino)
        elif fmt == "csv":
            escribir_csv(regs, destino)
        elif fmt == "html":
            escribir_html(regs, destino)
        elif fmt == "xlsx":
            if _XLSX_DISPONIBLE:
                escribir_xlsx(regs, destino)
            else:
                destino = f"{base}.csv"
                escribir_csv(regs, destino)
                print("  · 'openpyxl' no está instalado; genero CSV en su lugar "
                      "(instálalo con: pip install openpyxl).", file=sys.stderr)
        generados.append(destino)

    print(f"\nProcesados {n_total} .gpx | Wikiloc: {len(regs)} | "
          f"Ignorados: {n_ignorados} | Errores: {n_errores}", file=sys.stderr)
    print("Generado(s): " + ", ".join(dict.fromkeys(generados)), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
