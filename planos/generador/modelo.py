# -*- coding: utf-8 -*-
"""Geometría de los planos de la panadería industrial (nave nueva, 2.139,75 m²).
Coordenadas en metros. Origen: esquina interior inferior izquierda de la nave.
Hojas: 'dist' (planta baja), 'alt' (altillo, coordenadas locales), 'emp' (emplazamiento)."""
import math

K = {}  # hoja -> lista de entidades
def _add(sheet, e):
    K.setdefault(sheet, []).append(e); return e

def pl(sheet, layer, pts, closed=False, color='#000000', lw=0.25, fill=None, ls='-', z=2, alpha=1.0):
    return _add(sheet, dict(k='pl', layer=layer, pts=[tuple(p) for p in pts], closed=closed,
                             color=color, lw=lw, fill=fill, ls=ls, z=z, alpha=alpha))
def rect(sheet, layer, x0, y0, x1, y1, **kw):
    return pl(sheet, layer, [(x0, y0), (x1, y0), (x1, y1), (x0, y1)], closed=True, **kw)
def circ(sheet, layer, c, r, color='#000000', lw=0.25, fill=None, z=3):
    return _add(sheet, dict(k='c', layer=layer, c=tuple(c), r=r, color=color, lw=lw, fill=fill, z=z))
def txt(sheet, layer, p, s, pt=6, color='#000000', ha='center', va='center', bold=False, rot=0, z=6, bg=False):
    return _add(sheet, dict(k='t', layer=layer, p=tuple(p), s=s, pt=pt, color=color, ha=ha, va=va,
                             bold=bold, rot=rot, z=z, bg=bg))
def dim(sheet, p1, p2, off, color='#a01818', pt=6, txt_=None):
    """Cota lineal horizontal o vertical. off = desplazamiento (m) perpendicular."""
    return _add(sheet, dict(k='dim', layer='COTAS', p1=tuple(p1), p2=tuple(p2), off=off, color=color, pt=pt,
                             s=txt_))

def fmt(v, d=2):
    s = f"{v:,.{d}f}"
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')

def plen(pts):
    return sum(math.dist(pts[i], pts[i + 1]) for i in range(len(pts) - 1))

# ---------------------------------------------------------------- colores
C = dict(CGBT='#c62828', PRO='#1f5fbf', HOR='#e8681a', FRI='#11868a', ALU='#c99a06', ALT='#7b2d8e',
         EXT='#2e8b3a', BANDEJA='#f39c12', EMER='#2e7d32', COTA='#a01818', MURO='#000000')
CUADRO_COLOR = {'CGBT': C['CGBT'], 'CS-PRO': C['PRO'], 'CS-HOR': C['HOR'], 'CS-FRI': C['FRI'],
                'CS-ALU': C['ALU'], 'CS-ALT': C['ALT'], 'CS-EXT': C['EXT']}

# ---------------------------------------------------------------- nave
NAVE = [(0, 0), (53.5, 0), (53.5, 64.5), (34.5, 64.5), (34.5, 26.5), (0, 26.5)]
NAVE_EXT = [(-0.2, -0.2), (53.7, -0.2), (53.7, 64.7), (34.3, 64.7), (34.3, 26.7), (-0.2, 26.7)]
def area(pts):
    return abs(sum(pts[i][0] * pts[(i + 1) % len(pts)][1] - pts[(i + 1) % len(pts)][0] * pts[i][1]
                   for i in range(len(pts)))) / 2
A_NAVE = area(NAVE)
ALT_W, ALT_H = 20.0, 15.0           # altillo 300 m²
ALT_ORIG = (0.0, 11.5)              # esquina del altillo en coordenadas de nave

ZONAS = [  # nombre, (x0,y0,x1,y1), color
    ('SALA TÉCNICA', (0, 0, 12, 11.5), '#e4eaf1'),
    ('ALMACÉN\nAUXILIAR', (12, 0, 21.5, 11.5), '#f0efea'),
    ('RECEPCIÓN Y\nMATERIAS PRIMAS', (21.5, 0, 34.5, 11.5), '#f6efe1'),
    ('EXPEDICIÓN', (34.5, 0, 43.5, 14), '#e6f0f7'),
    ('CÁMARAS', (43.5, 0, 53.5, 11), '#d9e7f5'),
    ('ANTECÁMARA', (43.5, 11, 53.5, 14), '#f1f1f1'),
    ('BAJO ALTILLO', (0, 11.5, 20, 26.5), '#ebebeb'),
    ('AMASADO', (20, 11.5, 34.5, 26.5), '#fcefcf'),
    ('ENVASADO Y\nPAN RALLADO', (34.5, 14, 53.5, 26.5), '#e2f0df'),
    ('FERMENTACIÓN', (34.5, 26.5, 44, 45), '#fbe5cf'),
    ('ENFRIAMIENTO', (44, 26.5, 53.5, 45), '#e1ebf5'),
    ('HORNOS', (34.5, 45, 53.5, 64.5), '#f9dccf'),
]
# posición de la etiqueta (si no, centro)
ZONA_LBL = {'CÁMARAS': (46.1, 4.8), 'ANTECÁMARA': (46.2, 12.5), 'EXPEDICIÓN': (37.6, 8.6),
            'AMASADO': (27.0, 18.0), 'ENVASADO Y\nPAN RALLADO': (44.5, 18.6), 'HORNOS': (44.0, 55.0),
            'FERMENTACIÓN': (40.5, 41.5), 'ENFRIAMIENTO': (48.75, 41.5), 'BAJO ALTILLO': (7.0, 18.5),
            'SALA TÉCNICA': (6.0, 7.5), 'ALMACÉN\nAUXILIAR': (16.75, 7.0), 'RECEPCIÓN Y\nMATERIAS PRIMAS': (28.0, 6.5)}

# Puertas (rectángulo sobre el muro) y su rótulo
PUERTAS = [  # (x0,y0,x1,y1), rotulo, pos rotulo
    ((29.5, -0.2, 33.5, 0.0), 'Recepción', (31.5, -1.6)),
    ((37.0, -0.2, 41.0, 0.0), 'Expedición', (39.0, -1.6)),
    ((15.5, 26.5, 17.0, 26.7), 'Acceso personal', (21.8, 27.6)),
    ((34.3, 45.5, 34.5, 47.0), 'Salida', (32.6, 46.25)),
    ((53.5, 12.0, 53.7, 13.5), 'Salida', (55.4, 12.75)),
    ((53.5, 48.0, 53.7, 49.5), 'Salida', (55.4, 48.75)),
]
VENTANAS = [  # muro, desde, hasta
    ('S', 5.5, 8.5), ('S', 22.0, 25.0), ('S', 44.5, 47.5),
    ('N', 5.5, 8.5), ('N', 22.0, 25.0),
    ('W', 3.5, 10.0), ('W', 15.5, 22.0),
    ('E', 3.5, 10.0), ('E', 14.5, 21.0), ('E', 36.5, 43.0), ('E', 53.5, 60.0),
    ('WW', 29.5, 36.0), ('WW', 55.5, 62.0),
]
def ventana_rect(w, a, b):
    if w == 'S': return (a, -0.2, b, 0.0)
    if w == 'N': return (a, 26.5, b, 26.7)
    if w == 'W': return (-0.2, a, 0.0, b)
    if w == 'E': return (53.5, a, 53.7, b)
    if w == 'WW': return (34.3, a, 34.5, b)

# ---------------------------------------------------------------- receptores
# nº, descripción, modelo, uds, P total kW (Receptores_panaderia_v2.xlsx), cuadro
RECEPTORES = [
    (1, 'Horno de convección', 'Unox XEFT-06EU-EMRV', '4', 41.20, 'CS-HOR'),
    (2, 'Amasadora espiral 130 kg', 'Bongard Spiral EVO 200', '1', 6.80, 'CS-PRO'),
    (3, 'Amasadora espiral 80 kg', 'Bongard Spiral EVO 110', '1', 4.55, 'CS-PRO'),
    (4, 'Batidora planetaria 60 L', 'Bongard Saturne EVO 60', '1', 3.10, 'CS-PRO'),
    (5, 'Cámara de fermentación', 'Sveba Dahlen F500 HIK', '2', 8.00, 'CS-HOR'),
    (6, 'Congelación −20 °C, 32 m³', 'Intarcon BCV-NG-3 096', '2', 4.78, 'CS-FRI'),
    (7, 'Conservación 0 °C, 51 m³', 'Intarcon MCV-NY-3 108', '1', 2.48, 'CS-FRI'),
    (8, 'Conservación 0 °C, 40 m³', 'Intarcon MCV-NY-3 074', '1', 1.89, 'CS-FRI'),
    (9, 'Ensacadora', 'Por definir', '1', 5.50, 'CS-PRO'),
    (10, 'Trituradora de pan', 'Por definir', '1', 6.00, 'CS-PRO'),
    (11, 'Compresor de tornillo', 'Atlas Copco GA 5', '1', 6.14, 'CGBT'),
    (12, 'Cassette climatización', 'Mitsubishi PLA-RP140BA', '3', 15.63, 'CS-FRI/ALT'),
    (13, 'Alumbrado producción', 'Philips WT120C LED40S/840', '328*', 39.69, 'CS-ALU'),
    (14, 'Alumbrado altillo', 'Philips DN065B LED20S/830', '99', 1.88, 'CS-ALT'),
    (15, 'Alumbrado exterior', 'Philips BRP102 LED74/740', '13', 0.61, 'CS-EXT'),
    (16, 'Alumbrado emergencia', 'Legrand URA21 660611PL', '60', 0.14, 'CS-ALU/ALT'),
    (17, 'Línea de tomas 16 A', 'Schneider PKF16G423', '3', 11.04, 'CS-PRO/HOR'),
    (18, 'Puerta de muelle', 'Hörmann WA300R S4/SE', '2', 0.50, 'CS-FRI'),
    (19, 'Punto de recarga 32 A', 'Schneider EVB3S07N4A', '4', 29.60, 'CS-EXT'),
]
P_TOTAL = sum(r[4] for r in RECEPTORES)

# Máquinas en planta: nº, rectángulo, posición del marcador, discontinua (equipo en techo)
MAQ = [
    (1, (37.05, 62.7, 37.95, 63.7), (37.5, 61.0), False),
    (1, (41.55, 62.7, 42.45, 63.7), (42.0, 61.0), False),
    (1, (46.05, 62.7, 46.95, 63.7), (46.5, 61.0), False),
    (1, (50.55, 62.7, 51.45, 63.7), (51.0, 61.0), False),
    (2, (23.3, 21.0, 24.7, 21.9), (24.0, 23.5), False),
    (3, (26.9, 21.0, 28.1, 21.8), (27.5, 23.5), False),
    (4, (30.6, 21.0, 31.4, 21.9), (31.0, 23.5), False),
    (5, (37.0, 29.0, 38.0, 31.0), (39.6, 30.0), False),
    (5, (37.0, 35.0, 38.0, 37.0), (39.6, 36.0), False),
    (6, (44.0, 0.5, 47.6, 4.1), (45.8, 2.3), False),
    (6, (49.4, 0.5, 53.0, 4.1), (51.2, 2.3), False),
    (7, (44.0, 5.5, 48.5, 10.0), (46.25, 7.75), False),
    (8, (49.0, 6.0, 53.0, 10.0), (51.0, 8.0), False),
    (9, (46.75, 20.5, 48.25, 21.5), (49.75, 21.0), False),
    (10, (39.4, 20.5, 40.6, 21.5), (42.1, 21.0), False),
    (11, (2.0, 2.0, 3.2, 2.8), (2.6, 4.4), False),
    (12, (36.5, 16.5, 37.45, 17.45), (38.9, 17.0), True),
    (12, (51.5, 16.5, 52.45, 17.45), (50.1, 17.0), True),
    (17, (28.4, 25.0, 29.6, 26.2), (26.9, 25.6), False),
    (17, (52.0, 25.0, 53.2, 26.2), (50.5, 25.6), False),
    (17, (35.6, 43.2, 36.8, 44.4), (38.3, 43.8), False),
]
MARC_PUERTA = [(18, (31.5, 1.3)), (18, (39.0, 1.3))]

CUADROS = {  # rect, posición rótulo
    'CGBT': ((9.5, 0.2, 11.5, 0.8), (10.5, 1.6)),
    'CS-ALU': ((0.2, 12.2, 1.1, 14.6), (2.6, 15.2)),
    'CS-PRO': ((20.3, 12.0, 22.9, 12.9), (24.6, 10.2)),
    'CS-FRI': ((40.5, 12.6, 43.1, 13.5), (41.8, 14.4)),
    'CS-HOR': ((34.7, 47.5, 35.6, 50.1), (37.4, 50.8)),
}
SUBIDA_ALT = (12.3, 11.7, 13.5, 12.9)

BANDEJA = [[(10.5, 0.8), (10.5, 11.0), (35.1, 11.0), (35.1, 47.3)], [(35.1, 11.0), (41.8, 11.0)]]

# Circuitos: receptor nº -> lista de recorridos (cuadro -> equipo)
CIRC = {
    1: [[(35.6, 49.0), (36.2, 49.0), (36.2, 60.0), (x, 60.0), (x, 62.7)] for x in (37.5, 42.0, 46.5, 51.0)],
    5: [[(35.4, 47.5), (35.4, 36.0), (37.0, 36.0)], [(35.4, 47.5), (35.4, 30.0), (37.0, 30.0)]],
    2: [[(21.0, 12.9), (21.0, 13.6), (24.0, 13.6), (24.0, 21.0)]],
    3: [[(21.6, 12.9), (21.6, 14.2), (27.5, 14.2), (27.5, 21.0)]],
    4: [[(22.2, 12.9), (22.2, 14.8), (31.0, 14.8), (31.0, 21.0)]],
    10: [[(22.9, 12.45), (33.9, 12.45), (33.9, 20.15), (40.0, 20.15), (40.0, 20.5)]],
    9: [[(22.9, 12.65), (33.6, 12.65), (33.6, 19.9), (47.5, 19.9), (47.5, 20.5)]],
    17: [[(22.6, 12.9), (22.6, 15.4), (33.0, 15.4), (33.0, 25.6), (29.6, 25.6)],
         [(22.9, 12.25), (34.2, 12.25), (34.2, 24.0), (52.6, 24.0), (52.6, 25.0)],
         [(35.4, 47.5), (35.4, 43.8), (35.6, 43.8)]],
    12: [[(40.8, 13.5), (40.8, 15.5), (36.975, 15.5), (36.975, 16.5)], [(42.8, 13.5), (42.8, 15.8), (51.975, 15.8), (51.975, 16.5)]],
    6: [[(43.1, 12.8), (43.3, 12.8), (43.3, 2.3), (44.0, 2.3)],
        [(43.1, 13.0), (48.8, 13.0), (48.8, 2.3), (49.4, 2.3)]],
    7: [[(43.1, 13.2), (43.65, 13.2), (43.65, 7.75), (44.0, 7.75)]],
    8: [[(43.1, 13.35), (51.0, 13.35), (51.0, 10.0)]],
    18: [[(41.2, 12.6), (41.2, 2.0), (39.0, 2.0)], [(41.6, 12.6), (41.6, 2.6), (31.5, 2.6), (31.5, 2.05)]],
    11: [[(9.5, 0.5), (2.6, 0.5), (2.6, 2.0)]],
}
CIRC_CUADRO = {1: 'CS-HOR', 5: 'CS-HOR', 2: 'CS-PRO', 3: 'CS-PRO', 4: 'CS-PRO', 9: 'CS-PRO', 10: 'CS-PRO',
               12: 'CS-FRI', 6: 'CS-FRI', 7: 'CS-FRI', 8: 'CS-FRI', 18: 'CS-FRI', 11: 'CGBT'}
def circ_color(n, i):
    if n == 17: return C['HOR'] if i == 2 else C['PRO']
    return CUADRO_COLOR[CIRC_CUADRO[n]]

# Derivaciones desde el CGBT (por bandeja) a los cuadros secundarios
ALIM = {
    'CS-PRO': [(10.5, 0.8), (10.5, 11.0), (21.6, 11.0), (21.6, 12.0)],
    'CS-FRI': [(10.5, 0.8), (10.5, 11.0), (41.8, 11.0), (41.8, 12.6)],
    'CS-HOR': [(10.5, 0.8), (10.5, 11.0), (35.1, 11.0), (35.1, 47.5)],
    'CS-ALU': [(10.5, 0.8), (10.5, 11.0), (0.65, 11.0), (0.65, 12.2)],
    'CS-ALT': [(10.5, 0.8), (10.5, 11.0), (12.9, 11.0), (12.9, 11.7)],
}
ALIM_DIBUJO = {'CS-ALU': [(10.5, 11.0), (0.65, 11.0), (0.65, 12.2)],
               'CS-ALT': [(12.9, 11.0), (12.9, 11.7)]}

# Alumbrado de nave (provisional hasta DIALux de la nave nueva)
FILAS_MAIN = [1.5 + 3 * i for i in range(9)]           # y
LUM_X_MAIN = [1.75 + 2.5 * j for j in range(21)]        # x centro
FILAS_ALA = [28.5 + 3 * i for i in range(12)]
LUM_X_ALA = [36.0 + 2.5 * j for j in range(7)]
N_LUM_NAVE = len(FILAS_MAIN) * len(LUM_X_MAIN) + len(FILAS_ALA) * len(LUM_X_ALA)
COLECTOR_MAIN = [(0.65, FILAS_MAIN[0]), (0.65, FILAS_MAIN[-1])]
COLECTOR_ALA = [(34.8, FILAS_MAIN[-1]), (34.8, FILAS_ALA[-1])]
L_ALU = plen([(0.65, 14.6), (0.65, FILAS_MAIN[-1]), (34.8, FILAS_MAIN[-1]), (34.8, FILAS_ALA[-1]),
              (LUM_X_ALA[-1] + 0.6, FILAS_ALA[-1])])

# Emergencias en nave
EMER = []
EMER += [(x, y) for x in (4.5, 13.5, 22.5, 31.5, 40.5, 49.5) for y in (3.0, 9.0, 15.0, 21.0)]
EMER += [(x, y) for x in (38.5, 49.5) for y in (30.0, 36.0, 42.0, 48.0, 54.0, 60.0)]
EMER += [(28.8, 0.9), (36.4, 0.9), (16.25, 25.6), (36.0, 46.25), (52.6, 12.75), (52.6, 48.75)]  # salidas
EMER += [(12.6, 0.9), (1.9, 13.0), (19.6, 12.45), (44.2, 13.05), (36.4, 49.5)]               # cuadros
EMER += [(45.8, 3.4), (51.2, 3.4), (46.25, 9.2), (51.0, 9.2), (16.4, 21.5)]                # cámaras y escalera

# ---------------------------------------------------------------- altillo (coordenadas locales 20 x 15)
ALT_ZONAS = [('OFICINAS', (0, 0, 13, 15), '#eef3f8'), ('VESTUARIOS', (13, 0, 20, 9), '#eaf3e6'),
             ('ALMACÉN\nLIGERO', (13, 9, 17.2, 15), '#f3efe6'), ('', (17.2, 9, 20, 15), '#ffffff')]
ALT_LUM = [(0.909 + 1.818 * i, 0.833 + 1.667 * j) for i in range(11) for j in range(9)]  # DIALux (99 uds)
ALT_EMER = [(2.5, 2.8), (7.5, 2.8), (2.5, 11.8), (7.5, 11.8), (14.5, 2.8), (18.5, 2.8), (14.2, 14.5), (18.5, 7.8)]
ALT_CS = (11.9, 0.2, 12.8, 0.6)
ALT_CASSETTE = (8.6, 5.55, 9.55, 6.5)
ALT_CIRC = [(11.9, 0.4), (9.075, 0.4), (9.075, 5.55)]

N_EMER = len(EMER) + len(ALT_EMER)

# ---------------------------------------------------------------- emplazamiento
PARCELA = (-8.0, -16.0, 58.0, 68.0)
CGP = (9.75, -16.0, 11.25, -15.5)
MEDIDA = (11.5, -1.1, 12.9, -0.4)
LGA = [(10.5, -15.5), (10.5, -3.0), (12.2, -3.0), (12.2, -1.1)]
DI = [(12.2, -0.4), (12.2, 0.0), (11.0, 0.0), (11.0, 0.2)]
CSEXT = (25.0, 48.7, 26.0, 49.7)
L_CSEXT = [(10.0, 0.2), (10.0, -1.4), (-1.0, -1.4), (-1.0, 40.7), (25.5, 40.7), (25.5, 48.7)]
E1 = [(25.5, 49.7), (25.5, 57.7), (-6.5, 57.7), (-6.5, -14.0), (20.0, -14.0)]
E2 = [(26.0, 49.2), (30.3, 49.2), (30.3, 66.5), (56.5, 66.5), (56.5, -14.0), (44.0, -14.0)]
LUZ_EXT = [(-7.0, 52.0), (-7.0, 36.0), (-7.0, 20.0), (-7.0, 4.0), (-7.0, -10.0), (20.0, -14.6),
           (30.3, 62.5), (56.9, 60.0), (56.9, 44.0), (56.9, 28.0), (56.9, 12.0), (56.9, -4.0), (44.0, -14.6)]
Y0P = 26.7  # muro exterior norte del cuerpo principal
FILA1 = Y0P + 6.0
FILA2 = FILA1 + 5.0 + 6.0
FILA3 = FILA2 + 5.0 + 1.0
