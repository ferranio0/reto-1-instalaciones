# -*- coding: utf-8 -*-
"""Construye las entidades de las tres vistas a partir de modelo.py."""
import math
from modelo import *

def tri(sheet, layer, p, s=0.7, color=C['EMER']):
    x, y = p
    return pl(sheet, layer, [(x - s / 2, y - s * 0.4), (x + s / 2, y - s * 0.4), (x, y + s * 0.5)], closed=True,
              color=color, lw=0.3, fill=color, z=5)

def marcador(sheet, n, p, r=0.8, pt=5):
    circ(sheet, 'MAQUINAS', p, r, color='#333333', lw=0.4, fill='#ffffff', z=7)
    txt(sheet, 'TEXTOS', p, str(n), pt=pt, bold=True, z=8)

# ================================================================ PLANTA BAJA
def build_dist():
    S = 'dist'
    # muros
    pl(S, 'CERRAMIENTOS', NAVE_EXT, closed=True, color='#000000', lw=0.5, fill='#1a1a1a', z=0.5)
    pl(S, 'CERRAMIENTOS', NAVE, closed=True, color='#000000', lw=0.5, fill='#ffffff', z=0.6)
    for nombre, (x0, y0, x1, y1), col in ZONAS:
        rect(S, 'ZONAS', x0, y0, x1, y1, color='#8a8a8a', lw=0.3, ls=(0, (6, 2, 1, 2)), fill=col, z=0.8)
        lx, ly = ZONA_LBL.get(nombre, ((x0 + x1) / 2, (y0 + y1) / 2))
        a = (x1 - x0) * (y1 - y0)
        if nombre in ('CÁMARAS', 'ANTECÁMARA'):
            txt(S, 'TEXTOS', (lx, ly), nombre, pt=5.2, bold=True, color='#333333', z=9, bg=True)
            ap = (51.2, 4.8) if nombre == 'CÁMARAS' else (lx, ly - 0.95)
            txt(S, 'TEXTOS', ap, fmt(a, 2) + ' m²', pt=4.5, color='#555555', z=9, bg=True)
        elif nombre:
            txt(S, 'TEXTOS', (lx, ly), nombre, pt=6.5, bold=True, color='#333333', z=9, bg=True)
            dy = 1.25 if '\n' not in nombre else 1.85
            txt(S, 'TEXTOS', (lx, ly - dy), fmt(a, 2) + ' m²', pt=5, color='#555555', z=9, bg=True)
    # altillo sobre planta baja
    ax, ay = ALT_ORIG
    rect(S, 'ALTILLO', ax, ay, ax + ALT_W, ay + ALT_H, color=C['ALT'], lw=0.6, ls=(0, (4, 2)), z=1.5)
    for i in range(13):
        yy = ay + 9.0 + 0.5 * i
        pl(S, 'ALTILLO', [(ax + 17.2, yy), (ax + 20.0, yy)], color='#555555', lw=0.3, z=1.6)
    pl(S, 'ALTILLO', [(ax + 17.2, ay + 9.0), (ax + 17.2, ay + 15.0)], color='#555555', lw=0.3, z=1.6)
    txt(S, 'TEXTOS', (ax + 18.6, ay + 8.2), 'Escalera', pt=4.5, color='#444444')
    # ventanas y puertas
    for w, a, b in VENTANAS:
        x0, y0, x1, y1 = ventana_rect(w, a, b)
        rect(S, 'VENTANAS', x0, y0, x1, y1, color='#1e63c6', lw=0.6, fill='#9cc3f0', z=1.0)
    for (x0, y0, x1, y1), lbl, lp in PUERTAS:
        rect(S, 'PUERTAS', x0, y0, x1, y1, color='#b0127c', lw=0.6, fill='#e57ac0', z=1.0)
        txt(S, 'TEXTOS', lp, lbl, pt=5, color='#b0127c')
    # alumbrado (provisional)
    for y in FILAS_MAIN:
        pl(S, 'ALUMBRADO', [(0.65, y), (LUM_X_MAIN[-1] + 0.6, y)], color=C['ALU'], lw=0.35, z=2)
        for x in LUM_X_MAIN:
            pl(S, 'ALUMBRADO', [(x - 0.6, y), (x + 0.6, y)], color='#7a5c00', lw=1.4, z=2.1)
    for y in FILAS_ALA:
        pl(S, 'ALUMBRADO', [(34.8, y), (LUM_X_ALA[-1] + 0.6, y)], color=C['ALU'], lw=0.35, z=2)
        for x in LUM_X_ALA:
            pl(S, 'ALUMBRADO', [(x - 0.6, y), (x + 0.6, y)], color='#7a5c00', lw=1.4, z=2.1)
    pl(S, 'ALUMBRADO', COLECTOR_MAIN, color=C['ALU'], lw=0.5, z=2)
    pl(S, 'ALUMBRADO', COLECTOR_ALA, color=C['ALU'], lw=0.5, z=2)
    for p in EMER:
        tri(S, 'EMERGENCIA', p)
    # bandeja
    for b in BANDEJA:
        pl(S, 'BANDEJA', b, color=C['BANDEJA'], lw=2.2, z=3)
    for k, pts in ALIM_DIBUJO.items():
        pl(S, 'CIRCUITOS', pts, color=CUADRO_COLOR[k], lw=0.8, z=3.2)
    # circuitos
    for n, rutas in CIRC.items():
        for i, r in enumerate(rutas):
            pl(S, 'CIRCUITOS', r, color=circ_color(n, i), lw=0.6, z=3.5)
    # máquinas
    for n, (x0, y0, x1, y1), mp, techo in MAQ:
        rect(S, 'MAQUINAS', x0, y0, x1, y1, color='#222222', lw=0.5, fill='#ffffff',
             ls=(0, (3, 1.5)) if techo else '-', z=4)
        marcador(S, n, mp)
    for n, p in MARC_PUERTA:
        marcador(S, n, p, r=0.75)
    # cuadros
    for k, ((x0, y0, x1, y1), lp) in CUADROS.items():
        rect(S, 'CUADROS', x0, y0, x1, y1, color='#000000', lw=0.4, fill=CUADRO_COLOR[k], z=6)
        txt(S, 'TEXTOS', lp, k, pt=5.5, bold=True, color=CUADRO_COLOR[k], z=9, bg=True)
    rect(S, 'CUADROS', *SUBIDA_ALT, color='#000000', lw=0.4, fill=C['ALT'], z=6)
    txt(S, 'TEXTOS', (14.6, 10.2), 'Subida a CS-ALT', pt=4.5, color=C['ALT'], z=9, bg=True)
    # cotas exteriores y en cadena
    dim(S, (0, 0), (53.5, 0), -5.4)
    dim(S, (0, 0), (0, 26.5), -6.6)
    dim(S, (53.5, 26.5), (53.5, 64.5), 6.6)
    dim(S, (34.5, 64.5), (53.5, 64.5), 2.6)
    dim(S, (0, 26.5), (34.5, 26.5), 3.0)
    for a, b in ((0, 12), (12, 21.5), (21.5, 34.5), (34.5, 43.5), (43.5, 53.5)):
        dim(S, (a, 0), (b, 0), -3.2, pt=5)
    for a, b in ((0, 11.5), (11.5, 26.5)):
        dim(S, (0, a), (0, b), -3.6, pt=5)
    for a, b in ((0, 14), (14, 26.5), (26.5, 45), (45, 64.5)):
        dim(S, (53.5, a), (53.5, b), 3.6, pt=5)

# ================================================================ ALTILLO
def build_alt():
    S = 'alt'
    pl(S, 'ALTILLO', [(0, 0), (ALT_W, 0), (ALT_W, ALT_H), (0, ALT_H)], closed=True, color='#000000', lw=1.2,
       fill='#ffffff', z=0.5)
    for nombre, (x0, y0, x1, y1), col in ALT_ZONAS:
        rect(S, 'ZONAS', x0, y0, x1, y1, color='#555555', lw=0.5, fill=col, z=0.8)
    for i in range(13):
        yy = 9.0 + 0.5 * i
        pl(S, 'ALTILLO', [(17.2, yy), (20.0, yy)], color='#555555', lw=0.3, z=1)
    txt(S, 'TEXTOS', (18.6, 12.0), 'Escalera', pt=5, rot=90, color='#444444', z=9, bg=True)
    for p in ALT_LUM:
        circ(S, 'ALUMBRADO', p, 0.17, color='#8a6d00', lw=0.3, fill='#f5c518', z=2)
    for p in ALT_EMER:
        tri(S, 'EMERGENCIA', p, s=0.55)
    rect(S, 'MAQUINAS', *ALT_CASSETTE, color='#222222', lw=0.5, fill='#ffffff', ls=(0, (3, 1.5)), z=4)
    marcador(S, 12, (9.075, 7.5), r=0.55, pt=5)
    pl(S, 'CIRCUITOS', ALT_CIRC, color=C['ALT'], lw=0.7, z=3)
    rect(S, 'CUADROS', *ALT_CS, color='#000000', lw=0.4, fill=C['ALT'], z=6)
    txt(S, 'TEXTOS', (13.9, 1.2), 'CS-ALT', pt=5.5, bold=True, color=C['ALT'], z=9)
    txt(S, 'TEXTOS', (6.5, 13.6), 'OFICINAS', pt=6.5, bold=True, color='#333333', z=9, bg=True)
    txt(S, 'TEXTOS', (16.5, 5.5), 'VESTUARIOS', pt=6.5, bold=True, color='#333333', z=9, bg=True)
    txt(S, 'TEXTOS', (15.1, 12.4), 'ALMACÉN\nLIGERO', pt=5.5, bold=True, color='#333333', z=9, bg=True)
    for nombre, (x0, y0, x1, y1), col in ALT_ZONAS[:3]:
        lx = {'OFICINAS': (6.5, 12.6), 'VESTUARIOS': (16.5, 4.5), 'ALMACÉN\nLIGERO': (15.1, 10.9)}[nombre]
        txt(S, 'TEXTOS', lx, fmt((x1 - x0) * (y1 - y0), 2) + ' m²', pt=4.8, color='#555555', z=9, bg=True)
    dim(S, (0, 0), (ALT_W, 0), -1.6)
    dim(S, (0, 0), (0, ALT_H), -1.6)

# ================================================================ EMPLAZAMIENTO
def build_emp():
    S = 'emp'
    x0, y0, x1, y1 = PARCELA
    rect(S, 'PARCELA', x0, y0, x1, y1, color='#000000', lw=0.7, fill='#ffffff', z=0.2)
    rect(S, 'PARCELA', x0, y0, x1, y1, color='#000000', lw=0.9, z=0.9)
    # pavimentos
    gris = '#e3e3e3'
    rect(S, 'PAVIMENTO', x0, y0, x1, -0.2, color=gris, lw=0.1, fill=gris, z=0.3)        # patio de carga
    rect(S, 'PAVIMENTO', x0, -0.2, -0.2, FILA3 + 5 + 6, color=gris, lw=0.1, fill=gris, z=0.3)  # vial oeste
    rect(S, 'PAVIMENTO', -0.2, Y0P, 34.3, FILA3 + 5 + 6, color=gris, lw=0.1, fill=gris, z=0.3)  # aparcamiento
    rect(S, 'PAVIMENTO', 53.7, -0.2, x1, y1, color=gris, lw=0.1, fill=gris, z=0.3)         # vial este
    rect(S, 'PAVIMENTO', 34.3, 64.7, x1, y1, color=gris, lw=0.1, fill=gris, z=0.3)         # trasera ala
    rect(S, 'PAVIMENTO', x0, FILA3 + 11, 34.3, y1, color='#9cbf8c', lw=0.3, fill='#edf4e8', z=0.3)
    txt(S, 'TEXTOS', (13.0, (FILA3 + 11 + y1) / 2), 'Zona ajardinada / reserva', pt=6, color='#4d6b3c')
    # nave
    pl(S, 'CERRAMIENTOS', NAVE_EXT, closed=True, color='#000000', lw=0.5, fill='#1a1a1a', z=0.5)
    pl(S, 'CERRAMIENTOS', NAVE, closed=True, color='#000000', lw=0.5, fill='#ffffff', z=0.6)
    for w, a, b in VENTANAS:
        xa, ya, xb, yb = ventana_rect(w, a, b)
        rect(S, 'VENTANAS', xa, ya, xb, yb, color='#1e63c6', lw=0.6, fill='#9cc3f0', z=1.0)
    for (xa, ya, xb, yb), lbl, lp in PUERTAS:
        rect(S, 'PUERTAS', xa, ya, xb, yb, color='#b0127c', lw=0.6, fill='#e57ac0', z=1.0)
    ax, ay = ALT_ORIG
    rect(S, 'ALTILLO', ax, ay, ax + ALT_W, ay + ALT_H, color='#2e8b3a', lw=0.6, ls=(0, (4, 2)), z=1.5)
    txt(S, 'TEXTOS', (ax + 10, ay + 7.5), 'ALTILLO\n300 m²', pt=6.5, bold=True, color='#2e8b3a')
    txt(S, 'TEXTOS', (27.0, 6.0), 'PLANTA BAJA\n' + fmt(A_NAVE, 2) + ' m²', pt=9, bold=True)
    txt(S, 'TEXTOS', (44.0, 45.5), 'ALA', pt=8, bold=True)
    # aparcamiento
    def plaza(xa, ya, w, tipo):
        fill = {'n': '#ffffff', 've': '#a5d6a7', 'pre': '#fff3b0', 'pmr': '#b3cde8'}[tipo]
        rect(S, 'APARCAMIENTO' if tipo == 'n' else 'RECARGA', xa, ya, xa + w, ya + 5.0, color='#555555', lw=0.35,
             fill=fill, z=1.2)
        lbl = {'ve': 'VE', 'pre': 'Pre', 'pmr': 'PMR'}.get(tipo)
        if lbl:
            txt(S, 'RECARGA', (xa + w / 2, ya + 2.5), lbl, pt=4.8, bold=True, color='#333333')
    xs = 0.0
    for i in range(10):
        w = 3.7 if i == 5 else 2.5
        plaza(xs, FILA1, w, 'pmr' if i == 5 else 'n'); xs += w
    for fy in (FILA2, FILA3):
        for i in range(10):
            tipo = {7: 'pre', 8: 've', 9: 've'}.get(i, 'n')
            plaza(2.5 * i, fy, 2.5, tipo)
    rect(S, 'APARCAMIENTO', 0, FILA2 + 5, 25.0, FILA3, color='#555555', lw=0.35, fill='#d0d0d0', z=1.2)
    for cx, inst in ((18.75, False), (21.25, True), (23.75, True)):
        rect(S, 'RECARGA', cx - 0.45, FILA2 + 5.05, cx + 0.45, FILA3 - 0.05, color='#1b5e20', lw=0.4,
             fill='#1b5e20' if inst else '#ffffff', z=1.4)
    for xx in (15.5, 17.0):
        pl(S, 'APARCAMIENTO', [(xx, Y0P), (xx, FILA1)], color='#555555', lw=0.35, ls=(0, (2, 1.5)), z=1.2)
    for yy in (Y0P + 3.0, FILA1 + 6.8, FILA3 + 6.6):
        txt(S, 'TEXTOS', (8.0, yy), 'Pasillo 6 m', pt=5.5, color='#444444')
    txt(S, 'TEXTOS', (30.3, 40.0), 'Vial 8 m', pt=5.5, rot=90, color='#444444')
    txt(S, 'TEXTOS', (-4.1, 10.0), 'Vial 7,80 m', pt=5.5, rot=90, color='#444444')
    txt(S, 'TEXTOS', (27.0, -9.0), 'Patio de carga', pt=6, color='#444444')
    # accesos y calle
    for xa, xb, lbl in ((-8.0, -2.0, 'Acceso turismos 6 m'), (30.0, 38.0, 'Acceso camiones 8 m')):
        pl(S, 'PARCELA', [(xa, y0), (xb, y0)], color='#e3e3e3', lw=2.2, z=0.95)
        for xx in (xa, xb):
            pl(S, 'PARCELA', [(xx, y0 - 0.9), (xx, y0 + 0.9)], color='#000000', lw=0.6, z=0.96)
        txt(S, 'TEXTOS', ((xa + xb) / 2, y0 - 1.6), lbl, pt=5, color='#333333')
    txt(S, 'TEXTOS', (25.0, y0 - 5.0), 'CALLE (SUPUESTA)', pt=6, color='#555555')
    # acometida y cuadros
    for r, lbl, lp in ((CGP, 'CGP', (13.0, -15.2)), (MEDIDA, 'Medida', (15.3, -0.75)),
                       ((9.5, 0.2, 11.5, 0.8), 'CGBT', (10.5, 2.0))):
        rect(S, 'CUADROS', *r, color='#000000', lw=0.4, fill=C['CGBT'], z=6)
        txt(S, 'TEXTOS', lp, lbl, pt=5.5, bold=True, color=C['CGBT'], z=9)
    pl(S, 'CIRCUITOS', LGA, color=C['CGBT'], lw=0.9, z=5)
    pl(S, 'CIRCUITOS', DI, color=C['CGBT'], lw=0.9, z=5)
    txt(S, 'TEXTOS', (8.6, -6.0), 'LGA', pt=5.5, bold=True, color=C['CGBT'])
    rect(S, 'CUADROS', *CSEXT, color='#000000', lw=0.4, fill=C['EXT'], z=6)
    txt(S, 'TEXTOS', (29.6, 50.6), 'CS-EXT', pt=5.5, bold=True, color=C['EXT'], z=9)
    pl(S, 'CIRCUITOS', L_CSEXT, color=C['EXT'], lw=0.9, z=5)
    pl(S, 'CIRCUITOS', E1, color=C['EXT'], lw=0.6, ls=(0, (4, 2)), z=5)
    pl(S, 'CIRCUITOS', E2, color=C['EXT'], lw=0.6, ls=(0, (4, 2)), z=5)
    txt(S, 'TEXTOS', (-5.0, 59.0), 'E1', pt=5.5, bold=True, color=C['EXT'])
    txt(S, 'TEXTOS', (55.0, 67.4), 'E2', pt=5.5, bold=True, color=C['EXT'])
    for p in LUZ_EXT:
        circ(S, 'ALUMBRADO', p, 0.75, color='#8a6d00', lw=0.4, fill='#f5c518', z=6)
    # cotas
    dim(S, (x0, y1), (x1, y1), 3.2)
    dim(S, (x0, y0), (x0, y1), -4.0)
    dim(S, (0, -0.2), (53.5, -0.2), -3.0, pt=5.5)
    dim(S, (34.5, 64.7), (53.5, 64.7), -1.6, pt=5)
    dim(S, (53.5, 26.5), (53.5, 64.5), -2.4, pt=5.5)
    dim(S, (-0.2, 0), (-0.2, 26.5), -1.8, pt=5)
    dim(S, (x1, y0), (x1, -0.2), 3.6, pt=5)
    dim(S, (x0, 30.0), (-0.2, 30.0), 0.0, pt=5)

# ---------------------------------------------------------------- longitudes
def longitudes():
    L = {}
    for n, rutas in CIRC.items():
        L[n] = max(plen(r) for r in rutas)
    L[12] = max(L[12], plen(ALT_CIRC))
    L[13] = L_ALU
    Lc = {k: plen(v) for k, v in ALIM.items()}
    Lc['CS-EXT'] = plen(L_CSEXT)
    Le = dict(LGA=plen(LGA), DI=plen(DI), E1=plen(E1), E2=plen(E2), CSEXT=plen(L_CSEXT))
    return L, Lc, Le

def construir():
    K.clear(); build_dist(); build_alt(); build_emp()
    return K
