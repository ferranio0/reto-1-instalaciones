# -*- coding: utf-8 -*-
"""Dibuja las hojas A3 (PDF) y exporta el DXF."""
import math, sys, os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Circle, Rectangle
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib import font_manager
import matplotlib.image as mpimg
from construir import *

plt.rcParams['font.family'] = 'Liberation Sans'
plt.rcParams['pdf.fonttype'] = 42
HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = mpimg.imread(os.path.join(HERE, 'upv_logo.png'))
FECHA = '12/11/2026'
AUTORES = ['Ferran Bertomeu Peñarrocha', 'Álvaro Rubio Ordoñez', 'Néstor Ordiñana Fenollar']

class Vista:
    def __init__(self, ax, ox, oy, esc):
        self.ax, self.ox, self.oy, self.k = ax, ox, oy, 1000.0 / esc
    def P(self, p):
        return (self.ox + p[0] * self.k, self.oy + p[1] * self.k)

def dibujar(v, ents):
    ax = v.ax
    for e in sorted(ents, key=lambda e: e.get('z', 0)):
        k = e['k']
        if k == 'pl':
            pts = [v.P(p) for p in e['pts']]
            if e['closed'] and e['fill']:
                ax.add_patch(Polygon(pts, closed=True, facecolor=e['fill'], edgecolor=e['color'], lw=e['lw'],
                                     ls=e['ls'], zorder=e['z'], joinstyle='miter'))
            elif e['closed']:
                ax.add_patch(Polygon(pts, closed=True, fill=False, edgecolor=e['color'], lw=e['lw'],
                                     ls=e['ls'], zorder=e['z']))
            else:
                xs, ys = zip(*pts)
                ax.plot(xs, ys, color=e['color'], lw=e['lw'], ls=e['ls'], zorder=e['z'],
                        solid_capstyle='butt', solid_joinstyle='miter')
        elif k == 'c':
            ax.add_patch(Circle(v.P(e['c']), e['r'] * v.k, facecolor=e['fill'] or 'none', edgecolor=e['color'],
                                lw=e['lw'], zorder=e['z']))
        elif k == 't':
            ax.text(*v.P(e['p']), e['s'], fontsize=e['pt'], color=e['color'], ha=e['ha'], va=e['va'],
                    fontweight='bold' if e['bold'] else 'normal', rotation=e['rot'], zorder=e['z'],
                    linespacing=1.05, bbox=dict(boxstyle='square,pad=0.08', fc='white', ec='none', alpha=0.8) if e.get('bg') else None)
        elif k == 'dim':
            cota(v, e)

def cota(v, e):
    ax = v.ax
    (x1, y1), (x2, y2), off = e['p1'], e['p2'], e['off']
    horiz = abs(y1 - y2) < 1e-9
    val = abs(x2 - x1) if horiz else abs(y2 - y1)
    s = e['s'] or fmt(val, 2)
    col = e['color']
    if horiz:
        a, b = (x1, y1 + off), (x2, y2 + off)
        ext = [((x1, y1), a), ((x2, y2), b)]
    else:
        a, b = (x1 + off, y1), (x2 + off, y2)
        ext = [((x1, y1), a), ((x2, y2), b)]
    for p, q in ext:
        P, Q = v.P(p), v.P(q)
        # línea de referencia con pequeño hueco y prolongación de 1 mm
        dx, dy = Q[0] - P[0], Q[1] - P[1]
        L = math.hypot(dx, dy) or 1
        ux, uy = dx / L, dy / L
        ax.plot([P[0] + ux * 0.8, Q[0] + ux * 1.2], [P[1] + uy * 0.8, Q[1] + uy * 1.2], color=col, lw=0.25,
                zorder=10)
    A, B = v.P(a), v.P(b)
    ax.plot([A[0], B[0]], [A[1], B[1]], color=col, lw=0.3, zorder=10)
    for X in (A, B):
        ax.plot([X[0] - 0.9, X[0] + 0.9], [X[1] - 0.9, X[1] + 0.9], color=col, lw=0.6, zorder=10)
    m = ((A[0] + B[0]) / 2, (A[1] + B[1]) / 2)
    if horiz:
        sgn = 1 if off >= 0 else -1
        ax.text(m[0], m[1] + 0.7 * (1 if sgn > 0 else 1), s, fontsize=e['pt'], color=col, ha='center', va='bottom',
                zorder=11, bbox=dict(boxstyle='square,pad=0.05', fc='white', ec='none', alpha=0.85))
    else:
        ax.text(m[0] - 0.7, m[1], s, fontsize=e['pt'], color=col, ha='right', va='center', rotation=90,
                rotation_mode='anchor', zorder=11,
                bbox=dict(boxstyle='square,pad=0.05', fc='white', ec='none', alpha=0.85))

# ---------------------------------------------------------------- página
def pagina():
    fig = plt.figure(figsize=(420 / 25.4, 297 / 25.4))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 420); ax.set_ylim(0, 297); ax.set_aspect('equal'); ax.axis('off')
    ax.add_patch(Rectangle((10, 10), 400, 277, fill=False, lw=1.0, ec='#000000', zorder=20))
    return fig, ax

def cajetin(ax, plano, escala, num):
    X0, X1, X2, X3 = 247, 304, 375, 410
    Y0, Y1, Y2 = 10, 34, 57
    lw = 0.6
    def R(x0, y0, x1, y1):
        ax.add_patch(Rectangle((x0, y0), x1 - x0, y1 - y0, fc='white', ec='black', lw=lw, zorder=20))
    R(X0, Y0, X3, Y2)
    R(X0, Y1, X1, Y2); R(X0, Y0, X1, Y1)
    R(X1, Y1, X2, Y2); R(X1, Y0, X2, Y1)
    R(X2, 41, X3, Y2); R(X2, 26, X3, 41); R(X2, Y0, X3, 26)
    ax.imshow(LOGO, extent=(X0 + 3, X1 - 3, Y1 + 2.5, Y1 + 2.5 + (X1 - X0 - 6) * LOGO.shape[0] / LOGO.shape[1]),
              zorder=21, interpolation='antialiased')
    T = lambda x, y, s, pt=5, b=False, ha='left', va='top': ax.text(x, y, s, fontsize=pt, fontweight='bold' if b else
                                                                    'normal', ha=ha, va=va, zorder=22)
    T(X0 + 1.5, Y1 - 1.2, 'EPS Alcoy', 6.5, True)
    T(X0 + 1.5, Y1 - 5.2, 'GRADO EN INGENIERÍA MECÁNICA', 5)
    ax.plot([X0, X1], [Y1 - 8.5, Y1 - 8.5], color='black', lw=0.4, zorder=21)
    T(X0 + 1.5, Y1 - 9.6, 'AUTORES:', 4.2)
    for i, a in enumerate(AUTORES):
        T(X0 + 1.5, Y1 - 13.0 - 3.2 * i, a, 5)
    T(X1 + 1.5, Y2 - 1.2, 'TÍTULO:', 4.2)
    T((X1 + X2) / 2, (Y1 + Y2) / 2 - 1, 'PROYECTO DE LA INSTALACIÓN ELÉCTRICA DE BAJA\nTENSIÓN DE UNA PANADERÍA '
      'INDUSTRIAL', 6.5, False, 'center', 'center')
    T(X1 + 1.5, Y1 - 1.2, 'PLANO:', 4.2)
    T((X1 + X2) / 2, (Y0 + Y1) / 2 - 1, plano, 6.8, True, 'center', 'center')
    T(X2 + 1.5, Y2 - 1.2, 'FECHA:', 4.2); T((X2 + X3) / 2, 47.5, FECHA, 6.5, False, 'center', 'center')
    T(X2 + 1.5, 41 - 1.2, 'ESCALA:', 4.2); T((X2 + X3) / 2, 32.5, escala, 6.5, False, 'center', 'center')
    T(X2 + 1.5, 26 - 1.2, 'PLANO Nº:', 4.2); T((X2 + X3) / 2, 16.5, str(num), 12, True, 'center', 'center')

def tabla(ax, x, y, titulo, cols, filas, anchos, pt=5.2, h=3.9, negrita_ult=False, swatch=None):
    """Tabla con cabecera gris; devuelve la y inferior."""
    ax.text(x, y, titulo, fontsize=7.5, fontweight='bold', va='top', ha='left', zorder=22)
    y -= 5.0
    W = sum(anchos)
    def fila(vals, yy, head=False, bold=False, sw=None):
        ax.add_patch(Rectangle((x, yy - h), W, h, fc='#dde3ea' if head else 'white', ec='#7d8793', lw=0.3,
                               zorder=21))
        xx = x
        for i, (v, w) in enumerate(zip(vals, anchos)):
            ax.text(xx + 1.0, yy - h / 2, v, fontsize=pt, va='center', ha='left', zorder=22,
                    fontweight='bold' if (head or bold) else 'normal')
            xx += w
        if sw:
            ax.add_patch(Rectangle((x + W - 6.5, yy - h + 0.8), 5.5, h - 1.6, fc=sw, ec='black', lw=0.3, zorder=22))
    fila(cols, y, head=True); y -= h
    for i, f in enumerate(filas):
        fila(f, y, bold=(negrita_ult and i == len(filas) - 1), sw=swatch[i] if swatch else None); y -= h
    return y

def leyenda(ax, x, y, items, titulo='LEYENDA'):
    ax.text(x, y, titulo, fontsize=7.5, fontweight='bold', va='top', ha='left', zorder=22)
    y -= 6.5
    for kind, spec, lbl in items:
        cx, cy = x + 4, y
        if kind == 'rect':
            fc, ec = spec
            ax.add_patch(Rectangle((cx - 2.6, cy - 1.4), 5.2, 2.8, fc=fc, ec=ec, lw=0.4, zorder=22))
        elif kind == 'rectd':
            ax.add_patch(Rectangle((cx - 2.0, cy - 1.4), 4.0, 2.8, fc='white', ec='black', lw=0.4, ls=(0, (3, 1.5)),
                                   zorder=22))
        elif kind == 'line':
            col, lw, ls = spec
            ax.plot([cx - 3.5, cx + 3.5], [cy, cy], color=col, lw=lw, ls=ls, zorder=22)
        elif kind == 'lum':
            ax.plot([cx - 3.5, cx + 3.5], [cy, cy], color=C['ALU'], lw=0.35, zorder=22)
            ax.plot([cx - 2, cx + 2], [cy, cy], color='#7a5c00', lw=1.4, zorder=22)
        elif kind == 'circ':
            fc, ec = spec
            ax.add_patch(Circle((cx, cy), 1.3, fc=fc, ec=ec, lw=0.4, zorder=22))
        elif kind == 'tri':
            ax.add_patch(Polygon([(cx - 1.4, cy - 1.0), (cx + 1.4, cy - 1.0), (cx, cy + 1.3)], fc=C['EMER'],
                                 ec=C['EMER'], zorder=22))
        elif kind == 'marc':
            ax.add_patch(Circle((cx, cy), 1.5, fc='white', ec='#333333', lw=0.4, zorder=22))
            ax.text(cx, cy, 'n', fontsize=5, ha='center', va='center', fontweight='bold', zorder=23)
        ax.text(x + 10, y, lbl, fontsize=5.3, va='center', ha='left', zorder=22)
        y -= 4.6
    return y

def notas(ax, x, y, lineas, pt=5.0):
    for l in lineas:
        ax.text(x, y, l, fontsize=pt, va='top', ha='left', zorder=22, color='#222222')
        y -= 3.4
    return y

# ================================================================ HOJA 1 · EMPLAZAMIENTO
def hoja_emplazamiento(pdf, L, Lc, Le):
    fig, ax = pagina()
    v = Vista(ax, 60.0, 82.0, 400)
    dibujar(v, K['emp'])
    ax.text(18, 280, 'EMPLAZAMIENTO · E 1:400', fontsize=10, fontweight='bold', va='top')
    # flecha de "calle"
    x0, y0, x1, y1 = PARCELA
    pa, pb = (x1 - x0), (y1 - y0)
    xT = 252
    y = 281
    y = tabla(ax, xT, y, 'EXTERIORES', ['Concepto', 'Valor'], [
        [f'Parcela ({fmt(pa, 0)} × {fmt(pb, 0)} m)', fmt(pa * pb, 0) + ' m²'],
        ['Planta baja interior', fmt(A_NAVE, 2) + ' m²'],
        ['Altillo (20,00 × 15,00 m)', f'300 m² ({fmt(300 / A_NAVE * 100, 1)} % de la planta baja)'],
        ['Plazas de aparcamiento (2,50 × 5,00 m)', '30'],
        ['Cargadores dobles instalados (C1 y C2)', '2'],
        ['Plazas con punto de recarga', '4'],
        ['Cargador doble previsto (C3)', '1 (2 plazas)'],
        ['Plazas con conducciones', '6 (20 %)'],
        ['Plaza accesible', '1'],
        ['Luminarias exteriores', '13'],
    ], [92, 64])
    y -= 6
    y = tabla(ax, xT, y, 'LÍNEAS (PROPUESTA)', ['Línea', 'L en planta'], [
        ['LGA: CGP a equipo de medida', f"{fmt(Le['LGA'], 0)} m"],
        ['Derivación individual: medida a CGBT', f"{fmt(Le['DI'], 1)} m"],
        ['CGBT a CS-EXT', f"{fmt(Le['CSEXT'], 0)} m"],
        ['Alumbrado exterior E1', f"{fmt(Le['E1'], 0)} m"],
        ['Alumbrado exterior E2', f"{fmt(Le['E2'], 0)} m"],
    ], [92, 64])
    y -= 6
    y = leyenda(ax, xT, y, [
        ('rect', ('#a5d6a7', '#555555'), 'Plaza con punto de recarga'),
        ('rect', ('#fff3b0', '#555555'), 'Plaza con cargador previsto'),
        ('rect', ('#b3cde8', '#555555'), 'Plaza accesible (PMR)'),
        ('rect', ('#1b5e20', '#1b5e20'), 'Cargador doble instalado'),
        ('rect', ('#ffffff', '#1b5e20'), 'Cargador doble previsto'),
        ('circ', ('#f5c518', '#8a6d00'), 'Punto de luz exterior'),
        ('rect', (C['CGBT'], '#000000'), 'CGP, medida o cuadro'),
        ('line', (C['CGBT'], 0.9, '-'), 'LGA y derivación individual'),
        ('line', (C['EXT'], 0.9, '-'), 'Línea CGBT a CS-EXT'),
        ('line', (C['EXT'], 0.6, (0, (4, 2))), 'Alumbrado exterior (enterrado)'),
        ('rect', ('#edf4e8', '#9cbf8c'), 'Zona ajardinada'),
    ])
    y -= 2
    notas(ax, xT, y, ['Cotas en metros. Trazados y posición de luminarias: propuesta.',
                      'Se supone que la calle y la acometida están al sur del dibujo; falta el',
                      'emplazamiento real y la orientación del norte.'])
    cajetin(ax, 'PLANO DE EMPLAZAMIENTO', '1:400', 1)
    pdf.savefig(fig); fig.savefig(os.path.join(HERE, "out", "p1.png"), dpi=110); plt.close(fig)

# ================================================================ HOJA 2 · DISTRIBUCIÓN
def hoja_distribucion(pdf, L, Lc, Le):
    fig, ax = pagina()
    vd = Vista(ax, 40.0, 36.0, 300)
    dibujar(vd, K['dist'])
    va = Vista(ax, 42.0, 151.0, 200)
    dibujar(va, K['alt'])
    ax.text(18, 280, 'DISTRIBUCIÓN DE LA INSTALACIÓN', fontsize=10, fontweight='bold', va='top')
    ax.text(42, 237.5, 'PLANTA ALTILLO (300 m²) · E 1:200', fontsize=7.5, fontweight='bold', va='bottom')
    ax.text(45, 128.5, f'PLANTA BAJA ({fmt(A_NAVE, 2)} m²) · E 1:300', fontsize=7.5, fontweight='bold',
            va='bottom')
    xT = 250
    filas = []
    for n, d, m, u, p, c in RECEPTORES:
        if n in L:
            l = str(math.ceil(L[n]))
        elif n in (15, 19):
            l = 'Pl. 1'
        else:
            l = '—'
        filas.append([str(n), d, m, u, fmt(p, 2), c, l])
    filas.append(['', 'POTENCIA NOMINAL TOTAL', '', '', fmt(P_TOTAL, 2), '', ''])
    y = tabla(ax, xT, 281, 'RECEPTORES', ['Nº', 'Receptor', 'Modelo', 'Uds', 'P (kW)', 'Cuadro', 'L (m)'],
              filas, [6.5, 37, 42, 10, 12, 20, 11], pt=5.0, h=3.75, negrita_ult=True)
    y -= 5
    cq = [['CGBT', 'Cuadro general de baja tensión', '—'],
          ['CS-PRO', 'Amasado, envasado y tomas', str(math.ceil(Lc['CS-PRO']))],
          ['CS-HOR', 'Hornos, fermentación y tomas', str(math.ceil(Lc['CS-HOR']))],
          ['CS-FRI', 'Cámaras, puertas y climatización', str(math.ceil(Lc['CS-FRI']))],
          ['CS-ALU', 'Alumbrado y emergencia de nave', str(math.ceil(Lc['CS-ALU']))],
          ['CS-ALT', 'Altillo (alumbrado, climatización)', str(math.ceil(Lc['CS-ALT'])) + ' + subida'],
          ['CS-EXT', 'Recarga y alumbrado exterior', 'Pl. 1']]
    sw = [CUADRO_COLOR[r[0]] for r in cq]
    y = tabla(ax, xT, y, 'CUADROS (PROPUESTA)', ['Cuadro', 'Alimenta', 'L desde CGBT (m)', ''], cq,
              [17, 62, 30, 29.5], pt=5.0, h=3.75, swatch=sw)
    y -= 5
    xL = xT
    yl = leyenda(ax, xL, y, [
        ('lum', None, f'Luminaria estanca ({len(FILAS_MAIN)}+{len(FILAS_ALA)} líneas, {N_LUM_NAVE} uds)'),
        ('tri', None, f'Luminaria de emergencia ({N_EMER} uds)'),
        ('circ', ('#f5c518', '#8a6d00'), 'Downlight altillo (99 uds, DIALux)'),
        ('line', (C['BANDEJA'], 2.2, '-'), 'Bandeja principal'),
        ('line', ('#555555', 0.6, '-'), 'Circuito (color del cuadro)'),
        ('marc', None, 'Receptor nº n de la tabla'),
    ])
    leyenda(ax, xL + 78, y, [
        ('rect', ('#ffffff', '#222222'), 'Máquina o equipo'),
        ('rectd', None, 'Cassette en techo'),
        ('rect', ('#e57ac0', '#b0127c'), 'Puerta'),
        ('rect', ('#9cc3f0', '#1e63c6'), 'Ventana'),
        ('line', (C['ALT'], 0.6, (0, (4, 2))), 'Límite del altillo sobre planta baja'),
    ], titulo='')
    notas(ax, xL, yl - 1.0, [
        'Cotas en metros. Ubicación de máquinas, cuadros y trazados: propuesta.',
        f'* Alumbrado de planta baja provisional: {N_LUM_NAVE} luminarias dibujadas frente a 328 de la tabla',
        '   de receptores; ambas se cierran con el DIALux de la nave nueva.',
        'L: longitud en planta del circuito más largo, sin subidas ni bajadas a equipos.'], pt=4.9)
    cajetin(ax, 'PLANO DE DISTRIBUCIÓN DE LA INSTALACIÓN\nPLANTA BAJA Y ALTILLO', '1:300 / 1:200', 2)
    pdf.savefig(fig); fig.savefig(os.path.join(HERE, 'out', 'p2.png'), dpi=110); plt.close(fig)

# ================================================================ DXF (R12)
ACI = {'CERRAMIENTOS': 7, 'ZONAS': 8, 'ALTILLO': 6, 'VENTANAS': 5, 'PUERTAS': 6, 'ALUMBRADO': 2, 'EMERGENCIA': 3,
       'BANDEJA': 30, 'CIRCUITOS': 1, 'MAQUINAS': 7, 'CUADROS': 1, 'TEXTOS': 7, 'COTAS': 1, 'PARCELA': 7,
       'PAVIMENTO': 8, 'APARCAMIENTO': 8, 'RECARGA': 3}

def hex2aci(h):
    tabla_ = {C['CGBT']: 1, C['PRO']: 5, C['HOR']: 30, C['FRI']: 4, C['ALU']: 2, C['ALT']: 6, C['EXT']: 3}
    return tabla_.get(h)

def exportar_dxf(ruta):
    out = []
    w = lambda c, v: out.extend([str(c), str(v)])
    w(0, 'SECTION'); w(2, 'HEADER')
    w(9, '$ACADVER'); w(1, 'AC1009')
    w(9, '$DWGCODEPAGE'); w(3, 'ANSI_1252')
    w(9, '$INSBASE'); w(10, 0); w(20, 0); w(30, 0)
    w(9, '$EXTMIN'); w(10, -15); w(20, -20); w(30, 0)
    w(9, '$EXTMAX'); w(10, 90); w(20, 175); w(30, 0)
    w(0, 'ENDSEC')
    w(0, 'SECTION'); w(2, 'TABLES')
    w(0, 'TABLE'); w(2, 'LTYPE'); w(70, 2)
    w(0, 'LTYPE'); w(2, 'CONTINUOUS'); w(70, 0); w(3, 'Solid line'); w(72, 65); w(73, 0); w(40, 0.0)
    w(0, 'LTYPE'); w(2, 'DASHED'); w(70, 0); w(3, '__ __ __'); w(72, 65); w(73, 2); w(40, 0.75); w(49, 0.5)
    w(49, -0.25)
    w(0, 'ENDTAB')
    capas = sorted(ACI)
    w(0, 'TABLE'); w(2, 'LAYER'); w(70, len(capas))
    for c in capas:
        w(0, 'LAYER'); w(2, c); w(70, 0); w(62, ACI[c]); w(6, 'CONTINUOUS')
    w(0, 'ENDTAB')
    w(0, 'TABLE'); w(2, 'STYLE'); w(70, 1)
    w(0, 'STYLE'); w(2, 'STANDARD'); w(70, 0); w(40, 0.0); w(41, 1.0); w(50, 0.0); w(71, 0); w(42, 0.25)
    w(3, 'arial.ttf'); w(4, '')
    w(0, 'ENDTAB')
    w(0, 'ENDSEC')
    w(0, 'SECTION'); w(2, 'ENTITIES')
    def color(e):
        a = hex2aci(e.get('color', ''))
        if a: w(62, a)
    def poly(layer, pts, closed, e=None, ltype=None):
        w(0, 'POLYLINE'); w(8, layer)
        if e: color(e)
        if ltype: w(6, ltype)
        w(66, 1); w(10, 0); w(20, 0); w(30, 0); w(70, 1 if closed else 0)
        for x, y in pts:
            w(0, 'VERTEX'); w(8, layer); w(10, f'{x:.4f}'); w(20, f'{y:.4f}'); w(30, 0)
        w(0, 'SEQEND'); w(8, layer)
    def text(layer, p, s, hm, ha='center', va='center', rot=0):
        lines = s.split('\n')
        for i, l in enumerate(lines):
            yy = p[1] + (len(lines) - 1) / 2 * hm * 1.4 - i * hm * 1.4
            w(0, 'TEXT'); w(8, layer); w(10, f'{p[0]:.4f}'); w(20, f'{yy:.4f}'); w(30, 0); w(40, f'{hm:.3f}')
            w(1, l); w(50, rot)
            w(72, {'left': 0, 'center': 1, 'right': 2}[ha]); w(11, f'{p[0]:.4f}'); w(21, f'{yy:.4f}'); w(31, 0)
            w(73, {'bottom': 1, 'center': 2, 'baseline': 0, 'top': 3}[va])
    OFF = {'dist': ((0, 0), 300), 'alt': ((62, 11.5), 200), 'emp': ((0, 100), 400)}
    for sh, ((dx, dy), esc) in OFF.items():
        T = lambda p: (p[0] + dx, p[1] + dy)
        for e in K[sh]:
            k = e['k']
            if k == 'pl':
                lt = 'DASHED' if e['ls'] != '-' else None
                poly(e['layer'], [T(p) for p in e['pts']], e['closed'], e, lt)
            elif k == 'c':
                w(0, 'CIRCLE'); w(8, e['layer']); color(e)
                c = T(e['c']); w(10, f'{c[0]:.4f}'); w(20, f'{c[1]:.4f}'); w(30, 0); w(40, f"{e['r']:.4f}")
            elif k == 't':
                hm = e['pt'] * 0.3528 * 0.72 * esc / 1000.0   # altura de mayúscula equivalente
                text(e['layer'], T(e['p']), e['s'], hm, e['ha'], e['va'], e['rot'])
            elif k == 'dim':
                (x1, y1), (x2, y2), off = e['p1'], e['p2'], e['off']
                horiz = abs(y1 - y2) < 1e-9
                a, b = ((x1, y1 + off), (x2, y2 + off)) if horiz else ((x1 + off, y1), (x2 + off, y2))
                poly('COTAS', [T(a), T(b)], False)
                poly('COTAS', [T((x1, y1)), T(a)], False); poly('COTAS', [T((x2, y2)), T(b)], False)
                hm = e['pt'] * 0.3528 * 0.72 * esc / 1000.0
                tk = hm * 0.6
                for X in (a, b):
                    X = T(X); poly('COTAS', [(X[0] - tk, X[1] - tk), (X[0] + tk, X[1] + tk)], False)
                val = abs(x2 - x1) if horiz else abs(y2 - y1)
                m = T(((a[0] + b[0]) / 2, (a[1] + b[1]) / 2))
                if horiz:
                    text('COTAS', (m[0], m[1] + hm * 0.4), e['s'] or fmt(val, 2), hm, 'center', 'bottom')
                else:
                    text('COTAS', (m[0] - hm * 0.4, m[1]), e['s'] or fmt(val, 2), hm, 'center', 'bottom', 90)
    # títulos de vista en el DXF
    text('TEXTOS', (0, -12), f'PLANTA BAJA ({fmt(A_NAVE, 2)} m2)', 1.2, 'left')
    text('TEXTOS', (62, 28.5), 'PLANTA ALTILLO (300 m2)', 0.8, 'left')
    text('TEXTOS', (-8, 175), 'EMPLAZAMIENTO', 1.6, 'left')
    w(0, 'ENDSEC'); w(0, 'EOF')
    data = '\r\n'.join(out) + '\r\n'
    with open(ruta, 'wb') as f:
        f.write(data.replace('m²', 'm2').replace('−', '-').encode('cp1252'))

if __name__ == '__main__':
    os.makedirs(os.path.join(HERE, 'out'), exist_ok=True)
    construir()
    L, Lc, Le = longitudes()
    print('A nave', A_NAVE, 'P total', round(P_TOTAL, 2), 'lum', N_LUM_NAVE, 'emer', N_EMER)
    print({k: round(v, 1) for k, v in L.items()}); print({k: round(v, 1) for k, v in Lc.items()})
    print({k: round(v, 1) for k, v in Le.items()})
    with PdfPages(os.path.join(HERE, 'out', 'Planos_panaderia_A3.pdf')) as pdf:
        hoja_emplazamiento(pdf, L, Lc, Le)
        hoja_distribucion(pdf, L, Lc, Le)
        d = pdf.infodict(); d['Title'] = 'Planos de emplazamiento y distribución - Panadería industrial'
        d['Author'] = ', '.join(AUTORES)
    exportar_dxf(os.path.join(HERE, 'out', 'Plano_emplazamiento_y_distribucion.dxf'))
    print('ok')
