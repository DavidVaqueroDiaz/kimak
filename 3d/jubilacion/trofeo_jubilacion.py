# -*- coding: utf-8 -*-
"""
Trofeo conmemorativo "Feliz jubilacion" - reconstruccion CAD parametrica
a partir de una fotografia de referencia.

Salidas (junto a este fichero):
    trofeo_jubilacion.step  -> ensamblaje con colores  (SolidWorks: Abrir > STEP)
    trofeo_jubilacion.stl   -> malla                    (impresion 3D / visor)
    trofeo_jubilacion.svg   -> vista previa vectorial

Cotas en milimetros. Edita el bloque PARAMETROS y vuelve a ejecutar:
    pip install cadquery && python trofeo_jubilacion.py
"""

import math
import os
import cadquery as cq
from cadquery import exporters

# ============================================================== PARAMETROS ==
# --- Peana -----------------------------------------------------------------
BASE_L      = 200.0     # largo   (X)
BASE_W      = 86.0      # fondo   (Y)
BASE_H      = 35.0      # alto    (Z)
BASE_SLOPE  = 8.0       # retranqueo de la cara frontal inclinada
TEXT_DEPTH  = 0.8       # profundidad del grabado
LINEA_1     = u"POR TODA UNA VIDA DEDICADA A LAS ÑAPAS"
LINEA_2     = u"¡FELIZ JUBILACIÓN!"
H_LINEA_1   = 7.2
H_LINEA_2   = 11.5
FONT        = "DejaVu Sans"     # cambia a "Arial" en Windows

# --- Colores (RGBA 0-1) ----------------------------------------------------
C_BASE   = (0.10, 0.10, 0.12, 1.0)
C_TEXTO  = (0.85, 0.60, 0.18, 1.0)
C_MONO   = (0.13, 0.35, 0.68, 1.0)
C_PIEL   = (0.94, 0.76, 0.64, 1.0)
C_PELO   = (0.22, 0.17, 0.14, 1.0)
C_NEGRO  = (0.08, 0.08, 0.09, 1.0)
C_METAL  = (0.72, 0.74, 0.77, 1.0)
C_ROJO   = (0.45, 0.16, 0.13, 1.0)
C_MADERA = (0.55, 0.30, 0.16, 1.0)
C_VERDE  = (0.11, 0.35, 0.24, 1.0)

FIG_Y = 7.0             # retranqueo de la figura hacia el fondo
Z0 = BASE_H             # cara superior de la peana


# =================================================================== PEANA ==
def peana():
    """Prisma con la cara frontal inclinada (perfil extruido en X)."""
    y0, y1 = -BASE_W / 2.0, BASE_W / 2.0
    pts = [(y0, 0.0), (y1, 0.0), (y1, BASE_H), (y0 + BASE_SLOPE, BASE_H)]
    s = (cq.Workplane("YZ").polyline(pts).close()
         .extrude(BASE_L)
         .translate((-BASE_L / 2.0, 0, 0)))
    return s.edges("|X and >Z").fillet(1.5)


def plano_frontal():
    """Plano de trabajo sobre la cara frontal inclinada."""
    ln = math.hypot(BASE_SLOPE, BASE_H)
    normal = (0.0, -BASE_H / ln, BASE_SLOPE / ln)
    origen = (0.0, -BASE_W / 2.0 + BASE_SLOPE / 2.0, BASE_H / 2.0)
    return cq.Plane(origin=origen, xDir=(1, 0, 0), normal=normal)


def texto_solidos():
    """Solidos del rotulo (rellenan exactamente el grabado de la peana)."""
    pl = plano_frontal()
    t1 = (cq.Workplane(pl).center(0, 3.0)
          .text(LINEA_1, H_LINEA_1, -TEXT_DEPTH, font=FONT, kind="bold",
                halign="center", valign="center", combine=False))
    t2 = (cq.Workplane(pl).center(0, -10.0)
          .text(LINEA_2, H_LINEA_2, -TEXT_DEPTH, font=FONT, kind="bold",
                halign="center", valign="center", combine=False))
    return t1.union(t2)


# ================================================================== FIGURA ==
def _loft(secciones):
    """secciones = [(z, semieje_x, semieje_y, cx, cy), ...] -> solido."""
    z0, a0, b0, cx0, cy0 = secciones[0]
    wp = cq.Workplane("XY").workplane(offset=z0).center(cx0, cy0).ellipse(a0, b0)
    zp, cxp, cyp = z0, cx0, cy0
    for z, a, b, cx, cy in secciones[1:]:
        wp = wp.workplane(offset=z - zp).center(cx - cxp, cy - cyp).ellipse(a, b)
        zp, cxp, cyp = z, cx, cy
    return wp.loft(ruled=False)


def pierna(signo):
    x = signo * 11.0
    return _loft([
        (Z0 + 13, 6.5, 7.5, x, -1.0),          # tobillo
        (Z0 + 55, 8.5, 9.5, x, 0.0),           # rodilla
        (Z0 + 78, 10.5, 11.0, x * 0.95, 0.5),
        (Z0 + 95, 12.0, 12.0, x * 0.80, 0.5),  # cadera
    ])


def torso():
    return _loft([
        (Z0 + 88, 25.0, 13.5, 0, 0.5),         # caderas
        (Z0 + 108, 23.0, 12.5, 0, 0.5),        # cintura
        (Z0 + 120, 25.0, 14.0, 0, 0.5),
        (Z0 + 134, 26.5, 14.8, 0, 0.5),        # pecho
        (Z0 + 146, 26.5, 14.0, 0, 0.5),        # hombros
        (Z0 + 152, 21.0, 11.5, 0, 0.5),
    ])


def brazo(signo):
    """Devuelve (manga_azul, mano_piel)."""
    L = 64.0
    manga = _loft([
        (0.0, 4.8, 4.8, 0, 0),                 # puno
        (L * 0.40, 5.8, 5.8, 0, 0),            # codo
        (L * 0.80, 7.6, 8.0, 0, 0),
        (L, 9.5, 9.8, 0, 0),                   # hombro (empotrado en el torso)
    ])
    mano = cq.Workplane("XY").sphere(5.4).translate((0, 0.5, -3.5))
    ang = -signo * 7.0
    manga = manga.rotate((0, 0, L), (0, 1, 0), ang)
    mano = mano.rotate((0, 0, L), (0, 1, 0), ang)
    off = (signo * 22.0, 1.0, Z0 + 86.0)
    return manga.translate(off), mano.translate(off)


def bota(signo):
    return (cq.Workplane("XY")
            .box(13.0, 27.0, 13.0, centered=(True, True, False))
            .edges("|Z").fillet(3.5).edges(">Z").fillet(1.5)
            .translate((signo * 11.0, -3.0, Z0)))


def cabeza():
    zc = Z0 + 174.0
    craneo = cq.Workplane("XY").sphere(12.0).translate((0, 0, zc))
    mandibula = (cq.Workplane("XY").workplane(offset=zc - 12.0)
                 .ellipse(8.5, 9.0).workplane(offset=8.0).ellipse(11.5, 11.8)
                 .loft())
    nariz = cq.Workplane("XY").sphere(2.6).translate((0, -11.0, zc - 1.0))
    orejas = (cq.Workplane("XY").sphere(2.8).translate((11.0, 1.0, zc + 0.5))
              .union(cq.Workplane("XY").sphere(2.8)
                     .translate((-11.0, 1.0, zc + 0.5))))
    cuello = (cq.Workplane("XY").workplane(offset=Z0 + 148)
              .circle(7.6).extrude(16.0))
    return craneo.union(mandibula).union(nariz).union(orejas).union(cuello)


def pelo():
    zc = Z0 + 174.0
    return (cq.Workplane("XY").sphere(12.4).translate((0, 0, zc))
            .intersect(cq.Workplane("XY")
                       .box(40, 40, 14, centered=(True, True, False))
                       .translate((0, 0, zc + 1.5))))


def camiseta():
    """Cuello de la camiseta bajo el mono."""
    return (cq.Workplane("XY").workplane(offset=Z0 + 143)
            .ellipse(9.0, 7.0).workplane(offset=7.0).ellipse(7.6, 6.0).loft())


def cinturon():
    return (cq.Workplane("XY").workplane(offset=Z0 + 104)
            .ellipse(24.2, 13.6).extrude(7.0))


def funda():
    """Funda porta-destornillador en la cadera derecha."""
    f = (cq.Workplane("XY").workplane(offset=Z0 + 92)
         .ellipse(5.0, 4.0).workplane(offset=18).ellipse(6.0, 4.6).loft()
         .translate((23.0, -9.0, 0)))
    d = (cq.Workplane("XY").workplane(offset=Z0 + 108)
         .circle(1.6).extrude(13.0).translate((23.0, -9.0, 0)))
    return f, d


def bolsillos():
    p = cq.Workplane("XY").box(13.0, 3.0, 10.0, centered=(True, True, False))
    return (p.translate((-11.0, -14.0, Z0 + 130))
            .union(p.translate((11.0, -14.0, Z0 + 130))))


# ================================================================== UTILES ==
# Cada util se construye apoyado en Z=0 y luego se coloca sobre la peana.

def caja_herramientas():
    """Devuelve (cuerpo, tapa+asa). Caja de 58 x 30 mm."""
    cuerpo = (cq.Workplane("XY").box(58, 30, 15, centered=(True, True, False))
              .edges("|Z").fillet(2.5))
    tapa = (cq.Workplane("XY").workplane(offset=15)
            .rect(58, 30).workplane(offset=7).rect(48, 22).loft()
            .edges(">Z").fillet(1.0))
    asa = (cq.Workplane("XZ").workplane(offset=-2.0)
           .center(0, 27.0).rect(26, 3).extrude(4.0)
           .union(cq.Workplane("XZ").workplane(offset=-2.0)
                  .center(-11.5, 24.0).rect(3, 6).extrude(4.0))
           .union(cq.Workplane("XZ").workplane(offset=-2.0)
                  .center(11.5, 24.0).rect(3, 6).extrude(4.0)))
    return cuerpo, tapa.union(asa)


def llave_inglesa():
    """Llave ajustable tumbada (largo 62, espesor 6)."""
    t = 6.0
    mango = (cq.Workplane("XY").center(-12, 0).rect(44, 9).extrude(t)
             .edges("|Z").fillet(3.0))
    cabeza_ = (cq.Workplane("XY").center(18, 1).rect(20, 20).extrude(t)
               .edges("|Z").fillet(4.5))
    mordaza = (cq.Workplane("XY").center(24, -6.5).rect(13, 6).extrude(t)
               .edges("|Z").fillet(1.5))
    boca = cq.Workplane("XY").center(24, 5).rect(18, 8).extrude(t)
    tornillo = cq.Workplane("XY").center(4, 0).circle(3.0).extrude(t)
    return mango.union(cabeza_).union(mordaza).cut(boca).cut(tornillo)


def llave_fija():
    """Llave fija de dos bocas (largo 46, espesor 5)."""
    t = 5.0
    cuerpo = cq.Workplane("XY").rect(30, 7).extrude(t).edges("|Z").fillet(2.0)
    c1 = cq.Workplane("XY").center(-15, 0).circle(7).extrude(t)
    c2 = cq.Workplane("XY").center(15, 0).circle(8).extrude(t)
    s = cuerpo.union(c1).union(c2)
    b1 = (cq.Workplane("XY").center(-18, 0).rect(10, 6).extrude(t)
          .union(cq.Workplane("XY").center(-15, 0).circle(3.4).extrude(t)))
    b2 = (cq.Workplane("XY").center(18.5, 0).rect(10, 7.5).extrude(t)
          .union(cq.Workplane("XY").center(15, 0).circle(4.2).extrude(t)))
    return s.cut(b1).cut(b2)


def serrucho():
    """Devuelve (hoja_dentada, mango). Longitud total ~92, hoja de 2.2 mm."""
    t = 2.2
    hoja = (cq.Workplane("XY")
            .polyline([(-28, 0), (34, 0), (34, 9), (-28, 15)]).close()
            .extrude(t))
    n, paso = 30, 2.05
    dientes = cq.Workplane("XY")
    for i in range(n):
        x = -27.5 + i * paso
        dientes = dientes.union(
            cq.Workplane("XY")
            .polyline([(x - paso / 2, -0.1), (x + paso / 2, -0.1), (x, 2.0)])
            .close().extrude(t))
    hoja = hoja.cut(dientes)
    mango = (cq.Workplane("XY").center(-38, 7).rect(22, 22).extrude(7.0)
             .edges("|Z").fillet(7.0)
             .cut(cq.Workplane("XY").center(-39, 8).rect(9, 11).extrude(7.0)
                  .edges("|Z").fillet(3.5)))
    return hoja, mango


def taladro():
    """Taladro tumbado. Devuelve (cuerpo+empunadura, portabrocas, broca)."""
    r = 11.0
    cuerpo = (cq.Workplane("YZ").workplane(offset=-22).circle(r).extrude(38)
              .edges().fillet(2.0).translate((0, 0, r)))
    empun = (cq.Workplane("XY").center(-13, 0).rect(15, 17).extrude(2 * r - 2)
             .edges("|Z").fillet(4.0).translate((0, 0, 1.0)))
    porta = (cq.Workplane("YZ").workplane(offset=16).circle(8)
             .workplane(offset=11).circle(6).loft().translate((0, 0, r)))
    broca = (cq.Workplane("YZ").workplane(offset=27).circle(2.0).extrude(20)
             .translate((0, 0, r)))
    return cuerpo.union(empun), porta, broca


# ============================================================== ENSAMBLAJE ==
def _poner(shape, x, y, giro=0.0):
    return shape.rotate((0, 0, 0), (0, 0, 1), giro).translate((x, y, Z0))


def construir():
    # --- peana + rotulo grabado
    base = peana()
    txt = texto_solidos()
    base = base.cut(txt)

    # --- figura
    mg_d, mn_d = brazo(1)
    mg_i, mn_i = brazo(-1)
    cam = camiseta()
    mono = (torso().union(pierna(1)).union(pierna(-1)).union(bolsillos())
            .union(mg_d).union(mg_i).cut(cam))
    piel = cabeza().union(mn_d).union(mn_i).union(cam)
    pelo_ = pelo()
    piel = piel.cut(pelo_)
    fnd, dst = funda()
    negro = bota(1).union(bota(-1)).union(cinturon()).union(fnd)
    mono = mono.translate((0, FIG_Y, 0))
    piel = piel.translate((0, FIG_Y, 0))
    pelo_ = pelo_.translate((0, FIG_Y, 0))
    negro = negro.translate((0, FIG_Y, 0))
    dst = dst.translate((0, FIG_Y, 0))

    # --- utiles sobre la peana
    caja_c, caja_t = caja_herramientas()
    caja_c = _poner(caja_c, -66, 20, 0)
    caja_t = _poner(caja_t, -66, 20, 0)

    lli = _poner(llave_inglesa(), -62, -24, 9)
    llf = _poner(llave_fija(), -10, -26, -5)

    hoja, mango = serrucho()
    hoja = _poner(hoja, 64, -25, -4)
    mango = _poner(mango, 64, -25, -4)

    tal_c, tal_p, tal_b = taladro()
    tal_c = _poner(tal_c, 50, 20, -8)
    tal_p = _poner(tal_p, 50, 20, -8)
    tal_b = _poner(tal_b, 50, 20, -8)

    asm = cq.Assembly(name="TrofeoJubilacion")
    asm.add(base,   name="peana",           color=cq.Color(*C_BASE))
    asm.add(txt,    name="rotulo",          color=cq.Color(*C_TEXTO))
    asm.add(mono,   name="mono_trabajo",    color=cq.Color(*C_MONO))
    asm.add(piel,   name="cabeza_y_manos",  color=cq.Color(*C_PIEL))
    asm.add(pelo_,  name="pelo",            color=cq.Color(*C_PELO))
    asm.add(negro,  name="botas_cinturon",  color=cq.Color(*C_NEGRO))
    asm.add(dst,    name="destornillador",  color=cq.Color(*C_METAL))
    asm.add(caja_c, name="caja_cuerpo",     color=cq.Color(*C_METAL))
    asm.add(caja_t, name="caja_tapa",       color=cq.Color(*C_ROJO))
    asm.add(lli,    name="llave_inglesa",   color=cq.Color(*C_METAL))
    asm.add(llf,    name="llave_fija",      color=cq.Color(*C_METAL))
    asm.add(hoja,   name="serrucho_hoja",   color=cq.Color(*C_METAL))
    asm.add(mango,  name="serrucho_mango",  color=cq.Color(*C_MADERA))
    asm.add(tal_c,  name="taladro_cuerpo",  color=cq.Color(*C_VERDE))
    asm.add(tal_p,  name="taladro_porta",   color=cq.Color(*C_METAL))
    asm.add(tal_b,  name="taladro_broca",   color=cq.Color(*C_METAL))
    return asm


if __name__ == "__main__":
    out = os.path.dirname(os.path.abspath(__file__))
    asm = construir()
    asm.export(os.path.join(out, "trofeo_jubilacion.step"))
    comp = asm.toCompound()
    exporters.export(comp, os.path.join(out, "trofeo_jubilacion.stl"),
                     tolerance=0.05, angularTolerance=0.2)
    exporters.export(comp, os.path.join(out, "trofeo_jubilacion.svg"),
                     opt={"width": 900, "height": 900,
                          "projectionDir": (0.6, -1.0, 0.45),
                          "showAxes": False, "strokeWidth": 0.35})
    bb = comp.BoundingBox()
    print("OK  largo=%.1f  fondo=%.1f  alto=%.1f mm" % (bb.xlen, bb.ylen, bb.zlen))
