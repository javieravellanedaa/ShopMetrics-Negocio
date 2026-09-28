# -*- coding: utf-8 -*-
"""Cierra el presupuesto financiero para el segundo avance.

Hace cuatro cosas sobre el archivo, en este orden, y deja escrito por que:

1. Devuelve el formato de la plantilla a las tres hojas que se habian armado
   con formato propio --Mod. inversion, Amortizaciones y Presupuesto
   financiero--: se clonan celda por celda desde el ejemplo de la catedra
   (fuentes, rellenos, bordes, anchos, combinaciones) y sobre ese formato se
   escribe el contenido de ShopMetrics. El Anexo vuelve a la disposicion
   exacta de la plantilla, tomada del archivo original del 22/09, y conserva
   las dos filas de capacidad y holgura debajo de cada anio, que es lo que el
   profesor pidio ver el 31/08.

2. Aplica el modelo que hace rentable el negocio, tal como se decidio despues
   de correr los escenarios sobre copias:
     - precios: altas 40/120/450 y abonos 33/65/175 (Basico/Vidriera/Cadena).
       El alta cubre el kit que se instala; el abono se justifica contra la
       competencia (6.2.2 del informe) y un gasto de referencia del comercio
       de USD 960 al anio en software de gestion, facturacion y analitica.
     - volumen: las altas de 2027 se multiplican por 1,5 y las de 2028 por
       1,25, y despues todas por 1,11 (ajuste final para VAN positivo), con la promocion que lo sostiene (mas 1.500 y 2.000 por mes en
       campanias). Se llega al 19% del mercado meta a fin de 2028.
     - estructura: 1 vendedor en 2026, 2 en 2027, 2 y luego 3 en 2028; el
       segundo tecnico entra en agosto de 2027, cuando el anexo lo pide; el desarrollador sale de
       la dotacion y el mantenimiento se contrata afuera (20/40/60 h por mes);
       los fundadores cobran 1.200 brutos en 2026 y 2027 y 2.000 desde 2028.
     - costo del hardware: cada alta Vidriera lleva USD 34 de kit y cada alta
       Cadena USD 170. El anexo lo calculaba y costos variables no lo tomaba.

3. Carga el modelo de inversion: el desarrollo de la plataforma (85.000, del
   cronograma en Project) y el equipamiento, con la referencia de precio al
   lado, como en el ejemplo. Y las amortizaciones que se derivan.

4. Carga la tasa de corte, 30% en dolares.

    python3 scripts/cerrar_presupuesto.py
    python3 <skill>/recalc.py documento/Presupuesto\\ financiero\\ ShopMetrics.xlsx
    python3 scripts/restaurar_objetos.py documento/Presupuesto\\ financiero\\ ShopMetrics.xlsx

Los tres pasos van siempre en ese orden: openpyxl escribe, LibreOffice
recalcula, y restaurar_objetos devuelve lo que openpyxl borra.
"""
from __future__ import annotations

import os
import re
import sys
from copy import copy

import openpyxl
from openpyxl.utils import get_column_letter as L

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO = os.path.join(RAIZ, "documento", "Presupuesto financiero ShopMetrics.xlsx")
EJEMPLO = os.path.join(RAIZ, "documento", "Presupuesto financiero EJEMPLO V1.xlsx")
ORIGINAL = os.path.join(RAIZ, "scripts", "objetos-plantilla", "original-22-09.xlsx")

MESES = list(range(3, 15))                 # C..N en Costos RRHH y Costos fijos (D..O)
PV_MESES = list(range(5, 28, 2))           # E,G,...,AA en Proy. ventas y Costos variables


# ------------------------------------------------------------------ utilidades
def clonar(src, dst):
    """Copia una hoja entera --valores, formulas y formato-- sobre otra vacia."""
    for k, d in src.column_dimensions.items():
        dst.column_dimensions[k].width = d.width
        dst.column_dimensions[k].hidden = d.hidden
    for k, d in src.row_dimensions.items():
        dst.row_dimensions[k].height = d.height
    for row in src.iter_rows():
        for c in row:
            n = dst.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                n.font = copy(c.font); n.fill = copy(c.fill); n.border = copy(c.border)
                n.alignment = copy(c.alignment); n.number_format = c.number_format
                n.protection = copy(c.protection)
    for r in src.merged_cells.ranges:
        dst.merge_cells(str(r))
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines
    dst.sheet_properties.tabColor = src.sheet_properties.tabColor
    dst.freeze_panes = src.freeze_panes


def rehacer(wb, nombre, desde):
    """Reemplaza la hoja `nombre` por un clon de `desde`, en la misma posicion."""
    idx = wb.sheetnames.index(nombre)
    wb.remove(wb[nombre])
    ws = wb.create_sheet(nombre, idx)
    clonar(desde, ws)
    return ws


def estilo(desde, hasta):
    hasta.font = copy(desde.font); hasta.fill = copy(desde.fill); hasta.border = copy(desde.border)
    hasta.alignment = copy(desde.alignment); hasta.number_format = desde.number_format


def descombinar(ws, filas, cols=None):
    """Descombina toda combinacion que toque esas filas (y columnas, si se dan)."""
    for r in list(ws.merged_cells.ranges):
        if r.max_row >= min(filas) and r.min_row <= max(filas) and \
                (cols is None or (r.max_col >= min(cols) and r.min_col <= max(cols))):
            ws.unmerge_cells(str(r))


REF = re.compile(r"(?<![!A-Za-z0-9_\.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d\(])")


def desplazar_formulas(ws, desde_fila, cuanto=1):
    """Corre una fila hacia abajo toda referencia interna a filas >= desde_fila.
    No toca referencias a otras hojas (van precedidas de '!')."""
    def sub(m):
        fila = int(m.group(4))
        if fila >= desde_fila:
            fila += cuanto
        return "%s%s%s%d" % (m.group(1), m.group(2), m.group(3), fila)
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = REF.sub(sub, c.value)


def insertar_fila(ws, en):
    """Inserta una fila vacia en `en` corriendo todo lo de abajo, incluidas las
    combinaciones, las alturas y las formulas internas. openpyxl trae
    insert_rows pero no actualiza formulas ni combinaciones, asi que no sirve."""
    max_r = ws.max_row
    merges = [str(r) for r in ws.merged_cells.ranges]
    for r in merges:
        ws.unmerge_cells(r)
    for r in range(max_r, en - 1, -1):
        for col in range(1, ws.max_column + 1):
            s = ws.cell(row=r, column=col); d = ws.cell(row=r + 1, column=col)
            d.value = s.value
            if s.has_style: estilo(s, d)
            else: d.style = "Normal"
        if r in ws.row_dimensions and ws.row_dimensions[r].height:
            ws.row_dimensions[r + 1].height = ws.row_dimensions[r].height
    for col in range(1, ws.max_column + 1):
        c = ws.cell(row=en, column=col); c.value = None
    desplazar_formulas(ws, en)
    for r in merges:
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", r)
        a, b = int(m.group(2)), int(m.group(4))
        if a >= en: a += 1
        if b >= en: b += 1
        ws.merge_cells("%s%d:%s%d" % (m.group(1), a, m.group(3), b))


# ------------------------------------------------------------ 1) el formato
def formato_de_plantilla(wb, ej, orig):
    mi = rehacer(wb, "Mod. inversión", ej["Mod. inversión"])
    am = rehacer(wb, "Amortizaciones", ej["Amortizaciones"])
    pf = rehacer(wb, "Presupuesto financiero", ej["Presupuesto financiero"])
    an = rehacer(wb, "Anexo capacidad operativa", orig["Anexo capacidad operativa"])
    return mi, am, pf, an


# --------------------------------------------------------- 2) el contenido
CONCEPTOS = [   # (nombre, cantidad por bloque [anio0, 2026, 2027, 2028], precio, referencia)
    ("Desarrollo de la plataforma", [1, 0, 0, 0], 85000,
     "Cronograma en Microsoft Project: cuatro roles contratados, 2.324 h (ver informe de recursos abajo)"),
    ("Notebooks", [2, 2, 2, 1], 940,
     "Acer Aspire Go 15, Ryzen 7, 16 GB / 512 GB. Mercado Libre 9/2026: $1.450.000 a $1.545 por dólar"),
    ("Escritorios", [2, 2, 2, 1], 100,
     "Escritorio de oficina 120×60. Mercado Libre 9/2026, en el orden de $150.000"),
    ("Sillas de escritorio", [2, 2, 2, 1], 85,
     "Silla ergonómica de malla. Mercado Libre 9/2026, entre $90.000 y $156.000"),
    ("Equipamiento de red y periféricos", [1, 0, 0, 0], 205,
     "Router TP-Link Archer AX23 ($96.190) + switch 8 puertos + UPS 650 VA"),
    ("Registración de marca y constitución legal", [1, 0, 0, 0], 370,
     "INPI, 2 clases con honorarios (~$170.000) + constitución de SAS (~$400.000)"),
]
BLOQUES_INV = [(11, 28), (32, 45), (50, 63), (68, 81)]     # primera fila de items, fila del total


def modelo_inversion(mi):
    for f0, ftot in BLOQUES_INV:
        descombinar(mi, range(f0, f0 + 10), range(2, 7))
    for (f0, ftot), i_bloque in zip(BLOQUES_INV, range(4)):
        for k in range(10):                                  # las diez filas del ejemplo
            r = f0 + k
            for col in "BCDF":
                mi["%s%d" % (col, r)].value = None
            mi["E%d" % r].value = "=C%d*D%d" % (r, r)
        for k, (nombre, cant, precio, ref) in enumerate(CONCEPTOS):
            r = f0 + k
            mi["B%d" % r] = nombre
            mi["C%d" % r] = cant[i_bloque]
            mi["D%d" % r] = precio
            if i_bloque == 0:
                mi["F%d" % r] = ref
        mi["E%d" % ftot] = "=SUM(E%d:E%d)" % (f0, ftot - 1)
    for col, ftot in zip("GHIJ", (28, 45, 63, 81)):
        mi["%s5" % col] = "=$E$%d" % ftot
    mi["B4"], mi["C4"], mi["D4"] = 2026, 2027, 2028
    mi["H4"], mi["I4"], mi["J4"] = 2026, 2027, 2028


# fila -> (rubro, concepto, fila en Mod. inversion, anio de compra 0..3, vida util)
BIENES = [
    ("Informática y comunicaciones", "Notebooks", 12, 0, 3), (None, "Notebooks", 33, 1, 3),
    (None, "Notebooks", 51, 2, 3), (None, "Notebooks", 69, 3, 3),
    ("Muebles de oficina", "Escritorios", 13, 0, 10), (None, "Escritorios", 34, 1, 10),
    (None, "Escritorios", 52, 2, 10), (None, "Escritorios", 70, 3, 10),
    (None, "Sillas de escritorio", 14, 0, 10), (None, "Sillas de escritorio", 35, 1, 10),
    (None, "Sillas de escritorio", 53, 2, 10), (None, "Sillas de escritorio", 71, 3, 10),
    ("Equipamiento de red", "Red y periféricos", 15, 0, 3), (None, "Red y periféricos", 36, 1, 3),
]
COL_ADQ = {0: "G", 1: "H", 2: "I", 3: "J"}


def amortizaciones(am):
    descombinar(am, range(11, 27))
    for r in range(11, 27):
        for col in "BCDEFGHIJKLM":
            am["%s%d" % (col, r)].value = None
    for k, (rubro, concepto, fila_inv, anio, vida) in enumerate(BIENES):
        r = 11 + k
        am["B%d" % r] = rubro
        am["C%d" % r] = concepto
        am["D%d" % r] = "='Mod. inversión'!C%d" % fila_inv
        am["E%d" % r] = "='Mod. inversión'!D%d" % fila_inv
        am["F%d" % r] = vida
        for a in range(4):
            am["%s%d" % (COL_ADQ[a], r)] = "=D%d*E%d" % (r, r) if a == anio else 0
        for j, col in enumerate("KLM"):                       # cuotas 2026, 2027, 2028
            cuota = j + 1
            activo = anio <= cuota and (cuota - anio) < vida
            am["%s%d" % (col, r)] = "=$%s$%d/$F$%d" % (COL_ADQ[anio], r, r) if activo else 0
    for col in "KLM":
        am["%s27" % col] = "=SUM(%s11:%s26)" % (col, col)
    for a, b in ((11, 14), (15, 22), (23, 26)):                # los rubros, combinados como en la plantilla
        am.merge_cells("B%d:B%d" % (a, b))
    am["B4"], am["C4"], am["D4"] = 2026, 2027, 2028
    am["I4"], am["J4"], am["K4"] = 2026, 2027, 2028
    am["C6"], am["D6"] = "=Hipótesis!$D$25", "=Hipótesis!$D$26"


def presupuesto(pf):
    descombinar(pf, range(19, 36), range(2, 8))
    pf["E19"] = 0                                     # primer anio: no hay saldo anterior
    pf["F19"] = "=MAX(0,$L$23)*0.35"                  # 2027 paga sobre el imponible de 2026
    pf["G19"] = "=MAX(0,$M$23)*0.35"                  # 2028 sobre el de 2027
    pf["G33"] = 0.30                                  # tasa de corte en dolares
    pf["G34"] = "=NPV($G$33,E22:G22)+D22"
    pf["G35"] = "=IRR(D22:G22,$G$33)"
    for r in range(24, 30):
        v = pf["B%d" % r].value
        if isinstance(v, str):
            pf["B%d" % r] = v.replace("relacionadoc", "relacionados").replace("relacionado con", "relacionados con")


def anexo(an, costos_rrhh="Costos RRHH"):
    """Sobre la disposicion de la plantilla, agrega debajo de cada anio la
    capacidad de la dotacion y la holgura, mes a mes."""
    HORAS = [4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26]
    descombinar(an, [87, 88, 98, 99, 109, 112, 113])
    # la conclusion baja dos filas para hacer lugar al bloque de 2028
    merges = [str(r) for r in an.merged_cells.ranges if r.min_row in (110, 111)]
    spans = {}
    for r in merges:
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", r); spans[int(m.group(2))] = (m.group(1), m.group(3))
        an.unmerge_cells(r)
    for src, dst in ((111, 113), (110, 112)):
        s, d = an.cell(row=src, column=1), an.cell(row=dst, column=1)
        d.value = s.value; estilo(s, d); s.value = None
        if src in spans:
            an.merge_cells("%s%d:%s%d" % (spans[src][0], dst, spans[src][1], dst))
        an.row_dimensions[dst].height = an.row_dimensions[src].height
        an.row_dimensions[src].height = None
    for ftot, fdot, f1, f2 in ((86, 29, 87, 88), (97, 47, 98, 99), (108, 65, 109, 110)):
        a1 = an.cell(row=f1, column=1, value="Capacidad de los técnicos en dotación (h)")
        a2 = an.cell(row=f2, column=1, value="Holgura (h) — negativo = capacidad excedida")
        estilo(an.cell(row=ftot, column=1), a1); estilo(an.cell(row=ftot, column=1), a2)
        for i, col in enumerate(HORAS):
            c1 = an.cell(row=f1, column=col, value="='%s'!%s%d*$B$14" % (costos_rrhh, L(3 + i), fdot))
            c2 = an.cell(row=f2, column=col, value="=%s%d-%s%d" % (L(col), f1, L(col), ftot))
            estilo(an.cell(row=ftot, column=col), c1); estilo(an.cell(row=ftot, column=col), c2)
        b1 = an.cell(row=f1, column=2, value="=" + "+".join("%s%d" % (L(col), f1) for col in HORAS))
        b2 = an.cell(row=f2, column=2, value="=B%d-B%d" % (f1, ftot))
        estilo(an.cell(row=ftot, column=2), b1); estilo(an.cell(row=ftot, column=2), b2)


def hipotesis(h, comercios):
    for i, v in zip((58, 59, 60, 61, 62, 63), (40, 120, 450, 33, 65, 175)):
        h["C%d" % i] = v
    h["C12"] = 960
    for f, (n, share) in zip((24, 25, 26), comercios):
        h["B%d" % f] = "%d comercios" % n
        h["C%d" % f] = round(share, 4)
    for ref in ("B29", "B31", "B13", "B16"):
        v = h[ref].value
        if isinstance(v, str) and ("480" in v or "15%" in v or "Fudo" in v):
            v = v.replace("USD 480", "USD 960").replace("480", "960").replace("15%", "19%").replace("15 %", "19 %")
            h[ref] = v
    h["B31"] = ("Detalle: el gasto promedio anual de referencia son USD 960 (USD 80 mensuales), lo que un "
                "comercio minorista destina a software de gestión, facturación electrónica y analítica. "
                "Referencias 2026: planes de punto de venta con facturación electrónica y stock multi-tienda "
                "en el orden de USD 79 mensuales, y el plan de entrada de Fudo, uno de los sistemas con los "
                "que la plataforma se integra, en USD 40 mensuales por sucursal.")


def proyeccion(pv):
    """Escala las altas mensuales de 2027 y 2028 y devuelve los comercios
    activos a diciembre de cada anio (altas acumuladas)."""
    tot = {}
    for anio, f0, esc in ((2026, 19, 1.11), (2027, 82, 1.5 * 1.11), (2028, 144, 1.25 * 1.11)):
        s = 0
        for f in (f0, f0 + 1, f0 + 2):
            for c in PV_MESES:
                v = pv.cell(row=f, column=c).value
                if isinstance(v, (int, float)):
                    v = int(round(v * esc)); pv.cell(row=f, column=c).value = v; s += v
        tot[anio] = s
    acum = [tot[2026], tot[2026] + tot[2027], tot[2026] + tot[2027] + tot[2028]]
    return [(n, n / 3400) for n in acum]


def costos_rrhh(r):
    for c in MESES:
        r.cell(row=30, column=c).value = 1                                  # vendedores 2026
        r.cell(row=48, column=c).value = 2                                  # 2027
        r.cell(row=66, column=c).value = 2 if c - 2 <= 6 else 3             # 2028: 2, y 3 desde julio
        r.cell(row=47, column=c).value = 1 if c - 2 <= 7 else 2             # tecnicos 2027: 2 desde agosto, cuando el anexo lo pide
        r.cell(row=65, column=c).value = 2                                  # 2028
        for f in (28, 46, 64):                                              # desarrollador: contratado
            r.cell(row=f, column=c).value = 0
    # sueldo de los fundadores por anio, en el area libre a la derecha de la tabla de puestos
    r["M13"] = "Sueldo bruto de los fundadores por año"
    estilo(r["B13"], r["M13"])
    for ref, txt in (("M14", "Año"), ("N14", "Sueldo bruto"), ("O14", "Costo mensual")):
        r[ref] = txt; estilo(r["C14"], r[ref])
    for i, (anio, bruto) in enumerate(((2026, 1200), (2027, 1200), (2028, "=$D$16"))):
        f = 15 + i
        r["M%d" % f] = anio; estilo(r["C16"], r["M%d" % f])
        r["N%d" % f] = bruto; estilo(r["D16"], r["N%d" % f])
        r["O%d" % f] = "=N%d*(1+$K$16/$D$16)" % f; estilo(r["L16"], r["O%d" % f])
    r["M18"] = "Los dos primeros años los fundadores cobran menos que el valor de su puesto; la diferencia es aporte de trabajo de los socios."
    estilo(r["B13"], r["M18"]); r["M18"].font = copy(r["B16"].font)
    for f_costo, f_dot, celda in ((34, 26, "$O$15"), (35, 27, "$O$15"), (52, 44, "$O$16"), (53, 45, "$O$16")):
        for c in MESES:
            r.cell(row=f_costo, column=c).value = "=%s%d*%s" % (L(c), f_dot, celda)


def costos_fijos(cf):
    """Una linea nueva por bloque para el mantenimiento contratado, y mas
    promocion en 2027 y 2028 para sostener el volumen."""
    for en, monto in ((60, 1800), (42, 1200), (25, 600)):     # de abajo hacia arriba
        insertar_fila(cf, en)
        estilo(cf.cell(row=en - 1, column=2), cf.cell(row=en, column=2))
        for col in range(3, 18):
            estilo(cf.cell(row=en - 1, column=col), cf.cell(row=en, column=col))
        cf.cell(row=en, column=3).value = "Desarrollo y mantenimiento tercerizado"
        for col in range(4, 16):
            cf.cell(row=en, column=col).value = monto
        cf.cell(row=en, column=16).value = "=SUM(D%d:O%d)" % (en, en)
        cf.cell(row=en, column=17).value = "Free lance, 20/40/60 h por mes según el año. Sin cargas patronales."
        tot = en + 1
        for col in range(4, 16):
            cf.cell(row=tot, column=col).value = "=SUM(%s%d:%s%d)" % (L(col), en - 12, L(col), en)
        cf.cell(row=tot, column=16).value = "=SUM(D%d:O%d)" % (tot, tot)
    # promocion: campanias base
    filas = [row[0].row for row in cf.iter_rows(min_col=3, max_col=3)
             if row[0].value and "Campañas" in str(row[0].value)]
    for f, extra in zip(filas, (0, 1500, 2000)):
        for col in range(4, 16):
            v = cf.cell(row=f, column=col).value
            if isinstance(v, (int, float)):
                cf.cell(row=f, column=col).value = v + extra


def costos_variables(cv):
    """El kit que se instala en el local es insumo (08:20 de la clase): va a
    costo variable por alta, no a inversion ni a amortizaciones."""
    for f_rep, f_kit in ((20, 25), (43, 48), (66, 71)):
        for k, (nombre, unit, termino) in enumerate((
                ("Kit de instalación Plan Vidriera (sensor, contador, gateway, montaje)", 34, 0),
                ("Kit de instalación Plan Cadena (cinco locales)", 170, 1))):
            r = f_kit + k
            cv["B%d" % r] = nombre
            cv["C%d" % r] = unit
            for col in range(1, 30):
                estilo(cv.cell(row=f_rep, column=col), cv.cell(row=r, column=col))
            for col in range(4, 28):
                f = cv.cell(row=f_rep, column=col).value
                if not isinstance(f, str): continue
                if col % 2 == 0:                           # cantidad: un solo termino de la suma
                    partes = f.lstrip("=").split("+")
                    cv.cell(row=r, column=col).value = "=" + partes[termino]
                else:                                      # costo
                    cv.cell(row=r, column=col).value = re.sub(r"(?<=[A-Z\$])%d(?!\d)" % f_rep, str(r), f)
            cv.cell(row=r, column=28).value = re.sub(r"(?<=[A-Z\$])%d(?!\d)" % f_rep, str(r), cv.cell(row=f_rep, column=28).value)
        tot = f_kit + 7
        for col in list(range(5, 28, 2)) + [28]:
            cv.cell(row=tot, column=col).value = "=SUM(%s%d:%s%d)" % (L(col), f_rep - 5, L(col), tot - 1)


def main() -> int:
    wb = openpyxl.load_workbook(ARCHIVO)
    ej = openpyxl.load_workbook(EJEMPLO)
    orig = openpyxl.load_workbook(ORIGINAL)

    mi, am, pf, an = formato_de_plantilla(wb, ej, orig)
    modelo_inversion(mi)
    amortizaciones(am)
    presupuesto(pf)
    anexo(an)

    comercios = proyeccion(wb["Proy. ventas"])
    hipotesis(wb["Hipótesis"], comercios)
    costos_rrhh(wb["Costos RRHH"])
    costos_fijos(wb["Costos fijos"])
    costos_variables(wb["Costos variables"])

    wb.save(ARCHIVO)
    print("escrito. comercios a diciembre: %s" % ", ".join("%d (%.1f%%)" % (n, 100 * s) for n, s in comercios))
    return 0


if __name__ == "__main__":
    sys.exit(main())
