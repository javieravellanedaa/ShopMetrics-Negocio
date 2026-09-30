# -*- coding: utf-8 -*-
"""Aplica al presupuesto las correcciones de la clase del 29/09/2026 (video IMG_9216).

Cada bloque cita el minuto del video que lo pide. Se edita con openpyxl y despues
hay que recalcular y restaurar los objetos, como siempre:

    python3 scripts/correcciones_clase_29-09.py [archivo]
    python3 <skill>/recalc.py documento/Presupuesto\\ financiero\\ ShopMetrics.xlsx
    python3 scripts/restaurar_objetos.py documento/Presupuesto\\ financiero\\ ShopMetrics.xlsx
"""
from __future__ import annotations

import json
import os
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO = sys.argv[1] if len(sys.argv) > 1 else os.path.join(RAIZ, "documento", "Presupuesto financiero ShopMetrics.xlsx")
ALTAS = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "altas_29-09.json")))

PV_MESES = ["E", "G", "I", "K", "M", "O", "Q", "S", "U", "W", "Y", "AA"]     # Proy. ventas / Costos variables (cantidad)
RRHH_MESES = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]     # Costos RRHH
ANEXO_MESES = ["D", "F", "H", "J", "L", "N", "P", "R", "T", "V", "X", "Z"]    # Anexo (horas)

ROSA = PatternFill("solid", fgColor="FFF4CCCC")      # cubierta por otro puesto (leyenda de la plantilla)
AMARILLO = PatternFill("solid", fgColor="FFFFF2CC")  # tercerizada


def estilo(desde, hasta):
    hasta.font = copy(desde.font); hasta.fill = copy(desde.fill); hasta.border = copy(desde.border)
    hasta.alignment = copy(desde.alignment); hasta.number_format = desde.number_format


# ----------------------------------------------------------------- 1) Hipótesis
def hipotesis(h):
    """02:07-04:50: «yo tengo que entender qué número me están presentando»; que
    cada número del mercado diga qué es acá, no en el Word. 06:48-09:10: la
    cantidad de comercios sale de la facturación, no al revés."""
    h["B13"] = ("Qué es cada número: «Relevados» son los locales comerciales que el IDECBA contó en los 53 ejes "
                "comerciales de la Ciudad de Buenos Aires (tercer cuatrimestre de 2025); «Ocupados» son los que "
                "tenían actividad (el resto estaba vacío); «Indumentaria» son los ocupados del rubro indumentaria, "
                "textiles y calzado, el 24,3 %. Ese rubro es el segmento: se adopta 3.400 comercios como universo de "
                "trabajo. El mercado total es ese universo por el gasto anual de referencia (USD 960).")
    h["B16"] = ("Participación del 22% del segmento objetivo al cierre del tercer año, con una progresión de 3% el "
                "primer año, 12% el segundo y 22% el tercero. La progresión es moderada porque se trata de un "
                "emprendimiento que recién arranca y porque el segmento ya está acotado por la segmentación: son "
                "3.400 comercios de un solo rubro en los ejes relevados de la Ciudad de Buenos Aires. El salto del "
                "segundo al tercer año duplica la cartera, por lo que requiere un esfuerzo adicional de promoción "
                "que se refleja en el modelo de egresos.")
    # La cantidad de comercios es la que hace falta para esa facturación (comercios abonados al cierre).
    h["B24"] = "='Proy. ventas'!AA36&\" comercios\""
    h["B25"] = "='Proy. ventas'!AA99&\" comercios\""
    h["B26"] = "='Proy. ventas'!AA161&\" comercios\""
    h["B29"] = ("El valor del mercado captado es el objetivo de facturación de cada año: participación por mercado "
                "total. La proyección de ventas se arma para alcanzarlo (columna «Facturación proyectada», que "
                "sale de la lista de precios y de las altas mes a mes), y de ahí se obtiene la cantidad de "
                "comercios abonados que hace falta al cierre de cada año (columna «Cantidad»).")


# ------------------------------------------------------------ 2) Proy. ventas
def proyeccion(pv):
    """06:48-14:30 y 26:02: «yo sé que me tiene que dar 105, y vos me diste 47».
    La facturación de cada año tiene que dar el objetivo de participación; la
    cantidad de altas se deduce de ahí. 27:04: la progresión mensual va sumando
    suscripciones. Los precios no cambian (23:00: sólo si el análisis de precios
    del punto 6.2 lo justificara)."""
    filas = {"2026": {"basico": 19, "vidriera": 20, "cadena": 21},
             "2027": {"basico": 82, "vidriera": 83, "cadena": 84},
             "2028": {"basico": 144, "vidriera": 145, "cadena": 146}}
    for anio, planes in ALTAS.items():
        for plan, altas in planes.items():
            f = filas[anio][plan]
            for col, cant in zip(PV_MESES, altas):
                pv["%s%d" % (col, f)] = int(cant)
    pv["G6"] = ("La proyección se arma desde la facturación objetivo, que es el valor del mercado captado de la "
                "hoja Hipótesis (participación por mercado total: USD 105.754 en 2026, 385.805 en 2027 y "
                "705.677 en 2028). Con la lista de precios fija, se calcula cuántas altas por mes hacen falta "
                "para llegar a ese número: el alta e instalación se cobra una sola vez y el abono es mensual y "
                "se acumula (el comercio que se da de alta en marzo sigue pagando en abril), de modo que el "
                "ingreso crece mes a mes aunque las altas sean parejas. El primer trimestre de 2026 tiene pocas "
                "altas porque se destina a los primeros comercios del eje Avellaneda; a partir de allí la "
                "captación crece de manera sostenida. Los porcentajes de cada mes son la proporción del total "
                "anual que se factura ese mes.")
    # Capacidad operativa (bloques de comentarios de la misma hoja): estructura plana.
    pv["B41"] = ("* Se arranca sólo con los dos fundadores (gerencia general y gerencia de sistemas), que cubren "
                 "entre ambos todos los puestos de la estructura. La instalación y el soporte se contratan por "
                 "hora a técnicos freelance según las altas de cada mes, la venta se paga por comisión sobre cada "
                 "alta y el desarrollo, la contabilidad y el marketing están tercerizados en los costos fijos.")
    for f, texto in ((46, "Tercerizado (costo fijo)"), (47, "Freelance por hora (costo variable)"), (48, "Comisión por alta (costo variable)")):
        for col in PV_MESES:
            pv["%s%d" % (col, f)] = texto if col == "E" else "Sin cambios"
    for f in (108, 109, 110, 170, 171, 172):
        for col in PV_MESES:
            pv["%s%d" % (col, f)] = "Sin cambios"
    pv["B104"] = "* Sigue la estructura plana: los dos fundadores y todo lo demás tercerizado o por hora."
    for celda in ("B166", "B167", "C166"):
        v = pv[celda].value
        if isinstance(v, str) and v.startswith("*"):
            pv[celda] = pv["B104"].value
            break
    # Infraestructura: sólo el equipamiento de los dos fundadores, que ya está en el año cero.
    pv["E63"] = "Sin cambios"; pv["E64"] = "Sin cambios"; pv["E65"] = "Sin cambios"; pv["Q63"] = "Sin cambios"


# -------------------------------------------------------------- 3) Costos RRHH
def costos_rrhh(r):
    """15:12-17:30 y 29:14-35:00: «¿por qué necesito 5 personas? no lo pueden hacer los
    dos socios?»; estructura lo más plana posible, tercerizar; 34:20: el sueldo anual
    complementario se paga en junio y diciembre; 35:00: el despliegue de todos los
    puestos de toda área, marcando quién los cubre."""
    # Sólo los fundadores tienen dotación: técnico y vendedor pasan a costos variables.
    for f in (28, 29, 30, 46, 47, 48, 64, 65, 66):
        for col in RRHH_MESES:
            r["%s%d" % (col, f)] = 0
    # SAC: media cuota más en junio y diciembre, como en la plantilla (Junio + SAC).
    for encabezado in (33, 51, 69):
        r["H%d" % encabezado] = "Junio + SAC"
        r["N%d" % encabezado] = "Diciembre + SAC"
    for f in list(range(34, 39)) + list(range(52, 57)) + list(range(70, 75)):
        for col in ("H", "N"):
            v = r["%s%d" % (col, f)].value
            if isinstance(v, str) and v.startswith("=") and "*1.5" not in v:
                r["%s%d" % (col, f)] = v + "*1.5"
    # Leyenda de la plantilla.
    r["B9"] = "Posición cubierta por otro puesto"; r["B9"].fill = ROSA
    r["B10"] = "Posición tercerizada"; r["B10"].fill = AMARILLO
    r["M18"] = ("Los dos primeros años los fundadores cobran menos que el valor de su puesto; la diferencia es "
                "su aporte al emprendimiento. Son las dos únicas personas en relación de dependencia: el resto "
                "de los puestos de la estructura se cubre entre ambos o se terceriza (ver despliegue a la derecha).")
    # El despliegue de todos los puestos, por área, con quién lo cubre y dónde está el costo.
    puestos = [
        ("Gerencia", "Gerente general", "Fundador (CEO)", "Costos de RRHH"),
        ("Gerencia", "Gerente de sistemas", "Fundador (CTO)", "Costos de RRHH"),
        ("Administración", "Administrativo contable y de pagos", "Cubierto por el CEO", "—"),
        ("Administración", "Liquidación de sueldos e impuestos", "Tercerizado: estudio contable", "Costos fijos"),
        ("Administración", "Administrativo de cobranzas", "Cubierto por el CEO", "—"),
        ("Comercial", "Ejecutivo de ventas", "Freelance por comisión sobre cada alta", "Costos variables"),
        ("Comercial", "Atención al cliente y posventa", "Cubierto por el CEO", "—"),
        ("Marketing", "Gestión de campañas y contenidos", "Tercerizado: agencia de marketing", "Costos fijos"),
        ("Sistemas", "Arquitecto / desarrollador backend", "Tercerizado: mantenimiento evolutivo", "Costos fijos"),
        ("Sistemas", "Desarrollador frontend", "Tercerizado: mantenimiento evolutivo", "Costos fijos"),
        ("Sistemas", "Científico de datos (modelos)", "Cubierto por el CTO", "—"),
        ("Sistemas", "DevOps / infraestructura cloud", "Cubierto por el CTO", "—"),
        ("Servicio técnico", "Técnico de instalación", "Freelance por hora, según altas del mes", "Costos variables"),
        ("Servicio técnico", "Soporte técnico al comercio", "Freelance por hora, según cartera", "Costos variables"),
        ("Servicio técnico", "Operador de monitoreo de sensores", "Cubierto por el CTO", "—"),
    ]
    r["Q13"] = "Despliegue de todos los puestos de la estructura y quién los cubre"
    r["Q13"].font = Font(bold=True)
    for col, t in zip("QRST", ("Área", "Puesto", "Quién lo cubre", "Dónde está el costo")):
        c = r["%s14" % col]; c.value = t; estilo(r["B14"], c)
    for i, (area, puesto, quien, donde) in enumerate(puestos, start=15):
        for col, v in zip("QRST", (area, puesto, quien, donde)):
            c = r["%s%d" % (col, i)]; c.value = v; estilo(r["C16"], c)
            c.fill = ROSA if quien.startswith("Cubierto") else AMARILLO if not quien.startswith("Fundador") else copy(r["C16"].fill)
    for col, ancho in (("Q", 16), ("R", 36), ("S", 40), ("T", 20)):
        r.column_dimensions[col].width = ancho


# ---------------------------------------------------- 4) Anexo capacidad operativa
def anexo(an):
    """Las horas de instalación y soporte las cubren técnicos freelance contratados
    por mes entero según lo que pide la cartera (29:50: «todo lo que pueda tercerizar
    lo voy a tercerizar»)."""
    for fila in (87, 98, 109):
        an["A%d" % fila] = "Horas contratadas a técnicos freelance (por mes completo de %d h, a demanda)" % an["B14"].value
        req = fila - 1
        for col in ANEXO_MESES:
            an["%s%d" % (col, fila)] = "=ROUNDUP(%s%d/$B$14,0)*$B$14" % (col, req)


# ------------------------------------------------------------ 5) Costos variables
def costos_variables(cv):
    """35:21: los insumos IoT son variables y se calculan por instalación; 33:01: el
    freelance «lo voy pagando según las empresas que vaya» va a costos variables."""
    for f, f_anexo in ((27, 87), (50, 98), (73, 109)):
        for col in range(1, 29):
            estilo(cv.cell(f - 1, col), cv.cell(f, col))
        cv["A%d" % f] = "Operación"
        cv["B%d" % f] = "Técnico freelance de instalación y soporte (por hora, según el Anexo de capacidad)"
        cv["C%d" % f] = 7
        for col_cv, col_an in zip(PV_MESES, ANEXO_MESES):
            cant = openpyxl.utils.get_column_letter(openpyxl.utils.column_index_from_string(col_cv) - 1)
            cv["%s%d" % (cant, f)] = "='Anexo capacidad operativa'!%s%d" % (col_an, f_anexo)
            cv["%s%d" % (col_cv, f)] = "=%s%d*$C%d" % (cant, f, f)
        cv["AB%d" % f] = "=" + "+".join("%s%d" % (c, f) for c in PV_MESES)
    for f in (15, 38, 61):
        cv["B%d" % f] = "Comisión del vendedor freelance por alta concretada"
        cv["C%d" % f] = 40
    for f in (25, 48, 71):
        cv["B%d" % f] = "Insumos IoT por instalación - Plan Vidriera (sensor de vidriera, contador de puerta, gateway, montaje)"
    for f in (26, 49, 72):
        cv["B%d" % f] = "Insumos IoT por instalación - Plan Cadena (cinco locales)"


# -------------------------------------------------------------- 6) Mod. inversión
# ------------------------------------------------------------------ 6) Costos fijos
def costos_fijos(cf):
    """24:19-25:15: con una facturación pretendida alta, el esfuerzo de promoción
    tiene que acompañar: «por arriba del 10, 15 %» de las ventas. Se refuerzan las
    campañas de 2027, el año en que la cartera se duplica."""
    for fila, extra in ((41, 1000), (60, 2000)):   # Campañas base en Meta y Google, 2027 y 2028
        for col in range(4, 16):                   # D..O, los doce meses
            c = cf.cell(fila, col)
            if isinstance(c.value, (int, float)):
                c.value = c.value + extra


def modelo_inversion(mi):
    """40:18: «háganmela fácil: dejen solamente lo que entra en cada uno de los años y
    cuánta inversión, porque están [las filas en] cero». Con la estructura plana no
    hay altas de equipamiento en 2026 ni 2027; en 2028 se renuevan las dos notebooks
    (amortizadas a los tres años)."""
    for bloque in (32, 50):
        for f in range(bloque, bloque + 6):
            for col in "BCDE":
                mi["%s%d" % (col, f)] = None
        mi["B%d" % bloque] = "No se prevén inversiones en el año: la estructura es la de los dos fundadores, equipados en el año cero."
    for f in (68, 70, 71, 72, 73):
        for col in "BCDE":
            mi["%s%d" % (col, f)] = None
    mi["B69"] = "Notebooks (renovación a los 3 años)"
    mi["C69"] = 2; mi["D69"] = 940; mi["E69"] = "=C69*D69"
    # y ninguna fila vacía con «$ 0,00»: el total suma el rango y tolera blancos
    for bloque, total in ((32, 45), (50, 63), (68, 81)):
        for f in range(bloque, total):
            if mi["B%d" % f].value is None:
                mi["E%d" % f] = None


def main() -> int:
    wb = openpyxl.load_workbook(ARCHIVO)
    hipotesis(wb["Hipótesis"])
    proyeccion(wb["Proy. ventas"])
    costos_rrhh(wb["Costos RRHH"])
    anexo(wb["Anexo capacidad operativa"])
    costos_variables(wb["Costos variables"])
    costos_fijos(wb["Costos fijos"])
    modelo_inversion(wb["Mod. inversión"])
    wb.save(ARCHIVO)
    print("escrito:", ARCHIVO)
    return 0


if __name__ == "__main__":
    sys.exit(main())
