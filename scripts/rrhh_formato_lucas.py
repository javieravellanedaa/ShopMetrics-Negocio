#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rehace la hoja 'Costos RRHH' con el formato del template, tal como lo
respeto Lucas Fraguaga en 'PresupuestoFinanciero - Lucas Fraguaga.xlsx'.

Lo que corrige respecto de la version anterior:

  1. El despliegue de TODOS los puestos de TODAS las areas va DENTRO del cuadro
     'Costo mensual por puesto', no en una tabla aparte a la derecha.  Cada
     puesto queda marcado con el color de la leyenda (rosa = cubierto por otro
     puesto, amarillo = tercerizado, sin relleno = en relacion de dependencia).
     Scali, 30:21: "aca el toco todo lo que yo le deje... todos los puestos,
     todas las areas"; 35:02: "quiero el despliegue de todos los puestos de
     toda area, esto que hizo el aca".

  2. Un unico bloque por anio, con Cantidad y Costo intercalados por mes y las
     columnas 'Junio + SAC' / 'Diciembre + SAC', como el template.  Antes
     estaba partido en dos bloques (Dotacion / Costo), que no es el formato.

  3. El SAC se calcula como el template: costo del mes + medio sueldo del mes
     anterior, en lugar de multiplicar por 1,5.
"""
import os
import sys
from copy import copy

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter as L

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO = os.path.join(RAIZ, "documento", "Presupuesto financiero ShopMetrics.xlsx")

ROSA = "FFF4CCCC"        # posicion cubierta por otro puesto
AMARILLO = "FFFFF2CC"    # posicion tercerizada

DEP, CUB, TER = "dep", "cub", "ter"

# (area, puesto, estado, sueldo bruto, quien lo cubre / donde esta el costo)
PUESTOS = [
    ("Gerencia", "Gerente General (Fundador / CEO)", DEP, 1200,
     "Sueldo de fundador; ver la nota al pie del cuadro"),
    ("Gerencia", "Gerente de Sistemas (Fundador / CTO)", DEP, 1200,
     "Sueldo de fundador; ver la nota al pie del cuadro"),
    ("Administración", "Administrativo contable y de pagos", CUB, 900,
     "Lo cubre el Gerente General"),
    ("Administración", "Liquidación de sueldos e impuestos", TER, 800,
     "Estudio contable, abono mensual en costos fijos"),
    ("Administración", "Administrativo de cobranzas", CUB, 850,
     "Lo cubre el Gerente General"),
    ("Comercial", "Ejecutivo comercial de altas", TER, 1000,
     "Freelance por comisión sobre cada alta, en costos variables"),
    ("Comercial", "Atención al cliente y posventa", CUB, 900,
     "Lo cubre el Gerente General"),
    ("Marketing", "Gestión de campañas y contenidos", TER, 1100,
     "Agencia de marketing, abono mensual en costos fijos"),
    ("Sistemas", "Arquitecto de software", TER, 2200,
     "Contratado por proyecto; el desarrollo inicial es inversión"),
    ("Sistemas", "Desarrollador", TER, 1800,
     "Mantenimiento evolutivo contratado, en costos fijos"),
    ("Sistemas", "Científico de datos", CUB, 2000,
     "Lo cubre el Gerente de Sistemas"),
    ("Sistemas", "DevOps / infraestructura cloud", CUB, 1900,
     "Lo cubre el Gerente de Sistemas"),
    ("Servicio técnico", "Técnico de instalación y soporte", TER, 1100,
     "Freelance por hora según el Anexo, en costos variables"),
    ("Servicio técnico", "Operador de monitoreo de sensores", CUB, 900,
     "Lo cubre el Gerente de Sistemas"),
]

# dotacion mes a mes por anio, solo para los puestos en relacion de dependencia
UNO = [1] * 12
DOTACION = {                       # solo los puestos en relacion de dependencia
    0: {2026: UNO, 2027: UNO, 2028: UNO},      # Gerente General  (fundador)
    1: {2026: UNO, 2027: UNO, 2028: UNO},      # Gerente de Sistemas (fundador)
}

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio + SAC", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre + SAC"]
ANIOS = (2026, 2027, 2028)

CANT = [3 + 2 * i for i in range(12)]        # C,E,G,I,K,M,O,Q,S,U,W,Y
COSTO = [4 + 2 * i for i in range(12)]       # D,F,H,J,L,N,P,R,T,V,X,Z
TOTAL_ANUAL = 27                             # AA

APORTES = [("Jubilación", 0.1047, "tasa"), ("Ley 19.032", 0.0154, "tasa"),
           ("Obra social", 0.06, "tasa"), ("F. Nac. Empleo", 0.0092, "tasa"),
           ("Seg. de Vida", 1.2, "fijo"), ("ART", 0.03, "tasa")]


def estilo(desde, hasta):
    hasta.font = copy(desde.font)
    hasta.fill = copy(desde.fill)
    hasta.border = copy(desde.border)
    hasta.alignment = copy(desde.alignment)
    hasta.number_format = desde.number_format


def main() -> int:
    wb = openpyxl.load_workbook(ARCHIVO)
    r = wb["Costos RRHH"]

    # --- modelos de formato, tomados de la hoja actual antes de limpiarla ----
    M = {}
    for nombre, ref in (("bloque", "B13"), ("encabezado", "C14"), ("area", "B16"),
                        ("puesto", "C16"), ("bruto", "D16"), ("aporte", "E16"),
                        ("alicuota", "E15"), ("rotulo", "B15"),
                        ("sumaport", "K16"), ("mensual", "L16"),
                        ("cantidad", "C26"), ("costo", "C34"),
                        ("tot_rot", "B39"), ("tot_num", "C39"),
                        ("leyenda", "B9"), ("mes", "C33")):
        c = r[ref]
        M[nombre] = {"font": copy(c.font), "fill": copy(c.fill), "border": copy(c.border),
                     "alignment": copy(c.alignment), "number_format": c.number_format}
    nota_font = copy(r["M18"].font) if r["M18"].value else copy(r["C16"].font)

    def poner(ref, valor, modelo, fill=None):
        c = r[ref] if isinstance(ref, str) else ref
        c.value = valor
        m = M[modelo]
        c.font = copy(m["font"]); c.border = copy(m["border"])
        c.alignment = copy(m["alignment"]); c.number_format = m["number_format"]
        c.fill = PatternFill("solid", fgColor=fill) if fill else copy(m["fill"])
        return c

    # --- limpieza: todo desde la fila 9 hacia abajo ------------------------
    for rango in list(r.merged_cells.ranges):
        if rango.max_row >= 9:
            r.unmerge_cells(str(rango))
    vacio = PatternFill()
    from openpyxl.styles import Border, Font, Alignment
    for fila in r.iter_rows(min_row=9, max_row=max(r.max_row, 100), max_col=max(r.max_column, 31)):
        for c in fila:
            c.value = None
            c.fill = copy(vacio); c.border = Border(); c.font = Font()
            c.alignment = Alignment(); c.number_format = "General"

    # --- leyenda ----------------------------------------------------------
    poner("B9", "Posición cubierta por otro puesto", "leyenda", ROSA)
    poner("B10", "Posición tercerizada", "leyenda", AMARILLO)
    r.merge_cells("B9:D9"); r.merge_cells("B10:D10")

    # --- cuadro de puestos -------------------------------------------------
    F_ENC = 13                      # encabezados
    F_ALI = 14                      # alicuotas patronales
    F0 = 15                         # primer puesto
    poner("B12", "Costo mensual por puesto (USD)", "bloque")
    r.merge_cells("B12:L12")

    cols = ["Área", "Puesto", "Sueldo bruto"] + [a[0] for a in APORTES] + \
           ["Total aportes", "Costo mensual para la empresa"]
    for i, t in enumerate(cols):
        poner(r.cell(row=F_ENC, column=2 + i), t, "encabezado")
    poner(r.cell(row=F_ENC, column=13), "Quién lo cubre y dónde se imputa el costo",
          "encabezado")
    r.merge_cells(start_row=F_ENC, start_column=13,
                  end_row=F_ENC, end_column=TOTAL_ANUAL)

    poner("B%d" % F_ALI, "Alícuotas patronales", "rotulo")
    for i, (_, v, tipo) in enumerate(APORTES):
        c = poner(r.cell(row=F_ALI, column=5 + i), v, "alicuota")
        c.number_format = '"$ "#,##0.00' if tipo == "fijo" else "0.00%"

    fila_de = {}
    area_previa = None
    for i, (area, puesto, estado, bruto, detalle) in enumerate(PUESTOS):
        f = F0 + i
        fila_de[i] = f
        relleno = {CUB: ROSA, TER: AMARILLO}.get(estado)
        poner("B%d" % f, area if area != area_previa else None, "area", relleno)
        area_previa = area
        poner("C%d" % f, puesto, "puesto", relleno)
        poner("D%d" % f, bruto, "bruto", relleno)
        for j, (_, _, tipo) in enumerate(APORTES):
            ref = "%s%d" % (L(5 + j), f)
            formula = "=$%s$%d" % (L(5 + j), F_ALI) if tipo == "fijo" else \
                      "=$D%d*$%s$%d" % (f, L(5 + j), F_ALI)
            poner(ref, formula, "aporte", relleno)
        poner("K%d" % f, "=SUM(E%d:J%d)" % (f, f), "sumaport", relleno)
        poner("L%d" % f, "=D%d+K%d" % (f, f), "mensual", relleno)
        if detalle:
            c = poner("M%d" % f, detalle, "puesto", None)
            c.font = copy(nota_font)
            r.merge_cells(start_row=f, start_column=13,
                          end_row=f, end_column=TOTAL_ANUAL)

    F_NOTA = F0 + len(PUESTOS) + 1
    c = poner("B%d" % F_NOTA, "Los dos fundadores cobran un sueldo de fundador de USD 1.200, plano los tres años y por debajo del valor de mercado de su puesto; la diferencia es aporte de trabajo de los socios. Las posiciones sin color están en relación de dependencia y son las únicas que generan costo en esta hoja: lo tercerizado se imputa en costos fijos o variables, según la columna de la derecha.", "puesto", None)
    c.font = copy(nota_font)
    r.merge_cells(start_row=F_NOTA, start_column=2,
                  end_row=F_NOTA, end_column=TOTAL_ANUAL)

    # --- un bloque por anio -------------------------------------------------
    primeros = {}
    totales = {}
    f = F_NOTA + 3
    for anio in ANIOS:
        poner("B%d" % f, "Costos de RRHH - Año %d" % anio, "bloque")
        r.merge_cells("B%d:%s%d" % (f, L(TOTAL_ANUAL), f))
        f_enc, f_sub, f_ini = f + 1, f + 2, f + 3
        primeros[anio] = f_ini

        poner("B%d" % f_enc, "Puesto", "encabezado")
        r.merge_cells("B%d:B%d" % (f_enc, f_sub))
        for i, mes in enumerate(MESES):
            poner(r.cell(row=f_enc, column=CANT[i]), mes, "mes")
            estilo(r.cell(row=f_enc, column=CANT[i]), r.cell(row=f_enc, column=COSTO[i]))
            r.merge_cells(start_row=f_enc, start_column=CANT[i],
                          end_row=f_enc, end_column=COSTO[i])
            poner(r.cell(row=f_sub, column=CANT[i]), "Cantidad", "encabezado")
            poner(r.cell(row=f_sub, column=COSTO[i]), "Costo", "encabezado")
        poner(r.cell(row=f_enc, column=TOTAL_ANUAL), "Total anual", "mes")
        r.merge_cells(start_row=f_enc, start_column=TOTAL_ANUAL,
                      end_row=f_sub, end_column=TOTAL_ANUAL)

        for i, (_, puesto, estado, _, _) in enumerate(PUESTOS):
            fp = f_ini + i
            ref_costo = "$L$%d" % fila_de[i]
            poner("B%d" % fp, puesto, "puesto")
            dot = DOTACION.get(i, {}).get(anio, [0] * 12)
            for m in range(12):
                poner(r.cell(row=fp, column=CANT[m]), dot[m], "cantidad")
                cant = "%s%d" % (L(CANT[m]), fp)
                if m in (5, 11):          # junio y diciembre llevan medio sueldo mas
                    previo = "%s%d" % (L(COSTO[m - 1]), fp)
                    formula = "=%s*%s+%s/2" % (cant, ref_costo, previo)
                else:
                    formula = "=%s*%s" % (cant, ref_costo)
                poner(r.cell(row=fp, column=COSTO[m]), formula, "costo")
            suma = "+".join("%s%d" % (L(c), fp) for c in COSTO)
            poner(r.cell(row=fp, column=TOTAL_ANUAL), "=" + suma, "costo")

        f_tot = f_ini + len(PUESTOS)
        totales[anio] = f_tot
        poner("B%d" % f_tot, "Totales", "tot_rot")
        for col in CANT + COSTO + [TOTAL_ANUAL]:
            c = poner(r.cell(row=f_tot, column=col),
                      "=SUM(%s%d:%s%d)" % (L(col), f_ini, L(col), f_tot - 1), "tot_num")
            if col in CANT:
                c.number_format = "0"
        f = f_tot + 3

    # --- totales de la cabecera --------------------------------------------
    for i, anio in enumerate(ANIOS):
        r["%s6" % L(8 + i)] = "=$%s$%d" % (L(TOTAL_ANUAL), totales[anio])

    for i in range(12):
        r.column_dimensions[L(CANT[i])].width = 9
        r.column_dimensions[L(COSTO[i])].width = 12
    r.column_dimensions[L(TOTAL_ANUAL)].width = 14
    for col, ancho in (("B", 30), ("C", 37), ("D", 13), ("K", 12), ("L", 14)):
        r.column_dimensions[col].width = ancho

    wb.save(ARCHIVO)
    print("Costos RRHH rehecha con el formato de Lucas.")
    print("  puestos desplegados : %d en %d áreas" % (
        len(PUESTOS), len({p[0] for p in PUESTOS})))
    print("  en dependencia      : %d" % sum(1 for p in PUESTOS if p[2] == DEP))
    print("  cubiertas por otro  : %d" % sum(1 for p in PUESTOS if p[2] == CUB))
    print("  tercerizadas        : %d" % sum(1 for p in PUESTOS if p[2] == TER))
    print("  bloques anuales     : %s" % ", ".join(
        "%d en f%d-f%d" % (a, primeros[a], totales[a]) for a in ANIOS))
    return 0


if __name__ == "__main__":
    sys.exit(main())
