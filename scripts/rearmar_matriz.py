# -*- coding: utf-8 -*-
"""Vuelve a poner la Matriz de riesgos en el formato del EJEMPLO de la catedra.

    python3 scripts/rearmar_matriz.py

El contenido --los diez riesgos, sus causas, efectos y puntajes-- es el de
ShopMetrics y no se toca. Lo que cambia es la ropa: encabezado verde fuerte,
filas blancas sin combinar, la columna de nivel en naranja, probabilidad e
impacto repetidos en cada renglon de cada riesgo, y el bloque de escalas de la
derecha tal cual lo trae la plantilla, incluida su fila 13.
"""
from __future__ import annotations

import os
import sys
from copy import copy

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
DOC = os.path.join(os.path.dirname(AQUI), "documento")
LIBRO = os.path.join(DOC, "Presupuesto financiero ShopMetrics.xlsx")
EJEMPLO = os.path.join(DOC, "Presupuesto financiero EJEMPLO V1.xlsx")
HOJA = "Matriz de riesgos"
ULTIMA_FILA = 32              # 10 riesgos de 3 renglones, desde la fila 3


def estilo(desde, hasta):
    hasta.font = copy(desde.font)
    hasta.fill = copy(desde.fill)
    hasta.border = copy(desde.border)
    hasta.alignment = copy(desde.alignment)
    hasta.number_format = desde.number_format


def main() -> int:
    ejemplo = openpyxl.load_workbook(EJEMPLO)[HOJA]
    libro = openpyxl.load_workbook(LIBRO)
    hoja = libro[HOJA]

    # 1. El contenido propio, antes de tocar nada.
    cuerpo = {}
    for r in range(2, ULTIMA_FILA + 1):
        cuerpo[r] = [hoja.cell(r, c).value for c in range(2, 10)]
    grilla = {(r, c): hoja.cell(r, c).value for r in range(4, 9) for c in range(15, 20)}

    # 2. Limpiar: combinaciones, valores y estilos del area que se rearma.
    for rango in list(hoja.merged_cells.ranges):
        hoja.unmerge_cells(str(rango))
    vacio = openpyxl.styles.PatternFill(fill_type=None)
    for r in range(1, 41):
        for c in range(1, 21):
            cel = hoja.cell(r, c)
            cel.value = None
            cel.fill = vacio
            cel.border = openpyxl.styles.Border()
            cel.font = openpyxl.styles.Font()
            cel.alignment = openpyxl.styles.Alignment()

    # 3. La ropa del ejemplo. Encabezado de su fila 2; el cuerpo de su fila 3,
    #    que es la misma para todos sus renglones de datos.
    for c in range(2, 10):
        estilo(ejemplo.cell(2, c), hoja.cell(2, c))
    for r in range(3, ULTIMA_FILA + 1):
        for c in range(2, 10):
            estilo(ejemplo.cell(3, c), hoja.cell(r, c))
    for r in range(1, 14):
        for c in range(11, 20):
            estilo(ejemplo.cell(r, c), hoja.cell(r, c))
    for letra, dim in ejemplo.column_dimensions.items():
        hoja.column_dimensions[letra].width = dim.width
    hoja.row_dimensions[2].height = ejemplo.row_dimensions[2].height
    for r in range(3, ULTIMA_FILA + 1):
        hoja.row_dimensions[r].height = ejemplo.row_dimensions[3].height
    for r in range(1, 14):
        if ejemplo.row_dimensions[r].height and r > ULTIMA_FILA:
            hoja.row_dimensions[r].height = ejemplo.row_dimensions[r].height
    for rango in ejemplo.merged_cells.ranges:
        hoja.merge_cells(str(rango))
    hoja.sheet_view.showGridLines = ejemplo.sheet_view.showGridLines

    # 4. El contenido de vuelta. Probabilidad, impacto y nivel van en cada
    #    renglon del riesgo, como en el ejemplo, y el nivel es formula.
    for c, v in zip(range(2, 10), cuerpo[2]):
        hoja.cell(2, c).value = v
    p = i = None
    for r in range(3, ULTIMA_FILA + 1):
        fila = cuerpo[r]
        if fila[0]:                       # arranca un riesgo
            p, i = fila[5], fila[6]
        for c, v in zip(range(2, 7), fila[:5]):
            hoja.cell(r, c).value = v
        hoja.cell(r, 7).value = p
        hoja.cell(r, 8).value = i
        hoja.cell(r, 9).value = "=+G%d*H%d" % (r, r)
        # El nivel se pinta por banda, como en el ejemplo: verde hasta 6,
        # celeste hasta 12, naranja hasta 20, rojo de 21 a 25. Son los
        # mismos colores de la leyenda de la derecha.
        nivel = (p or 0) * (i or 0)
        color = ("92D050" if nivel <= 6 else "00B0F0" if nivel <= 12
                 else "FFC000" if nivel <= 20 else "FF0000")
        hoja.cell(r, 9).fill = openpyxl.styles.PatternFill("solid", fgColor=color)

        # La fila crece con el texto. El ejemplo tambien lo hace (41,25 y
        # 54,75): lo que no hace es dejar texto cortado.
        lineas = 1
        for c, ancho in ((3, 29.7), (4, 24.1), (5, 26.9), (6, 30.6)):
            texto = str(hoja.cell(r, c).value or "")
            if texto:
                por_linea = max(1, int(ancho * 1.15))
                lineas = max(lineas, sum(-(-len(s) // por_linea) for s in texto.split("\n")))
        hoja.row_dimensions[r].height = max(41.25, lineas * 12.0 + 6)

    # 5. El bloque de la derecha: escalas y bandas de la catedra, tal cual;
    #    la grilla con los riesgos de ShopMetrics.
    from openpyxl.cell.cell import MergedCell
    for r in range(1, 14):
        for c in range(11, 20):
            if isinstance(hoja.cell(r, c), MergedCell):
                continue              # la leyenda O9:S12 ya esta combinada
            if (r, c) in grilla:
                hoja.cell(r, c).value = grilla[(r, c)]
            else:
                hoja.cell(r, c).value = ejemplo.cell(r, c).value

    libro.save(LIBRO)
    print("  %s: %d riesgos, formato del ejemplo" % (HOJA, sum(1 for r in cuerpo.values() if r[0] and r[0] != "Riesgo")))
    return 0


if __name__ == "__main__":
    sys.exit(main())
