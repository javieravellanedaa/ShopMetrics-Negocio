# -*- coding: utf-8 -*-
"""Arma las cuatro hojas de escenarios y el plan de contingencia.

    python3 scripts/armar_escenarios.py

La maqueta se clona celda por celda de `Escenario 1` del EJEMPLO de la catedra:
mismos estilos, mismos anchos, mismas combinaciones. Lo unico propio es el
contenido. Despues hay que correr LibreOffice para recalcular y
`restaurar_objetos.py` para devolver los botones y el logo.

Como funciona un escenario, segun la clase (ver documento/clase-riesgos):
se trae el presupuesto financiero entero por referencia y se le aplica el
riesgo encima. Todo lo que se sigue de aplicarlo --los costos variables que
bajan con las ventas, el recurso que ese nivel ya no justifica, el puesto que
no se arma-- va en el escenario y no en el plan de contingencia.

El impuesto a las ganancias se recalcula dentro de cada hoja y no se copia del
presupuesto: si las ventas bajan, baja la utilidad y por lo tanto el impuesto.
Scali lo pide explicito: "te va a cambiar la utilidad antes del impuesto; si
esta bien planteado todo, lo va a cambiar".
"""
from __future__ import annotations

import os
import shutil
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Alignment

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
DOC = os.path.join(RAIZ, "documento")
LIBRO = os.path.join(DOC, "Presupuesto financiero ShopMetrics.xlsx")
EJEMPLO = os.path.join(DOC, "Presupuesto financiero EJEMPLO V1.xlsx")

PF = "'Presupuesto financiero'"
ANIOS = ("E", "F", "G")          # 2026, 2027, 2028 en las dos hojas

# Las amortizaciones del ejercicio, que se restan para el monto imponible.
# Son las mismas celdas que usa el presupuesto en K16, L16 y M16.
AMORT = ("Amortizaciones!$I$5", "Amortizaciones!$J$5", "Amortizaciones!$K$5")

# Costo anual con cargas de los puestos que un escenario de menos actividad no
# incorpora. Sale de 'Costos RRHH': sueldo bruto por el coeficiente de cargas
# (2.000 -> 2.589, o sea 1,2945) por doce meses.
COEF, MESES = 1.2945, 12
COMERCIAL = round(1000 * COEF * MESES)     # Ejecutivo comercial de altas
TECNICO = round(1100 * COEF * MESES)       # Tecnico de instalacion y soporte


class Escenario:
    """Un riesgo aplicado al presupuesto.

    `ventas`, `variables` y `rrhh` son el factor que multiplica cada renglon
    del presupuesto; `mas_variables`, `menos_rrhh`, `mas_fijos` y `menos_inv`
    son montos que se suman o se restan. Todo por ejercicio: el primer valor
    es 2026, el segundo 2027, el tercero 2028.
    """

    def __init__(self, hoja, titulo, riesgo, causa, cuant, cual, ventas=None,
                 variables=None, mas_variables=None, menos_rrhh=None,
                 mas_fijos=None, menos_inv=None, nota=None):
        self.hoja, self.titulo, self.riesgo = hoja, titulo, riesgo
        self.causa, self.cuant, self.cual, self.nota = causa, cuant, cual, nota
        ceros = (0, 0, 0)
        unos = (1, 1, 1)
        self.ventas = ventas or unos
        self.variables = variables or unos
        self.mas_variables = mas_variables or ceros
        self.menos_rrhh = menos_rrhh or ceros
        self.mas_fijos = mas_fijos or ceros
        self.menos_inv = menos_inv or ceros

    def formula(self, fila_pf, i, factor=1.0, suma=0.0):
        base = "=+%s!%s%d" % (PF, ANIOS[i], fila_pf)
        if factor != 1.0:
            base += "*%s" % repr(round(factor, 4))
        if suma:
            base += " %s %d" % ("+" if suma > 0 else "-", abs(int(round(suma))))
        return base


def copiar_maqueta(origen, destino, filas=24, columnas=16):
    """Clona estilos, anchos, altos y combinaciones, sin contenido."""
    # Se copian los objetos de estilo y no `_style`: ese atributo guarda
    # indices a las tablas del libro de origen, que en otro libro apuntan a
    # cualquier otra cosa --o a nada, y entonces merge_cells revienta--.
    for fila in origen.iter_rows(min_row=1, max_row=filas, max_col=columnas):
        for celda in fila:
            nueva = destino.cell(celda.row, celda.column)
            if not celda.has_style:
                continue
            nueva.font = copy(celda.font)
            nueva.fill = copy(celda.fill)
            nueva.border = copy(celda.border)
            nueva.alignment = copy(celda.alignment)
            nueva.number_format = celda.number_format
            nueva.protection = copy(celda.protection)
    for letra, dim in origen.column_dimensions.items():
        d = destino.column_dimensions[letra]
        d.width, d.hidden = dim.width, dim.hidden
    for idx, dim in origen.row_dimensions.items():
        if idx <= filas:
            destino.row_dimensions[idx].height = dim.height
    for rango in list(origen.merged_cells.ranges):
        if rango.max_row <= filas and rango.max_col <= columnas:
            try:
                destino.merge_cells(str(rango))
            except ValueError:
                pass
    destino.sheet_view.showGridLines = origen.sheet_view.showGridLines


def escribir(hoja, esc: Escenario):
    p = hoja.cell
    p(2, 2).value = "Presupuesto financiero"
    p(3, 2).value = "Ejercicio"
    p(3, 4).value = "Inicio"
    for i, anio in enumerate((2026, 2027, 2028)):
        p(3, 5 + i).value = anio

    p(4, 2).value = "Ingresos"
    p(5, 2).value = "Egresos"
    p(5, 3).value = "Costos fijos"
    p(6, 3).value = "Costos variables"
    p(7, 3).value = "Costos de RRHH"
    p(8, 2).value = "Utilidad Antes de Impuestos (UAII)"
    p(9, 2).value = "Impuesto Ingresos Brutos "
    p(9, 3).value = 0.03
    p(10, 2).value = "Impuesto a las ganancias 35%"
    p(12, 2).value = "Utilidad despues de Impuestos (UDII)"
    p(13, 2).value = "Inversión"
    p(14, 2).value = "Flujo de Fondos (FF)"
    for fila in (4, 5, 6, 7, 8, 9, 12, 13, 14):
        p(fila, 4).value = "-"

    for i in range(3):
        col = 5 + i
        p(4, col).value = esc.formula(12, i, esc.ventas[i])
        p(5, col).value = esc.formula(13, i, suma=esc.mas_fijos[i])
        p(6, col).value = esc.formula(14, i, esc.variables[i], esc.mas_variables[i])
        p(7, col).value = esc.formula(15, i, suma=-esc.menos_rrhh[i])
        L = ANIOS[i]
        p(8, col).value = "=+{0}4-{0}5-{0}6-{0}7".format(L)
        p(9, col).value = "=+%s4*$C$9" % L
        # El monto imponible del ejercicio, neto de amortizaciones, igual que
        # en el presupuesto. El impuesto se paga al ano siguiente.
        p(22, col).value = "=+{0}8-{0}9-{1}".format(L, AMORT[i])
        p(10, col).value = ("0" if i == 0
                            else "=MAX(0,$%s$22)*0.35" % ANIOS[i - 1])
        p(12, col).value = "=+{0}8-{0}9-{0}10".format(L)
        p(13, col).value = esc.formula(21, i, suma=esc.menos_inv[i])
        p(14, col).value = "=+{0}12+{0}13".format(L)

    p(13, 4).value = "=+%s!D21" % PF
    p(14, 4).value = "=+D13"
    p(21, 2).value = "Monto imponible del ejercicio (neto de amortizaciones)"

    p(18, 5).value = "Tasa de Corte"
    p(18, 6).value = "=+%s!$G$33" % PF
    p(19, 5).value = "VAN"
    p(19, 6).value = "=NPV($F$18,E14:G14)+D14"
    p(20, 5).value = "TIR"
    p(20, 6).value = "=IRR(D14:G14,$F$18)"

    p(3, 9).value = "RIESGO: %s" % esc.titulo
    p(4, 9).value = esc.riesgo
    p(4, 10).value = esc.causa
    p(4, 11).value = esc.cuant
    p(4, 12).value = esc.cual
    if esc.nota:
        # Debajo del cuadro: en la fila 7 --donde el ejemplo pone "Año 2" y
        # "Año 3"-- una nota larga se desborda y pisa el renglón de RRHH.
        celda = p(16, 9)
        celda.value = esc.nota
        celda.alignment = Alignment(wrap_text=True, vertical="top")
        hoja.merge_cells(start_row=16, start_column=9, end_row=18, end_column=12)


# ---------------------------------------------------------------------------
# Los tres escenarios y el plan. Los riesgos y sus efectos salen de la hoja
# `Matriz de riesgos`, no de ningun numero inventado aca.
# ---------------------------------------------------------------------------

ESCENARIOS = [
    Escenario(
        "Escenario 1",
        "R1 · No alcanzar el volumen mínimo de comercios adheridos",
        "R 1",
        "- Tasa de conversión por visita menor a la proyectada\n"
        "- Resistencia del comerciante a instalar el sensor\n"
        "- Densidad efectiva del eje menor a la estimada",
        "- Altas del año 3 un 25% por debajo de las 510 proyectadas\n"
        "- Ingreso 2027 y 2028 un 25% menor\n"
        "- Los costos variables acompañan la baja en la misma proporción",
        "- Estructura sobredimensionada frente al ingreso\n"
        "- Pérdida de la masa crítica de pares que da valor al benchmarking",
        ventas=(1, 0.75, 0.75),
        variables=(1, 0.75, 0.75),
        menos_rrhh=(0, COMERCIAL, COMERCIAL),
        menos_inv=(0, 1200, 1200),
        nota="A ese nivel de altas no se incorpora el segundo ejecutivo "
             "comercial ni se arma su puesto de trabajo."),

    Escenario(
        "Escenario 2",
        "R2 · Rotación de comercios antes de amortizar el equipamiento",
        "R 2",
        "- Alta rotación propia del comercio minorista de barrio\n"
        "- Bajo uso de la plataforma tras las primeras semanas\n"
        "- Estacionalidad del rubro indumentaria",
        "- Ingreso recurrente 2027 y 2028 un 15% menor por las bajas\n"
        "- Reposición y reacondicionamiento del equipamiento retirado: "
        "USD 12.000 por ejercicio",
        "- Erosión de la base instalada\n"
        "- Mayor esfuerzo comercial para sostener el mismo stock de clientes",
        ventas=(1, 0.85, 0.85),
        variables=(1, 0.85, 0.85),
        mas_variables=(0, 12000, 12000),
        nota="La dotación no baja: sostener la base rotando exige el mismo "
             "trabajo de instalación y posventa."),

    Escenario(
        "Escenario 3",
        "Combinado · R1 volumen + R2 rotación + R6 salto cambiario",
        "R1+R2+R6",
        "- Las altas no alcanzan el volumen previsto\n"
        "- Los comercios que entran rotan antes del año\n"
        "- El abono en pesos se licúa entre la facturación y la cobranza",
        "- Ingreso 2027 y 2028 un 40% por debajo del presupuesto\n"
        "- Reposición de equipamiento retirado: USD 15.000 por ejercicio\n"
        "- Los costos variables acompañan la baja de ventas",
        "- El proyecto deja de cubrir su costo de capital\n"
        "- Dependencia del aporte de los socios para sostener la operación",
        ventas=(1, 0.60, 0.60),
        variables=(1, 0.60, 0.60),
        mas_variables=(0, 15000, 15000),
        menos_rrhh=(0, COMERCIAL, COMERCIAL),
        menos_inv=(0, 1200, 1200),
        nota="Los tres riesgos a la vez. No se incorpora el segundo ejecutivo "
             "comercial, pero el técnico de instalación se mantiene: con "
             "rotación hay más altas y más retiros que atender, no menos. Es "
             "el peor de los tres y es el que sostiene el plan."),

    Escenario(
        "Plan de Contingencia",
        "Plan sobre el escenario 3, que es el que no resiste",
        "R1+R2+R6",
        "- Escenario 3 con las medidas de contingencia aplicadas",
        "1ª medida: campaña de retención sobre la base instalada y segundo "
        "canal de captación por referidos. Suma USD 6.000 por ejercicio al "
        "costo fijo y recupera 8 puntos de la caída de ventas.\n"
        "2ª medida: extensión del plazo de permanencia bonificada a cambio "
        "de un abono menor. Suma USD 3.000 por ejercicio.",
        "- El proyecto vuelve a cubrir su costo de capital\n"
        "- Se preserva la masa crítica que da sentido al benchmarking",
        ventas=(1, 0.66, 0.66),
        variables=(1, 0.66, 0.66),
        mas_variables=(0, 15000, 15000),
        menos_rrhh=(0, COMERCIAL, COMERCIAL),
        mas_fijos=(0, 9000, 9000),
        menos_inv=(0, 1200, 1200),
        nota="El plan nivela, no mejora: la TIR queda por encima de la tasa "
             "de corte y muy por debajo de la del presupuesto original."),
]


def main() -> int:
    if not os.path.exists(LIBRO):
        print("no está %s" % LIBRO); return 1
    ejemplo = openpyxl.load_workbook(EJEMPLO)
    modelo = ejemplo["Escenario 1"]
    libro = openpyxl.load_workbook(LIBRO)

    for esc in ESCENARIOS:
        hoja = libro[esc.hoja]
        # Las combinaciones se deshacen antes de limpiar: en una celda
        # combinada todas menos la de arriba a la izquierda son de solo
        # lectura, y asignarles None revienta.
        for rango in list(hoja.merged_cells.ranges):
            hoja.unmerge_cells(str(rango))
        for fila in hoja.iter_rows(min_row=1, max_row=hoja.max_row,
                                   max_col=hoja.max_column):
            for celda in fila:
                celda.value = None
        copiar_maqueta(modelo, hoja)
        escribir(hoja, esc)
        print("  %-22s %s" % (esc.hoja, esc.titulo))

    libro.save(LIBRO)
    print("guardado: %s" % os.path.basename(LIBRO))
    return 0


if __name__ == "__main__":
    sys.exit(main())
