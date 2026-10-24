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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from clonar_hoja_ejemplo import clonar

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
                 mas_fijos=None, menos_inv=None, objetivos=None, plan=None):
        self.hoja, self.titulo, self.riesgo = hoja, titulo, riesgo
        self.causa, self.cuant, self.cual = causa, cuant, cual
        self.objetivos, self.plan = objetivos or (), plan
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


def _cuadro(hoja, esc, base, fila_riesgo, fila_anios):
    """El cuadro del presupuesto con el riesgo aplicado.

    `base` es la fila donde arranca el bloque: 2 en un escenario y 10 en el
    plan de contingencia, que antes lleva el titulo y los objetivos. Las filas
    de adentro guardan siempre la misma distancia entre si, que es la del
    ejemplo de la catedra.
    """
    p = hoja.cell
    r = lambda n: base + n          # 0 = "Presupuesto financiero"

    p(r(0), 2).value = "Presupuesto financiero"
    p(r(1), 2).value = "Ejercicio"
    p(r(1), 4).value = "Inicio"
    for i, anio in enumerate((2026, 2027, 2028)):
        p(r(1), 5 + i).value = anio

    p(r(2), 2).value = "Ingresos"
    p(r(3), 2).value = "Egresos"
    p(r(3), 3).value = "Costos fijos"
    p(r(4), 3).value = "Costos variables"
    p(r(5), 3).value = "Costos de RRHH"
    p(r(6), 2).value = "Utilidad Antes de Impuestos (UAII)"
    p(r(7), 2).value = "Impuesto Ingresos Brutos "
    p(r(7), 3).value = 0.03
    p(r(8), 2).value = "Impuesto a las ganancias 35%"
    p(r(10), 2).value = "Utilidad despues de Impuestos (UDII)"
    p(r(11), 2).value = "Inversión"
    p(r(12), 2).value = "Flujo de Fondos (FF)"
    for n in (2, 3, 4, 5, 6, 7, 10, 11, 12):
        p(r(n), 4).value = "-"

    imponible = r(20)               # fuera de la vista, como el monto del PF
    for i in range(3):
        col = 5 + i
        L = ANIOS[i]
        p(r(2), col).value = esc.formula(12, i, esc.ventas[i])
        p(r(3), col).value = esc.formula(13, i, suma=esc.mas_fijos[i])
        p(r(4), col).value = esc.formula(14, i, esc.variables[i], esc.mas_variables[i])
        p(r(5), col).value = esc.formula(15, i, suma=-esc.menos_rrhh[i])
        p(r(6), col).value = "=+{0}{1}-{0}{2}-{0}{3}-{0}{4}".format(
            L, r(2), r(3), r(4), r(5))
        p(r(7), col).value = "=+%s%d*$C$%d" % (L, r(2), r(7))
        # El monto imponible del ejercicio, neto de amortizaciones. El impuesto
        # se paga al ano siguiente, igual que en el presupuesto.
        p(imponible, col).value = "=+{0}{1}-{0}{2}-{3}".format(
            L, r(6), r(7), AMORT[i])
        p(r(8), col).value = (0 if i == 0
                              else "=MAX(0,$%s$%d)*0.35" % (ANIOS[i - 1], imponible))
        p(r(10), col).value = "=+{0}{1}-{0}{2}-{0}{3}".format(L, r(6), r(7), r(8))
        p(r(11), col).value = esc.formula(21, i, suma=esc.menos_inv[i])
        p(r(12), col).value = "=+{0}{1}+{0}{2}".format(L, r(10), r(11))

    p(r(11), 4).value = "=+%s!D21" % PF
    p(r(12), 4).value = "=+D%d" % r(11)
    p(imponible, 2).value = "Monto imponible del ejercicio (neto de amortizaciones)"

    # La tasa, el VAN y la TIR, con la misma formula que usa el presupuesto:
    # el desembolso inicial no se descuenta, se suma.
    tasa = r(16)
    p(tasa, 5).value = "Tasa de Corte"
    p(tasa, 6).value = "=+%s!$G$33" % PF
    p(tasa + 1, 5).value = "VAN"
    p(tasa + 1, 6).value = "=NPV($F$%d,E%d:G%d)+D%d" % (tasa, r(12), r(12), r(12))
    p(tasa + 2, 5).value = "TIR"
    p(tasa + 2, 6).value = "=IRR(D%d:G%d,$F$%d)" % (r(12), r(12), tasa)

    # El riesgo, en las mismas celdas que el ejemplo.
    p(fila_riesgo, 9).value = "RIESGO: %s" % esc.titulo
    p(fila_riesgo + 1, 9).value = esc.riesgo
    p(fila_riesgo + 1, 10).value = esc.causa
    p(fila_riesgo + 1, 11).value = esc.cuant
    p(fila_riesgo + 1, 12).value = esc.cual
    p(fila_anios, 11).value = "Año 2"
    p(fila_anios, 12).value = "Año 3"


def escribir_escenario(hoja, esc: Escenario):
    _cuadro(hoja, esc, base=2, fila_riesgo=3, fila_anios=7)


def escribir_plan(hoja, esc: Escenario):
    p = hoja.cell
    p(2, 2).value = "PLAN DE CONTINGENCIA"
    p(4, 2).value = "Objetivos del Plan de Contingencia"
    for i, objetivo in enumerate(esc.objetivos):
        p(5 + i, 2).value = objetivo
    _cuadro(hoja, esc, base=10, fila_riesgo=11, fila_anios=15)
    p(17, 9).value = "PLAN DE CONTINGENCIA"
    p(18, 9).value = esc.plan


# ---------------------------------------------------------------------------
# Los tres escenarios y el plan. Los riesgos y sus efectos salen de la hoja
# `Matriz de riesgos`, no de ningun numero inventado aca.
# ---------------------------------------------------------------------------

ESCENARIOS = [
    Escenario(
        "Escenario 1",
        "R1 · Caída de la plataforma o de la infraestructura en la nube",
        "R 1",
        "- Falla del proveedor de nube o de la base de datos\n"
        "- Error en un despliegue que deja la API sin servicio\n"
        "- Ataque de denegación de servicio",
        "- 72 horas sin servicio para toda la base: se bonifica el abono y una de "
        "cada diez cuentas se da de baja\n"
        "- Ingreso 2027 y 2028 un 10% menor\n"
        "- Recuperación y horas de DevOps: USD 3.000 por ejercicio\n"
        "- Los costos variables acompañan la baja de ventas",
        "- Pérdida de confianza en un servicio vendido como tiempo real\n"
        "- Daño reputacional en el corredor",
        ventas=(1, 0.90, 0.90),
        variables=(1, 0.90, 0.90),
        mas_variables=(0, 3000, 3000)),

    Escenario(
        "Escenario 2",
        "R2 · Baja anticipada de comercios adheridos",
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
        mas_variables=(0, 12000, 12000)),

    Escenario(
        "Escenario 3",
        "Combinado · R1 caída de plataforma + R2 baja anticipada + R10 competidor",
        "R1+R2+R10",
        "- Una caída prolongada de la plataforma\n"
        "- Los comercios que entran se van antes del año\n"
        "- Un competidor local con precio agresivo",
        "- Ingreso 2027 y 2028 un 40% por debajo: la caída hace perder cuentas, "
        "las bajas no se reponen y el competidor presiona el precio\n"
        "- Reposición y reinstalación de equipamiento: USD 15.000 por ejercicio\n"
        "- Los costos variables acompañan la baja de ventas",
        "- El proyecto deja de cubrir su costo de capital\n"
        "- Se pierde la masa crítica que da sentido al benchmarking",
        ventas=(1, 0.60, 0.60),
        variables=(1, 0.60, 0.60),
        mas_variables=(0, 15000, 15000),
        menos_rrhh=(0, COMERCIAL, COMERCIAL),
        menos_inv=(0, 1200, 1200)),

    Escenario(
        "Plan de Contingencia",
        "Plan sobre el escenario 3, que es el que no resiste",
        "R1+R2+R10",
        "- Escenario 3 con las medidas aplicadas",
        "- Segunda nube y monitoreo: +USD 4.000 de costo fijo\n"
        "- Retención y referidos: +USD 5.000",
        "- Vuelve a cubrir el costo de capital\n"
        "- Conserva la masa crítica",
        ventas=(1, 0.66, 0.66),
        variables=(1, 0.66, 0.66),
        mas_variables=(0, 15000, 15000),
        menos_rrhh=(0, COMERCIAL, COMERCIAL),
        mas_fijos=(0, 9000, 9000),
        menos_inv=(0, 1200, 1200),
        objetivos=(
            "· Sostener el servicio aunque falle la nube principal.",
            "· Devolver el proyecto por encima de su costo de capital.",
            "· Proteger la base instalada frente al competidor."),
        plan=(
            "El plan se aplica sobre el escenario 3, que es el único de los tres "
            "que no resiste: con la caída de la plataforma, las bajas anticipadas "
            "y el competidor a la vez, la TIR cae a 22,54% contra una tasa de "
            "corte del 30% y el VAN queda negativo.\n"
            "\n"
            "1ª medida: respaldo en una segunda nube con conmutación automática y "
            "monitoreo permanente, para que una caída del proveedor principal no "
            "deje a la base sin servicio. Suma USD 4.000 por ejercicio al costo "
            "fijo.\n"
            "2ª medida: campaña de retención sobre la base instalada y canal de "
            "captación por referidos, que es donde el competidor no llega. Suma "
            "USD 5.000 por ejercicio.\n"
            "\n"
            "Entre las dos recuperan 6 de los 40 puntos de caída de ventas. La "
            "dotación se mantiene en la del escenario: no se incorpora el segundo "
            "ejecutivo comercial.\n"
            "\n"
            "El plan nivela, no mejora: la TIR queda en 31,34%, apenas por encima "
            "de la tasa de corte y muy por debajo del 91,70% del presupuesto "
            "original.")),
]


def main() -> int:
    if not os.path.exists(LIBRO):
        print("no está %s" % LIBRO); return 1
    ejemplo = openpyxl.load_workbook(EJEMPLO)
    libro = openpyxl.load_workbook(LIBRO)

    for esc in ESCENARIOS:
        es_plan = esc.hoja == "Plan de Contingencia"
        # Cada una se clona de la hoja del ejemplo que le corresponde: el plan
        # tiene maqueta propia --titulo, objetivos, la tabla corrida ocho
        # filas-- y no es un escenario con otro nombre.
        modelo = ejemplo["Plan de Contingencia" if es_plan else "Escenario 1"]
        hoja = libro[esc.hoja]
        clonar(modelo, hoja, con_valores=False, filas=30, columnas=14)
        (escribir_plan if es_plan else escribir_escenario)(hoja, esc)
        print("  %-22s %s" % (esc.hoja, esc.titulo))

    libro.save(LIBRO)
    print("guardado: %s" % os.path.basename(LIBRO))
    return 0


if __name__ == "__main__":
    sys.exit(main())
