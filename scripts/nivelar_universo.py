# -*- coding: utf-8 -*-
"""Nivela el Excel y el Word sobre un mismo universo de mercado, ampliado.

Tres correcciones que salen de la clase del 29/09 y de la revisión posterior:

1. La tabla del mercado en Hipótesis mostraba la desagregación (relevados,
   ocupados, indumentaria) y eso fue lo que el profesor no entendió: «vos me
   estás desagregando los datos acá» (02:07), «¿cómo entiendo estos varios
   acá?» (02:23), y recién al llegar al número final, «bien, y esto ya es el
   segmento» (04:42). Su plantilla muestra tres celdas: clientes, gasto
   promedio anual y total mercado. Se deja igual, con la composición en el Word.

2. El universo decía 3.400 comercios de indumentaria de los ejes de la Ciudad,
   pero el mercado meta del punto 4.5 habla del AMBA. El propio informe ya tiene
   el universo ampliado: el segmento B del punto 4.4 son 10.600 locales de otros
   rubros con vidriera en los mismos ejes, con prioridad alta en la matriz de
   crecimiento. El universo pasa a ser el total de locales ocupados con vidriera
   de los 53 ejes relevados, 14.000, con indumentaria como rubro de arranque.

3. El punto 6.2.5 del informe declara que su lista de precios «alimenta
   directamente la hoja de hipótesis» y tenía la mitad de los valores del Excel
   (altas 25/60/250 contra 40/120/450; abonos 15/29/79 contra 33/65/175). El
   profesor lo avisó: «y si lo cambiás, acordate de cambiarlo en el 6.2»
   (23:16). Se lleva el informe a los precios vigentes del modelo.

    python3 scripts/nivelar_universo.py
"""
from __future__ import annotations

import os
import re
import sys

import docx
import openpyxl

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(RAIZ, "documento", "Presupuesto financiero ShopMetrics.xlsx")
DOCX = os.path.join(RAIZ, "documento", "STF_Gomez_Javier_E1_v1.docx")

UNIVERSO = 14000          # locales ocupados de los 53 ejes relevados: 3.400 indumentaria + 10.600 otros rubros con vidriera
GASTO = 960               # gasto anual de referencia por comercio, USD
PARTICIPACION = (0.008, 0.029, 0.052)     # sobre el universo ampliado, para la misma facturación objetivo

ALTAS = (40, 120, 450)
ABONOS = (33, 65, 175)


# ------------------------------------------------------------------- Excel
def excel():
    wb = openpyxl.load_workbook(XLSX)
    h = wb["Hipótesis"]

    # 1) fuera la desagregación: la tabla queda como la de la plantilla
    for fila in (9, 10):
        for col in "BCDE":
            h["%s%d" % (col, fila)] = None

    h["B8"] = ("Comercios minoristas con local a la calle y vidriera a la vereda, de entre uno y cinco locales y "
               "hasta nueve ocupados, ubicados en los ejes comerciales de mayor densidad del AMBA, que ya operan "
               "con alguno de los siete sistemas de punto de venta cubiertos por los conectores propios. El rubro "
               "de arranque es indumentaria, textiles y calzado, y la cabecera de entrada es el eje Avellaneda.")
    h["B12"] = UNIVERSO
    # la cantidad de comercios sale de la proyección, con separador de miles
    for fila, ref in ((24, "AA36"), (25, "AA99"), (26, "AA161")):
        h["B%d" % fila] = '=SUBSTITUTE(TEXT(\'Proy. ventas\'!%s,"#,##0"),",",".")&" comercios"' % ref
    h["B13"] = ("Los 14.000 comercios son los locales ocupados que el IDECBA releva en los 53 ejes comerciales de "
                "mayor densidad (tercer cuatrimestre de 2025). Todos tienen local a la calle y vidriera, que es la "
                "condición para instalar el servicio. El detalle por rubro y la ampliación al Gran Buenos Aires "
                "están en el punto 4 del informe.")

    h["B16"] = ("Participación del 5,2% del universo al cierre del tercer año, con una progresión del 0,8% el "
                "primer año, 2,9% el segundo y 5,2% el tercero. Es una progresión moderada: el rubro de arranque "
                "concentra el esfuerzo comercial del primer año y la ampliación a los otros rubros con vidriera "
                "de los mismos ejes recién se aborda una vez validado ese rubro, de modo que el universo crece "
                "más rápido que la captación.")
    for fila, p in zip((24, 25, 26), PARTICIPACION):
        h["C%d" % fila] = p
        h["C%d" % fila].number_format = "0.0%"      # 0,8% y no 1%

    h["B29"] = ("El valor del mercado captado es el objetivo de facturación de cada año: participación por mercado "
                "total. La proyección de ventas se arma para alcanzarlo, y de ahí se obtiene la cantidad de "
                "comercios abonados que hace falta al cierre de cada año. La cantidad es mayor que la que "
                "resultaría de dividir el mercado captado por el gasto de referencia, porque el ticket promedio "
                "de la lista de precios es menor que ese gasto.")

    wb.save(XLSX)
    print("   Excel: universo %s · participación %s" % (UNIVERSO, ", ".join("%.1f%%" % (p * 100) for p in PARTICIPACION)))


# -------------------------------------------------------------------- Word
def texto(p, nuevo):
    """Reemplaza el texto de un párrafo conservando el formato del primer run."""
    if not p.runs:
        p.add_run(nuevo); return
    p.runs[0].text = nuevo
    for r in p.runs[1:]:
        r.text = ""


def celda(c, nuevo):
    if c.paragraphs and c.paragraphs[0].runs:
        c.paragraphs[0].runs[0].text = nuevo
        for r in c.paragraphs[0].runs[1:]:
            r.text = ""
        for p in c.paragraphs[1:]:
            for r in p.runs:
                r.text = ""
    else:
        c.text = nuevo


def word():
    d = docx.Document(DOCX)
    cambios = []

    def buscar(prefijo):
        for p in d.paragraphs:
            if p.text.strip().startswith(prefijo):
                return p
        return None

    # --- 4.4 segmento A: el universo del rubro de arranque no cambia, pero se dice que es el rubro y no el universo
    p = buscar("Segmento A. Indumentaria con vidriera")
    if p:
        texto(p, "Segmento A. Indumentaria con vidriera en ejes de alta densidad del AMBA. Comercios de "
                 "indumentaria, textiles y calzado ubicados en los corredores de mayor densidad. Representan el "
                 "24,3% de los locales ocupados relevados, unos 3.400 comercios solo en los ejes de la Ciudad de "
                 "Buenos Aires. El eje Avellaneda concentra 1.073 locales relevados con la ocupación más alta del "
                 "operativo, 96,6%. Es el rubro de arranque, no el universo completo: los otros rubros con vidriera "
                 "de los mismos ejes, descritos en el segmento B, comparten la infraestructura de medición y se "
                 "incorporan al mercado meta una vez validado el rubro inicial.")
        cambios.append("4.4 segmento A")

    # --- 4.5 mercado meta: el universo es el comercio con vidriera, con indumentaria como cabecera
    for p in d.paragraphs:
        t = p.text.strip()
        if t.startswith("Comercios minoristas de indumentaria, textiles y calzado, de entre uno y cinco locales"):
            texto(p, "Comercios minoristas con local a la calle y vidriera a la vereda, de entre uno y cinco "
                     "locales y hasta nueve ocupados, ubicados en los ejes comerciales de mayor densidad del AMBA, "
                     "que ya operan con alguno de los siete sistemas de punto de venta cubiertos por los conectores "
                     "propios. El rubro de arranque es indumentaria, textiles y calzado, y la cabecera de entrada "
                     "es el eje Avellaneda.")
            cambios.append("4.5 mercado meta")
            break

    # --- 4.5 valorización: el universo cuantificado
    for p in d.paragraphs:
        if p.text.strip().startswith("Valorización del segmento."):
            texto(p, "Valorización del segmento. El universo del mercado meta son los 14.000 locales ocupados que "
                     "el relevamiento cuenta en los 53 ejes de mayor densidad: 3.400 de indumentaria, textiles y "
                     "calzado, que es el rubro de arranque, y unos 10.600 de los otros rubros con vidriera a la "
                     "vereda, que comparten la misma lógica de embudo y la misma infraestructura de instalación. "
                     "La cantidad no es un supuesto: surge del relevamiento oficial. El valor del mercado se "
                     "obtiene multiplicando ese universo por el gasto anual de referencia del comercio en "
                     "herramientas de gestión, facturación electrónica y analítica, USD 960. El universo del Gran "
                     "Buenos Aires amplía esta cifra, pero no existe un operativo equivalente que lo mida con la "
                     "misma fuente, de modo que queda fuera del dimensionamiento y dentro del plan de expansión.")
            cambios.append("4.5 valorización")
            break

    # --- tablas: 4.3 dimensionamiento, 4.4 ficha, 6.4 competencia, 6.5 lista de precios
    for t in d.tables:
        enc = [c.text.strip() for c in t.rows[0].cells]
        # tabla 4.3
        if enc and enc[0] == "Universo" and any("ARR" in e for e in enc):
            celda(t.rows[0].cells[1], "Locales")
            celda(t.rows[0].cells[2], "Mercado a USD 960 anuales por comercio")
            for i in range(3, len(enc)):
                celda(t.rows[0].cells[i], "")
            filas = [("Indumentaria, textiles y calzado en los ejes relevados, rubro de arranque", "3.400", "USD 3.264.000"),
                     ("Otros rubros con vidriera en los mismos ejes", "10.600", "USD 10.176.000"),
                     ("Universo del mercado meta", "14.000", "USD 13.440.000"),
                     ("Participación pretendida a tres años, 5,2% del universo", "1.033 comercios abonados", "USD 705.506")]
            for r, datos in zip(t.rows[1:], filas):
                for c, v in zip(r.cells[:3], datos):
                    celda(c, v)
                for c in r.cells[3:]:
                    celda(c, "")
            if len(t.rows) - 1 > len(filas):
                for r in t.rows[len(filas) + 1:]:
                    for c in r.cells:
                        celda(c, "")
            cambios.append("tabla 4.3")
        # ficha 4.4
        if enc and enc[0] == "Dimensión":
            for r in t.rows[1:]:
                k = r.cells[0].text.strip()
                if k.startswith("Universo"):
                    celda(r.cells[0], "Universo del mercado meta")
                    celda(r.cells[1], "14.000 locales ocupados con vidriera en los 53 ejes de mayor densidad "
                                      "relevados: 3.400 de indumentaria, que es el rubro de arranque, y 10.600 de "
                                      "los otros rubros con vidriera.")
                if k.startswith("Participación"):
                    celda(r.cells[1], "5,2% del universo al cierre del tercer año, con una progresión de 0,8%, "
                                      "2,9% y 5,2%.")
                if k.startswith("Tipo de cliente"):
                    celda(r.cells[1], "Comercio minorista con local a la calle y vidriera a la vereda. Rubro de "
                                      "arranque: indumentaria, textiles y calzado.")
            cambios.append("ficha 4.4")
        # tabla 6.5 lista de precios
        if enc and enc[0] == "Servicio" and any("Cobro" in e for e in enc):
            precios = dict(zip(("Alta e instalación Plan Básico", "Alta e instalación Plan Vidriera",
                                "Alta e instalación Plan Cadena", "Abono Plan Básico", "Abono Plan Vidriera",
                                "Abono Plan Cadena"), ALTAS + ABONOS))
            for r in t.rows[1:]:
                k = r.cells[0].text.strip()
                if k in precios:
                    celda(r.cells[1], str(precios[k]))
            cambios.append("tabla 6.5")
        # tabla 6.4 competencia
        if enc and enc[0] == "Oferta":
            nuestros = {"ShopMetrics Plan Básico": ABONOS[0], "ShopMetrics Plan Vidriera": ABONOS[1],
                        "ShopMetrics Plan Cadena": ABONOS[2]}
            for r in t.rows[1:]:
                k = r.cells[0].text.strip()
                if k in nuestros:
                    celda(r.cells[1], str(nuestros[k]))
            cambios.append("tabla 6.4")

    # --- textos del punto 6.2 que citan los precios viejos
    reemplazos = [
        (r"se escalona en tres niveles \(USD 15, 29 y 79\)", "se escalona en tres niveles (USD %d, %d y %d)" % ABONOS),
        (r"un abono de USD 15 a 79 según el plan", "un abono de USD %d a %d según el plan" % (ABONOS[0], ABONOS[2])),
        (r"un abono de USD 15 podría sugerir", "un abono de USD %d podría sugerir" % ABONOS[0]),
    ]
    for p in d.paragraphs:
        for pat, nuevo in reemplazos:
            if re.search(pat, p.text):
                texto(p, re.sub(pat, nuevo, p.text))
                cambios.append("6.2 " + nuevo[:28])

    # --- 6.2.1.1: el equipo propio son los dos fundadores
    for p in d.paragraphs:
        if p.text.strip().startswith("Los costos fijos representan el componente principal"):
            texto(p, "Los costos fijos representan el componente estructural de la empresa y se mantienen lo más "
                     "planos posible. En relación de dependencia están únicamente los dos socios fundadores, el "
                     "Gerente General y el Gerente de Sistemas, que cubren entre ambos los puestos de la "
                     "estructura. Todo lo demás se contrata por fuera: el desarrollo y el mantenimiento evolutivo "
                     "de la plataforma, la contabilidad y la gestión de campañas figuran como costo fijo mensual; "
                     "la instalación, el soporte y la venta se pagan según el volumen de cada mes y por eso son "
                     "costo variable. Completan los fijos la infraestructura cloud base, el alquiler de la "
                     "oficina, los servicios, las licencias de software y la auditoría de seguridad de la "
                     "información.")
            cambios.append("6.2.1.1 costos fijos")
            break

    d.save(DOCX)
    print("   Word: %d bloques nivelados -> %s" % (len(cambios), ", ".join(cambios)))


def main() -> int:
    """Sin argumentos hace las dos partes. «--word» sólo el informe, que es lo que
    hace falta cuando el Excel ya está nivelado y recalculado: volver a escribirlo
    con openpyxl le borraría los valores calculados."""
    if "--word" not in sys.argv:
        excel()
    word()
    return 0


if __name__ == "__main__":
    sys.exit(main())
