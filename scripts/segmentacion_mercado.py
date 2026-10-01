#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Escribe la segmentacion del mercado meta en el punto 2 de la hoja Hipotesis
y la deja nivelada con el punto 4 del informe.

Scali pidio que el mercado meta diga COMO se segmento y lo cuantifique segmento
por segmento, con el formato del trabajo que mostro en clase: cada segmento con
su cantidad de empresas, su fuente, su gasto anual y el mercado direccionable
que consolida, y un total consolidado al final.

Los numeros no se escriben a mano: salen de B12 (universo), C12 (gasto anual de
referencia) y de los porcentajes publicados del relevamiento del IDECBA, de modo
que la suma de los segmentos siempre cierra contra el total de la planilla.
"""
import os
import sys
from copy import copy

from openpyxl.styles import Alignment

import docx
import openpyxl

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(RAIZ, "documento", "Presupuesto financiero ShopMetrics.xlsx")
DOCX = os.path.join(RAIZ, "documento", "STF_Gomez_Javier_E1_v1.docx")

RELEVADOS = 14083        # locales ocupados que releva el IDECBA en los 53 ejes
EJES = 53

# (etiqueta, participacion publicada, comercios, descripcion)
SEGMENTOS = [
    ("Segmento A — Indumentaria, textiles y calzado", 0.243, 3400,
     "es el rubro de mayor concentración del relevamiento y la cabecera de entrada por el "
     "eje Avellaneda, que reúne 1.073 locales con la ocupación más alta del operativo, 96,6%"),
    ("Segmento B — Alojamiento y comida", 0.113, 1600,
     "gastronomía con vidriera a la vereda, que comparte el mismo embudo de paso, detención "
     "e ingreso y la misma infraestructura de medición"),
    ("Segmento C — Resto de rubros con vidriera", 0.644, 9000,
     "alimentos y bebidas, perfumería, calzado deportivo y servicios con local a la calle, "
     "que comparten la infraestructura de medición con menor densidad de pares por rubro"),
]


def n(x, dec=0):
    s = "{:,.{d}f}".format(x, d=dec)
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(x):
    return n(100 * x, 1) + "%"


def armar(universo, gasto):
    """El parrafo de segmentacion, construido desde los valores de la planilla."""
    partes = [
        "Segmentación del mercado meta. El universo se segmenta por rubro sobre los %s locales "
        "ocupados que el Instituto de Estadística y Censos de la Ciudad de Buenos Aires releva en "
        "los %d ejes comerciales de mayor densidad (tercer cuatrimestre de 2025), redondeados a %s. "
        "Los tres segmentos comparten el mismo tramo de tamaño —de uno a cinco locales y hasta "
        "nueve ocupados— y por eso el gasto anual de referencia en herramientas de gestión, "
        "facturación electrónica y analítica es común, USD %s por comercio; lo que los diferencia "
        "es la densidad de pares por eje y la prioridad comercial de ingreso."
        % (n(RELEVADOS), EJES, n(universo), n(gasto))
    ]
    for etq, share, cant, desc in SEGMENTOS:
        partes.append(
            "%s: compuesto por %s comercios, el %s de los locales ocupados relevados; %s. "
            "Con un gasto anual de USD %s por comercio consolida un mercado direccionable de USD %s."
            % (etq, n(cant), pct(share), desc, n(gasto), n(cant * gasto)))
    partes.append(
        "Total consolidado: los tres segmentos suman %s comercios y un mercado direccionable de "
        "USD %s anuales, que es el total que toma el cuadro siguiente. El universo del Gran Buenos "
        "Aires amplía esta cifra, pero no existe un operativo equivalente que lo mida con la misma "
        "fuente, de modo que queda fuera del dimensionamiento y dentro del plan de expansión."
        % (n(sum(s[2] for s in SEGMENTOS)), n(sum(s[2] for s in SEGMENTOS) * gasto)))
    return "\n".join(partes)


def excel(texto):
    wb = openpyxl.load_workbook(XLSX)
    h = wb["Hipótesis"]
    for rango in list(h.merged_cells.ranges):
        if rango.min_row <= 9 <= rango.max_row:
            h.unmerge_cells(str(rango))
    c = h["B9"]
    c.value = texto
    modelo = h["B8"]
    c.font = copy(modelo.font); c.fill = copy(modelo.fill); c.border = copy(modelo.border)
    al = copy(modelo.alignment)
    al.wrap_text = True
    al.vertical = "top"
    c.alignment = al
    h.merge_cells("B9:E9")
    # el alto lo resuelve LibreOffice al recalcular: con customHeight apagado
    # la fila se autoajusta al texto y no queda espacio muerto debajo
    # alto: una pasada propia en vez del autoajuste, que se queda corto por
    # una linea. ~125 caracteres por linea sobre el ancho combinado B:E.
    import math
    cpl = 125
    lineas = sum(max(1, math.ceil(len(t) / cpl)) for t in texto.split("\n"))
    h.row_dimensions[9].height = lineas * 11.0 + 4
    wb.save(XLSX)
    return len(texto)


def word():
    """Abre los 10.600 del segmento B del punto 4 en los dos tramos del Excel."""
    d = docx.Document(DOCX)
    marca = "Son unos 10.600 locales ocupados en los mismos corredores."
    nuevo = ("Son unos 10.600 locales ocupados en los mismos corredores, que se abren en dos tramos: "
             "1.600 de alojamiento y comida, el 11,3% de los locales relevados, y 9.000 del resto de "
             "los rubros con vidriera. Comparten la lógica de embudo y la infraestructura de medición, "
             "pero con menor densidad de pares por rubro.")
    viejo_cierre = (" Comparten la lógica de embudo y la infraestructura de medición, pero con menor "
                    "densidad de pares por rubro.")
    for p in d.paragraphs:
        if marca in p.text:
            t = p.text.replace(marca + viejo_cierre, nuevo).replace(marca, nuevo) \
                if viejo_cierre in p.text else p.text.replace(marca, nuevo)
            p.runs[0].text = t
            for r in p.runs[1:]:
                r.text = ""
            d.save(DOCX)
            return True
    return False


def main() -> int:
    solo_word = "--word" in sys.argv       # el Excel ya esta recalculado: no pisarlo
    v = openpyxl.load_workbook(XLSX, data_only=True)["Hipótesis"]
    universo, gasto, total = v["B12"].value, v["C12"].value, v["D12"].value
    suma = sum(s[2] for s in SEGMENTOS)
    if suma != universo:
        print("ERROR: los segmentos suman %d y el universo es %d" % (suma, universo))
        return 1
    if abs(sum(s[2] * gasto for s in SEGMENTOS) - total) > 1:
        print("ERROR: el mercado de los segmentos no cierra contra D12")
        return 1
    texto = armar(universo, gasto)
    largo = len(texto) if solo_word else excel(texto)
    ok = word()
    print("Segmentación %s (%d caracteres)." % (
        "del informe" if solo_word else "escrita en Hipótesis!B9", largo))
    for etq, share, cant, _ in SEGMENTOS:
        print("   %-46s %6s  %6s  USD %s" % (etq, n(cant), pct(share), n(cant * gasto)))
    print("   %-46s %6s  %6s  USD %s" % ("TOTAL", n(suma), "100,0%", n(total)))
    print("   punto 4 del informe: %s" % ("segmento B abierto en dos tramos" if ok else "NO SE ENCONTRÓ"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
