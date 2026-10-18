# -*- coding: utf-8 -*-
"""Reescribe las fichas de casos de uso (10.5.3) con la estructura que pide la
guia de catedra (T12347, Vilaboa 2025).

Que cambia respecto de la version anterior:

  - la cabecera de cada ficha pasa a ser un titulo, para que entre en el indice
    con su numero de pagina (lo pide el 10.5.2)
  - se agregan los campos "Identificacion del Caso de Uso", "Nombre del Caso de
    Uso" y "Puntos de extension y condiciones"
  - se renombran tres campos a como los nombra la guia
  - Pre y Post Condicion pasan a estar numeradas
  - el escenario principal se desglosa en pasos atomicos, contados desde el
    usuario: sin tokens, sin endpoints, sin SQL. La propia guia lo pide en el
    10.5.3.12: "debe evitar sumar a la respuesta del sistema los hechos que este
    realiza para responder"
  - los caminos alternativos pasan a tener letra y a colgar de un paso concreto

Se corre desde la raiz de ShopMetrics-Negocio:  python3 scripts/reestructurar_cu.py
"""
from __future__ import annotations

import copy
import os
import re
import sys

from docx import Document
from docx.shared import Pt
from docx.text.paragraph import Paragraph

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORD = os.path.join(RAIZ, "documento", "STF_Gomez_Javier_E2_v2.docx")

CABECERA = re.compile(r"^(CU-\d{3}-\d{3})\s*[:·]\s*(.+)$")

# Los campos que se conservan tal cual estaban, con su nombre nuevo.
RENOMBRES = {
    "Objetivo": "Objetivo",
    "Descripción": "Descripción del Caso de Uso",
    "Actores Primarios": "Actores Primarios",
    "Actores Secundarios": "Actores Secundarios",
    "Evento Disparador": "Eventos disparadores del caso de uso",
}


def leer_fichas(doc):
    """Ubica cada ficha: (indice_cabecera, id, nombre, campos, indice_prototipo)."""
    els = list(doc.element.body.iterchildren())
    marcas = []
    for i, e in enumerate(els):
        if not e.tag.endswith("}p"):
            continue
        m = CABECERA.match(Paragraph(e, doc).text.strip())
        if m:
            marcas.append((i, m.group(1), m.group(2)))
    fichas = []
    for k, (i, cu, nombre) in enumerate(marcas):
        fin = marcas[k + 1][0] if k + 1 < len(marcas) else len(els)
        campos, proto = {}, None
        for j in range(i + 1, fin):
            if not els[j].tag.endswith("}p"):
                continue
            t = Paragraph(els[j], doc).text.strip()
            if not t:
                continue
            if t.startswith("10.5."):
                fin = j
                break
            if t.startswith("Prototipo"):
                proto = j
            m = re.match(r"^([A-ZÁÉÍÓÚ][^:]{2,36}):\s*(.*)$", t)
            if m and m.group(1) in RENOMBRES:
                campos[m.group(1)] = m.group(2)
        fichas.append({"ini": i, "fin": fin, "cu": cu, "nombre": nombre,
                       "campos": campos, "proto": proto})
    return els, fichas


def parrafo_como(modelo, etiqueta, texto, sangria=False):
    """Clona el formato de un parrafo existente y escribe etiqueta + texto."""
    nuevo = copy.deepcopy(modelo._p)
    for hijo in list(nuevo):
        if hijo.tag.endswith("}r") or hijo.tag.endswith("}hyperlink"):
            nuevo.remove(hijo)
    p = Paragraph(nuevo, modelo._parent)
    if etiqueta:
        r = p.add_run(etiqueta + ": ")
        r.bold = True
    if texto:
        p.add_run(texto)
    if sangria:
        p.paragraph_format.left_indent = Pt(18)
    return nuevo


def reescribir(doc, els, ficha, nuevo, escenario, caminos):
    """Reemplaza el cuerpo de una ficha. Conserva la cabecera --restilada a
    titulo-- y el parrafo del prototipo, que lleva el hipervinculo a Figma."""
    cab = Paragraph(els[ficha["ini"]], doc)
    cab.style = doc.styles["Heading 4"]
    cab.text = ""
    r = cab.add_run("%s · %s" % (ficha["cu"], ficha["nombre"]))
    r.bold = True

    c = ficha["campos"]
    modelo = Paragraph(els[ficha["ini"] + 1], doc)
    bloques = []

    def campo(etiqueta, texto):
        bloques.append(parrafo_como(modelo, etiqueta, texto))

    def lista(etiqueta, items):
        campo(etiqueta, "")
        for n, x in enumerate(items, 1):
            bloques.append(parrafo_como(modelo, "", "%d. %s" % (n, x), sangria=True))

    campo("Objetivo", c.get("Objetivo", ""))
    campo("Identificación del Caso de Uso", ficha["cu"])
    campo("Nombre del Caso de Uso", "%s (Espejo %s)" % (nuevo["nombre"], nuevo["pantalla"]))
    campo("Descripción del Caso de Uso", c.get("Descripción", ""))
    lista("Pre-Condición", nuevo["pre"])
    lista("Post Condición", nuevo["post"])
    campo("Actores Primarios", c.get("Actores Primarios", ""))
    campo("Actores Secundarios", c.get("Actores Secundarios", ""))
    campo("Eventos disparadores del caso de uso", c.get("Evento Disparador", ""))
    campo("Puntos de extensión y condiciones", nuevo["extension"])

    campo("Escenario Principal de Éxito", "")
    for paso in escenario:
        bloques.append(parrafo_como(modelo, "", paso, sangria=True))

    campo("Caminos Alternativos", "")
    for titulo, lineas in caminos:
        bloques.append(parrafo_como(modelo, "", titulo, sangria=True))
        for x in lineas:
            bloques.append(parrafo_como(modelo, "", x, sangria=True))

    bloques.append(parrafo_como(modelo, "Prototipo de la Interfaz de usuario", ""))

    # insertar lo nuevo justo antes del parrafo del prototipo (que se conserva),
    # y borrar lo viejo que queda en el medio
    ancla = els[ficha["proto"]] if ficha["proto"] else els[ficha["fin"]]
    viejos = [els[j] for j in range(ficha["ini"] + 1, ficha["fin"])
              if j != ficha["proto"] and els[j].tag.endswith("}p")]
    for b in bloques:
        ancla.addprevious(b)
    for v in viejos:
        v.getparent().remove(v)
    return len(bloques)


def main():
    sys.path.insert(0, os.path.join(RAIZ, "..", "ShopMetrics-Tecnologia",
                                    "documentacion", "scripts", "secuencia"))
    from casos import CASOS as FICHAS
    from generar import ficha_de
    doc = Document(WORD)
    els, fichas = leer_fichas(doc)
    print("fichas halladas en el documento: %d" % len(fichas))
    hechas, sin_datos = 0, []
    solo = set(sys.argv[1:])
    for f in reversed(fichas):              # de atras hacia adelante: los indices no se mueven
        if solo and f["cu"] not in solo:
            continue
        if f["cu"] not in FICHAS:
            sin_datos.append(f["cu"]); continue
        d = FICHAS[f["cu"]]
        esc, cam = ficha_de(f["cu"], d)
        reescribir(doc, els, f, d, esc, cam)
        hechas += 1
    doc.save(WORD)
    print("reescritas: %d" % hechas)
    if sin_datos:
        print("sin contenido nuevo (quedaron como estaban): %s" % ", ".join(sorted(sin_datos)))
    return 0


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    sys.exit(main())
