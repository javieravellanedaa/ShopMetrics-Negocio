# -*- coding: utf-8 -*-
"""Devuelve al presupuesto los objetos de la plantilla que openpyxl no conserva.

openpyxl es lo que se usa para escribir formulas y valores, pero al guardar
tira las formas y las imagenes de cada hoja: los botones "Volver al indice",
el logo de la portada, y ademas reescribe los graficos a su manera. Este script
se corre DESPUES de cualquier edicion con openpyxl y vuelve a poner esas partes
tal como estan en la plantilla, sin tocar ninguna celda.

Las partes viven en scripts/objetos-plantilla/, sacadas del archivo original
del 22/09 que esta en git. Se quitaron las dos imagenes que la plantilla traia
en Mod. inversion porque son del ejemplo del profesor --un grafico con sus
numeros y la captura de su Project--, y en este archivo van las de ShopMetrics.

    python3 scripts/restaurar_objetos.py documento/Presupuesto\\ financiero\\ ShopMetrics.xlsx

Escribe encima del archivo y deja una copia .antes.xlsx por si algo sale mal.
"""
from __future__ import annotations

import io
import os
import re
import shutil
import sys
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
DONANTE = os.path.join(AQUI, "objetos-plantilla")

TIPOS = {
    "drawing": "application/vnd.openxmlformats-officedocument.drawing+xml",
    "chart": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
}
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
# el elemento <drawing> tiene que ir antes de estos, si estan
DESPUES_DE_DRAWING = ("<legacyDrawing", "<legacyDrawingHF", "<picture", "<oleObjects",
                      "<controls", "<webPublishItems", "<tableParts", "<extLst")


def leer_mapa():
    mapa = {}
    with io.open(os.path.join(DONANTE, "mapa.txt"), encoding="utf-8") as f:
        for linea in f:
            if linea.strip():
                hoja, dib = linea.rstrip("\n").split("\t")
                mapa[hoja] = dib
    return mapa


def hojas_del_libro(partes):
    """[(nombre, 'xl/worksheets/sheetN.xml')] en el orden del libro."""
    wb = partes["xl/workbook.xml"].decode("utf-8")
    rels = partes["xl/_rels/workbook.xml.rels"].decode("utf-8")
    destino = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        e = m.group(0)
        i = re.search(r'Id="([^"]+)"', e); t = re.search(r'Target="([^"]+)"', e)
        if i and t and "worksheets/" in t.group(1):
            destino[i.group(1)] = "xl/" + t.group(1).lstrip("/").replace("xl/", "")
    out = []
    for m in re.finditer(r"<sheet\b[^>]*>", wb):
        e = m.group(0)
        n = re.search(r'name="([^"]+)"', e); r = re.search(r'r:id="([^"]+)"', e)
        out.append((n.group(1).replace("&amp;", "&"), destino[r.group(1)]))
    return out


def main(ruta: str) -> int:
    copia = ruta.replace(".xlsx", ".antes.xlsx")
    shutil.copy(ruta, copia)
    with zipfile.ZipFile(ruta) as z:
        partes = {n: z.read(n) for n in z.namelist()}
        orden = z.namelist()

    # 1) afuera lo que openpyxl escribio de dibujos, graficos e imagenes
    #    (los comentarios y su vml se quedan: esos si los maneja bien)
    for n in list(partes):
        if (n.startswith("xl/drawings/") and "vmlDrawing" not in n) or n.startswith("xl/charts/") \
                or n.startswith("xl/media/"):
            del partes[n]

    # 2) adentro las partes de la plantilla
    for carpeta in ("drawings", "charts", "media"):
        base = os.path.join(DONANTE, carpeta)
        for raiz, _, archivos in os.walk(base):
            for a in archivos:
                rel = os.path.relpath(os.path.join(raiz, a), DONANTE).replace(os.sep, "/")
                partes["xl/" + rel] = open(os.path.join(raiz, a), "rb").read()

    # 3) cada hoja apunta a su dibujo
    mapa = leer_mapa()
    puestos = 0
    for nombre, ruta_hoja in hojas_del_libro(partes):
        xml = partes[ruta_hoja].decode("utf-8")
        rels_ruta = ruta_hoja.replace("worksheets/", "worksheets/_rels/") + ".rels"
        rels = partes.get(rels_ruta, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                          b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>').decode("utf-8")
        # sacar la referencia de dibujo que haya dejado openpyxl
        xml = re.sub(r"<drawing\b[^>]*/>", "", xml)
        rels = re.sub(r'<Relationship\b[^>]*Type="%s"[^>]*/>' % re.escape(REL_DRAWING), "", rels)
        dib = mapa.get(nombre)
        if dib:
            usados = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
            rid = "rId%d" % (max(usados) + 1 if usados else 1)
            rels = rels.replace("</Relationships>",
                                '<Relationship Id="%s" Type="%s" Target="../drawings/%s"/></Relationships>'
                                % (rid, REL_DRAWING, dib))
            etiqueta = '<drawing r:id="%s"/>' % rid
            pos = min([xml.find(t) for t in DESPUES_DE_DRAWING if xml.find(t) >= 0] + [xml.rfind("</worksheet>")])
            xml = xml[:pos] + etiqueta + xml[pos:]
            if 'xmlns:r=' not in xml[:600]:
                xml = xml.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
            puestos += 1
        partes[ruta_hoja] = xml.encode("utf-8")
        partes[rels_ruta] = rels.encode("utf-8")

    # 4) tipos de contenido: sacar los de partes que ya no estan, agregar los nuevos
    ct = partes["[Content_Types].xml"].decode("utf-8")
    ct = re.sub(r'<Override\b[^>]*PartName="/xl/(?:drawings/(?!vmlDrawing)|charts/|media/)[^"]*"[^>]*/>', "", ct)
    nuevos = []
    for n in partes:
        if n.startswith("xl/drawings/drawing") and n.endswith(".xml"):
            nuevos.append('<Override PartName="/%s" ContentType="%s"/>' % (n, TIPOS["drawing"]))
        elif n.startswith("xl/charts/chart") and n.endswith(".xml"):
            nuevos.append('<Override PartName="/%s" ContentType="%s"/>' % (n, TIPOS["chart"]))
    if 'Extension="png"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="png" ContentType="image/png"/><Default ', 1)
    ct = ct.replace("</Types>", "".join(nuevos) + "</Types>")
    partes["[Content_Types].xml"] = ct.encode("utf-8")

    # 5) escribir, [Content_Types].xml primero como corresponde
    orden_final = ["[Content_Types].xml"] + [n for n in orden if n in partes and n != "[Content_Types].xml"] \
                  + [n for n in partes if n not in orden]
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        for n in orden_final:
            if n in partes:
                z.write if False else z.writestr(n, partes[n])

    with zipfile.ZipFile(ruta) as z:
        assert z.testzip() is None
        nombres = z.namelist()
    print("hojas con dibujo: %d   dibujos: %d   graficos: %d   imagenes: %d"
          % (puestos, sum(1 for n in nombres if re.match(r"xl/drawings/drawing\d+\.xml$", n)),
             sum(1 for n in nombres if re.match(r"xl/charts/chart\d+\.xml$", n)),
             sum(1 for n in nombres if n.startswith("xl/media/"))))
    print("copia previa en: %s" % os.path.basename(copia))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
