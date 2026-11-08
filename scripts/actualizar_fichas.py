# -*- coding: utf-8 -*-
"""Pone al dia fichas del 10.5.3 y diagramas del 10.5.4 y 10.5.7 que ya existen.

A diferencia de reestructurar_cu.py --que armo las fichas una vez, leyendo los
campos con sus nombres viejos-- esto toca solo lo que cambia cuando el sistema
cambia: el parrafo de puntos de extension, el escenario principal y los caminos
alternativos. Objetivo, descripcion, actores y disparador quedan como estan.
Cada parrafo nuevo se clona de uno que ya esta en la ficha, para que no se
distinga de lo anterior. Las imagenes se reemplazan en su lugar, sin mover
nada del documento.

    python3 scripts/actualizar_fichas.py --seco CU-031-001 ...   informa y no guarda
    python3 scripts/actualizar_fichas.py CU-031-001 ...          actualiza
"""
from __future__ import annotations

import copy
import os
import re
import sys

from docx import Document
from docx.text.paragraph import Paragraph

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORD = os.path.join(RAIZ, "documento", "STF_Gomez_Javier_E2_v2.docx")
TECNOLOGIA = os.path.join(RAIZ, "..", "ShopMetrics-Tecnologia", "documentacion")
SEQ = os.path.join(TECNOLOGIA, "modelo-ea", "export", "seq")
CLS = os.path.join(TECNOLOGIA, "modelo-ea", "export", "clases")
sys.path.insert(0, os.path.join(TECNOLOGIA, "scripts", "secuencia"))

CABECERA = re.compile(r"^(CU-\d{3}-\d{3})\s*·\s*(.+)$")


def _texto(p):
    return p.text.strip()


def _reemplazar_texto(p: Paragraph, etiqueta: str, texto: str) -> None:
    """Deja la etiqueta en negrita y el texto despues, en el formato del parrafo."""
    runs = p.runs
    modelo = copy.deepcopy(runs[0]._r) if runs else None
    for r in runs:
        r._r.getparent().remove(r._r)
    a = p.add_run(etiqueta + ": ")
    a.bold = True
    p.add_run(texto)


def _clonar(modelo: Paragraph, texto: str, negrita: str = ""):
    nuevo = copy.deepcopy(modelo._p)
    for hijo in list(nuevo):
        if hijo.tag.endswith("}r") or hijo.tag.endswith("}hyperlink"):
            nuevo.remove(hijo)
    p = Paragraph(nuevo, modelo._parent)
    if negrita:
        r = p.add_run(negrita)
        r.bold = True
    if texto:
        p.add_run(texto)
    return nuevo


def fichas(doc):
    pars = doc.paragraphs
    marcas = [(i, m.group(1)) for i, p in enumerate(pars) if (m := CABECERA.match(_texto(p)))]
    # la primera aparicion de cada CU con "Objetivo" despues es la ficha
    salida = {}
    for k, (i, cu) in enumerate(marcas):
        fin = marcas[k + 1][0] if k + 1 < len(marcas) else len(pars)
        cuerpo = pars[i + 1:fin]
        if any(_texto(p).startswith("Objetivo") for p in cuerpo[:3]) and cu not in salida:
            salida[cu] = (i, fin)
    return salida


def actualizar_ficha(doc, cu, ini, fin, datos, seco):
    from generar import ficha_de
    pars = doc.paragraphs[ini:fin]
    ext = next(p for p in pars if _texto(p).startswith("Puntos de extensión"))
    esc_i = next(k for k, p in enumerate(pars) if _texto(p).startswith("Escenario Principal"))
    cam_i = next(k for k, p in enumerate(pars) if _texto(p).startswith("Caminos Alternativos"))
    pro_i = next(k for k, p in enumerate(pars) if _texto(p).startswith("Prototipo"))
    paso_modelo = pars[esc_i + 1]
    camino_modelo = pars[cam_i + 1] if cam_i + 1 < pro_i else paso_modelo
    escenario, caminos = ficha_de(cu, datos)
    print("  %s: %d pasos en el Word -> %d; caminos %d -> %d" % (
        cu, cam_i - esc_i - 1, len(escenario),
        sum(1 for p in pars[cam_i + 1:pro_i] if _texto(p).startswith("Camino")),
        len(caminos)))
    if seco:
        return
    _reemplazar_texto(ext, "Puntos de extensión y condiciones", datos["extension"])
    # se borra lo viejo entre "Escenario" y "Prototipo", salvo el rotulo de caminos
    viejos = [p for p in pars[esc_i + 1:pro_i] if p is not pars[cam_i]]
    ancla_caminos = pars[cam_i]._p
    nuevos_pasos = [_clonar(paso_modelo, x) for x in escenario]
    nuevos_caminos = []
    for titulo, lineas in caminos:
        nuevos_caminos.append(_clonar(camino_modelo, titulo))
        nuevos_caminos += [_clonar(camino_modelo, x) for x in lineas]
    for n in nuevos_pasos:
        ancla_caminos.addprevious(n)
    ancla_prototipo = pars[pro_i]._p
    for n in nuevos_caminos:
        ancla_prototipo.addprevious(n)
    for v in viejos:
        v._p.getparent().remove(v._p)


def reemplazar_imagen(doc, cu, seccion, png, seco):
    """La imagen bajo el titulo numerado del CU en la seccion: «10.5.4.32 CU-031-001 · ...»."""
    pars = doc.paragraphs
    titulo = re.compile(r"^%s\.\d+ %s " % (re.escape(seccion), re.escape(cu)))
    for i, p in enumerate(pars):
        if titulo.match(_texto(p)):
            for q in pars[i + 1:i + 3]:
                blips = q._p.xpath(".//a:blip")
                if blips:
                    rid = blips[0].get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    parte = doc.part.related_parts[rid]
                    print("  %s en %s: imagen %s (%d bytes -> %d)" % (
                        cu, seccion, rid, len(parte.blob), os.path.getsize(png)))
                    if not seco:
                        import io
                        from PIL import Image
                        viejo = Image.open(io.BytesIO(parte.blob)).size
                        parte._blob = open(png, "rb").read()
                        _ajustar_alto(q, png, viejo)
                    return True
    print("  %s en %s: NO se encontró la imagen" % (cu, seccion))
    return False


def _ajustar_alto(parrafo, png, viejo):
    """Le da al recuadro la proporcion de la imagen nueva sin agrandarlo: si el
    diagrama crecio --un camino alternativo mas-- se conserva el alto que
    tenia y se angosta, porque mas alto se sale de la pagina y se corta."""
    from PIL import Image
    ancho, alto = Image.open(png).size
    v_ancho, v_alto = viejo
    for ext in parrafo._p.xpath(".//wp:extent | .//a:ext"):
        cx, cy = int(ext.get("cx")), int(ext.get("cy"))
        alto_viejo = cx * v_alto / v_ancho
        nuevo_cy = cx * alto / ancho
        if nuevo_cy > alto_viejo:
            ext.set("cy", str(int(alto_viejo)))
            ext.set("cx", str(int(alto_viejo * ancho / alto)))
        else:
            ext.set("cy", str(int(nuevo_cy)))


# Los diagramas que cambiaron. Las fichas pueden cambiar solo en el texto de
# extension sin tocar el diagrama; estos son los que se redibujaron.
SECUENCIAS = ["CU-022-001", "CU-024-001", "CU-003-001", "CU-004-001", "CU-034-001", "CU-008-001", "CU-009-001", "CU-012-001", "CU-027-001", "CU-031-001"]
CLASES = ["CU-034-001", "CU-009-001", "CU-012-001", "CU-031-001"]


def main(argumentos):
    seco = "--seco" in argumentos
    pedidos = [a for a in argumentos if a.startswith("CU-")]
    from casos import CASOS
    doc = Document(WORD)
    ubicadas = fichas(doc)
    for cu in reversed(sorted(pedidos)):       # de atras hacia adelante
        if cu not in ubicadas:
            print("  %s: no está en el Word" % cu)
            continue
        actualizar_ficha(doc, cu, *ubicadas[cu], CASOS[cu], seco)
    for cu in pedidos:
        if cu in SECUENCIAS:
            reemplazar_imagen(doc, cu, "10.5.4", os.path.join(SEQ, cu + ".png"), seco)
        if cu in CLASES:
            reemplazar_imagen(doc, cu, "10.5.7", os.path.join(CLS, cu + ".png"), seco)
    if not seco:
        doc.save(WORD)
        print("guardado")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
