# -*- coding: utf-8 -*-
"""Suma al informe los casos de uso nuevos y las columnas que el codigo agrego.

Toca cuatro lugares del punto 10.5, que son los que tienen que quedar
nivelados entre si:

  10.5.2  el indice de casos de uso          (una fila por caso de uso)
  10.5.3  las fichas                          (la ficha completa, con sus caminos)
  10.5.4  los diagramas de secuencia          (el titulo y la imagen)
  10.5.7  los diagramas de clases             (idem)
  10.5.8  el diccionario de datos             (una fila por columna nueva)

El formato no se inventa: cada parrafo, cada fila y cada imagen se clonan de
una que ya esta en el documento y se les cambia el texto. Es la unica forma de
que lo agregado no se distinga de lo que ya estaba.

    python3 scripts/agregar_cu.py
"""
from __future__ import annotations

import copy
import os
import re
import sys

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORD = os.path.join(RAIZ, "documento", "STF_Gomez_Javier_E2_v2.docx")
TECNOLOGIA = os.path.join(RAIZ, "..", "ShopMetrics-Tecnologia")
SEQ = os.path.join(TECNOLOGIA, "documentacion", "modelo-ea", "export", "seq")
CLS = os.path.join(TECNOLOGIA, "documentacion", "modelo-ea", "export", "clases")

# Los casos de uso que se agregan, con el modulo que les toca en el indice.
NUEVOS = [
    ("CU-032-001", "032: Padrón de locatarios"),
    ("CU-033-001", "033: Zonas del centro"),
]

# Las columnas que la migracion 008 agrego, por entidad. La entidad se
# identifica por una columna que ya tiene, porque los titulos de las tablas
# del diccionario estan numerados y no nombran la entidad en el texto.
COLUMNAS = {
    ("id", "nombre", "email", "password_hash", "activo"): [
        ("centro_activo_id", "UUID", "FK", "Centro sobre el que está trabajando. Nulo o "
         "apuntando a un centro no asignado vale como «el primero de los suyos»."),
        ("silencio_desde", "TIME", "", "Hora desde la que no quiere recibir avisos. Nula "
         "significa «avisarme siempre»."),
        ("silencio_hasta", "TIME", "", "Hora hasta la que no quiere recibir avisos. El "
         "rango puede cruzar la medianoche."),
    ],
    ("id", "centro_id", "nombre", "rubro", "plan"): [
        # DT-18: la cuenta del portal. La migracion 005 la agrego a la base en
        # la semana 8 --es lo que vuelve cumplible la pre-condicion "el
        # locatario tiene cuenta activa en el portal" del CU-027-001 y del
        # CU-021-001-- y el diccionario nunca la sumo.
        ("usuario_id", "UUID", "FK", "Cuenta del portal con la que entra este locatario. "
         "Opcional y única: un locatario puede existir sin cuenta --lo carga el centro y "
         "todavía no entró-- y una cuenta atañe a lo sumo un locatario."),
        ("estado", "VARCHAR(160)", "", "Estado del locatario (activo, dado de baja). Nulo "
         "cuenta como activo. La baja es lógica: su facturación histórica se conserva."),
        ("onboarding_paso", "SMALLINT", "", "Por dónde va el asistente de alta, para poder "
         "retomarlo donde se dejó."),
    ],
    # DT-05: el paso 2 del CU-018-001 muestra la descripcion del rol y el
    # diagrama de clases la tiene; el diccionario era el unico que no.
    ("id", "nombre"): [
        ("descripcion", "TEXT", "", "Texto que explica el alcance del rol. El paso 2 del "
         "CU-018-001 lo muestra al lado del nombre: dos roles pueden llamarse parecido y "
         "hacer cosas muy distintas."),
    ],
    ("id", "centro_id", "nombre", "superficie"): [
        ("estado", "VARCHAR(160)", "", "Estado de la zona (activa, dada de baja). Nulo "
         "cuenta como activa. No se borra: el histórico de tráfico cuelga de ella."),
    ],
    ("id", "usuario_id", "tipo", "formato", "hash", "fecha", "centro_id"): [
        ("desde", "DATE", "", "Primer día del período exportado. Sin el período el hash no "
         "verifica nada, porque no se puede repetir la misma consulta."),
        ("hasta", "DATE", "", "Último día del período exportado."),
        ("alcance", "VARCHAR(160)", "", "Qué datos incluye, cuando es una exportación para "
         "auditoría externa (datos POS, tráfico, log de auditoría)."),
        ("filas", "INTEGER", "", "Cuántos registros entraron en el archivo."),
    ],
}

# Y las dos columnas de recomendacion, que se identifican aparte porque la
# entidad tiene muchas columnas y la clave de arriba quedaria ilegible.
RECOMENDACION = [
    ("fundamento", "TEXT", "", "Con qué números se sostiene la recomendación. Se guarda y no "
     "se recalcula al leer: la recomendación se tomó sobre los números de un momento."),
    ("detalle", "TEXT", "", "La acción concreta. El tipo dice «horarios» y el detalle dice "
     "qué horario."),
]


# ------------------------------------------------------------------ utilidades


def parrafos(doc):
    return list(doc.element.body.iterchildren())


def texto_de(doc, el):
    return Paragraph(el, doc).text.strip() if el.tag.endswith("}p") else ""


def clonar_despues(modelo, destino):
    """Copia el XML de `modelo` y lo inserta despues de `destino`."""
    nuevo = copy.deepcopy(modelo)
    destino.addnext(nuevo)
    return nuevo


def poner_texto(doc, el, texto):
    """Deja el parrafo con un solo run, conservando el formato del primero."""
    p = Paragraph(el, doc)
    if not p.runs:
        p.add_run(texto)
        return
    p.runs[0].text = texto
    for r in p.runs[1:]:
        r._element.getparent().remove(r._element)


def imagen_de(el, archivo):
    """Reapunta la imagen del parrafo clonado al PNG que corresponde.

    No se puede reemplazar el blob sin tocar las relaciones del paquete, asi
    que se devuelve el nombre del archivo esperado y el llamador avisa si hay
    que cambiarlo a mano. Para los casos de uso nuevos la imagen se agrega con
    add_picture sobre un parrafo vacio, que si crea la relacion.
    """
    return os.path.basename(archivo)


# --------------------------------------------------------- 10.5.4 y 10.5.7


def seccion_de_diagramas(doc, prefijo, carpeta, titulo_de):
    """Agrega los diagramas que falten en 10.5.4 o en 10.5.7.

    Busca el ultimo titulo con ese prefijo, clona su parrafo de titulo, y
    agrega la imagen con add_picture para que el paquete cree la relacion.
    """
    from docx.shared import Cm

    els = parrafos(doc)
    numerados = []
    pat = re.compile(r"^%s\.(\d+)\s+(CU-\d{3}-\d{3})" % re.escape(prefijo))
    for i, e in enumerate(els):
        m = pat.match(texto_de(doc, e))
        if m:
            numerados.append((i, int(m.group(1)), m.group(2)))
    if not numerados:
        return []

    ya = {cu for _i, _n, cu in numerados}
    ultimo_i, ultimo_n, _cu = numerados[-1]
    modelo_titulo = els[ultimo_i]
    # el parrafo de la imagen es el que sigue al titulo
    ancla = els[ultimo_i + 1]

    agregados = []
    n = ultimo_n
    sys.path.insert(0, os.path.join(TECNOLOGIA, "documentacion", "scripts", "secuencia"))
    from casos import CASOS

    for cu in sorted(CASOS):
        if cu in ya:
            continue
        png = os.path.join(carpeta, cu + ".png")
        if not os.path.exists(png):
            continue
        n += 1
        titulo = clonar_despues(modelo_titulo, ancla)
        poner_texto(doc, titulo, "%s.%d %s · %s" % (prefijo, n, cu, CASOS[cu]["nombre"]))
        vacio = clonar_despues(ancla, titulo)
        # se vacia el parrafo clonado y se le pone la imagen nueva
        p = Paragraph(vacio, doc)
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
        p.add_run().add_picture(png, width=Cm(16))
        ancla = vacio
        agregados.append("%s.%d %s" % (prefijo, n, cu))
    return agregados


# ------------------------------------------------------------------ 10.5.2


def fila_del_indice(doc, cu, modulo):
    """Agrega la fila del caso de uso al indice, si no esta."""
    sys.path.insert(0, os.path.join(TECNOLOGIA, "documentacion", "scripts", "secuencia"))
    from casos import CASOS

    for t in doc.tables:
        if not t.rows or t.rows[0].cells[0].text.strip() != "ID":
            continue
        if not t.rows[1].cells[0].text.strip().startswith("CU-"):
            continue
        if any(f.cells[0].text.strip() == cu for f in t.rows):
            return False
        nueva = copy.deepcopy(t.rows[-1]._tr)
        t.rows[-1]._tr.addnext(nueva)
        fila = Table(t._tbl, t._parent).rows[-1]
        for celda, texto in zip(fila.cells, (cu, CASOS[cu]["nombre"], modulo)):
            poner_texto(doc, celda.paragraphs[0]._p, texto)
            for p in celda.paragraphs[1:]:
                p._p.getparent().remove(p._p)
        return True
    return False


# ------------------------------------------------------------------ 10.5.8


def filas_del_diccionario(doc):
    """Agrega al diccionario las columnas que la migracion 008 sumo."""
    agregadas = []
    for t in doc.tables:
        if not t.rows or t.rows[0].cells[0].text.strip() != "Campo":
            continue
        campos = tuple(f.cells[0].text.strip() for f in list(t.rows)[1:])
        nuevas = None
        for clave, filas in COLUMNAS.items():
            # Se compara por prefijo y no por igualdad, porque la tabla crece
            # a medida que se le agregan columnas y exigir el largo exacto
            # haria que el script dejara de encontrarla despues de la primera
            # corrida.
            #
            # Pero un prefijo corto no alcanza para identificar una entidad:
            # ("id", "nombre") es el principio de rol, de centro_comercial y
            # de indicador. Cuando la clave tiene menos de tres campos se
            # exige ademas que los campos que faltan sean solo los que se van
            # a agregar.
            if campos[:len(clave)] != clave:
                continue
            if len(clave) < 3:
                sobrantes = set(campos[len(clave):]) - {c[0] for c in filas}
                if sobrantes:
                    continue
            nuevas = filas
            break
        if nuevas is None and campos and campos[0] == "id" and "rubro_objetivo" in campos:
            nuevas = RECOMENDACION
        if nuevas is None:
            continue
        for campo, tipo, clave, descripcion in nuevas:
            if campo in campos:
                continue
            nueva = copy.deepcopy(t.rows[-1]._tr)
            t.rows[-1]._tr.addnext(nueva)
            fila = Table(t._tbl, t._parent).rows[-1]
            for celda, texto in zip(fila.cells, (campo, tipo, clave, descripcion)):
                poner_texto(doc, celda.paragraphs[0]._p, texto)
                for p in celda.paragraphs[1:]:
                    p._p.getparent().remove(p._p)
            agregadas.append(campo)
    return agregadas


# ------------------------------------------------------------------ 10.5.3


def ficha_nueva(doc, cu):
    """Clona la ficha de la ultima y la reescribe con los datos del caso nuevo.

    Clonar y reescribir, en vez de armar los parrafos desde cero, es lo que
    garantiza que la ficha nueva tenga exactamente el mismo formato: el mismo
    estilo de titulo, la misma sangria, el mismo parrafo de prototipo con su
    hipervinculo. Despues `reescribir()` --el de reestructurar_cu.py, que es
    el que ya le dio forma a las otras 32-- le pone el contenido.
    """
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, os.path.join(TECNOLOGIA, "documentacion", "scripts", "secuencia"))
    from casos import CASOS
    from generar import ficha_de
    from reestructurar_cu import leer_fichas, reescribir

    _els, fichas = leer_fichas(doc)
    if any(f["cu"] == cu for f in fichas):
        return False
    ultima = fichas[-1]
    els = parrafos(doc)

    # El bloque entero de la ultima ficha, clonado y pegado a continuacion.
    bloque = [els[j] for j in range(ultima["ini"], ultima["fin"])]
    ancla = bloque[-1]
    copias = []
    for el in bloque:
        ancla = clonar_despues(el, ancla)
        copias.append(ancla)
    poner_texto(doc, copias[0], "%s · %s" % (cu, CASOS[cu]["nombre"]))

    # Y ahora se la reescribe como a cualquier otra. Los cuatro campos que
    # `reescribir` lee del documento --objetivo, descripcion, actores y
    # disparador-- vienen clonados de la ficha anterior, asi que se
    # reemplazan por los de la fuente unica antes de escribir.
    _els, fichas = leer_fichas(doc)
    nueva = next(f for f in fichas if f["cu"] == cu)
    d = CASOS[cu]
    nueva["campos"] = {
        "Objetivo": d.get("objetivo", ""),
        "Descripción": d.get("descripcion", ""),
        "Actores Primarios": d.get("actores_primarios", d["actor"]),
        "Actores Secundarios": d.get("actores_secundarios", ""),
        "Evento Disparador": d.get("disparador", ""),
    }
    esc, cam = ficha_de(cu, d)
    reescribir(doc, parrafos(doc), nueva, d, esc, cam)
    return True


# ------------------------------------------------------------------ principal


def main():
    doc = Document(WORD)

    puestas = [cu for cu, modulo in NUEVOS if fila_del_indice(doc, cu, modulo)]
    print("10.5.2  filas agregadas al índice: %s" % (", ".join(puestas) or "ninguna"))

    seq = seccion_de_diagramas(doc, "10.5.4", SEQ, None)
    print("10.5.4  diagramas agregados: %s" % (", ".join(seq) or "ninguno"))

    cls = seccion_de_diagramas(doc, "10.5.7", CLS, None)
    print("10.5.7  diagramas agregados: %s" % (", ".join(cls) or "ninguno"))

    fichas = [cu for cu, _m in NUEVOS if ficha_nueva(doc, cu)]
    print("10.5.3  fichas agregadas: %s" % (", ".join(fichas) or "ninguna"))

    columnas = filas_del_diccionario(doc)
    print("10.5.8  columnas agregadas: %s" % (", ".join(columnas) or "ninguna"))

    doc.save(WORD)
    return 0


if __name__ == "__main__":
    sys.exit(main())
