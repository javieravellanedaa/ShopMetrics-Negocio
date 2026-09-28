# -*- coding: utf-8 -*-
"""Saca una captura cada vez que cambia lo que se proyecta.

La camara esta fija y la pantalla ocupa la franja de arriba del cuadro. Si se
detectaran cambios sobre el cuadro entero, cada persona que se mueve contaria
como un cambio y saldrian cientos de capturas iguales. Recortando a la pantalla
antes de comparar, lo unico que dispara una captura es que cambie la planilla:
una celda distinta, otra hoja, un cuadro de dialogo.
"""
from __future__ import annotations

import os
import subprocess
import sys

S = os.path.dirname(os.path.abspath(__file__))
VIDEO = ("/Users/javier/Pictures/Photos Library.photoslibrary/originals/7/"
         "75B6B906-251B-4BCA-8595-3F8B91390F82.mov")
DESTINO = os.path.join(S, "pantallas")
UMBRAL = 0.035          # cuanto tiene que cambiar para contar como pantalla nueva


def main() -> int:
    os.makedirs(DESTINO, exist_ok=True)
    filtro = (
        # la pantalla esta en el 42% de arriba; se recorta antes de comparar
        "crop=in_w:in_h*0.42:0:0,"
        "scale=960:-1,"
        "select='gt(scene,%s)',"
        "metadata=print:file=%s/tiempos.txt" % (UMBRAL, S)
    )
    subprocess.run(
        ["ffmpeg", "-y", "-i", VIDEO, "-vf", filtro,
         "-vsync", "vfr", "-q:v", "4",
         os.path.join(DESTINO, "p%04d.jpg")],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # los tiempos quedan en el archivo de metadatos, uno por captura
    tiempos = []
    with open(os.path.join(S, "tiempos.txt")) as f:
        for linea in f:
            if linea.startswith("frame:"):
                for parte in linea.split():
                    if parte.startswith("pts_time:"):
                        tiempos.append(float(parte.split(":")[1]))

    capturas = sorted(os.listdir(DESTINO))
    print("capturas: %d   tiempos: %d" % (len(capturas), len(tiempos)))

    # renombrar a minuto-segundo, que es como se busca despues
    for archivo, t in zip(capturas, tiempos):
        nuevo = "min-%02d-%02d-%02d.jpg" % (t // 3600, (t % 3600) // 60, t % 60)
        os.replace(os.path.join(DESTINO, archivo), os.path.join(DESTINO, nuevo))

    quedan = sorted(os.listdir(DESTINO))
    print("renombradas: %d" % len(quedan))
    if quedan:
        print("  de %s a %s" % (quedan[0], quedan[-1]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
