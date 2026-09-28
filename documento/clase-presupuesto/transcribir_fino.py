# -*- coding: utf-8 -*-
"""Transcribe la clase entera tomando precauciones contra la perdida de texto.

La primera pasada perdio cien segundos: el modelo se trabo repitiendo "No se"
y nadie se hubiera enterado, porque una transcripcion trabada se ve igual que
una transcripcion de un tramo en silencio. Este proceso toma cinco recaudos:

1. Modelo grande en vez del turbo. Es mas lento pero entiende mejor un audio
   de aula con ruido y varias personas hablando.
2. Audio filtrado: se corta lo que esta por debajo de la voz, se baja el ruido
   de fondo y se nivela el volumen.
3. Ventanas con solape. Si una ventana se traba, la de al lado cubre el mismo
   tramo, y una frase cortada al medio aparece entera en alguna de las dos.
4. Escape de temperatura. Es lo que rompe el bucle de repeticion: si el texto
   sale sospechoso el modelo reintenta con mas azar.
5. Control posterior: se buscan bucles y huecos y se informan, en vez de
   dejarlos pasar en silencio.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys

import mlx_whisper

S = os.path.dirname(os.path.abspath(__file__))
FUENTE = os.path.join(S, "audio.wav")
LIMPIO = os.path.join(S, "audio-limpio.wav")
MODELO = "mlx-community/whisper-large-v3-mlx"

VENTANA = 300          # segundos de audio por pasada
SOLAPE = 30            # cuanto se pisan dos ventanas seguidas


def preparar() -> float:
    """Filtra el audio y devuelve su duracion."""
    subprocess.run(
        ["ffmpeg", "-y", "-i", FUENTE,
         "-af", "highpass=f=80,afftdn=nf=-25,dynaudnorm=f=150:g=15",
         "-ar", "16000", "-ac", "1", LIMPIO],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    salida = subprocess.run(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration",
         "-of", "csv=p=0", LIMPIO],
        capture_output=True, text=True, check=True)
    return float(salida.stdout.strip())


def trozo(desde: float, hasta: float) -> str:
    ruta = os.path.join(S, "_trozo.wav")
    subprocess.run(["ffmpeg", "-y", "-ss", str(desde), "-to", str(hasta),
                    "-i", LIMPIO, ruta],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return ruta


def hms(t: float) -> str:
    return "%02d:%02d:%02d" % (t // 3600, (t % 3600) // 60, t % 60)


def main() -> int:
    dur = preparar()
    print("duracion: %s" % hms(dur), flush=True)

    segmentos = []
    inicio = 0.0
    n = 0
    while inicio < dur:
        fin = min(inicio + VENTANA, dur)
        ruta = trozo(inicio, fin)
        r = mlx_whisper.transcribe(
            ruta, path_or_hf_repo=MODELO, language="es", verbose=False,
            temperature=(0.0, 0.2, 0.4, 0.6, 0.8, 1.0),
            condition_on_previous_text=False,
            compression_ratio_threshold=2.2,
            logprob_threshold=-1.0,
            no_speech_threshold=0.5,
            word_timestamps=True)
        for s in r["segments"]:
            s["start"] += inicio
            s["end"] += inicio
            for w in s.get("words", []):
                w["start"] += inicio
                w["end"] += inicio
            segmentos.append(s)
        n += 1
        print("  ventana %2d  %s - %s  (%d segmentos)"
              % (n, hms(inicio), hms(fin), len(r["segments"])), flush=True)
        if fin >= dur:
            break
        inicio = fin - SOLAPE

    # --- unir: donde dos ventanas se pisan, se queda el texto de la primera
    segmentos.sort(key=lambda s: s["start"])
    unidos = []
    for s in segmentos:
        t = s["text"].strip()
        if not t:
            continue
        if unidos:
            ult = unidos[-1]
            mismo = t.lower() == ult["text"].strip().lower()
            pisado = s["start"] < ult["end"] - 0.6
            if mismo and s["start"] < ult["end"] + 3:
                continue
            if pisado and mismo:
                continue
        unidos.append(s)

    json.dump(unidos, open(os.path.join(S, "fino.json"), "w"), ensure_ascii=False)

    with open(os.path.join(S, "fino.txt"), "w") as f:
        for s in unidos:
            f.write("[%s] %s\n" % (hms(s["start"]), s["text"].strip()))

    # --- controles
    print("\ncontroles", flush=True)
    bucles, previo, repes = [], None, 0
    for s in unidos:
        t = s["text"].strip().lower()
        if t == previo:
            repes += 1
        else:
            if repes >= 4:
                bucles.append((previo, repes))
            repes, previo = 0, t
    if repes >= 4:
        bucles.append((previo, repes))
    print("  bucles de repeticion: %d" % len(bucles))
    for t, c in bucles[:6]:
        print("     %dx  %s" % (c + 1, t[:60]))

    huecos = []
    for a, b in zip(unidos, unidos[1:]):
        if b["start"] - a["end"] > 12:
            huecos.append((a["end"], b["start"]))
    print("  huecos de mas de 12 s: %d" % len(huecos))
    for a, b in huecos[:8]:
        print("     %s -> %s  (%.0f s)" % (hms(a), hms(b), b - a))

    palabras = sum(len(s["text"].split()) for s in unidos)
    print("\nsegmentos: %d   palabras: %d" % (len(unidos), palabras))
    return 0


if __name__ == "__main__":
    sys.exit(main())
