# -*- coding: utf-8 -*-
"""Dibuja el Gantt de fases del cronograma de desarrollo para el informe.

Lee el mismo XML que abre Microsoft Project, asi la figura dice exactamente lo
que dice el cronograma. Se dibuja a nivel de fase --once barras y el hito--
porque un Gantt de 41 filas no se lee a 159 mm de ancho, que es lo que entra
en la pagina. Cada barra lleva el color del rol que carga mas horas en esa
fase, que es la informacion que importa: quien trabaja, cuando y en paralelo
con que.

    python3 scripts/gantt_project.py
"""
from __future__ import annotations

import datetime as dt
import os
import xml.etree.ElementTree as ET
from collections import defaultdict

import matplotlib
matplotlib.use("Agg")
import matplotlib.ticker
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import Patch

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XML = os.path.join(RAIZ, "project", "ShopMetrics-desarrollo.xml")
SALIDA = os.path.join(RAIZ, "documento", "img_cierre", "project-gantt-fases.png")
NS = "{http://schemas.microsoft.com/project}"

COLOR = {                       # un color por rol; sobrio, se distingue en gris
    "Arquitecto de software (free lance)":        "#3a4a6b",
    "Desarrollador backend (free lance)":         "#5b8fb9",
    "Desarrollador frontend (free lance)":        "#9fbfd8",
    "Ingeniero de machine learning (free lance)": "#c98b4b",
}
CORTO = {k: k.replace(" (free lance)", "") for k in COLOR}


def texto(e, t):
    x = e.find(NS + t)
    return x.text if x is not None else None


def horas(pt):
    h, resto = pt[2:].split("H")
    return int(h) + int(resto.split("M")[0]) / 60


def main() -> int:
    r = ET.parse(XML).getroot()
    tareas = list(r.find(NS + "Tasks"))
    recursos = {int(texto(x, "UID")): texto(x, "Name") for x in r.find(NS + "Resources")}

    # horas por recurso en cada tarea hoja, para colorear la fase
    horas_tarea = defaultdict(lambda: defaultdict(float))
    for a in r.find(NS + "Assignments"):
        horas_tarea[int(texto(a, "TaskUID"))][int(texto(a, "ResourceUID"))] += horas(texto(a, "Work"))

    fases, hito = [], None
    fase_actual = None
    for t in tareas:
        nivel = texto(t, "OutlineLevel")
        if texto(t, "Milestone") == "1":
            hito = (texto(t, "Name"), dt.datetime.fromisoformat(texto(t, "Start")))
            continue
        if texto(t, "Summary") == "1" and nivel == "1":
            fase_actual = dict(nombre=texto(t, "Name"),
                               ini=dt.datetime.fromisoformat(texto(t, "Start")),
                               fin=dt.datetime.fromisoformat(texto(t, "Finish")),
                               horas=defaultdict(float))
            fases.append(fase_actual)
        elif nivel == "2" and fase_actual is not None:
            for u, h in horas_tarea[int(texto(t, "UID"))].items():
                fase_actual["horas"][recursos[u]] += h

    plt.rcParams.update({"font.family": ["Arial", "Liberation Sans", "DejaVu Sans"],
                         "font.size": 9})
    fig, ax = plt.subplots(figsize=(11.2, 4.6), dpi=220)

    n = len(fases)
    for i, f in enumerate(fases):
        y = n - i
        rol = max(f["horas"], key=f["horas"].get)
        ax.barh(y, f["fin"] - f["ini"], left=f["ini"], height=0.62,
                color=COLOR[rol], edgecolor="none")
        total = sum(f["horas"].values())
        ax.text(f["fin"] + dt.timedelta(days=2), y, "%d h" % total,
                va="center", ha="left", fontsize=7.5, color="#444444")
    if hito:
        ax.plot(hito[1], 0, marker="D", color="#222222", markersize=6, zorder=5)
        ax.text(hito[1] + dt.timedelta(days=2), 0, hito[0] + "  " + hito[1].strftime("%d/%m/%Y"),
                va="center", ha="left", fontsize=8, fontweight="bold")

    ax.set_yticks([n - i for i in range(n)] + [0])
    ax.set_yticklabels([f["nombre"] for f in fases] + [""], fontsize=8.5)
    ax.set_ylim(-0.8, n + 0.8)

    ini = min(f["ini"] for f in fases); fin = max(f["fin"] for f in fases)
    ax.set_xlim(ini - dt.timedelta(days=6), fin + dt.timedelta(days=60))
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    # los meses en castellano, sin depender del locale de la maquina
    MESES = ("ene", "feb", "mar", "abr", "may", "jun",
             "jul", "ago", "sep", "oct", "nov", "dic")
    ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(
        lambda x, _: "%s\n%d" % (MESES[mdates.num2date(x).month - 1], mdates.num2date(x).year)))
    ax.tick_params(axis="x", labelsize=8)
    ax.grid(axis="x", color="#dddddd", linewidth=0.7)
    ax.set_axisbelow(True)
    for lado in ("top", "right", "left"):
        ax.spines[lado].set_visible(False)

    ax.legend(handles=[Patch(color=c, label=CORTO[k]) for k, c in COLOR.items()],
              title="Rol que carga más horas en la fase", loc="upper right",
              fontsize=7.5, title_fontsize=7.5, frameon=False)

    fig.tight_layout()
    os.makedirs(os.path.dirname(SALIDA), exist_ok=True)
    fig.savefig(SALIDA, dpi=220)
    print("escrito: %s" % os.path.relpath(SALIDA, RAIZ))
    print("  fases: %d   desde %s hasta %s   hito %s"
          % (n, ini.strftime("%d/%m/%Y"), fin.strftime("%d/%m/%Y"),
             hito[1].strftime("%d/%m/%Y") if hito else "-"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
