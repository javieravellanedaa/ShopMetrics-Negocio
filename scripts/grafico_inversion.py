# -*- coding: utf-8 -*-
"""Dibuja el grafico de inversion por anio, que va pegado en Mod. inversion
del presupuesto (como en la plantilla de la catedra) y como figura del 8.5.

    python3 scripts/grafico_inversion.py
"""
from __future__ import annotations

import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as tk
import openpyxl

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(RAIZ, "documento", "Presupuesto financiero ShopMetrics.xlsx")
SALIDA = os.path.join(RAIZ, "documento", "img_punto8", "INV04_inversion_grafico.png")


def main() -> int:
    v = openpyxl.load_workbook(XLSX, data_only=True)["Mod. inversión"]
    vals = [v[c + "5"].value or 0 for c in "GHIJ"]
    plt.rcParams.update({"font.family": ["Arial", "DejaVu Sans"], "font.size": 9})
    fig, ax = plt.subplots(figsize=(6, 3.6), dpi=200)
    barras = ax.bar(["Inversión inicial", "2026", "2027", "2028"], vals, color="#4a86e8", width=0.6)
    for b, y in zip(barras, vals):
        ax.text(b.get_x() + b.get_width() / 2, y, "US$ {:,.0f}".format(y).replace(",", "."),
                ha="center", va="bottom", fontsize=8)
    ax.set_title("Inversión por año", loc="left", fontsize=11, color="#555555")
    ax.yaxis.set_major_formatter(tk.FuncFormatter(lambda x, _: "US$ {:,.0f}".format(x).replace(",", ".")))
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.grid(axis="y", color="#e5e5e5"); ax.set_axisbelow(True)
    fig.tight_layout(); fig.savefig(SALIDA, dpi=200)
    print("escrito:", os.path.relpath(SALIDA, RAIZ), [round(x) for x in vals])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
