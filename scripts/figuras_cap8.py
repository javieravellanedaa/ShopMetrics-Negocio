# -*- coding: utf-8 -*-
"""Genera las figuras del capitulo 8 del informe a partir del presupuesto.

Las figuras originales eran capturas de pantalla de Excel. Cuando el modelo
cambia hay que rehacerlas todas, asi que se generan desde la planilla con
render_hoja (recortes de rango con el formato de la celda) y matplotlib (los
graficos, calcando los de Excel). Se sobreescriben los archivos con el mismo
nombre que usa el Word, y se agregan los de los puntos 8.5, 8.6 y 8.7.

    python3 scripts/figuras_cap8.py
"""
from __future__ import annotations

import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as tk
import openpyxl

from render_hoja import rango, ARCHIVO, FUENTE

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG = os.path.join(RAIZ, "documento", "img_punto8")
EXCEL = ["#4a86e8", "#e34234", "#f1c232", "#3e9e4d", "#f57c00", "#3fbfbf", "#7b4fa0"]
plt.rcParams.update({"font.family": [FUENTE, "Arial", "DejaVu Sans"], "font.size": 9})


def usd(x, _=None):
    return "US$ {:,.0f}".format(x).replace(",", ".")


def guardar(fig, nombre):
    fig.savefig(os.path.join(IMG, nombre), dpi=200, facecolor="white", bbox_inches="tight")
    plt.close(fig)


def marco(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.grid(axis="y", color="#e5e5e5"); ax.set_axisbelow(True)


def main() -> int:
    v = openpyxl.load_workbook(ARCHIVO, data_only=True)
    h, mi, me, pf, inv, amo, an = (v["Hipótesis"], v["Mod. ingresos"], v["Mod. egresos"], v["Presupuesto financiero"],
                                   v["Mod. inversión"], v["Amortizaciones"], v["Anexo capacidad operativa"])
    anios = ["2026", "2027", "2028"]

    # ---------------------------------------------------------- 8.2 hipotesis
    rango("Hipótesis", "B3:E12", f"{IMG}/01_hipotesis_negocio_mercado.png")
    rango("Hipótesis", "B15:E16", f"{IMG}/02a_participacion_texto.png")
    rango("Hipótesis", "A22:E26", f"{IMG}/02b_participacion_tabla.png")
    rango("Hipótesis", "B56:E60", f"{IMG}/03a_precios_altas.png")
    rango("Hipótesis", "B61:E63", f"{IMG}/03b_precios_abonos.png")
    fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=200)
    mercado = [h["D%d" % r].value for r in (24, 25, 26)]; fact = [h["E%d" % r].value for r in (24, 25, 26)]
    xs = range(3)
    ax.bar([x - 0.2 for x in xs], mercado, 0.38, color=EXCEL[0], label="Valor del mercado captado (USD)")
    ax.bar([x + 0.2 for x in xs], fact, 0.38, color=EXCEL[4], label="Facturación proyectada (USD)")
    ax.set_xticks(list(xs)); ax.set_xticklabels(anios); ax.yaxis.set_major_formatter(tk.FuncFormatter(usd))
    ax.set_title("Evolución de la participación en el mercado", fontsize=11, loc="left"); ax.legend(frameon=False, fontsize=8); marco(ax)
    guardar(fig, "02c_participacion_grafico.png")

    # ---------------------------------------------------------- 8.3 ingresos
    rango("Mod. ingresos", "B15:E30", f"{IMG}/04_ingresos_tabla.png", quitar_vacias=True)
    etiquetas = [mi["B%d" % r].value for r in range(17, 23)]
    for i, col in enumerate("CDE"):
        vals = [mi["%s%d" % (col, r)].value or 0 for r in range(17, 23)]
        fig, ax = plt.subplots(figsize=(5.4, 3.6), dpi=200)
        ax.pie(vals, colors=EXCEL[:6], autopct=lambda p: "%d%%" % round(p) if p >= 3 else "", pctdistance=0.78,
               startangle=90, counterclock=False, textprops={"fontsize": 8})
        ax.set_title("Distribución de ingresos %s" % anios[i], fontsize=11)
        ax.legend(etiquetas, loc="lower center", bbox_to_anchor=(0.5, -0.32), ncol=2, frameon=False, fontsize=7.5)
        guardar(fig, "05_ingresos_grafico_%d.png" % (i + 1))

    # ---------------------------------------------------------- 8.4 egresos
    rango("Mod. egresos", "A8:D16", f"{IMG}/06_egresos_tabla.png")
    rango("Mod. egresos", "A18:D22", f"{IMG}/07_egresos_composicion.png")
    fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=200)
    base = [0, 0, 0]
    for i, (r, etq) in enumerate(((10, "Costos fijos"), (11, "Costos variables"), (12, "Costos de RRHH"))):
        vals = [me["%s%d" % (c, r)].value or 0 for c in "BCD"]
        ax.bar(anios, vals, 0.55, bottom=base, color=EXCEL[i], label=etq); base = [b + x for b, x in zip(base, vals)]
    ax.yaxis.set_major_formatter(tk.FuncFormatter(usd)); ax.set_title("Modelo de egresos", fontsize=11, loc="left")
    ax.legend(frameon=False, fontsize=8); marco(ax); guardar(fig, "15_egresos_grafico_1.png")
    fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=200)
    ing = [me["%s14" % c].value or 0 for c in "BCD"]; egr = [me["%s13" % c].value or 0 for c in "BCD"]
    ax.plot(anios, ing, marker="o", color=EXCEL[0], label="Ingresos proyectados"); ax.plot(anios, egr, marker="o", color=EXCEL[1], label="Egresos totales")
    for x, y in zip(anios, ing): ax.annotate(usd(y), (x, y), textcoords="offset points", xytext=(0, 7), ha="center", fontsize=7.5)
    for x, y in zip(anios, egr): ax.annotate(usd(y), (x, y), textcoords="offset points", xytext=(0, -13), ha="center", fontsize=7.5)
    ax.yaxis.set_major_formatter(tk.FuncFormatter(usd)); ax.set_title("Ingresos vs. egresos", fontsize=11, loc="left")
    ax.legend(frameon=False, fontsize=8); marco(ax); guardar(fig, "16_egresos_grafico_2.png")

    # ---------------------------------------------------------- 8.4.1 anexo: horas hombre y conclusion
    s1 = [1, 2] + list(range(3, 15)); s2 = [1, 2] + list(range(15, 27))
    for k, (r0, r1) in enumerate(((78, 88), (89, 99), (100, 110))):
        rango("Anexo capacidad operativa", "A%d:Z%d" % (r0, r1), f"{IMG}/AN%02d_hh_%s_s1.png" % (11 + 2 * k, anios[k]), columnas=s1, min_col_px=30)
        rango("Anexo capacidad operativa", "A%d:Z%d" % (r0, r1), f"{IMG}/AN%02d_hh_%s_s2.png" % (12 + 2 * k, anios[k]), columnas=s2, min_col_px=30)
    rango("Anexo capacidad operativa", "A112:N113", f"{IMG}/AN17_conclusion.png")

    # ---------------------------------------------------------- 8.5 inversion
    rango("Mod. inversión", "B9:F28", f"{IMG}/INV01_inversion_anio0.png", quitar_vacias=True)
    rango("Mod. inversión", "B30:E81", f"{IMG}/INV02_inversion_2026_2028.png", quitar_vacias=True)
    rango("Mod. inversión", "G3:J5", f"{IMG}/INV03_inversion_resumen.png")
    # INV04 (inversion por anio) lo escribe restaurar_objetos al pegarlo en el
    # Excel, e INV05 (Gantt) lo escribe gantt_project.py; los dos ya viven aca
    import shutil
    shutil.copy(os.path.join(RAIZ, "project", "capturas", "hoja-de-recursos.png"), f"{IMG}/INV06_project_recursos.png")

    # ---------------------------------------------------------- 8.6 amortizaciones
    rango("Amortizaciones", "B8:M27", f"{IMG}/AMO01_amortizaciones.png", quitar_vacias=True, min_col_px=34)
    rango("Amortizaciones", "I3:K5", f"{IMG}/AMO02_amortizaciones_resumen.png")

    # ---------------------------------------------------------- 8.7 presupuesto financiero
    rango("Presupuesto financiero", "B11:G22", f"{IMG}/PRE01_presupuesto.png")
    rango("Presupuesto financiero", "B24:G29", f"{IMG}/PRE02_ratios.png")
    rango("Presupuesto financiero", "J22:N23", f"{IMG}/PRE03_monto_imponible.png")
    rango("Presupuesto financiero", "F33:G35", f"{IMG}/PRE04_van_tir.png")
    fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=200)
    ff = [pf["%s22" % c].value or 0 for c in "DEFG"]; acum = []; s = 0
    for x in ff: s += x; acum.append(s)
    et = ["Año 0", "2026", "2027", "2028"]
    ax.bar(et, ff, 0.55, color=[EXCEL[1] if x < 0 else EXCEL[3] for x in ff], label="Flujo de fondos del ejercicio")
    ax.plot(et, acum, marker="o", color="#333333", label="Flujo acumulado")
    for x, y in zip(et, acum): ax.annotate(usd(y), (x, y), textcoords="offset points", xytext=(0, 8), ha="center", fontsize=7.5)
    ax.axhline(0, color="#999999", linewidth=0.8); ax.yaxis.set_major_formatter(tk.FuncFormatter(usd))
    ax.set_title("Flujo de fondos y acumulado", fontsize=11, loc="left"); ax.legend(frameon=False, fontsize=8); marco(ax)
    guardar(fig, "PRE05_flujo_grafico.png")
    print("figuras generadas en", os.path.relpath(IMG, RAIZ))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
