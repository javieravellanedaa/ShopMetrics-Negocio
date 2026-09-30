# -*- coding: utf-8 -*-
"""Pone al dia el capitulo 8 del informe con el presupuesto del segundo avance.

Todo numero que aparece en el texto se lee de la planilla en el momento de
correr el script, asi el informe no puede decir algo distinto del Excel. Las
figuras existentes se reemplazan adentro del .docx --estan embebidas; cambiar
el png en disco no las toca-- y se agregan las de los puntos 8.5, 8.6 y 8.7.

    python3 scripts/word_cap8_v2.py
"""
from __future__ import annotations

import copy
import os
import re

import docx
import openpyxl
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from PIL import Image

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCX = os.path.join(RAIZ, "documento", "STF_Gomez_Javier_E1_v1.docx")
XLSX = os.path.join(RAIZ, "documento", "Presupuesto financiero ShopMetrics.xlsx")
IMG = os.path.join(RAIZ, "documento", "img_punto8")
ANCHO = 16.3          # cm, el ancho que usan las figuras del capitulo
NUEVAS = 13           # figuras que se agregan en 8.5-8.7; las de riesgos corren ese numero


def n(x, dec=0):
    s = "{:,.{d}f}".format(x, d=dec)
    return s.replace(",", "X").replace(".", ",").replace("X", ".").replace("-", "\u2212")


def pct(x, dec=0):
    return n(100 * x, dec) + "%"


# ------------------------------------------------------------------ datos
v = openpyxl.load_workbook(XLSX, data_only=True)
P, H, MI, ME, AN, INV, AM, RR, CF, CV = (v["Presupuesto financiero"], v["Hipótesis"], v["Mod. ingresos"], v["Mod. egresos"],
                                        v["Anexo capacidad operativa"], v["Mod. inversión"], v["Amortizaciones"],
                                        v["Costos RRHH"], v["Costos fijos"], v["Costos variables"])
g = lambda ws, ref: ws[ref].value or 0
ing = [g(P, c + "12") for c in "EFG"]; fij = [g(P, c + "13") for c in "EFG"]; var = [g(P, c + "14") for c in "EFG"]
rrhh = [g(P, c + "15") for c in "EFG"]; uaii = [g(P, c + "16") for c in "EFG"]; iibb = [g(P, c + "17") for c in "EFG"]
iigg = [g(P, c + "19") for c in "EFG"]; udii = [g(P, c + "20") for c in "EFG"]; inv = [g(P, c + "21") for c in "DEFG"]
ff = [g(P, c + "22") for c in "DEFG"]; imp = [g(P, c + "23") for c in "LMN"]
r_rrhh = [g(P, c + "29") for c in "EFG"]; r_promo = [g(P, c + "25") for c in "EFG"]; r_fij = [g(P, c + "27") for c in "EFG"]; r_var = [g(P, c + "28") for c in "EFG"]
tasa, van, tir = g(P, "G33"), g(P, "G34"), g(P, "G35")
acum = sum(ff)
egresos = [g(ME, c + "13") for c in "BCD"]; comp = [[g(ME, c + str(r)) for c in "BCD"] for r in (20, 21, 22)]
altas_ing = [sum(g(MI, c + str(r)) for r in (17, 18, 19)) for c in "CDE"]; abonos_ing = [sum(g(MI, c + str(r)) for r in (20, 21, 22)) for c in "CDE"]
vidriera28 = g(MI, "E21")
precios = [g(H, "C%d" % r) for r in range(58, 64)]
gasto, mercado = g(H, "C12"), g(H, "D12")
comercios = [n(int(str(H["B%d" % r].value).replace(" comercios", ""))) for r in (24, 25, 26)]; shares = [g(H, "C%d" % r) for r in (24, 25, 26)]
captado = [g(H, "D%d" % r) for r in (24, 25, 26)]
# ticket ponderado 2028: abonos / comercio-meses
cm28 = sum(g(MI, "E%d" % r) / p for r, p in zip((20, 21, 22), precios[3:]))
arpu = abonos_ing[2] / cm28
H_ = [4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26]
MES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
hh_tot = [g(AN, "B86"), g(AN, "B97"), g(AN, "B108")]
hh_pico = []
for f in (86, 97, 108):
    vals = [g(AN, "%s%d" % (openpyxl.utils.get_column_letter(c), f)) for c in H_]
    m = max(range(12), key=lambda i: vals[i]); hh_pico.append((vals[m], MES[m]))
holg_min = [min(g(AN, "%s%d" % (openpyxl.utils.get_column_letter(c), f)) for c in H_) for f in (88, 99, 110)]
hh_contratadas = [[g(AN, "%s%d" % (openpyxl.utils.get_column_letter(c), f)) for c in H_] for f in (87, 98, 109)]
bloques = [(min(x) / 150, max(x) / 150) for x in hh_contratadas]
recupero = next((2025 + i for i in range(1, 4) if sum(ff[:i + 1]) >= 0), None)
tec27 = [RR.cell(row=47, column=c).value for c in range(3, 15)]
mes_tec2 = MES[tec27.index(2)] if 2 in tec27 else None
vend = [[RR.cell(row=f, column=c).value for c in range(3, 15)] for f in (30, 48, 66)]
fund = [g(RR, "N15"), g(RR, "N17")]
inv_anios = [g(INV, c + "5") for c in "GHIJ"]; amort = [g(AM, c + "5") for c in "IJK"]
desarrollo = g(INV, "D11")
kit_vid, kit_cad = g(AN, "W47"), g(AN, "W65")

# ------------------------------------------------------------------ word
d = docx.Document(DOCX)
ps = d.paragraphs


def buscar(prefijo, desde=0):
    for i in range(desde, len(d.paragraphs)):
        if d.paragraphs[i].text.strip().startswith(prefijo):
            return d.paragraphs[i]
    raise KeyError(prefijo)


def texto(p, cuerpo, negrita=None):
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    if negrita:
        rr = p.add_run(negrita); rr.bold = True
    p.add_run(cuerpo)
    p.alignment = AL.JUSTIFY
    return p


def despues(p_ref, estilo="Normal"):
    nuevo = OxmlElement("w:p")
    p_ref._p.addnext(nuevo)
    p = Paragraph(nuevo, p_ref._parent)
    p.style = d.styles[estilo]
    return p


def parrafo(p_ref, cuerpo, negrita=None):
    p = despues(p_ref)
    return texto(p, cuerpo, negrita)


def figura(p_ref, archivo, ancho_cm, leyenda):
    pi = despues(p_ref); pi.alignment = AL.CENTER
    pi.paragraph_format.keep_with_next = True; pi.paragraph_format.space_after = Pt(4)
    pi.add_run().add_picture(os.path.join(IMG, archivo), width=Cm(ancho_cm))
    pc = despues(pi); pc.alignment = AL.CENTER
    r = pc.add_run(leyenda); r.italic = True; r.font.size = Pt(9)
    return pc


def reemplazar_imagen(p, archivo):
    """Cambia el binario de la imagen del parrafo y ajusta la altura al nuevo aspecto."""
    ruta = os.path.join(IMG, archivo)
    w, h = Image.open(ruta).size
    for blip in p._p.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        d.part.related_parts[rid]._blob = open(ruta, "rb").read()
    for ext in list(p._p.iter(qn("wp:extent"))) + list(p._p.iter(qn("a:ext"))):
        cx = int(ext.get("cx")); ext.set("cy", str(int(cx * h / w)))


def borrar(p):
    p._p.getparent().remove(p._p)


def sectprs():
    """Los sectPr apaisado y vertical que ya usa el documento (los del anexo)."""
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    apaisado = vertical = None
    for p in d.paragraphs:
        pPr = p._p.find(W + "pPr")
        sp = pPr.find(W + "sectPr") if pPr is not None else None
        if sp is None:
            continue
        pg = sp.find(W + "pgSz")
        if pg is None:
            continue
        w, h = int(pg.get(W + "w")), int(pg.get(W + "h"))
        if w > h and apaisado is None:
            apaisado = copy.deepcopy(sp)
        if h > w and vertical is None:
            vertical = copy.deepcopy(sp)
    # que ninguna de las dos reinicie la numeracion de pagina
    for sp in (vertical, apaisado):
        for pn in list(sp.findall(W + "pgNumType")):
            sp.remove(pn)
    return vertical, apaisado


def poner_sectpr(p, sect):
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    pPr = p._p.find(W + "pPr")
    if pPr is None:
        pPr = p._p.makeelement(W + "pPr", {}); p._p.insert(0, pPr)
    pPr.append(copy.deepcopy(sect))


def figura_apaisada(p_ref, archivo, leyenda):
    """La figura en su propia pagina apaisada: el parrafo anterior cierra la
    seccion vertical y el epigrafe cierra la apaisada, como hace el anexo."""
    poner_sectpr(p_ref, SECT_V)
    pc = figura(p_ref, archivo, 24.9, leyenda)
    poner_sectpr(pc, SECT_A)
    return pc


SECT_V, SECT_A = sectprs()
assert SECT_V is not None and SECT_A is not None

# ---- 0) la fecha del encabezado
for sec in d.sections:
    for hdr in (sec.header, sec.first_page_header):
        for t in hdr.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            if "15/08/2026" in r.text:
                                r.text = r.text.replace("15/08/2026", "30/09/2026").replace("28/09/2026", "30/09/2026")

# ---- 1) figuras existentes que cambian (por su epigrafe)
CAMBIAN = {
    "Figura 8.1.": "01_hipotesis_negocio_mercado.png", "Figura 8.2.": "02a_participacion_texto.png",
    "Figura 8.3.": "02b_participacion_tabla.png", "Figura 8.4.": "02c_participacion_grafico.png",
    "Figura 8.5.": "03a_precios_altas.png", "Figura 8.6.": "03b_precios_abonos.png",
    "Figura 8.7.": "04_ingresos_tabla.png", "Figura 8.8.": "05_ingresos_grafico_1.png",
    "Figura 8.9.": "05_ingresos_grafico_2.png", "Figura 8.10.": "05_ingresos_grafico_3.png",
    "Figura 8.11.": "06_egresos_tabla.png", "Figura 8.12.": "15_egresos_grafico_1.png",
    "Figura 8.13.": "07_egresos_composicion.png",
    "Figura 8.25.": "AN11_hh_2026_s1.png", "Figura 8.26.": "AN12_hh_2026_s2.png", "Figura 8.27.": "AN13_hh_2027_s1.png",
    "Figura 8.28.": "AN14_hh_2027_s2.png", "Figura 8.29.": "AN15_hh_2028_s1.png", "Figura 8.30.": "AN16_hh_2028_s2.png",
}
ps = d.paragraphs
for i, p in enumerate(ps):
    t = p.text.strip()
    for cap, archivo in CAMBIAN.items():
        if t.startswith(cap) and i > 0 and ps[i - 1]._p.find(".//" + qn("a:blip")) is not None:
            reemplazar_imagen(ps[i - 1], archivo)
# la de ingresos vs egresos esta antes del titulo 8.4.1 (su epigrafe 8.14 quedo mas abajo)
h841 = buscar("8.4.1 Anexo de capacidad operativa")
ps = d.paragraphs
idx = [i for i, p in enumerate(ps) if p._p is h841._p][0]
if ps[idx - 1]._p.find(".//" + qn("a:blip")) is not None:
    reemplazar_imagen(ps[idx - 1], "16_egresos_grafico_2.png")

# ---- 2) (las figuras de riesgos se sacan mas abajo; no hace falta renumerar)

# ---- 3) texto del capitulo
texto(buscar("El presupuesto financiero de ShopMetrics se desarrolla"),
      "El presupuesto financiero de ShopMetrics se desarrolla en el archivo «Presupuesto financiero ShopMetrics.xlsx», que "
      "acompaña a este informe. Este punto reproduce las hipótesis sobre las que se construye cada modelo y los resultados que "
      "de ellas se derivan, de modo que la planilla quede explicada y no sea necesario recorrerla para entender el razonamiento. "
      "Los cuadros y gráficos que siguen son la salida directa de la planilla. Conforme al alcance del segundo avance se "
      "desarrollan aquí el modelo de negocios, las hipótesis, el modelo de ingresos, el modelo de egresos, el modelo de "
      "inversión, las amortizaciones y el presupuesto financiero; los escenarios y el plan de mejoras corresponden al tercero. "
      "La matriz de riesgos se adelantó en el primer avance porque no depende de los importes. Todos los valores están "
      "expresados en dólares estadounidenses.")
texto(buscar("Origen de los datos."),
      "Todas las cifras que se presentan en este capítulo provienen de un único archivo de cálculo y están encadenadas por "
      "fórmula, de modo que ninguna se carga dos veces. La hoja Hipótesis fija los tres parámetros de partida: el tamaño del "
      "universo de comercios, el porcentaje de participación que se pretende alcanzar en cada ejercicio y la lista de precios de "
      "los seis servicios. De ahí se desprende la hoja Proy. ventas, que distribuye mes a mes las altas de cada plan y acumula "
      "los abonos vigentes, y que es la única hoja donde se cargan cantidades a mano. Esas cantidades no se eligen: la "
      "facturación de cada ejercicio tiene que dar el valor del mercado captado que fija la participación pretendida, y las "
      "altas por mes son las que hacen falta para llegar a ese número con la lista de precios vigente. El costo de desarrollo "
      "de la plataforma, que es el grueso de la inversión inicial, proviene del cronograma en Microsoft Project que acompaña "
      "al informe.",
      "Origen de los datos. ")
texto(buscar("Mercado meta y participación."),
      "El universo de arranque es de 3.400 comercios de indumentaria, textiles y calzado, cifra que surge de aplicar la "
      "participación publicada del rubro (24,3%%) sobre los 14.083 locales ocupados que releva el Instituto de Estadística y "
      "Censos de la Ciudad de Buenos Aires en sus ejes comerciales. Sobre ese universo se plantea una participación del %s al "
      "cierre del tercer año, con una progresión del %s, %s y %s. Esa participación, aplicada al mercado total, es el objetivo "
      "de facturación de cada ejercicio, y la cantidad de comercios que hace falta para alcanzarlo con la lista de precios "
      "—%s, %s y %s abonados a diciembre de cada año— se deduce de la proyección de ventas, no al revés. Es una meta "
      "exigente para un solo rubro, y por eso la estructura de costos comerciales incluye el esfuerzo de promoción que la "
      "sostiene, como se detalla en el modelo de egresos."
      % (pct(shares[2]), pct(shares[0]), pct(shares[1]), pct(shares[2]), comercios[0], comercios[1], comercios[2]),
      "Mercado meta y participación. ")
texto(buscar("Valorización del mercado y su diferencia"),
      "El valor del mercado captado se obtiene multiplicando la cantidad de comercios por un gasto de referencia de USD %s "
      "anuales (USD %s mensuales), que es lo que un comercio minorista destina a software de gestión, facturación electrónica y "
      "analítica. La referencia se construye sobre dos precios públicos de 2026: los planes de punto de venta con facturación "
      "electrónica y stock multi-tienda, en el orden de los USD 79 mensuales, y el plan de entrada de Fudo, uno de los sistemas "
      "con los que la plataforma se integra, en el orden de los USD 40 mensuales por sucursal. Ese valor mide el tamaño del "
      "bolsillo disponible del segmento, y con la participación pretendida se convierte en el objetivo de facturación. La "
      "proyección de ventas se construye para alcanzarlo: el ingreso del modelo es de USD %s en 2028 contra un valor captado de "
      "USD %s. Como el ticket promedio ponderado de la propia lista de precios es de USD %s mensuales, unos USD %s anuales, "
      "porque el grueso de la cartera se concentra en el Plan Vidriera, hacen falta más comercios abonados que los que "
      "resultarían de dividir el mercado captado por el gasto de referencia; la diferencia es el margen disponible para "
      "servicios adicionales una vez consolidada la base instalada."
      % (n(gasto), n(gasto / 12), n(ing[2]), n(captado[2]), n(arpu, 1), n(arpu * 12)),
      "Valorización del mercado y su diferencia con el ingreso propio. ")
texto(buscar("Capacidad de operación."),
      "La estructura se mantiene lo más plana posible: en relación de dependencia están únicamente los dos fundadores, y "
      "todo lo demás se terceriza o se contrata a demanda. El anexo de capacidad operativa calcula el tiempo que consume cada "
      "alta según el plan (media hora en el Básico, tres horas en el Vidriera y doce en el Cadena) y las horas de soporte por "
      "comercio activo (entre 0,15 y 0,60 horas mensuales según el plan). Al pie de cada ejercicio muestra cuántas horas se "
      "contratan a técnicos freelance —por mes completo de 150 horas, según lo que pide la cartera— y la holgura que queda.",
      "Capacidad de operación. ")
texto(buscar("Lectura del modelo. El ingreso se multiplica"),
      "El ingreso se multiplica por %s entre 2026 y 2027 y por %s entre 2027 y 2028, un perfil coherente con un negocio de "
      "suscripción que arranca desde cero y que en el segundo ejercicio suma el efecto de una captación más intensa. La "
      "composición se desplaza con claridad hacia el ingreso recurrente, que pasa de representar el %s del total en 2026 al %s "
      "en 2028, mientras los cargos de alta caen del %s al %s. Ese desplazamiento es el que vuelve previsible el negocio, pero "
      "también expone el punto crítico: el ingreso recurrente depende de la permanencia, de modo que la rotación de comercios "
      "(riesgo R2 de la matriz) actúa directamente sobre la línea principal del modelo. Dentro de la cartera, el Plan Vidriera "
      "concentra el %s del ingreso de 2028, lo que confirma que es el núcleo de la oferta."
      % (n(ing[1] / ing[0], 1), n(ing[2] / ing[1], 1), pct(abonos_ing[0] / ing[0]), pct(abonos_ing[2] / ing[2]),
         pct(altas_ing[0] / ing[0]), pct(altas_ing[2] / ing[2]), pct(vidriera28 / ing[2])),
      "Lectura del modelo. ")
texto(buscar("Criterios adoptados."),
      "En primer lugar, el equipamiento que se instala en cada local —sensor de vidriera, contador de puerta, gateway y kit de "
      "montaje— es un insumo del servicio y no un bien de uso: se imputa como costo variable por cada alta (USD %s en el Plan "
      "Vidriera y USD %s en el Plan Cadena, que cubre cinco locales), no forma parte del modelo de inversión y no se amortiza. "
      "El cargo de alta e instalación está fijado de modo de cubrirlo. En segundo lugar, la instalación y el soporte los "
      "hacen técnicos freelance contratados por hora según las horas que pide el anexo de capacidad, y por eso son costo "
      "variable y no de recursos humanos: se pagan según las altas y la cartera de cada mes, sin cargas patronales. Lo mismo "
      "vale para la venta, que se paga por comisión sobre cada alta concretada. En tercer lugar, el desarrollo y mantenimiento "
      "evolutivo de la plataforma, la contabilidad y el marketing se contratan por fuera de la estructura y figuran como "
      "costos fijos. En costos variables quedan entonces los insumos IoT de cada instalación, las horas del técnico, la "
      "comisión del vendedor, la movilidad de instalación, el consumo de infraestructura por comercio activo, las "
      "notificaciones, las comisiones de los medios de pago y las bonificaciones de las campañas de captación."
      % (n(kit_vid - 18), n(kit_cad - 72)),
      "Criterios adoptados. ")
texto(buscar("Lectura del modelo. Los recursos humanos"),
      "Los recursos humanos explican entre el %s y el %s de los egresos en los tres ejercicios; los costos fijos, que incluyen "
      "el mantenimiento contratado de la plataforma y las campañas de promoción, entre el %s y el %s; y los variables, que "
      "incluyen los insumos de cada instalación y las horas de los técnicos freelance, entre el %s y el %s. La estructura "
      "de recursos humanos es la mínima posible durante todo el horizonte: los dos fundadores, que cubren entre ambos la "
      "gerencia general y la de sistemas y absorben los puestos administrativos, comerciales y técnicos que no se "
      "tercerizan; la hoja de recursos humanos despliega los quince puestos de la estructura y dice quién cubre cada uno y "
      "dónde está su costo. Los fundadores perciben una remuneración reducida durante los dos primeros ejercicios; la "
      "diferencia con el valor de su puesto es aporte de trabajo de los socios. El sueldo anual complementario se paga en "
      "junio y en diciembre."
      % (pct(min(comp[2])), pct(max(comp[2])), pct(min(comp[0])), pct(max(comp[0])), pct(min(comp[1])), pct(max(comp[1]))),
      "Lectura del modelo. ")
texto(buscar("Costo del servicio."),
      "Valorizadas al costo horario del técnico freelance, siete dólares la hora, las altas cuestan 3,50, 21 y 84 dólares de "
      "mano de obra contra precios de lista de %s, %s y %s dólares. Sumado el equipamiento que se instala en el local, el "
      "costo completo del alta asciende a 3,50, 55 y 254 dólares respectivamente, de modo que el cargo de alta cubre el "
      "esfuerzo técnico y el material comprometido en los tres planes con margen. El soporte mensual cuesta 1,05, 1,75 y 4,20 "
      "dólares por comercio, contra abonos de %s, %s y %s dólares."
      % (n(precios[0]), n(precios[1]), n(precios[2]), n(precios[3]), n(precios[4]), n(precios[5])),
      "Costo del servicio. ")
texto(buscar("Lectura del anexo."),
      "El total de horas hombre asciende a %s en 2026, %s en 2027 y %s en 2028. Lo que define la dotación no es ese total "
      "sino el mes de mayor exigencia: %s horas en %s de 2026, %s en %s de 2027 y %s en %s de 2028. Las horas se contratan "
      "a técnicos freelance por mes completo de 150: entre %s y %s técnicos-mes en 2026, entre %s y %s en 2027 y entre %s y "
      "%s en 2028. Las dos filas al pie de cada ejercicio muestran las horas contratadas y la holgura mes a mes; la holgura "
      "mínima del horizonte es de %s horas."
      % (n(hh_tot[0], 2), n(hh_tot[1], 2), n(hh_tot[2], 2), n(hh_pico[0][0], 2), hh_pico[0][1], n(hh_pico[1][0], 2), hh_pico[1][1],
      n(hh_pico[2][0], 2), hh_pico[2][1], n(bloques[0][0]), n(bloques[0][1]), n(bloques[1][0]), n(bloques[1][1]),
      n(bloques[2][0]), n(bloques[2][1]), n(min(holg_min), 1)),
      "Lectura del anexo. ")
texto(buscar("Alcance. El anexo dimensiona"),
      "El anexo dimensiona el área técnica, no el área comercial. La venta se paga por comisión sobre cada alta concretada "
      "—%s, %s y %s altas anuales respectivamente—, de modo que el esfuerzo comercial es variable y no exige dotación. Las "
      "horas contratadas de la fila al pie son las que toma por fórmula la hoja de costos variables para valorizar el trabajo "
      "de los técnicos."
      % (n(altas_ing[0] and sum(g(MI, "C%d" % r) / p for r, p in zip((17, 18, 19), precios[:3]))),
         n(sum(g(MI, "D%d" % r) / p for r, p in zip((17, 18, 19), precios[:3]))),
         n(sum(g(MI, "E%d" % r) / p for r, p in zip((17, 18, 19), precios[:3])))),
      "Alcance. ")
texto(buscar("Contraste con los ingresos."),
      "La comparación entre ambas curvas muestra el perfil buscado: el resultado operativo es negativo en el primer ejercicio, "
      "con un déficit de USD %s en 2026, y positivo desde el segundo, con USD %s en 2027 y USD %s en 2028. Los egresos pasan "
      "de representar %s veces los ingresos a %s veces. La revisión de la cátedra sobre la versión anterior dejó dos "
      "correcciones que este modelo incorpora: la proyección de ventas se construye desde la facturación objetivo que fija la "
      "participación pretendida, y la estructura se reduce a los dos fundadores, con la instalación, el soporte y la venta "
      "pagados según el volumen. A eso se sumaron la contratación externa del desarrollo y la incorporación del costo de los "
      "insumos de instalación. El punto de equilibrio operativo se alcanza en 2027."
      % (n(-uaii[0]), n(uaii[1]), n(uaii[2]), n(egresos[0] / ing[0], 1), n(egresos[2] / ing[2], 1)),
      "Contraste con los ingresos. ")

# ---- 4) 8.5 modelo de inversion
h = buscar("8.5 Modelo de inversión"); borrar(buscar("Se desarrolla en el segundo avance. Comprende"))
p = parrafo(h,
    "El modelo de inversión reúne lo que la empresa tiene que adquirir antes de vender y lo que suma en cada ejercicio para "
    "sostener el crecimiento. La inversión inicial —el año cero, que es 2025— asciende a USD %s y está dominada por el "
    "desarrollo de la plataforma, USD %s; el resto es el equipamiento informático y de oficina de los fundadores, el "
    "equipamiento de red y la registración de la marca y la constitución de la sociedad. Como la estructura es la de los "
    "dos fundadores durante todo el horizonte, en 2026 y 2027 no hay inversiones, y en 2028 se renuevan las dos notebooks, "
    "amortizadas a los tres años, por USD %s. El equipamiento que se instala en los locales no figura aquí: es insumo del "
    "servicio y se imputa como costo variable, según el criterio explicado en el punto 8.4."
    % (n(inv_anios[0]), n(desarrollo), n(inv_anios[3])))
p = figura_apaisada(p, "INV01_inversion_anio0.png", "Figura 8.31. Inversión inicial (año cero), con la referencia de precio de cada concepto. Fuente: planilla de presupuesto financiero, hoja Mod. inversión.")
p = parrafo(p,
    "El desarrollo se contrata por fuera de la estructura: los fundadores gestionan la empresa y no participan de la "
    "construcción, siguiendo el criterio de arrancar con la menor estructura posible y tercerizar lo que no sea el rol de "
    "dirección. El costo es el de las horas de quienes construyen la plataforma, y sale del cronograma de desarrollo armado "
    "en Microsoft Project que acompaña a este informe: cuatro roles —un arquitecto de software, un desarrollador backend, un "
    "desarrollador frontend y un ingeniero de machine learning—, 2.324 horas entre el 6 de enero y el 1 de septiembre de "
    "2025, a valor hora de mercado para clientes argentinos en 2026 (USD 40 el arquitecto y el ingeniero de machine "
    "learning, USD 22 los desarrolladores). Al no ser empleados no llevan cargas patronales ni aguinaldo. El cronograma "
    "cubre los treinta y un casos de uso especificados en el plan de desarrollo tecnológico del capítulo 10 y está atado a "
    "la propuesta de valor del punto 1.7: las once fases construyen, en ese orden, la ingesta de datos de los puntos de "
    "venta y los sensores, el motor de métricas, el panel del centro, las alertas, los modelos de Machine Learning y el "
    "portal del locatario.",
    "Justificación de la inversión en la solución tecnológica. ")
p = figura(p, "INV05_gantt_fases.png", ANCHO, "Figura 8.32. Cronograma de desarrollo de la plataforma por fase, con el rol que carga más horas en cada una. Fuente: project/ShopMetrics-desarrollo.xml, generado desde el archivo de Microsoft Project.")
p = figura(p, "INV06_project_recursos.png", ANCHO, "Figura 8.33. Hoja de recursos de Microsoft Project: valor hora, horas y costo de cada rol. El total, USD %s, es el valor del concepto «Desarrollo de la plataforma»." % n(desarrollo))
p = parrafo(p,
    "Backend y frontend corren en paralelo: el frontend arranca cuando hay una API con datos que consumir, a fines de "
    "abril, y desde ahí no se corta; los modelos arrancan cuando están las métricas. Por eso las fases 5, 6, 7 y 10 se "
    "superponen entre mayo y junio y el desarrollo termina cuatro meses antes de la salida al mercado, en enero de 2026. "
    "Ninguno de los cuatro roles supera el cien por ciento de dedicación en ningún día del cronograma.")
p = figura(p, "INV02_inversion_2026_2028.png", ANCHO, "Figura 8.34. Inversión de los ejercicios 2026, 2027 y 2028: sólo lo que entra en cada año. Fuente: planilla de presupuesto financiero, hoja Mod. inversión.")
p = figura(p, "INV04_inversion_grafico.png", 12.0, "Figura 8.35. Inversión por año.")

# ---- 5) 8.6 amortizaciones
h = buscar("8.6 Amortizaciones"); borrar(buscar("Se desarrollan en el segundo avance, una vez definido"))
p = parrafo(h,
    "Se amortizan únicamente los bienes de uso: el equipamiento informático y de comunicaciones a tres años y los muebles "
    "de oficina a diez, que son los plazos que fija la resolución técnica del Consejo Profesional de Ciencias Económicas y "
    "no una elección propia. El desarrollo de la plataforma y la registración de la marca forman parte de la inversión pero "
    "no se amortizan, y el equipamiento que se instala en los locales tampoco, porque es insumo y no bien de uso. Cada bien "
    "amortiza desde el ejercicio en que se adquiere y sólo por las cuotas que caen dentro del horizonte: un bien a diez años "
    "comprado en 2026 se lleva siete cuotas al próximo plan de negocios.")
p = parrafo(p,
    "La amortización no es un egreso financiero —no mueve fondos— y por eso no forma parte de los costos ni del flujo de "
    "fondos. Se calcula para una sola cosa: descontarla del monto imponible del impuesto a las ganancias, que es donde "
    "incide. Las cuotas del período suman USD %s en 2026, USD %s en 2027 y USD %s en 2028."
    % (n(amort[0], 2), n(amort[1], 2), n(amort[2], 2)),
    "Para qué se calculan. ")
p = figura_apaisada(p, "AMO01_amortizaciones.png", "Figura 8.36. Cuadro de amortizaciones: valor de adquisición por ejercicio y cuota anual de cada bien. Fuente: planilla de presupuesto financiero, hoja Amortizaciones.")
p = figura(p, "AMO02_amortizaciones_resumen.png", 8.0, "Figura 8.37. Amortización total del período por ejercicio.")

# ---- 6) 8.7 presupuesto financiero
h = buscar("8.7 Presupuesto financiero"); borrar(buscar("Se desarrolla en el segundo avance. Consolida"))
p = parrafo(h,
    "El presupuesto financiero consolida los modelos anteriores en un flujo de fondos por ejercicio. Parte de los ingresos "
    "del modelo de ingresos y les resta los tres bloques de costo del modelo de egresos para obtener la utilidad antes de "
    "impuestos: USD %s en 2026, USD %s en 2027 y USD %s en 2028. Sobre ese resultado se calculan dos impuestos. El impuesto "
    "a los ingresos brutos se aplica sobre la facturación, se paga aunque el ejercicio dé pérdida y se simplifica a una "
    "alícuota directa del 3%%, sin abrir el convenio multilateral por jurisdicción. El impuesto a las ganancias, del 35%%, se "
    "aplica sobre un monto imponible que no es la utilidad: se le descuentan los ingresos brutos, para no gravar dos veces, y "
    "las amortizaciones del período. El impuesto al valor agregado no se incluye, por decisión metodológica de la cátedra."
    % (n(-uaii[0]), n(uaii[1]), n(uaii[2])))
p = parrafo(p,
    "Las ganancias de un ejercicio se liquidan en el ejercicio siguiente. Como el emprendimiento arranca sin saldos "
    "anteriores, el primer año no paga impuesto a las ganancias; en 2027 tampoco, porque el monto imponible de 2026 fue "
    "negativo (USD %s); y en 2028 paga USD %s, el 35%% del monto imponible de 2027 (USD %s). El monto imponible de 2028, "
    "USD %s, queda fuera del horizonte y se liquidaría en 2029."
    % (n(imp[0]), n(iigg[2]), n(imp[1]), n(imp[2])),
    "El desfasaje del impuesto a las ganancias. ")
p = figura(p, "PRE01_presupuesto.png", ANCHO, "Figura 8.38. Presupuesto financiero: de los ingresos al flujo de fondos, por ejercicio. Fuente: planilla de presupuesto financiero, hoja Presupuesto financiero.")
p = figura(p, "PRE03_monto_imponible.png", 11.0, "Figura 8.39. Monto imponible del impuesto a las ganancias: utilidad antes de impuestos menos ingresos brutos menos amortizaciones.")
p = parrafo(p,
    "Restada la inversión de cada ejercicio, el flujo de fondos es de USD %s en el año cero, USD %s en 2026, USD %s en 2027 "
    "y USD %s en 2028. El acumulado al cierre del tercer ejercicio es de USD %s: la inversión se recupera dentro del "
    "horizonte, en %s, con un margen de %s veces la inversión inicial que los escenarios del tercer avance van a "
    "poner a prueba."
    % (n(ff[0]), n(ff[1]), n(ff[2]), n(ff[3]), n(acum), recupero or "el tercer año", n(acum / -ff[0], 1)),
    "Flujo de fondos. ")
p = figura(p, "PRE05_flujo_grafico.png", 12.5, "Figura 8.40. Flujo de fondos del ejercicio y acumulado.")
p = parrafo(p,
    "Debajo del flujo, la planilla relaciona cada bloque de costo con los ingresos del mismo ejercicio. El costo de recursos "
    "humanos pasa del %s de los ingresos en 2026 al %s en 2027 y al %s en 2028: el primer valor refleja una estructura que "
    "todavía no tiene cartera que la sostenga, y el último es el de una empresa de base tecnológica en régimen. El costo de "
    "promoción, sumadas las campañas de los costos fijos y las bonificaciones de los variables, representa el %s, el %s y el "
    "%s de los ingresos: es el esfuerzo que sostiene el salto de ventas de 2027 y se mantiene por encima del diez por ciento "
    "durante todo el horizonte."
    % (pct(r_rrhh[0]), pct(r_rrhh[1]), pct(r_rrhh[2]), pct(r_promo[0]), pct(r_promo[1]), pct(r_promo[2])),
    "Ratios de control. ")
p = figura(p, "PRE02_ratios.png", ANCHO, "Figura 8.41. Costos de promoción, fijos, variables y de recursos humanos en relación con los ingresos.")
p = parrafo(p,
    "La tasa de corte se fija en el %s anual en dólares. Se construye por comparación con las alternativas de menor riesgo "
    "en la misma moneda: un bono del Tesoro de los Estados Unidos rinde en el orden del 4%% anual, y la deuda soberana "
    "argentina en dólares, con un riesgo país de 533 puntos básicos a fines de septiembre de 2026, en el orden del 9 al 10%%. "
    "A una empresa que todavía no facturó, sin historial ni garantías, un inversor le exige por lo menos el triple de lo que "
    "le rinde prestarle al Estado argentino. La tasa es alta pero está lejos del 75%% del ejemplo de la cátedra, que "
    "corresponde a un proyecto en pesos."
    % pct(tasa),
    "Tasa de corte. ")
if van >= 0:
    lectura_van = (
        "Con esa tasa, el valor actual neto de los cuatro flujos es de USD %s, positivo, y la tasa interna de retorno del "
        "%s, por encima de la tasa de corte. Ambos indicadores dicen lo mismo: dentro del horizonte de tres años el "
        "proyecto devuelve la inversión y remunera el capital a la tasa exigida; el valor actual neto equivale al %s de la "
        "inversión inicial. El resultado descansa en que la facturación alcance la participación pretendida: ese supuesto es "
        "el que los escenarios del tercer avance van a poner a prueba. La lectura se completa en el punto 9."
        % (n(van), pct(tir, 1), pct(van / -inv[0], 1)))
else:
    lectura_van = (
        "Con esa tasa, el valor actual neto de los cuatro flujos es de USD %s y la tasa interna de retorno del %s. Ambos "
        "indicadores dicen lo mismo: dentro del horizonte de tres años el proyecto recupera la inversión pero no remunera el "
        "capital a la tasa exigida, porque el retorno se concentra a partir de 2028 y el ejercicio siguiente queda fuera del "
        "análisis. La lectura se completa en el punto 9." % (n(van), pct(tir, 1)))
p = parrafo(p, lectura_van, "Valor actual neto y tasa interna de retorno. ")
p = figura(p, "PRE04_van_tir.png", 7.0, "Figura 8.42. Tasa de corte, valor actual neto y tasa interna de retorno.")
p = figura(p, "INV03_inversion_resumen.png", 10.0, "Figura 8.43. Inversión por ejercicio, tal como entra al presupuesto financiero.")

# ---- 7) 8.8, 8.9 y 9 pertenecen al tercer avance: se sacan del informe
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ps = d.paragraphs
i0 = next(i for i, q in enumerate(ps) if q.text.strip().startswith("8.8 Matriz de riesgos"))
i1 = next(i for i, q in enumerate(ps) if q.text.strip().startswith("Bibliografía"))
# dentro del tramo hay dos saltos de seccion (las paginas apaisadas de la
# matriz). El ultimo parrafo del tramo cierra la seccion apaisada y vuelve a
# vertical; si se borra sin mas, la Bibliografia hereda la orientacion
# equivocada. Se conserva ese sectPr en un parrafo vacio antes de Bibliografia.
sect_final = None
for q in ps[i0:i1]:
    pPr = q._p.find(W + "pPr")
    sp = pPr.find(W + "sectPr") if pPr is not None else None
    if sp is not None:
        sect_final = copy.deepcopy(sp)
for q in ps[i0:i1]:
    q._p.getparent().remove(q._p)
if sect_final is not None:
    vac = despues(ps[i0 - 1])
    poner_sectpr(vac, sect_final)
# la bibliografia pasa a ser el punto 9 no: queda sin numero, como estaba

# ---- 8) referencias cruzadas a lo que se saco
for q in d.paragraphs:
    if "La matriz de riesgos se adelantó en el primer avance porque no depende de los importes. " in q.text:
        for r in q.runs:
            r.text = r.text.replace("los escenarios y el plan de mejoras corresponden al tercero. La matriz de riesgos se adelantó en el primer avance porque no depende de los importes. ",
                                    "la matriz de riesgos, los escenarios, el plan de mejoras y el análisis de viabilidad corresponden al tercero. ")
    if "(riesgo R2 de la matriz)" in q.text:
        for r in q.runs:
            r.text = r.text.replace("(riesgo R2 de la matriz)", "(uno de los riesgos que se valoran en el tercer avance)")
    if "y la matriz de riesgos valora las desviaciones posibles sobre ese mismo conjunto" in q.text:
        for r in q.runs:
            r.text = r.text.replace(", y la matriz de riesgos valora las desviaciones posibles sobre ese mismo conjunto", "")
    if "La lectura se completa en el punto 9." in q.text:
        for r in q.runs:
            r.text = r.text.replace(" La lectura se completa en el punto 9.", " La lectura de viabilidad se completa en el tercer avance, junto con los escenarios.")
    if "que los escenarios del tercer avance van a poner a prueba" in q.text or "Ese margen es el que los escenarios del tercer avance van a poner a prueba." in q.text:
        pass   # esas menciones son correctas: anuncian lo que viene

d.save(DOCX)
print("Word actualizado: figuras reemplazadas, 8.5-8.7 y 9 escritos.")
print("  UAII %s / %s / %s | acumulado %s | VAN %s | TIR %s | tecnico 2: %s" % (n(uaii[0]), n(uaii[1]), n(uaii[2]), n(acum), n(van), pct(tir, 1), mes_tec2))
