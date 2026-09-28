# -*- coding: utf-8 -*-
"""Genera el cronograma de desarrollo en XML que abre Microsoft Project.

La inversion inicial del negocio es el costo de las horas de la gente que
construye la plataforma. Project lo calcula solo: se le dan los recursos con su
valor hora, las tareas con su duracion y quien trabaja en cada una.

El desarrollo se contrata afuera: los fundadores gestionan la empresa y no
participan de la construccion, asi que no figuran aca. Son cuatro roles, los
imprescindibles para este software: un arquitecto, que define y valida; un
backend, que hace la API, la base y las integraciones; un frontend, que hace
el panel y las vistas; y un ingeniero de machine learning para los modelos.
Backend y frontend son cadenas independientes y corren en paralelo. El valor
hora es el de mercado para clientes argentinos en 2026, sin cargas patronales
ni aguinaldo porque no son empleados.
"""
from __future__ import annotations

import datetime as dt
import os
import xml.etree.ElementTree as ET

AQUI = os.path.dirname(os.path.abspath(__file__))
SALIDA = os.path.join(AQUI, "ShopMetrics-desarrollo.xml")

# El desarrollo es previo a salir a vender. Los ingresos del presupuesto
# arrancan en 2026, asi que el ano cero --el del desarrollo-- es 2025.
ARRANQUE = dt.date(2025, 1, 6)          # primer lunes habil del ano
JORNADA = 8                             # horas por dia
HORAS_MES = 150                         # del anexo de capacidad operativa

# (nombre, valor hora USD de mercado, clave). Fuentes en LEEME.md.
EQUIPO = [
    ("Arquitecto de software (free lance)",       40.00, "ARQ"),
    ("Desarrollador backend (free lance)",        22.00, "BACK"),
    ("Desarrollador frontend (free lance)",       22.00, "FRONT"),
    ("Ingeniero de machine learning (free lance)", 40.00, "ML"),
]

# Cada tarea: (clave, nombre, dias, {recurso: dedicacion}, despues_de)
# despues_de = None -> la tarea anterior de la fase; la primera de una fase,
# la ultima de la fase anterior. Se indica explicito solo donde el orden no es
# el de la lista, que es lo que hace posible el solape.
PLAN = [
 ("1. Definición y arquitectura", [
   ("1.1", "Relevamiento funcional y alcance del producto", 8, {"ARQ":1.0}, None),
   ("1.2", "Arquitectura de la solución y elección del stack", 7, {"ARQ":1.0}, None),
   ("1.3", "Modelo de datos y diccionario", 5, {"ARQ":1.0,"BACK":0.5}, None),
 ]),
 ("2. Base de datos y cimientos", [
   ("2.1", "Esquema, migraciones y series temporales", 8, {"BACK":1.0,"ARQ":0.5}, None),
   ("2.2", "Entorno de desarrollo y automatización de pruebas", 7, {"BACK":1.0}, None),
 ]),
 ("3. Integración con sistemas externos", [
   ("3.1", "Conectores con puntos de venta", 12, {"BACK":1.0,"ARQ":0.5}, None),
   ("3.2", "Ingesta de sensores de conteo", 10, {"BACK":1.0}, None),
   ("3.3", "Monitoreo de salud de las integraciones", 8, {"BACK":1.0}, None),
 ]),
 ("4. Motor de métricas e indicadores", [
   ("4.1", "Cálculo de tráfico, conversión, ventas y vacancia", 12, {"BACK":1.0,"ARQ":0.5}, None),
   ("4.2", "Agregados por hora, zona y locatario", 8, {"BACK":1.0}, None),
 ]),
 # El frontend arranca cuando hay API con datos que consumir, y desde ahi
 # corre sin cortes en su propia cadena.
 ("5. Panel web del centro", [
   ("5.1", "Sistema de diseño y navegación", 6, {"FRONT":1.0}, ["4.1"]),
   ("5.2", "Tablero de indicadores y series", 12, {"FRONT":1.0}, None),
   ("5.3", "Mapa del centro y ficha de locatarios", 10, {"FRONT":1.0}, None),
 ]),
 ("6. Motor de reglas y alertas", [
   ("6.1", "Definición y evaluación de reglas por umbral", 10, {"BACK":1.0,"ARQ":0.5}, ["4.2"]),
   ("6.2", "Canales de notificación", 5, {"BACK":1.0}, None),
   ("6.3", "Centro de alertas y ciclo de atención", 10, {"FRONT":1.0}, ["5.3", "6.1"]),
 ]),
 # Los modelos necesitan las metricas calculadas; el resto es independiente.
 ("7. Modelos de Machine Learning", [
   ("7.1", "Predicción de riesgo de vacancia", 10, {"ML":1.0}, ["4.2"]),
   ("7.2", "Recomendación de mix de locatarios", 8, {"ML":1.0}, None),
   ("7.3", "Detección de anomalías en la operación", 12, {"ML":1.0}, None),
 ]),
 ("8. Portal del locatario", [
   ("8.1", "Tablero del local y comparación con la categoría", 12, {"FRONT":1.0}, ["6.3"]),
   ("8.2", "Consentimiento sobre el uso de datos", 5, {"FRONT":1.0}, None),
 ]),
 ("9. Vista móvil de operaciones", [
   ("9.1", "Resolución de alertas en terreno", 10, {"FRONT":1.0}, None),
   ("9.2", "Historial de turno del operario", 6, {"FRONT":1.0}, None),
 ]),
 ("10. Seguridad, permisos y auditoría", [
   ("10.1", "Autenticación y matriz de permisos", 8, {"BACK":1.0,"ARQ":0.5}, ["6.2"]),
   ("10.2", "Auditoría de accesos a datos sensibles", 7, {"BACK":1.0}, None),
 ]),
 ("11. Pruebas y salida a producción", [
   ("11.1", "Integración de los modelos de ML en el panel y las alertas", 6, {"BACK":1.0}, ["10.2", "7.3"]),
   ("11.2", "Pruebas integrales y corrección", 12, {"ARQ":1.0,"BACK":0.5,"FRONT":0.5}, ["11.1", "9.2"]),
   ("11.3", "Documentación y manual de uso", 5, {"ARQ":0.5,"BACK":0.5}, None),
   ("11.4", "Despliegue y puesta en marcha", 6, {"ARQ":1.0,"BACK":1.0}, None),
 ]),
]


def habil(f: dt.date) -> dt.date:
    while f.weekday() >= 5:
        f += dt.timedelta(days=1)
    return f


def sumar(f: dt.date, dias: int) -> dt.date:
    """Ultimo dia habil de una tarea que dura `dias` y arranca en `f`."""
    f = habil(f)
    for _ in range(dias - 1):
        f = habil(f + dt.timedelta(days=1))
    return f


def sig(f: dt.date) -> dt.date:
    return habil(f + dt.timedelta(days=1))


def dur(dias: float) -> str:
    return "PT%dH0M0S" % round(dias * JORNADA)


def sub(padre, etiqueta, texto=None):
    e = ET.SubElement(padre, etiqueta)
    if texto is not None:
        e.text = str(texto)
    return e


def programar():
    """Fechas de cada tarea a partir de sus predecesoras. Devuelve la lista
    de tareas hoja en orden, con uid, fase y fechas, y los limites por fase."""
    hojas, fin_de, ultima_de_fase = [], {}, None
    uid = 0
    for i_fase, (fase, items) in enumerate(PLAN):
        uid += 1                      # la fase ocupa un uid
        uid_fase = uid
        previa = None
        for clave, nombre, dias, quienes, deps in items:
            uid += 1
            if deps is None:
                deps = [previa] if previa else ([ultima_de_fase] if ultima_de_fase else [])
            ini = ARRANQUE if not deps else sig(max(fin_de[d] for d in deps))
            fin = sumar(ini, dias)
            fin_de[clave] = fin
            hojas.append(dict(uid=uid, fase=uid_fase, clave=clave, nombre=nombre, dias=dias,
                              quienes=quienes, deps=deps, ini=ini, fin=fin))
            previa = clave
        ultima_de_fase = previa
    return hojas, uid


def main() -> int:
    NS = "http://schemas.microsoft.com/project"
    ET.register_namespace("", NS)
    p = ET.Element("{%s}Project" % NS)

    sub(p, "SaveVersion", 14)
    sub(p, "Name", "ShopMetrics - Desarrollo de la plataforma")
    sub(p, "Title", "ShopMetrics - Desarrollo de la plataforma")
    sub(p, "Author", "Javier Gómez Avellaneda")
    sub(p, "Company", "ShopMetrics")
    sub(p, "ScheduleFromStart", 1)
    sub(p, "StartDate", "%sT08:00:00" % ARRANQUE.isoformat())
    sub(p, "FYStartDate", 1)
    sub(p, "CriticalSlackLimit", 0)
    sub(p, "CurrencyDigits", 2)
    sub(p, "CurrencySymbol", "US$")
    sub(p, "CurrencySymbolPosition", 0)
    sub(p, "CalendarUID", 1)
    sub(p, "DefaultStartTime", "08:00:00")
    sub(p, "DefaultFinishTime", "17:00:00")
    sub(p, "MinutesPerDay", JORNADA * 60)
    sub(p, "MinutesPerWeek", JORNADA * 60 * 5)
    sub(p, "DaysPerMonth", 20)
    sub(p, "DefaultTaskType", 0)
    sub(p, "DefaultFixedCostAccrual", 3)
    sub(p, "DefaultStandardRate", 0)
    sub(p, "DefaultOvertimeRate", 0)
    sub(p, "DurationFormat", 7)
    sub(p, "WorkFormat", 2)
    sub(p, "NewTasksEstimated", 0)
    sub(p, "SpreadActualCost", 0)
    sub(p, "MultipleCriticalPaths", 0)
    sub(p, "AutoAddNewResourcesAndTasks", 1)
    sub(p, "MicrosoftProjectServerURL", 1)
    sub(p, "Autolink", 0)
    sub(p, "NewTaskStartDate", 0)

    # ---- calendario: lunes a viernes de 8 a 17 con una hora de almuerzo
    cals = sub(p, "Calendars")
    cal = sub(cals, "Calendar")
    sub(cal, "UID", 1); sub(cal, "Name", "Standard"); sub(cal, "IsBaseCalendar", 1)
    sub(cal, "BaseCalendarUID", -1)
    dias = sub(cal, "WeekDays")
    for d in range(1, 8):
        wd = sub(dias, "WeekDay")
        sub(wd, "DayType", d)
        laborable = d not in (1, 7)          # 1 = domingo, 7 = sabado
        sub(wd, "DayWorking", 1 if laborable else 0)
        if laborable:
            horas = sub(wd, "WorkingTimes")
            for desde, hasta in (("08:00:00", "12:00:00"), ("13:00:00", "17:00:00")):
                wt = sub(horas, "WorkingTime")
                sub(wt, "FromTime", desde); sub(wt, "ToTime", hasta)

    # ---- recursos
    recursos = sub(p, "Resources")
    ident, tarifa = {}, {}
    for i, (nombre, mensual, clave) in enumerate(EQUIPO, start=1):
        ident[clave] = i
        tarifa[i] = round(mensual, 2)      # aca `mensual` ya es el valor hora
        r = sub(recursos, "Resource")
        sub(r, "UID", i); sub(r, "ID", i); sub(r, "Name", nombre)
        sub(r, "Type", 1); sub(r, "IsNull", 0)
        sub(r, "MaxUnits", "1")
        sub(r, "StandardRate", tarifa[i])
        sub(r, "StandardRateFormat", 2)          # por hora
        sub(r, "OvertimeRate", 0); sub(r, "OvertimeRateFormat", 2)
        sub(r, "CostPerUse", 0); sub(r, "AccrueAt", 3)
        sub(r, "CalendarUID", 1)

    # ---- tareas
    hojas, uid_max = programar()
    uid_por_clave = {h["clave"]: h["uid"] for h in hojas}
    tareas = sub(p, "Tasks")

    def enlace(nodo, pred_uid):
        e = sub(nodo, "PredecessorLink")
        sub(e, "PredecessorUID", pred_uid)
        sub(e, "Type", 1); sub(e, "CrossProject", 0)
        sub(e, "LinkLag", 0); sub(e, "LagFormat", 7)

    for i_fase, (fase, items) in enumerate(PLAN):
        de_la_fase = [h for h in hojas if h["clave"] in {it[0] for it in items}]
        uid_fase = de_la_fase[0]["fase"]
        ini_f = min(h["ini"] for h in de_la_fase)
        fin_f = max(h["fin"] for h in de_la_fase)

        nodo = sub(tareas, "Task")             # la fila resumen, sin vinculos
        sub(nodo, "UID", uid_fase); sub(nodo, "ID", uid_fase); sub(nodo, "Name", fase)
        sub(nodo, "Active", 1); sub(nodo, "Manual", 0)
        sub(nodo, "Type", 1); sub(nodo, "IsNull", 0)
        sub(nodo, "OutlineLevel", 1); sub(nodo, "Summary", 1); sub(nodo, "Milestone", 0)
        sub(nodo, "Start", "%sT08:00:00" % ini_f.isoformat())
        sub(nodo, "Finish", "%sT17:00:00" % fin_f.isoformat())
        sub(nodo, "DurationFormat", 7); sub(nodo, "ConstraintType", 0)

        for h in de_la_fase:
            t = sub(tareas, "Task")
            sub(t, "UID", h["uid"]); sub(t, "ID", h["uid"]); sub(t, "Name", h["nombre"])
            sub(t, "Active", 1); sub(t, "Manual", 0)
            sub(t, "Type", 0); sub(t, "IsNull", 0)
            sub(t, "OutlineLevel", 2); sub(t, "Summary", 0); sub(t, "Milestone", 0)
            sub(t, "Start", "%sT08:00:00" % h["ini"].isoformat())
            sub(t, "Finish", "%sT17:00:00" % h["fin"].isoformat())
            sub(t, "Duration", dur(h["dias"])); sub(t, "DurationFormat", 7)
            sub(t, "Work", dur(h["dias"] * sum(h["quienes"].values())))
            sub(t, "ConstraintType", 0); sub(t, "EffortDriven", 0)
            sub(t, "FixedCostAccrual", 3)
            for d in h["deps"]:
                enlace(t, uid_por_clave[d])

    # ---- hito de cierre: el mismo dia en que termina la ultima tarea
    ultima = max(hojas, key=lambda h: h["fin"])
    uid_hito = uid_max + 1
    hito = sub(tareas, "Task")
    sub(hito, "UID", uid_hito); sub(hito, "ID", uid_hito)
    sub(hito, "Name", "Plataforma lista para salir al mercado")
    sub(hito, "Active", 1); sub(hito, "Manual", 0)
    sub(hito, "Type", 1); sub(hito, "IsNull", 0)
    sub(hito, "OutlineLevel", 1); sub(hito, "Summary", 0); sub(hito, "Milestone", 1)
    sub(hito, "Start", "%sT17:00:00" % ultima["fin"].isoformat())
    sub(hito, "Finish", "%sT17:00:00" % ultima["fin"].isoformat())
    sub(hito, "Duration", "PT0H0M0S"); sub(hito, "DurationFormat", 7)
    sub(hito, "ConstraintType", 0)
    enlace(hito, ultima["uid"])

    # ---- asignaciones
    asig = sub(p, "Assignments")
    total, horas_por, n = 0.0, {}, 0
    for h in hojas:
        for clave, ded in h["quienes"].items():
            n += 1
            ur = ident[clave]
            horas = h["dias"] * JORNADA * ded
            costo = round(horas * tarifa[ur], 2)
            total += costo
            horas_por[ur] = horas_por.get(ur, 0) + horas
            a = sub(asig, "Assignment")
            sub(a, "UID", n); sub(a, "TaskUID", h["uid"]); sub(a, "ResourceUID", ur)
            sub(a, "Units", ded)
            sub(a, "Work", "PT%.0fH0M0S" % horas)
            sub(a, "RegularWork", "PT%.0fH0M0S" % horas)
            sub(a, "RemainingWork", "PT%.0fH0M0S" % horas)
            sub(a, "Cost", costo); sub(a, "RemainingCost", costo)
            sub(a, "CostRateTable", 0)
            sub(a, "Start", "%sT08:00:00" % h["ini"].isoformat())
            sub(a, "Finish", "%sT17:00:00" % h["fin"].isoformat())

    ET.ElementTree(p).write(SALIDA, encoding="UTF-8", xml_declaration=True)

    print("escrito: %s" % os.path.basename(SALIDA))
    print("  arranque          : %s" % ARRANQUE.strftime("%d/%m/%Y"))
    print("  fin del desarrollo: %s" % ultima["fin"].strftime("%d/%m/%Y"))
    print("  tareas            : %d hoja + %d fases + 1 hito = %d" % (len(hojas), len(PLAN), uid_hito))
    print("  asignaciones      : %d" % n)
    print()
    print("  %-32s %8s %10s %12s" % ("Recurso", "horas", "USD/h", "costo"))
    for _, mensual, clave in EQUIPO:
        u = ident[clave]; hs = horas_por.get(u, 0)
        print("  %-32s %8.0f %10.2f %12s" % (clave, hs, tarifa[u], "{:,.2f}".format(hs * tarifa[u])))
    print("  %-32s %8s %10s %12s" % ("INVERSIÓN INICIAL", "", "", "{:,.2f}".format(total)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
