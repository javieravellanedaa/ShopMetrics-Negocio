# -*- coding: utf-8 -*-
"""Genera el cronograma de desarrollo en XML que abre Microsoft Project.

La inversion inicial del negocio es el costo de las horas de la gente que
construye la plataforma. Project lo calcula solo: se le dan los recursos con su
valor hora, las tareas con su duracion y quien trabaja en cada una.

El valor hora no se inventa: sale de la hoja 'Costos RRHH' del presupuesto
--costo mensual con cargas patronales-- dividido por las 150 horas mensuales
que fija el anexo de capacidad operativa. Asi el numero que sale de aca es
coherente con el resto del plan, que es lo que se pide.

    python3 armar_project.py

Deja ShopMetrics-desarrollo.xml, que se abre con Archivo -> Abrir en Project.
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

# (nombre, costo mensual con cargas, iniciales)  -> hoja 'Costos RRHH' L16:L18
EQUIPO = [
    ("Gerente de Sistemas (CTO)", 2439.80, "CTO"),
    ("Desarrollador",             2195.94, "DEV"),
    ("Gerente General (CEO)",     2439.80, "CEO"),
]

# (fase, [(tarea, dias, {recurso: dedicacion})])
PLAN = [
 ("1. Definición y arquitectura", [
   ("Relevamiento funcional y alcance del producto", 8, {"CTO":1.0,"CEO":0.5}),
   ("Arquitectura de la solución y elección del stack", 7, {"CTO":1.0}),
   ("Modelo de datos y diccionario", 5, {"CTO":1.0,"DEV":0.5}),
 ]),
 ("2. Base de datos y cimientos", [
   ("Esquema, migraciones y series temporales", 8, {"DEV":1.0,"CTO":0.5}),
   ("Entorno de desarrollo y automatización de pruebas", 7, {"DEV":1.0}),
 ]),
 ("3. Integración con sistemas externos", [
   ("Conectores con puntos de venta", 12, {"DEV":1.0,"CTO":0.5}),
   ("Ingesta de sensores de conteo", 10, {"DEV":1.0}),
   ("Monitoreo de salud de las integraciones", 8, {"DEV":1.0}),
 ]),
 ("4. Motor de métricas e indicadores", [
   ("Cálculo de tráfico, conversión, ventas y vacancia", 12, {"DEV":1.0,"CTO":0.5}),
   ("Agregados por hora, zona y locatario", 8, {"DEV":1.0}),
 ]),
 ("5. Panel web del centro", [
   ("Sistema de diseño y navegación", 6, {"DEV":1.0}),
   ("Tablero de indicadores y series", 12, {"DEV":1.0}),
   ("Mapa del centro y ficha de locatarios", 10, {"DEV":1.0}),
 ]),
 ("6. Motor de reglas y alertas", [
   ("Definición y evaluación de reglas por umbral", 10, {"DEV":1.0,"CTO":0.5}),
   ("Centro de alertas y ciclo de atención", 10, {"DEV":1.0}),
   ("Canales de notificación", 5, {"DEV":1.0}),
 ]),
 ("7. Modelos de Machine Learning", [
   ("Detección de anomalías en la operación", 12, {"CTO":1.0,"DEV":0.5}),
   ("Predicción de riesgo de vacancia", 10, {"CTO":1.0}),
   ("Recomendación de mix de locatarios", 8, {"CTO":1.0}),
 ]),
 ("8. Portal del locatario", [
   ("Tablero del local y comparación con la categoría", 12, {"DEV":1.0}),
   ("Consentimiento sobre el uso de datos", 5, {"DEV":1.0}),
 ]),
 ("9. Vista móvil de operaciones", [
   ("Resolución de alertas en terreno", 10, {"DEV":1.0}),
   ("Historial de turno del operario", 6, {"DEV":1.0}),
 ]),
 ("10. Seguridad, permisos y auditoría", [
   ("Autenticación y matriz de permisos", 8, {"DEV":1.0,"CTO":0.5}),
   ("Auditoría de accesos a datos sensibles", 7, {"DEV":1.0}),
 ]),
 ("11. Pruebas y salida a producción", [
   ("Pruebas integrales y corrección", 12, {"DEV":1.0,"CTO":1.0}),
   ("Documentación y manual de uso", 5, {"DEV":1.0}),
   ("Despliegue y puesta en marcha", 6, {"CTO":1.0,"DEV":1.0,"CEO":0.25}),
 ]),
]


def habil(f: dt.date) -> dt.date:
    while f.weekday() >= 5:
        f += dt.timedelta(days=1)
    return f


def sumar(f: dt.date, dias: int) -> dt.date:
    """Devuelve el ultimo dia habil de una tarea que dura `dias`."""
    f = habil(f)
    for _ in range(dias - 1):
        f = habil(f + dt.timedelta(days=1))
    return f


def sig(f: dt.date) -> dt.date:
    return habil(f + dt.timedelta(days=1))


def dur(dias: int) -> str:
    return "PT%dH0M0S" % (dias * JORNADA)


def sub(padre, etiqueta, texto=None):
    e = ET.SubElement(padre, etiqueta)
    if texto is not None:
        e.text = str(texto)
    return e


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
    ident = {}
    for i, (nombre, mensual, clave) in enumerate(EQUIPO, start=1):
        ident[clave] = i
        r = sub(recursos, "Resource")
        sub(r, "UID", i); sub(r, "ID", i); sub(r, "Name", nombre)
        sub(r, "Type", 1); sub(r, "IsNull", 0)
        sub(r, "MaxUnits", "1")
        sub(r, "StandardRate", round(mensual / HORAS_MES, 2))
        sub(r, "StandardRateFormat", 2)          # por hora
        sub(r, "OvertimeRate", 0); sub(r, "OvertimeRateFormat", 2)
        sub(r, "CostPerUse", 0); sub(r, "AccrueAt", 3)
        sub(r, "CalendarUID", 1)

    # ---- tareas
    tareas = sub(p, "Tasks")
    asignaciones = []
    uid = 0
    ident_tarea = 0
    cursor = ARRANQUE
    fin_fase_previa = None

    for fase, items in PLAN:
        # la fase es una tarea resumen: sus fechas salen de lo que contiene
        uid += 1
        uid_fase = uid
        nodo_fase = sub(tareas, "Task")
        ident_tarea += 1
        inicio_fase = habil(cursor)
        primera = uid_fase

        detalle = []
        for nombre, dias, quienes in items:
            uid += 1
            ident_tarea += 1
            ini = habil(cursor)
            fin = sumar(ini, dias)
            detalle.append((uid, ident_tarea, nombre, dias, ini, fin, quienes,
                            uid - 1 if uid - 1 != uid_fase else None))
            cursor = sig(fin)
        fin_fase = detalle[-1][5]

        # -- la fila resumen
        sub(nodo_fase, "UID", uid_fase)
        sub(nodo_fase, "ID", uid_fase)
        sub(nodo_fase, "Name", fase)
        sub(nodo_fase, "Active", 1); sub(nodo_fase, "Manual", 0)
        sub(nodo_fase, "Type", 1); sub(nodo_fase, "IsNull", 0)
        sub(nodo_fase, "OutlineLevel", 1)
        sub(nodo_fase, "Summary", 1)
        sub(nodo_fase, "Milestone", 0)
        sub(nodo_fase, "Start", "%sT08:00:00" % inicio_fase.isoformat())
        sub(nodo_fase, "Finish", "%sT17:00:00" % fin_fase.isoformat())
        sub(nodo_fase, "DurationFormat", 7)
        sub(nodo_fase, "ConstraintType", 0)
        if fin_fase_previa is not None:
            enlace = sub(nodo_fase, "PredecessorLink")
            sub(enlace, "PredecessorUID", fin_fase_previa)
            sub(enlace, "Type", 1); sub(enlace, "CrossProject", 0)
            sub(enlace, "LinkLag", 0); sub(enlace, "LagFormat", 7)

        # -- las tareas de la fase
        for u, idt, nombre, dias, ini, fin, quienes, previa in detalle:
            t = sub(tareas, "Task")
            sub(t, "UID", u); sub(t, "ID", u); sub(t, "Name", nombre)
            sub(t, "Active", 1); sub(t, "Manual", 0)
            sub(t, "Type", 0); sub(t, "IsNull", 0)
            sub(t, "OutlineLevel", 2)
            sub(t, "Summary", 0); sub(t, "Milestone", 0)
            sub(t, "Start", "%sT08:00:00" % ini.isoformat())
            sub(t, "Finish", "%sT17:00:00" % fin.isoformat())
            sub(t, "Duration", dur(dias)); sub(t, "DurationFormat", 7)
            sub(t, "Work", dur(int(dias * sum(quienes.values()))))
            sub(t, "ConstraintType", 0)
            sub(t, "EffortDriven", 0)
            sub(t, "FixedCostAccrual", 3)
            if previa:
                enlace = sub(t, "PredecessorLink")
                sub(enlace, "PredecessorUID", previa)
                sub(enlace, "Type", 1); sub(enlace, "CrossProject", 0)
                sub(enlace, "LinkLag", 0); sub(enlace, "LagFormat", 7)
            for clave, ded in quienes.items():
                asignaciones.append((u, ident[clave], ded, dias, ini, fin))
        fin_fase_previa = uid_fase

    # ---- hito de cierre
    uid += 1
    hito = sub(tareas, "Task")
    cierre = habil(cursor)
    sub(hito, "UID", uid); sub(hito, "ID", uid)
    sub(hito, "Name", "Plataforma lista para salir al mercado")
    sub(hito, "Active", 1); sub(hito, "Manual", 0)
    sub(hito, "Type", 1); sub(hito, "IsNull", 0)
    sub(hito, "OutlineLevel", 1); sub(hito, "Summary", 0)
    sub(hito, "Milestone", 1)
    sub(hito, "Start", "%sT08:00:00" % cierre.isoformat())
    sub(hito, "Finish", "%sT08:00:00" % cierre.isoformat())
    sub(hito, "Duration", "PT0H0M0S"); sub(hito, "DurationFormat", 7)
    sub(hito, "ConstraintType", 0)
    enlace = sub(hito, "PredecessorLink")
    sub(enlace, "PredecessorUID", fin_fase_previa)
    sub(enlace, "Type", 1); sub(enlace, "CrossProject", 0)
    sub(enlace, "LinkLag", 0); sub(enlace, "LagFormat", 7)

    # ---- asignaciones
    asig = sub(p, "Assignments")
    total = 0.0
    tarifa = {ident[c]: m / HORAS_MES for _, m, c in EQUIPO}
    for i, (ut, ur, ded, dias, ini, fin) in enumerate(asignaciones, start=1):
        horas = dias * JORNADA * ded
        total += horas * tarifa[ur]
        a = sub(asig, "Assignment")
        sub(a, "UID", i); sub(a, "TaskUID", ut); sub(a, "ResourceUID", ur)
        sub(a, "Units", ded)
        sub(a, "Work", "PT%.0fH0M0S" % horas)
        sub(a, "RegularWork", "PT%.0fH0M0S" % horas)
        sub(a, "RemainingWork", "PT%.0fH0M0S" % horas)
        sub(a, "Cost", round(horas * tarifa[ur], 2))
        sub(a, "RemainingCost", round(horas * tarifa[ur], 2))
        sub(a, "CostRateTable", 0)
        sub(a, "Start", "%sT08:00:00" % ini.isoformat())
        sub(a, "Finish", "%sT17:00:00" % fin.isoformat())

    ET.ElementTree(p).write(SALIDA, encoding="UTF-8", xml_declaration=True)

    print("escrito: %s" % os.path.basename(SALIDA))
    print("  arranque        : %s" % ARRANQUE.strftime("%d/%m/%Y"))
    print("  fin del desarrollo: %s" % cierre.strftime("%d/%m/%Y"))
    print("  tareas          : %d en %d fases" % (uid, len(PLAN)))
    print("  asignaciones    : %d" % len(asignaciones))
    print()
    horas_por = {}
    for ut, ur, ded, dias, _i, _f in asignaciones:
        horas_por[ur] = horas_por.get(ur, 0) + dias * JORNADA * ded
    print("  %-32s %8s %10s %12s" % ("Recurso", "horas", "USD/h", "costo"))
    for _, mensual, clave in EQUIPO:
        u = ident[clave]
        h = horas_por.get(u, 0)
        print("  %-32s %8.0f %10.2f %12s"
              % (clave, h, mensual / HORAS_MES,
                 "{:,.0f}".format(h * mensual / HORAS_MES)))
    print("  %-32s %8s %10s %12s" % ("INVERSIÓN INICIAL", "", "", "{:,.0f}".format(total)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
