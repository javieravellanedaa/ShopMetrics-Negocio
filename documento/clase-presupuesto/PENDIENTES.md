# Segundo avance — estado al 28/09/2026

Archivos de la entrega, todos en el repo y en Drive (`STF/Primera entrega/`):

- `Presupuesto financiero ShopMetrics.xlsx` — el Excel, hasta el presupuesto
  financiero inclusive. 3.127 fórmulas, cero errores.
- `STF_Gomez_Javier_E1_v1.docx` y `.pdf` — el informe, con el capítulo 8
  hasta el presupuesto financiero (8.1 a 8.7). La matriz de riesgos, los
  escenarios y la viabilidad se sacaron: son del tercer avance.
- `project/ShopMetrics-desarrollo.xml` — el cronograma de desarrollo, con su
  verificación en Microsoft Project (`ShopMetrics-desarrollo-verificado.xml`
  y `project/capturas/`).

## Qué quedó hecho

**Excel.** Las tres hojas que faltaban (modelo de inversión, amortizaciones,
presupuesto financiero) con el formato clonado del ejemplo de la cátedra; el
Anexo en la disposición de la plantilla más las filas de capacidad y holgura;
el costo del kit de instalación como costo variable, que faltaba; la captura
del informe de recursos de Project y el gráfico de inversión pegados en
`Mod. inversión`, como en la plantilla; los botones, el logo y los gráficos
originales recuperados. Tasa de corte 30% en dólares.

**Modelo.** Rentable desde el segundo año, inversión recuperada en 2028, VAN
positivo al 30%:

| | 2026 | 2027 | 2028 |
|---|---:|---:|---:|
| Ingresos | 45.843 | 233.021 | 468.014 |
| UAII | −63.029 | +56.738 | +243.959 |
| Flujo de fondos | −66.701 | +47.178 | +211.397 |
| RRHH / ingresos | 139% | 36% | 26% |

Inversión año cero 67.705 (64.880 de desarrollo). Acumulado a fin de 2028:
+124.169. **VAN al 30%: +5.123 (7,6% de la inversión); TIR 32,4%.** El ajuste
final para que el VAN diera positivo fue abonos +10% (33/65/175) y volumen
+11%; la participación a fin de 2028 queda en 21,6%.

**Decisiones tomadas** (todas documentadas en `scripts/cerrar_presupuesto.py`):
precios 40/120/450 y 33/65/175; gasto de referencia del comercio 960/año con
fuente; altas de 2027 ×1,5 y de 2028 ×1,25, y todas ×1,11, con promoción del 15% de los
ingresos; participación final 21,6%; un vendedor en 2026, dos en 2027, tres
desde julio de 2028; segundo técnico en agosto de 2027 (cuando el anexo lo
pide); desarrollo y mantenimiento contratados afuera; fundadores a 1.200
brutos en 2026 y 2027.

**Project.** Cuatro roles contratados (arquitecto, backend, frontend, ML),
2.324 h, 06/01 → 01/09/2025, US$ 64.880 (arquitecto y ML a US$ 40/h, backend y frontend a US$ 22/h). Verificado tres veces en Project
2016; Project no movió ninguna fecha.

**Word.** Capítulo 8 puesto al día con los números nuevos (cada cifra del
texto se lee de la planilla al generar), 13 figuras nuevas en 8.5–8.7, las
figuras existentes reemplazadas adentro del .docx, encabezado con fecha
28/09/2026. Los puntos 8.8, 8.9 y 9 se quitaron por pertenecer al tercer
avance; el texto de 8.8 sigue en el original del primer avance en git
(843d78f) para retomarlo entonces.

## Qué queda para Javier

1. ~~Abrir el Excel en Excel~~ Hecho el 28/09 desde Windows: abre sin
   reparación, se ven el logo, las imágenes y los botones (ahora en las 19 hojas).
2. **Abrir el Word y actualizar el índice** (F9 sobre la tabla de contenido):
   los números de página del índice son los del primer avance.
3. **Revisar las decisiones del modelo**, sobre todo el 19,4% de participación
   y el sueldo reducido de los fundadores: son defendibles, pero son suyas.
4. **QA Tester como quinto rol** del cronograma: el ejemplo del profesor lo
   tiene; nuestro Project no. Si se agrega, se regenera y se vuelve a
   verificar en Windows.
5. El VAN positivo depende de sostener el volumen de 2027 (402 comercios a
   diciembre): es la hipótesis que los escenarios del tercer avance van a golpear.
   Holgura técnica mínima 6,7 h (diciembre de 2027): con más volumen, el segundo
   técnico tendría que entrar antes.

## Cómo regenerar todo

    python3 scripts/cerrar_presupuesto.py        # sobre el Excel previo al cierre
    python3 <skill>/recalc.py documento/Presupuesto\ financiero\ ShopMetrics.xlsx
    python3 scripts/restaurar_objetos.py documento/Presupuesto\ financiero\ ShopMetrics.xlsx
    python3 scripts/figuras_cap8.py
    git checkout HEAD~N -- documento/STF_Gomez_Javier_E1_v1.docx   # el del primer avance
    python3 scripts/word_cap8_v2.py
