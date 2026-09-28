# Segundo avance — estado al 28/09/2026

Archivos de la entrega, todos en el repo y en Drive (`STF/Primera entrega/`):

- `Presupuesto financiero ShopMetrics.xlsx` — el Excel, hasta el presupuesto
  financiero inclusive. 3.127 fórmulas, cero errores.
- `STF_Gomez_Javier_E1_v1.docx` y `.pdf` — el informe, con el capítulo 8
  completo (8.1 a 8.8) y el punto 9 con la viabilidad financiera.
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

**Modelo.** Rentable desde el segundo año, inversión recuperada dentro de 2028:

| | 2026 | 2027 | 2028 |
|---|---:|---:|---:|
| Ingresos | 41.263 | 208.244 | 415.751 |
| UAII | −68.326 | +23.852 | +168.045 |
| Flujo de fondos | −71.814 | +15.355 | +149.007 |
| RRHH / ingresos | 160% | 42% | 31% |
| Promoción / ingresos | 16% | 16% | 10% |

Inversión año cero 87.825 (85.000 de desarrollo). Acumulado a fin de 2028:
+4.723. VAN al 30%: −66.158; TIR 1,2% — el retorno cae después del horizonte,
y así se explica en el 8.7 y en el punto 9 con la frase del profesor (27:26).

**Decisiones tomadas** (todas documentadas en `scripts/cerrar_presupuesto.py`):
precios 40/120/450 y 30/59/159; gasto de referencia del comercio 960/año con
fuente; altas de 2027 ×1,5 y de 2028 ×1,25 con promoción del 16% de los
ingresos; participación final 19,4%; un vendedor en 2026, dos en 2027, tres
desde julio de 2028; segundo técnico en agosto de 2027 (cuando el anexo lo
pide); desarrollo y mantenimiento contratados afuera; fundadores a 1.200
brutos en 2026 y 2027.

**Project.** Cuatro roles contratados (arquitecto, backend, frontend, ML),
2.324 h, 06/01 → 01/09/2025, US$ 85.000. Verificado tres veces en Project
2016; Project no movió ninguna fecha.

**Word.** Capítulo 8 puesto al día con los números nuevos (cada cifra del
texto se lee de la planilla al generar), 13 figuras nuevas en 8.5–8.7, las
figuras existentes reemplazadas adentro del .docx, las de riesgos
renumeradas (8.44–8.51), encabezado con fecha 28/09/2026.

## Qué queda para Javier

1. **Abrir el Excel en Excel (Windows) y confirmar que no pide reparación.**
   LibreOffice lo abre y recalcula, pero Excel es más estricto y no se pudo
   probar desde la Mac.
2. **Abrir el Word y actualizar el índice** (F9 sobre la tabla de contenido):
   los números de página del índice son los del primer avance.
3. **Revisar las decisiones del modelo**, sobre todo el 19,4% de participación
   y el sueldo reducido de los fundadores: son defendibles, pero son suyas.
4. **QA Tester como quinto rol** del cronograma: el ejemplo del profesor lo
   tiene; nuestro Project no. Si se agrega, se regenera y se vuelve a
   verificar en Windows.
5. El margen con que se recupera la inversión (+4.723) es estrecho: los
   escenarios del tercer avance lo van a poner a prueba.

## Cómo regenerar todo

    python3 scripts/cerrar_presupuesto.py        # sobre el Excel previo al cierre
    python3 <skill>/recalc.py documento/Presupuesto\ financiero\ ShopMetrics.xlsx
    python3 scripts/restaurar_objetos.py documento/Presupuesto\ financiero\ ShopMetrics.xlsx
    python3 scripts/figuras_cap8.py
    git checkout HEAD~N -- documento/STF_Gomez_Javier_E1_v1.docx   # el del primer avance
    python3 scripts/word_cap8_v2.py
