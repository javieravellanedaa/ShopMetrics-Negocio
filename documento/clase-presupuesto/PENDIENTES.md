# Presupuesto financiero — qué quedó armado y qué falta

Archivo: `Presupuesto financiero ShopMetrics V2.xlsx` (el V1 queda intacto).

## Lo que ya está

Las tres hojas que estaban vacías quedaron armadas y con las fórmulas
conectadas. 2.852 fórmulas, cero errores de cálculo.

- **Mod. inversión** — cuatro bloques (año cero, 2026, 2027, 2028) con los
  conceptos de ShopMetrics. Los totales suben al encabezado y de ahí al
  presupuesto.
- **Amortizaciones** — sólo bienes de uso, con las vidas útiles de la
  resolución técnica: informática 3 años, muebles 10. Cada bien amortiza desde
  el año en que se compra y sólo por los años que caen dentro del horizonte.
- **Presupuesto financiero** — la cascada completa, los cuatro ratios de
  control y VAN/TIR.

**Código de colores:** azul = dato que cargás a mano · negro = fórmula ·
verde = viene de otra hoja · amarillo = celda que falta completar.

## Lo que tenés que completar

Son tres cosas, todas en amarillo:

1. **Cantidad y precio unitario** de cada concepto en `Mod. inversión`, en los
   cuatro bloques. Con eso se llenan solas las amortizaciones y la fila de
   inversión del presupuesto.
2. **La tasa de corte**, en `Presupuesto financiero!C31`. El VAN y la TIR
   aparecen al cargarla.
3. **Revisar las vidas útiles** en `Amortizaciones` columna F si agregás rubros.

## Dos advertencias

### Los sensores no van en inversión

El ejemplo del profesor los amortiza, y en clase dice que eso está mal (minuto
08:20): un sensor que se instala en el comercio para prestar el servicio es
**insumo**, no bien de uso. Tu modelo ya lo tiene bien: está como *Reposición
de equipamiento no recuperable* en costos variables. No lo muevas.

El desarrollo de la plataforma y la registración de marca **sí** van en el
modelo de inversión, pero **no se amortizan**. Por eso están en la hoja de
inversión y no en la de amortizaciones.

### La tasa de corte del ejemplo no te sirve

El 75 % del ejemplo es una tasa en pesos. Tu proyecto está en dólares —la
hipótesis, la facturación y los costos están todos en USD— así que la tasa
tiene que ser de dólares, bastante más baja. El profesor lo aclara en 37:26 y
en 39:10 insiste en que si el proyecto se define en dólares, **todo** tiene que
estar en dólares.

## El problema de fondo: el negocio no cierra a tres años

Con los costos que están cargados hoy, el flujo de fondos da negativo los tres
años y no hay TIR posible:

|  | 2026 | 2027 | 2028 |
|---|---:|---:|---:|
| Ingresos | 20.863 | 79.130 | 160.047 |
| Costos de RRHH | 122.985 | 150.323 | 183.153 |
| **RRHH / ingresos** | **589 %** | **190 %** | **114 %** |
| Flujo de fondos | −133.590 | −117.933 | −82.941 |

El ratio de RRHH sobre ingresos es exactamente el que el profesor dijo que
mira al corregir (30:30 y 32:06). En su ejemplo arranca en 71 % y le pareció
alto. Acá arranca en 589 % y **no baja de 100 % en ningún año**: la estructura
se come todos los ingresos durante todo el horizonte.

Dicho de otro modo: tal como está, el VAN va a dar negativo con cualquier tasa
de corte y el negocio no es viable. Hay tres salidas y son decisiones tuyas:

- **Arrancar con menos estructura.** Es lo que él sugiere (32:06): *"yo voy a
  intentar arrancar con la menor estructura posible"*. Menos gente el primer
  año, creciendo con las ventas.
- **Contratar el desarrollo afuera** en vez de tenerlo en la estructura
  (48:22). Sin aportes patronales ni sueldo anual complementario, y además el
  desarrollo pasa a ser inversión del año cero en vez de costo recurrente.
- **Revisar la proyección de ventas.** 102 comercios el primer año sobre un
  mercado meta de 3.400 es un 3 %; si el embudo da para más, los ingresos
  suben sin tocar los costos.

Lo que no conviene es dejarlo así: es el primer número que va a mirar.
