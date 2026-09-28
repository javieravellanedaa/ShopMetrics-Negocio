# Presupuesto financiero — qué quedó armado y qué falta

Archivo: `Presupuesto financiero ShopMetrics.xlsx` (el V1 queda intacto).

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

Revisadas las 21 hojas una por una, **faltan dos cosas**, las dos en amarillo:

1. **Cantidad y precio unitario** en `Mod. inversión`, 48 celdas en los cuatro
   bloques. De esas, la que manda es `D11` (*Desarrollo de la plataforma*),
   que sale del total del cronograma de Project.
2. **La tasa de corte**, en `Presupuesto financiero!C31`. El VAN y la TIR
   aparecen solos al cargarla. En dólares: el 75 % del ejemplo es de pesos.

El resto de las hojas del bloque presupuestario están completas y sus totales
viajan bien hasta el flujo de fondos. La única fórmula que hoy devuelve vacío
es el VAN, esperando la tasa.

> Corrección respecto de una versión anterior de este archivo: se habían
> marcado `Costos RRHH`, `Proy. ventas`, `Mod. ingresos` y `Costos variables`
> como incompletas comparando cantidad de celdas contra el ejemplo. Es mal
> indicador: el ejemplo es una empresa bastante más grande. Verificadas celda
> por celda, están completas —`Costos RRHH` tiene 360 celdas cargadas y
> ninguna vacía—.

## Un cabo suelto: el Anexo de capacidad operativa no lo usa nadie

En el ejemplo, el anexo calcula por plan cuántas horas lleva cada servicio y
cuánto cuesta esa unidad, y la hoja de costos variables **lee de ahí**: 252
celdas apuntan al anexo. El costo variable del servicio *es* el trabajo que
lleva prestarlo.

En ShopMetrics el anexo hace la misma cuenta —le da 3, 18 y 72 dólares por
instalación para los planes Básico, Vidriera y Cadena— pero **ninguna celda
del libro lo lee**. Los costos variables listan gastos de bolsillo con precios
escritos a mano (comisión, movilidad, nube, pasarela, medios de pago,
reposición) y toman las cantidades directo de la proyección de ventas.

No está mal: son dos formas distintas de costear. En el ejemplo la mano de obra
del servicio es costo variable; acá está en la estructura fija, como el técnico
de instalación y soporte. Pero deja el anexo calculando algo que no se usa, y
esa es la clase de hilo del que se tira en una corrección: si el anexo
justifica la dotación, conviene que se vea el vínculo, y si justifica el costo
unitario del servicio, conviene que los costos variables lo lean.

Lo dice en 09:08: «vas a tener un **costo variable unitario**, que está
compuesto por estos insumos, y el costo fijo es un total».

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
| Costos de RRHH | 122.985 | 144.954 | 168.386 |
| **RRHH / ingresos** | **589 %** | **183 %** | **105 %** |
| Flujo de fondos | −133.590 | −112.563 | −68.175 |

El ratio de RRHH sobre ingresos es exactamente el que el profesor dijo que
mira al corregir (30:30 y 32:06). En su ejemplo arranca en 71 % y le pareció
alto. Acá arranca en 589 % y **no baja de 100 % en ningún año**: la estructura
se come todos los ingresos durante todo el horizonte.

**Ajuste ya aplicado (28/09).** El anexo de capacidad operativa mostraba que los técnicos entraban un año antes de hacer falta: el segundo en septiembre de 2027 con un solo técnico cubriendo hasta diciembre, y el tercero en febrero de 2028 con dos cubriendo todo el año. Se corrió el segundo a enero de 2028 y se sacó el tercero del horizonte. Son US$ 20.137 menos en tres años, y el anexo ahora muestra mes a mes la capacidad de la dotación y la holgura, conectado por fórmula a `Costos RRHH`. Los números de la tabla ya reflejan el ajuste.

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
