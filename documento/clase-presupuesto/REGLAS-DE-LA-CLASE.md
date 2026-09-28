# Clase de presupuesto financiero — Scali, 22/09/2026

Grabación de 52:13. `transcripcion-cruda.txt` tiene el texto con marca de tiempo;
`capturas/` tiene la pantalla en los momentos donde muestra una fórmula o un
resultado. El audio es de aula y el transcriptor confunde palabras: donde el
texto crudo dice algo raro, vale lo que está acá, que está contrastado contra
lo que se ve en pantalla.

---

## Lo que hay que entregar en el segundo avance

Lo dice textual en el minuto **43:31**:

> hasta acá lo que tenemos que resolver para el segundo avance: tener el
> presupuesto hasta el presupuesto financiero resuelto, y tener el proyecto que
> justifique en el modelo de inversión esa inversión inicial con respecto a la
> solución tecnológica.

Y en **00:00**: el segundo avance *incluye* lo del primero, más toda la parte de
costos y de inversión.

**Lo que NO entra en el segundo avance.** En 49:16 arranca a hablar de la matriz
de riesgo, los escenarios, el plan de contingencia y las viabilidades — y lo
presenta como lo que sigue, no como parte de esta entrega. Además la viabilidad
legal depende de un curso que recién se dicta **el 28 de octubre** (49:56), así
que no puede estar exigida antes.

---

## Las cuatro reglas que más peso tienen

### 1. Sólo se amortizan bienes de uso

Amortizar es reconocer contablemente que un bien pierde valor con el uso. Se
amortizan **bienes de uso**: informática, muebles y útiles, instalaciones,
maquinaria, rodados, inmuebles.

**No se amortizan** (10:08 y 10:26): licencias, costos legales de armado de la
estructura, registración de marca. Nada que no sea un bien.

### 2. Los sensores del ejemplo están mal amortizados

Esto es lo más importante para ShopMetrics y lo dice sobre su propio archivo
(08:20):

> yo también quiero decirle que no hagan esto de acá, que tienen un ejemplo,
> porque amortizó los insumos (…) Esos son insumos. Si ustedes vienen a dar un
> servicio, esos son insumos que van a instalar (…) Son insumos, capital de
> trabajo, no tiene nada que ver con un bien de uso, no es amortizable.

El `Presupuesto financiero EJEMPLO V1.xlsx` amortiza cinco tipos de sensores y
la unidad central. **Está mal a propósito y lo aclara en clase.** Los sensores
que se instalan para prestar el servicio son insumo: van al **costo variable**,
no al modelo de inversión ni a amortizaciones.

Para ShopMetrics esto define dónde va cada cosa según quién sea el dueño del
sensor. Si el sensor queda instalado en el centro para dar el servicio, es
insumo. Si es equipamiento propio de la empresa, es bien de uso.

### 3. Las cuotas de amortización están reguladas

No se eligen (04:59). Las fija una resolución técnica del Consejo Profesional de
Ciencias Económicas:

| Rubro | Años |
|---|---|
| Informática y comunicaciones | 3 |
| Muebles y útiles | 10 |
| Maquinaria | 10 |
| Rodados | 10 |
| Inmuebles | 50 |

Si uno amortizara más rápido de lo permitido tendría pérdidas más grandes y
pagaría menos ganancias; por eso está regulado (05:54).

En 30:30 aclara el horizonte: el marco temporal es de **tres años**, así que un
bien a 3 años queda amortizado dentro del análisis y uno a 10 se lleva siete
cuotas al próximo plan de negocio. No se ponen las cuotas que caen fuera del
horizonte (18:09: *"nosotros no lo ponemos"*).

### 4. La amortización sirve para una sola cosa acá

No es un egreso financiero. El presupuesto es de flujo de fondos, y la
amortización no mueve plata. Se calcula **únicamente** para descontarla del
monto imponible del impuesto a las ganancias y pagar menos (07:07 y 18:30).

---

## Cómo se arma el presupuesto financiero

Todo lo que entra ya está calculado en las hojas anteriores. El orden es:

```
Ingresos                          ← Mod. ingresos
− Costos fijos                    ← Mod. egresos
− Costos variables
− Costos de RRHH
= Utilidad Antes de Impuestos (UAII)

− Impuesto a los Ingresos Brutos  = Ingresos × 3 %
− Impuesto a las Ganancias        = monto imponible del año anterior × 35 %
= Utilidad Después de Impuestos (UDII)

− Inversión                       ← Mod. inversión, en negativo
= Flujo de Fondos
```

### Ingresos brutos: 3 % sobre la facturación

Es provincial y cambia por jurisdicción (3 %, 3,5 %, 3,6 %). Con convenio
multilateral habría que repartirlo por provincia; acá se simplifica a un **3 %
directo sobre los ingresos** (21:53). Se paga aunque se venda a pérdida: *"¿facturaste? pagá
el impuesto"* (20:41).

### El IVA no se calcula

Decisión explícita (13:20). Un presupuesto profesional lleva todos los
impuestos, pero acá van sólo estos dos.

### Ganancias: 35 %, y con un año de desfasaje

Ésta es la parte que más se presta a error. Dos cosas:

**El monto imponible no es la utilidad.** Se le restan los ingresos brutos —si
no habría doble imposición— y las amortizaciones del período (23:05). En la
captura `min-00-23-20.jpg` se lee la fórmula en la barra: `=E16-E17-K16`, o sea
UAII − IIBB − amortizaciones.

**Se paga al año siguiente.** Las ganancias de un ejercicio se liquidan en abril
o mayo del año que viene (25:04). Como el emprendimiento arranca ahora y no hay
saldo anterior, **el primer año el impuesto a las ganancias es cero**, y el
último monto imponible queda fuera del horizonte. En el ejemplo la única celda
con impuesto es la de 2028 y toma el monto imponible de 2027: `=M23*0.35`.

---

## Los ratios de control

Debajo del flujo de fondos hay una tabla que divide cada costo por los ingresos
del mismo año. No es decorativa: es lo que él mira al corregir (35:35 — *"yo
generalmente los hago"*).

- **Costos de RRHH / ingresos.** En el ejemplo da 71 % el primer año. Un valor
  por encima de 100 % significa que la estructura se come todos los ingresos.
  Puede estar justificado en una empresa de base tecnológica, pero hay que
  justificarlo o arrancar con menos estructura (32:06).
- **Costo de promoción / ingresos.** Suma lo que está en promoción tanto en
  costos fijos (publicidad, ferias) como en variables (promociones de venta).
  Alrededor del 10–15 % le pareció razonable para ese negocio; para un
  e-commerce lo consideró bajo (34:22).
- **Coherencia entre promoción y salto de ventas** (34:41). Si las ventas casi
  se duplican de un año al otro, tiene que haber un esfuerzo de promoción que lo
  explique. Si no, el salto no se sostiene.

---

## Viabilidad: tasa de corte, VAN y TIR

**La tasa de corte** es el retorno mínimo que justifica invertir en este negocio
en lugar de una alternativa de bajo riesgo (36:43). Se arma mirando en qué otra
cosa podría ponerse la plata: plazo fijo, acciones, bonos, y la rentabilidad
promedio del sector.

**Depende de la moneda** (37:26). En el ejemplo la tasa de corte es **75 %**,
que es de pesos; para un proyecto en dólares sería mucho más baja. Y si el
proyecto se define en dólares, **todo** tiene que estar en dólares (39:10).

**VAN**: trae el flujo de fondos al momento cero. Si da positivo, el negocio es
viable. **TIR**: `=TIR(flujo de fondos; estimación)`. Si la TIR queda por encima
de la tasa de corte, el negocio se banca.

En el ejemplo (captura `min-00-43-05.jpg`): tasa de corte 75 %, VAN $1.592.952,
TIR 80,76 %, y la inversión se recupera en el tercer año.

Advertencia suya (42:07): si la TIR queda pegada a la tasa de corte, cualquier
escenario de riesgo va a dar mal. Conviene que haya holgura.

---

## Lo que tiene que justificar el modelo de inversión

El costo de la solución tecnológica tiene que estar atado a la propuesta de
valor del producto o servicio (00:38 y 45:13). Si se prometen ciertas
funcionalidades, la inversión y el tiempo de desarrollo tienen que dar para
construirlas.

**Cuidado con confundir inversión con estructura** (46:13). El equipo de
desarrollo es un costo de **RRHH**, no una inversión, y tiene que ser coherente
con el análisis de puestos que ya se hizo. Dos caminos:

- Ponerlo **dentro de la estructura**: paga aportes patronales y sueldo anual
  complementario.
- **Contratarlo afuera**: sin costos laborales.

Es una decisión a tomar y a justificar, no un detalle.

También (44:08): el desarrollo es **previo** a poder vender. Hay que ubicar en
el tiempo cuánto lleva; si arranca hoy, el año cero es el año de desarrollo.

---

## Estructura propia o free lance: la decisión sobre el equipo de desarrollo

Entre 46:00 y 48:46 responde una pregunta sobre dónde va el equipo que
construye el software. Es un tramo que en la primera transcripción se había
perdido —devolvía «No sé» repetido— y que se recuperó después; el audio tenía
el mismo nivel que el resto, así que no era silencio.

Lo primero que aclara es que el esfuerzo de desarrollo **sí es inversión**
(46:06):

> yo arranco con recursos mínimos una vez que inicio el proyecto, pero si
> requiero más recursos para el desarrollo, inversión, lo tengo que contemplar.

Lo segundo es que eso **no significa inflar la estructura permanente**
(46:13 y 47:48). Un departamento de desarrollo propio se justifica en una
empresa que ya creció y lanza servicios todo el tiempo, no en un
emprendimiento que arranca. Y advierte contra la tentación de pasar a la
estructura toda la gente que aparece en el cronograma.

Lo tercero es la disyuntiva concreta, y la deja explícitamente en manos del
alumno (48:17 y 48:26):

> ¿Lo puedo meter dentro de la estructura o lo puedo contratar como un free
> lance? Para no tener que pagar los costos laborales. Todo lo que agregue en
> esta estructura va a tener los aportes patronales y el sueldo anual
> complementario. (…) Ahí definirán ustedes cómo armar la estructura de costos.

Esto es directamente aplicable al problema del ratio de RRHH: contratar el
desarrollo afuera saca esos aportes de la estructura y baja el costo
recurrente, sin que deje de contarse la inversión del año cero.

---

## Qué está verificado y qué no

Un compañero pasó por mensaje tres cosas que habría pedido el profesor.
Contrastadas contra la grabación:

| Lo que se dijo | Estado |
|---|---|
| «El Excel hasta el presupuesto financiero» | **Confirmado.** Textual en 43:37. |
| «Pidió un Project para justificar la inversión inicial del software» | **Parcial.** El pedido está en 43:53 y lo que describe en 44:21–45:00 es un cronograma: tiempos, recursos, alcance y fecha de salida. Pero no dice «Microsoft Project» ni lo muestra en pantalla en ningún momento de los 52 minutos. |
| «La inversión va a ser el costo de las horas de los involucrados» | **No está dicho así.** Lo más cerca es 46:06, que pone los recursos de desarrollo del lado de la inversión. El método —horas por valor hora— es compatible con lo que explica, pero no aparece en esta clase. |

Los tres son razonables y el segundo y el tercero encajan con todo lo demás.
Conviene confirmarlos con el profesor antes de la entrega, porque el tercero
es el que define el número, y es el que menos respaldo tiene en la grabación.
