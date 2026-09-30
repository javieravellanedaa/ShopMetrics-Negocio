# Correcciones de Scali sobre nuestro Excel — clase del 29/09/2026

Grabación de 43:01 (`IMG_9216 2.MOV`). El profesor abre **nuestro** archivo
`Presupuesto financiero ShopMetrics.xlsx` (versión del 28/09) proyectado y lo
recorre hoja por hoja. `transcripcion-cruda.txt` tiene el texto con marca de
tiempo; `pantallas/` tiene lo que se veía en cada momento. Donde la transcripción
confunde palabras, vale lo que está acá, contrastado con la pantalla.

**Plazo:** «tienen hasta mañana hasta las 12 horas» (36:37), es decir **30/09
a las 12:00**. Amplió a dos intentos más la entrega en la plataforma.

## Lo que corrigió, en orden de peso

### 1. La proyección de ventas tiene que dar la facturación objetivo (06:48–14:30, 18:41–26:02)

Es la corrección central y la repite cinco veces. El valor del mercado captado de
la hoja Hipótesis (participación × mercado total: 105.754 / 385.805 / 705.677)
**es el objetivo de facturación**, y la proyección de ventas tiene que dar ese
número, no 47.409:

> yo sé que me tiene que dar 105 (…) te tiene que dar 105, parecido a los 105,
> 385 el segundo año y 705 el tercero. Y vos me diste 47. (07:36, 08:00)

> De vuelta: no me importa la cantidad de clientes, me importa la facturación.
> Hacelo a partir de la facturación. Olvidate la cantidad de clientes. Si
> necesitás X suscripciones más, serán X suscripciones más. (19:25)

> Entonces no eran 110 clientes, eran 200, 300 clientes. La facturación. Y de
> ahí sacamos la cantidad de clientes que necesitamos. (21:14)

Alternativas que aceptó: subir la lista de precios sólo si está justificado por
el análisis de precios del punto 6.2 del informe («acordate de cambiarlo en el
6.2», 23:16); o bajar la participación pretendida al 1 % si se mantienen los
110 clientes (21:35). **Se eligió respetar la participación y deducir las altas.**

Sobre los porcentajes mensuales (09:20–10:50, 27:04–29:00): el primer mes se
vende 0; el resto es una progresión (o la estacionalidad del negocio) y **la
suma tiene que dar 100 %**. En un negocio por suscripción la sumatoria de los
abonos que se van sumando es la que tiene que dar el total anual (11:47).

**Lo que se cambió:** las altas por mes de cada plan se recalcularon para que
cada año dé el objetivo (271 / 326 / 436 altas; 271 / 597 / 1.033 comercios
abonados al cierre). Facturación resultante: **105.865 / 386.897 / 705.506**.
La columna «Cantidad» de Hipótesis ahora sale de la proyección. Los precios no
cambiaron.

### 2. Cada número del mercado tiene que explicarse en el Excel (02:07–04:50)

> Si no tengo el Word no voy ahí. Por eso les estoy pidiendo que me lo pongan
> acá. (02:51) (…) yo tengo que entender, con algún detalle que ustedes me van
> a poner, qué número me están presentando. (03:10)

Preguntó qué eran «Relevados», «Ocupados», «Indumentaria» y si el segmento era
AMBA o Avellaneda. Y sobre los USD 960 anuales: «el mercado es muy chico, empieza
a hacer ruido por lo chico» (05:16), pero lo aceptó.

**Lo que se cambió:** la nota debajo de la tabla del mercado (Hipótesis B13)
dice qué es cada columna, de dónde sale, y que el segmento es el rubro
indumentaria de los ejes relevados de la Ciudad de Buenos Aires. El texto de
participación decía «8 % el segundo año» y la tabla 12 %: quedó 12 %.

### 3. Estructura lo más plana posible: sólo los dos fundadores (15:12–17:30, 29:14–34:00)

> ¿Por qué necesito 5 personas dentro de la organización? ¿No lo pueden hacer
> los dos? Supongo que los gerentes son los dos socios. (15:42)

> Voy a tratar de mantener la estructura de costos lo más plana posible, y los
> fijos más que nada. Todo lo que pueda tercerizar lo voy a tercerizar. (29:34)

Dónde va cada cosa (32:03–34:07): el estudio contable es un canon mensual en
**costos fijos**; desarrollo y testing tercerizados en **costos fijos**; el
freelance «que lo voy contratando y voy pagando según las empresas» va a
**costos variables**, sin aportes ni cargas. En recursos humanos quedan sólo
los que están en relación de dependencia, con bruto más aportes.

Y quiere ver **el despliegue de todos los puestos de todas las áreas** (30:30,
35:00) para comprobar que cada tarea está cubierta por alguien: por uno de los
dos socios o por un tercero.

**Lo que se cambió:** dotación sólo de los dos fundadores los tres años; el
técnico pasa a costo variable por hora (ligado a las horas del Anexo, contratado
por mes entero de 150 h); el vendedor pasa a comisión por alta (USD 40);
desarrollo, contabilidad y marketing ya estaban en fijos. En Costos RRHH se
agregó a la derecha el despliegue de los 15 puestos de la estructura, con la
leyenda de la plantilla (rosa: cubierto por otro puesto; amarillo: tercerizado)
y dónde está su costo.

### 4. El sueldo anual complementario (34:20)

> Acá lo sacaste el sueldo anual complementario. Eso se paga en junio y en
> diciembre, el medio sueldo que tenés que pagar. El template lo tiene
> calculado.

**Lo que se cambió:** en el costo mensual de RRHH, junio y diciembre llevan
media cuota más (`×1,5`), con los encabezados «Junio + SAC» y «Diciembre + SAC»
como en la plantilla.

### 5. Los insumos IoT son costo variable, por instalación (35:21)

> Vos tenés alguna solución de IoT, eso no son insumos… los insumos son
> variables, eso lo vas a tener que calcular a partir de las instalaciones que
> hagas.

Ya estaban como «kit de instalación»; **se renombraron** «Insumos IoT por
instalación» para que se vea de qué se trata.

### 6. Modelo de inversión: sólo lo que entra cada año (36:05–40:37)

Miró el desarrollo de la plataforma (64.880, del Project) y el registro de marca
y constitución legal. Sobre los bloques 2026-2028 con filas en cero:

> Háganmela fácil: dejen solamente lo que entra en cada uno de los años y
> cuánta inversión, porque están [las filas] en cero. (40:18)

**Lo que se cambió:** 2026 y 2027 dicen «no se prevén inversiones en el año»
(la estructura es la de los dos fundadores, ya equipados en el año cero); 2028
tiene sólo la renovación de las dos notebooks, amortizadas a los tres años.

### 7. Presupuesto financiero (40:52–41:47)

Buscó dónde se restaba la amortización y encontró la columna «Monto imponible»
(«ahí lo veo», 41:14): está bien como está. Sobre los ratios: los costos sumaban
232 % de las ventas en 2026 («puede quedar negativo, a estos niveles sí», 18:10);
con la facturación objetivo y la estructura plana quedan en 115 %. Y el
esfuerzo de promoción con una facturación pretendida alta tiene que estar
«por arriba del 10, 15 %» de las ventas (24:32): se reforzaron las campañas de
2027 y 2028 para que quede en 13 % / 12 % / 10 %.

### 8. Lo que aprobó sin cambios

- Tasa de corte 30 % en dólares: «es una buena tasa, interesante» (18:27). En
  pesos sería 60-70 %.
- La distinción instalación (una sola vez) / abono (se acumula) (13:12).
- Amortizaciones: «hiciste la amortización porque tenés bienes, bien» (40:38).
- El costo de desarrollo del Project, 64.880 (38:44).
- Si la inversión se recupera en el cuarto o quinto año, se puede plantear un
  plazo más largo (25:19), pero acá se recupera en el segundo.

## Resultado después de las correcciones

| | 2026 | 2027 | 2028 |
|---|---:|---:|---:|
| Ingresos | 105.865 | 386.897 | 705.506 |
| Costos fijos | 32.640 | 77.700 | 106.860 |
| Costos variables | 50.578 | 71.375 | 101.598 |
| Costos de RRHH | 38.061 | 38.061 | 63.435 |
| UAII | −15.414 | +199.762 | +433.613 |
| Flujo de fondos | −18.590 | +188.155 | +344.970 |

Inversión año cero 67.705. La inversión se recupera en 2027. VAN al 30 %:
+186.348; TIR 114 %. El primer año da negativo, como el profesor anticipó
(«seguramente el primer año va a dar negativo», 17:18).

## Qué queda para el Word (punto 6.2 y capítulo 8)

- Capítulo 8 regenerado con las figuras nuevas (`scripts/figuras_cap8.py` y
  `scripts/word_cap8_v2.py`).
- El punto 6.2 (precios) no cambia porque la lista de precios no se tocó.
