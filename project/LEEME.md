# Cronograma de desarrollo — para Microsoft Project

`ShopMetrics-desarrollo.xml` es el cronograma que justifica la inversión
inicial. Se abre en Project con **Archivo → Abrir**, eligiendo *XML* en el tipo
de archivo. No hace falta convertir nada ni instalar complementos.

## De dónde sale el número

La inversión inicial es el costo de las horas de quienes construyen la
plataforma. Project lo calcula solo a partir de los recursos con su valor hora,
las tareas con su duración y quién trabaja en cada una.

El desarrollo se contrata afuera. Los fundadores gestionan la empresa y no
participan de la construcción, así que no figuran; el relevamiento inicial y
la aceptación final quedan del lado del cliente, como en cualquier
contratación. Son cuatro roles, los imprescindibles para este software, y al
no ser empleados no llevan aportes patronales ni aguinaldo:

| Rol | Qué hace | USD/h | Horas | Costo |
|---|---|---:|---:|---:|
| Arquitecto de software | relevamiento, arquitectura, modelo de datos; acompaña al backend en lo crítico; lidera pruebas y despliegue | 40 | 524 | 20.960 |
| Desarrollador backend | API, base de series temporales, integraciones POS e IoT, métricas, reglas, seguridad | 22 | 944 | 20.768 |
| Desarrollador frontend | panel web, centro de alertas, portal del locatario, vista móvil | 22 | 616 | 13.552 |
| Ingeniero de machine learning | los tres modelos de la fase 7 | 40 | 240 | 9.600 |
| **Inversión inicial** | | | **2.324** | **64.880** |

Del 6 de enero al 1 de septiembre de 2025: 29 tareas en 11 fases más el hito
de salida, 41 filas. Backend y frontend son cadenas independientes: el
frontend arranca cuando hay API con datos que consumir (fines de abril) y
desde ahí corre sin cortes; los modelos arrancan cuando están las métricas.
Por eso las fases 5, 6, 7 y 10 se superponen entre mayo y junio. El backend
queda libre unas semanas entre fines de junio y mediados de agosto, esperando
que el frontend termine para las pruebas integrales: se contrata por bloques.

Valor hora de mercado para clientes argentinos en 2026: semi-senior US$ 25–45,
senior US$ 45–60. Fuentes: [Teclab](https://teclab.edu.ar/tecnologia-y-desarrollo/cuanto-cobra-un-programador-en-argentina/)
y [Cristian Tait](https://cristiantait.com/blog/programador-web-freelance-argentina-2026).

## Por qué el año cero es 2025

El desarrollo es previo a salir a vender: *"si no tenés esto, no podés
arrancar"*. Como los ingresos del presupuesto arrancan en 2026, el desarrollo
tiene que caer en 2025. Termina en diciembre y la plataforma sale al mercado en
enero de 2026.

Como el desarrollo es contratado, no hay superposición con `Costos RRHH`:
las horas de 2025 son una factura que va a la inversión, y la dotación de
2026 en adelante es la de operación.

## Qué revisar y cambiar en Project

Todo lo de abajo es una propuesta: está armado para que lo corrijas, no para
usarlo tal cual.

1. **Las duraciones.** Cada tarea tiene los días que me parecieron razonables
   para el alcance. Son lo primero que hay que ajustar.
2. **La dedicación.** En la columna de asignaciones, 1 es tiempo completo y 0,5
   es media jornada. Si el CTO no puede estar al 100 %, el plazo se estira.
3. **El alcance.** Las once fases cubren los 31 casos de uso especificados. Si
   salís al mercado con menos, sacá fases y el costo baja.
4. **Las dependencias.** Ya no van en cadena: la fase 7 se solapa con la 6 y
   la 8. Lo que queda encadenado es la línea del desarrollador (2, 3, 4, 5,
   6, 8, 9, 10, 11), que no se puede solapar consigo misma. Acortar más
   exige un segundo desarrollador, y eso sí cambia la inversión.

## Cómo devolverlo

En Project: **Archivo → Guardar como → XML de Project (\*.xml)**. Ese archivo
se puede leer y revisar acá. El `.mpp` no.

## Después

El total que dé Project va a `Mod. inversión` del presupuesto, en el bloque
*Inversión inicial - Año cero*, en el concepto **Desarrollo de la plataforma**.
Con eso se completan solas las amortizaciones y el presupuesto financiero.

Ojo: el desarrollo de la plataforma **no se amortiza** —igual que en el ejemplo
de la cátedra, donde el ítem "Sistema" está en el modelo de inversión pero no
en la hoja de amortizaciones—. Va a la inversión y de ahí al flujo de fondos.

## Regenerarlo

    python3 armar_project.py

Editar `PLAN` y `EQUIPO` dentro de `armar_project.py` y volver a correrlo
rehace el XML. Conviene si el cambio es grande; para ajustes finos es más
cómodo hacerlo directamente en Project.

## Verificación

Verificado tres veces en Microsoft Project 2016 (Windows), la última sobre esta
versión de cuatro roles: costo 64.880,00, fin 01/09/2025, cero sobreasignación,
y las 41 fechas coinciden al minuto con las del XML. El exportado desde Project
es `ShopMetrics-desarrollo-verificado.xml`; las capturas están en `capturas/`.
La hoja de recursos va pegada en `Mod. inversión` del presupuesto, como en la
plantilla de la cátedra.
