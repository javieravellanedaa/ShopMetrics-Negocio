# Cronograma de desarrollo — para Microsoft Project

`ShopMetrics-desarrollo.xml` es el cronograma que justifica la inversión
inicial. Se abre en Project con **Archivo → Abrir**, eligiendo *XML* en el tipo
de archivo. No hace falta convertir nada ni instalar complementos.

## De dónde sale el número

La inversión inicial es el costo de las horas de quienes construyen la
plataforma. Project lo calcula solo a partir de tres cosas: los recursos con su
valor hora, las tareas con su duración y quién trabaja en cada una.

El valor hora no está inventado: es el costo mensual con cargas patronales de
la hoja `Costos RRHH` dividido por las 150 horas mensuales que fija el anexo de
capacidad operativa. Por eso el resultado es coherente con el resto del plan.

| Recurso | USD/hora | Horas | Costo |
|---|---:|---:|---:|
| Gerente de Sistemas (CTO) | 16,27 | 744 | 12.104,88 |
| Desarrollador | 14,64 | 1.580 | 23.131,20 |
| Gerente General (CEO) | 16,27 | 44 | 715,88 |
| **Inversión inicial** | | **2.368** | **35.951,96** |

Del 6 de enero al 31 de octubre de 2025: 29 tareas en 11 fases más el hito de
salida, 41 filas en total. La fase 7 (Machine Learning) corre en paralelo con
la 6 y la 8 porque carga sobre el CTO y no compite por el desarrollador, que
es el recurso escaso; eso acorta cinco semanas sin cambiar el costo, porque
las horas son las mismas. Los vínculos van entre tareas, nunca entre fases.

## Por qué el año cero es 2025

El desarrollo es previo a salir a vender: *"si no tenés esto, no podés
arrancar"*. Como los ingresos del presupuesto arrancan en 2026, el desarrollo
tiene que caer en 2025. Termina en diciembre y la plataforma sale al mercado en
enero de 2026.

Esto además evita contar dos veces a la misma gente: en 2025 sus horas son
inversión, y desde 2026 su sueldo es costo de RRHH, que es como ya está cargado
en el presupuesto. La hoja `Costos RRHH` no tiene bloque 2025, así que no hay
superposición.

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
