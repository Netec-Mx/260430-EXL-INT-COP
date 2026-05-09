# Práctica 6: Uso de Microsoft Copilot en Excel

## Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 25 minutos |
| **Complejidad** | Alta |
| **Nivel de Bloom** | Crear |
| **Módulo** | 6 — Microsoft Copilot en Excel |
| **Versión de software requerida** | Microsoft 365 Apps (versión 2308 o superior) con licencia Copilot habilitada |

---

## Descripción General

En esta práctica el estudiante integrará Microsoft Copilot en un flujo de trabajo real de análisis de datos de ventas almacenado en OneDrive. Partiendo de la comprensión conceptual de la Lección 6.1 —arquitectura de Copilot, requisitos de licencia y flujo básico de activación— el estudiante usará el panel lateral de Copilot para automatizar tareas de formato, generar fórmulas complejas, producir resúmenes analíticos y crear gráficos mediante instrucciones en lenguaje natural. La práctica culmina con un escenario integrador donde el estudiante diseña y documenta sus propios prompts para limpiar, analizar y reportar un conjunto de datos sin procesar, evaluando críticamente la calidad de los resultados generados por la IA.

---

## Objetivos de Aprendizaje

Al completar esta práctica, el estudiante será capaz de:

- [ ] Verificar que Copilot está habilitado en su cuenta de Microsoft 365 y abrir el panel lateral de Copilot dentro de Excel.
- [ ] Formular prompts efectivos en español para automatizar tareas de formato, ordenación y creación de columnas calculadas.
- [ ] Solicitar a Copilot la generación de resúmenes analíticos, gráficos comparativos y fórmulas de búsqueda (BUSCARV / XLOOKUP) mediante lenguaje natural.
- [ ] Aplicar Copilot de manera autónoma en un escenario integrador para limpiar datos, generar un mini-reporte y documentar los prompts utilizados.
- [ ] Evaluar críticamente los resultados producidos por Copilot, identificando aciertos, limitaciones y necesidad de corrección manual.

---

## Prerrequisitos

### Conocimientos previos

| Área | Nivel requerido |
|---|---|
| Tablas de Excel (insertar, dar formato, filtrar) | Intermedio — Práctica 4 completada |
| Fórmulas condicionales y de búsqueda (SI, BUSCARV) | Intermedio — Prácticas 2 y 3 completadas |
| Creación y personalización de gráficos | Básico — Práctica 5 completada |
| Navegación general de la interfaz de Excel 365 | Intermedio — Prácticas 1–5 completadas |

### Acceso y configuración previa

- Cuenta de Microsoft 365 con **licencia de Copilot habilitada** (verificada por el administrador antes de la sesión).
- Archivo de práctica [Lab06_VentasAnuales (1).xlsx](./Lab06_VentasAnuales (1).xlsx)

- Conexión a Internet activa y estable (mínimo 10 Mbps) durante toda la práctica.
- Excel configurado en **idioma español** para que los nombres de funciones generados por Copilot coincidan con el idioma de la hoja.

> ⚠️ **Aviso crítico:** Si tu cuenta no tiene la licencia de Copilot activa, el botón de Copilot en la cinta no estará disponible o aparecerá atenuado. Notifica al instructor de inmediato para recibir acceso a la demostración proyectada como alternativa.

---

## Entorno de Laboratorio

### Hardware mínimo requerido

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 8ª gen / AMD Ryzen 5 (64 bits) | Intel Core i7 / AMD Ryzen 7 |
| Memoria RAM | 8 GB | 16 GB |
| Espacio en disco | 10 GB disponibles | 20 GB disponibles |
| Resolución de pantalla | 1366 × 768 px | 1920 × 1080 px |
| Conexión a Internet | 10 Mbps | 25 Mbps o superior |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | Microsoft 365 (≥ 2308) | Canal mensual actualizado |
| Microsoft Copilot en Excel | Integrado en M365 | Requiere licencia Copilot M365 |
| Microsoft OneDrive | Integrado con M365 | El archivo DEBE estar en OneDrive |
| Navegador web | Edge (recomendado) / Chrome 110+ | Para verificar licencia en portal M365 |
| Sistema Operativo | Windows 10 (21H2+) / Windows 11 | — |

### Configuración inicial del entorno

Realiza estos pasos **antes de comenzar** la práctica cronometrada:

```
1. Abre el navegador y accede a: https://portal.microsoft365.com
2. Inicia sesión con tu cuenta institucional de Microsoft 365.
3. Verifica que en "Aplicaciones" aparece el ícono de Copilot.
4. Navega a OneDrive y confirma que el archivo Lab06_VentasAnuales.xlsx
   está presente en la carpeta "Archivos de práctica / Módulo 6".
5. Abre el archivo directamente desde OneDrive (doble clic en OneDrive,
   NO descargues el archivo al disco local).
6. Una vez abierto en Excel, confirma que en la barra de título aparece
   el ícono de nube (☁) indicando que el archivo está en OneDrive.
```

> 💡 **Nota sobre idioma de fórmulas:** Esta práctica usa fórmulas en español (BUSCARV, SI, SUMA, etc.). Si tu Excel está en inglés, los equivalentes son: VLOOKUP, IF, SUM. Consulta al instructor si necesitas la tabla de equivalencias completa.

---

## Procedimiento Paso a Paso

---

### Paso 1: Verificar el acceso a Copilot y explorar el panel lateral

**Objetivo:** Confirmar que Copilot está habilitado en la cuenta y familiarizarse con la interfaz del panel antes de usarlo para tareas reales.

**Instrucciones:**

1. Con el archivo **`Lab06_VentasAnuales.xlsx`** abierto en Excel (desde OneDrive), observa la cinta de opciones en la parte superior.
2. Haz clic en la pestaña **Inicio**.
3. Localiza el botón **Copilot** en el extremo derecho del grupo de comandos (ícono con el logotipo multicolor de Copilot). Si no lo ves en **Inicio**, búscalo en la pestaña **Copilot** si existe en tu versión.


   > Si el botón aparece **atenuado (gris)**, detente aquí y notifica al instructor. No continúes hasta resolver el acceso.

4. Haz clic en el botón **Copilot**. Se abrirá un **panel lateral** a la derecha de la hoja de cálculo.

![Imagen práctica](../images/imagen%20146.png)

5. Observa los elementos del panel:
   - **Cuadro de texto** en la parte inferior (donde escribirás los prompts).
   - **Área de conversación** en la parte superior (donde aparecerán las respuestas).
   - **Sugerencias de prompts** predefinidas que Copilot muestra como punto de partida.
6. Lee las sugerencias que aparecen en el panel sin hacer clic en ninguna todavía.
7. En el cuadro de texto del panel, escribe el siguiente prompt de prueba y presiona **Enter**:

   ```
   ¿Qué tipo de datos contiene esta hoja de cálculo?
   ```

8. Lee la respuesta generada por Copilot. Observa que el asistente describe las columnas y el tipo de información presente en la tabla.

![Imagen práctica](../images/imagen%20147.png)


**Resultado esperado:**
El panel de Copilot está abierto y visible a la derecha. Copilot ha respondido describiendo el contenido del archivo: menciona columnas como Vendedor, Región, Producto, Mes, Unidades Vendidas e Ingresos Totales (o equivalentes según el archivo).

**Verificación:**
- [ ] El panel de Copilot está abierto sin mensajes de error.
- [ ] Copilot respondió en español describiendo la estructura de los datos.
- [ ] La respuesta menciona al menos 3 columnas del archivo correctamente.

---

### Paso 2: Convertir el rango en tabla y aplicar formato con Copilot

**Objetivo:** Usar Copilot para verificar que los datos están en formato de tabla (requisito técnico fundamental de la Lección 6.1) y aplicar un estilo de tabla profesional mediante lenguaje natural.

**Instrucciones:**

1. Haz clic en cualquier celda dentro del rango de datos de la hoja **"Ventas_2024"**.
2. Observa si los datos ya tienen formato de tabla (verás pestañas de filtro en los encabezados y la pestaña contextual **Diseño de tabla** en la cinta). Si no tienen formato de tabla, continúa con el paso 3; si ya son tabla, pasa al paso 5.
3. En el panel de Copilot, escribe el siguiente prompt:

   ```
   Convierte este rango de datos en una tabla de Excel con encabezados y aplica un estilo profesional de color azul.
   ```

4. Copilot mostrará una sugerencia o ejecutará la acción directamente. Si muestra un botón **"Aplicar"** o **"Insertar"**, haz clic en él para confirmar la acción.
5. Verifica que los datos ahora tienen el formato de tabla (encabezados con botones de filtro, filas alternadas con color).

![Imagen práctica](../images/imagen%20148.png)

6. Ahora escribe el siguiente prompt para ordenar los datos:

   ```
   Ordena la tabla por la columna "Ingresos Totales" de mayor a menor.
   ```

7. Copilot mostrará una vista previa o ejecutará la ordenación. Confirma la acción si se te solicita.

![Imagen práctica](../images/imagen%20149.png)

8. Escribe un tercer prompt para agregar una columna calculada:

   ```
   Agrega una nueva columna llamada "Ingreso Promedio por Unidad" que divida los Ingresos Totales entre las Unidades Vendidas para cada fila.
   ```

9. Copilot generará la fórmula y mostrará una vista previa. **Antes de aplicar**, lee la fórmula sugerida en el panel y anótala en tu hoja de respuestas.
10. Haz clic en **"Insertar columna"** o el botón equivalente que muestre Copilot para aplicar la columna calculada.

![Imagen práctica](../images/imagen%20150.png)

**Resultado esperado:**
- Los datos están en formato de tabla con estilo azul aplicado.
- La tabla está ordenada de mayor a menor por Ingresos Totales.
- Existe una nueva columna "Ingreso Promedio por Unidad" con la fórmula `=[Ingresos Totales]/[Unidades Vendidas]` (o equivalente) aplicada en todas las filas.

**Verificación:**
- [ ] La pestaña contextual **Diseño de tabla** aparece en la cinta cuando se selecciona la tabla.
- [ ] La columna "Ingreso Promedio por Unidad" existe y muestra valores numéricos (no errores).
- [ ] Los datos están ordenados correctamente (el valor más alto de Ingresos Totales aparece en la primera fila de datos).
- [ ] La fórmula generada por Copilot está anotada en la hoja de respuestas.

---

### Paso 3: Generar un resumen analítico con lenguaje natural

**Objetivo:** Solicitar a Copilot un análisis descriptivo de los datos de ventas que identifique tendencias, productos destacados y regiones con mejor desempeño, practicando la formulación de prompts analíticos efectivos.

**Instrucciones:**

1. Asegúrate de que el panel de Copilot sigue abierto. Si lo cerraste, vuelve a abrirlo desde la pestaña **Inicio**.
2. Haz clic en cualquier celda dentro de la tabla de ventas para que Copilot tenga contexto del conjunto de datos activo.
3. Escribe el siguiente prompt analítico en el panel de Copilot:

   ```
   Muéstrame un resumen de los datos de ventas: ¿cuáles son los 3 productos con mayores ingresos totales, qué región tiene el mejor desempeño y hay alguna tendencia visible por mes?
   ```

4. Lee detenidamente la respuesta generada por Copilot. El asistente producirá un resumen en texto con los hallazgos principales.

![Imagen práctica](../images/imagen%20151.png)

5. Evalúa la respuesta: ¿Los productos mencionados coinciden con lo que puedes verificar visualmente en la tabla ordenada del Paso 2? ¿La región mencionada parece correcta?
6. Si la respuesta es imprecisa o incompleta, refina el prompt. Escribe:

   ```
   Calcula el total de ingresos por región y muéstrame los resultados ordenados de mayor a menor.
   ```

7. Copilot generará una tabla resumen o un conjunto de datos agrupados. Si ofrece insertar los resultados en la hoja, haz clic en **"Insertar"** o **"Agregar a hoja"** para colocarlos en una nueva área de la hoja.
8. Escribe un tercer prompt para profundizar en tendencias temporales:

   ```
   ¿En qué mes se registraron los mayores ingresos totales? Muéstrame el total mensual de ingresos.
   ```
![Imagen práctica](../images/imagen%20152.png)

9. Anota en tu **Hoja de Respuestas** (ver Sección 6 de este laboratorio):
   - Los 3 productos con mayores ingresos según Copilot.
   - La región con mejor desempeño.
   - El mes con mayores ingresos.
   - Una observación personal sobre si los resultados parecen correctos o si detectaste alguna inconsistencia.

**Resultado esperado:**
Copilot ha generado al menos dos respuestas analíticas: un resumen descriptivo en texto y una tabla de totales por región o por mes. Los datos del resumen son verificables contra la tabla original.

**Verificación:**
- [ ] El panel de Copilot muestra al menos 3 intercambios de conversación (3 prompts y sus respuestas).
- [ ] Existe una tabla de resumen insertada en la hoja o visible en el panel con totales por región.
- [ ] La Hoja de Respuestas tiene completados los 4 puntos del paso 9.
- [ ] El estudiante ha verificado al menos uno de los valores del resumen contra la tabla original.

---

### Paso 4: Crear un gráfico comparativo mediante un prompt descriptivo

**Objetivo:** Generar un gráfico de barras comparativo de ventas por región usando únicamente una instrucción en lenguaje natural dirigida a Copilot, y evaluar la calidad del gráfico producido.

**Instrucciones:**

1. En el panel de Copilot, escribe el siguiente prompt:

   ```
   Crea un gráfico de barras que compare los ingresos totales por región. Usa colores distintos para cada región y agrega un título descriptivo al gráfico.
   ```

2. Copilot mostrará una vista previa del gráfico sugerido en el panel. Observa:
   - ¿El tipo de gráfico es el correcto (barras)?
   - ¿Las regiones están representadas en el eje correcto?
   - ¿Los valores parecen coherentes con los datos?
3. Haz clic en **"Agregar a hoja"** o el botón equivalente para insertar el gráfico en la hoja de cálculo.
4. Una vez insertado, haz clic en el gráfico para seleccionarlo. Observa las pestañas contextuales **Diseño de gráfico** y **Formato** que aparecen en la cinta.
5. Evalúa el gráfico generado según los siguientes criterios y anota tus observaciones en la Hoja de Respuestas:
   - ¿El título del gráfico es descriptivo y profesional?
   - ¿Los ejes tienen etiquetas claras?
   - ¿La leyenda es legible?
   - ¿Hay algún elemento que mejorarías manualmente?

![Imagen práctica](../images/imagen%20153.png)

6. Si el gráfico necesita mejoras, intenta primero pedírselas a Copilot. Escribe:

   ```
   Agrega etiquetas de datos al gráfico de barras que muestre los valores exactos de ingresos sobre cada barra.
   ```

7. Si Copilot no puede modificar el gráfico directamente, realiza esta mejora manualmente: haz clic derecho sobre las barras del gráfico → **Agregar etiquetas de datos**.

**Resultado esperado:**
Un gráfico de barras está insertado en la hoja de cálculo con:
- Título descriptivo (ej. "Ingresos Totales por Región — 2024").
- Barras de colores distintos por región.
- Etiquetas de datos visibles.

**Verificación:**
- [ ] El gráfico está insertado en la hoja (no solo en el panel de Copilot).
- [ ] El gráfico es de tipo barras (horizontal o vertical) y muestra datos por región.
- [ ] Las etiquetas de datos son visibles sobre las barras.
- [ ] Las observaciones de evaluación están anotadas en la Hoja de Respuestas.

---

### Paso 5: Generar fórmulas de búsqueda con Copilot (BUSCARV / XLOOKUP)

**Objetivo:** Solicitar a Copilot que genere y explique una fórmula BUSCARV o XLOOKUP para cruzar datos entre dos tablas, comprendiendo la lógica de la fórmula generada antes de aplicarla.

**Instrucciones:**

1. Navega a la hoja **"Catálogo_Productos"** del mismo libro (pestaña en la parte inferior de la pantalla). Esta hoja contiene una tabla con columnas: **ID_Producto**, **Nombre_Producto**, **Categoría** y **Precio_Lista**.
2. Regresa a la hoja **"Ventas_2024"**. Observa que existe una columna **ID_Producto** pero no hay columna de **Categoría** ni de **Precio_Lista**.
3. Haz clic en la celda del encabezado de la primera columna vacía a la derecha de la tabla de ventas.
4. En el panel de Copilot, escribe el siguiente prompt:

   ```
   Necesito traer la columna "Categoría" desde la hoja "Catálogo_Productos" hacia la tabla de ventas, usando el ID_Producto como clave de búsqueda. ¿Qué fórmula debo usar y cómo la escribo?
   ```

5. Copilot generará una explicación y una fórmula. La fórmula esperada será similar a una de estas dos opciones:

   ```excel
   =BUSCARV([@ID_Producto],Catálogo_Productos[#Todo],3,0)
   ```

   ```excel
   =XLOOKUP([@ID_Producto],Catálogo_Productos[ID_Producto],Catálogo_Productos[Categoría],"No encontrado")
   ```

6. **Antes de aplicar la fórmula**, léela detenidamente y responde en tu Hoja de Respuestas:
   - ¿Qué argumento representa el valor buscado?
   - ¿Qué argumento define dónde buscar?
   - ¿Qué argumento define qué devolver?
   - ¿Cuál es la diferencia principal entre BUSCARV y XLOOKUP según la explicación de Copilot?

7. Haz clic en el botón **"Insertar fórmula"** o **"Copiar fórmula"** que muestra Copilot, o copia manualmente la fórmula desde el panel.
8. Pega la fórmula en la celda del encabezado vacío que seleccionaste en el paso 3. Escribe primero el encabezado de columna: `Categoría`.
9. En la primera celda de datos de la nueva columna, escribe o pega la fórmula generada por Copilot.
10. Verifica que la fórmula se propague automáticamente a todas las filas de la tabla (las tablas de Excel hacen esto automáticamente).

![Imagen práctica](../images/imagen%20154.png)

11. Comprueba algunos valores: selecciona una fila al azar, anota el **ID_Producto** y ve manualmente a la hoja **Catálogo_Productos** para verificar que la categoría devuelta es correcta.
12. Si aparece el error `#N/A` en alguna fila, escribe en Copilot:

    ```
    La fórmula devuelve #N/A en algunas celdas. ¿Cómo puedo manejar los errores de búsqueda para que muestre "Sin categoría" en lugar del error?
    ```

**Resultado esperado:**
La tabla de ventas tiene una nueva columna **"Categoría"** con los valores correctos traídos desde la hoja **Catálogo_Productos** mediante BUSCARV o XLOOKUP. Los errores `#N/A`, si existían, han sido manejados con un valor de texto alternativo.

**Verificación:**
- [ ] La columna "Categoría" existe en la tabla de ventas con datos en todas las filas.
- [ ] Al menos 3 filas han sido verificadas manualmente contra la hoja Catálogo_Productos.
- [ ] No hay errores `#N/A` visibles (o están manejados con texto alternativo).
- [ ] Las 4 preguntas del paso 6 están respondidas en la Hoja de Respuestas.

---

### Paso 6: Escenario integrador — Análisis autónomo con Copilot

**Objetivo:** Aplicar de manera autónoma todas las capacidades de Copilot exploradas en los pasos anteriores para limpiar, analizar y reportar un conjunto de datos sin procesar, documentando el proceso de ingeniería de prompts y evaluando críticamente los resultados.

> ⏱️ **Tiempo asignado para este paso: 12 minutos.** Es el paso más extenso y requiere trabajo independiente.

**Instrucciones:**

1. Navega a la hoja **"Datos_Brutos"** del libro. Esta hoja contiene un conjunto de datos de ventas del primer semestre sin formato de tabla, con posibles inconsistencias (filas vacías, nombres de regiones con mayúsculas/minúsculas inconsistentes, columnas desordenadas).

2. **Fase de exploración (2 min):** Antes de usar Copilot, observa los datos durante 1–2 minutos y anota en la Hoja de Respuestas al menos **3 problemas de calidad de datos** que identifies visualmente (ej. filas vacías, inconsistencias en nombres, columnas sin encabezado, etc.).

3. **Fase de limpieza con Copilot (3 min):** Diseña y escribe tus propios prompts para limpiar los datos. Debes lograr al menos:
   - Convertir el rango en tabla de Excel.
   - Identificar y eliminar filas completamente vacías.
   - Estandarizar los nombres de las regiones (si hay inconsistencias).
   
   Documenta **cada prompt que uses** en la Hoja de Respuestas, incluyendo si el resultado fue satisfactorio o requirió ajuste manual.

   Ejemplo de prompt inicial (modifícalo según lo que necesites):
   ```
   Este rango tiene problemas de calidad. Conviértelo en tabla, identifica filas vacías y dime cómo estandarizar los nombres de la columna Región.
   ```

4. **Fase de análisis con Copilot (4 min):** Usa el panel de Copilot para generar:
   - Un **resumen ejecutivo** en texto (2–3 párrafos) de los datos del primer semestre: tendencias, picos de venta, producto o región destacada.
   - Una **tabla de resumen** con totales por mes.
   - Un **gráfico de líneas** que muestre la evolución de ingresos mes a mes.
   
   Para cada solicitud, escribe un prompt original (no copies los del paso 3). Documenta todos los prompts en la Hoja de Respuestas.

5. **Fase de construcción del mini-reporte (3 min):** Inserta una nueva hoja llamada **"Reporte_Semestre1"** y organiza en ella:
   - El resumen ejecutivo generado por Copilot (cópialo desde el panel o pídele a Copilot que lo inserte directamente).
   - La tabla de resumen mensual.
   - El gráfico de líneas con texto alternativo configurado.
   - Una celda con el texto: `Reporte generado con asistencia de Microsoft Copilot — [tu nombre] — [fecha]`.

   Para crear la nueva hoja:
   ```
   Clic derecho en la pestaña de hoja → Insertar → Hoja de cálculo → Aceptar
   Doble clic en la pestaña → Escribe: Reporte_Semestre1 → Enter
   ```

6. **Evaluación crítica (tiempo libre restante):** En la Hoja de Respuestas, responde las siguientes preguntas de reflexión:
   - ¿En qué tarea fue más útil Copilot durante esta práctica?
   - ¿Identificaste algún resultado incorrecto o impreciso generado por Copilot? Descríbelo.
   - ¿Qué harías diferente en la formulación de tus prompts si repitieras esta práctica?
   - ¿En qué situaciones profesionales reales aplicarías Copilot en Excel?

**Resultado esperado:**
La hoja **"Reporte_Semestre1"** contiene: un resumen ejecutivo en texto, una tabla de resumen mensual, un gráfico de líneas con texto alternativo y la celda de autoría. La Hoja de Respuestas tiene documentados todos los prompts utilizados en este paso y las respuestas a las 4 preguntas de reflexión.

**Verificación:**
- [ ] La hoja "Reporte_Semestre1" existe en el libro y contiene los 4 elementos requeridos.
- [ ] El gráfico de líneas tiene texto alternativo configurado.
- [ ] La celda de autoría con nombre, fecha y mención a Copilot está presente.
- [ ] Al menos 5 prompts originales del Paso 6 están documentados en la Hoja de Respuestas.
- [ ] Las 4 preguntas de reflexión están respondidas con al menos 2 oraciones cada una.

---

## Validación y Pruebas Finales

Antes de entregar el archivo, realiza las siguientes verificaciones globales del libro:

### Lista de verificación de entrega

| # | Criterio de verificación | ✓ |
|---|---|---|
| 1 | El archivo está guardado en **OneDrive** (ícono de nube visible en barra de título) | ☐ |
| 2 | La hoja **"Ventas_2024"** contiene una tabla de Excel con estilo aplicado | ☐ |
| 3 | La tabla de ventas tiene la columna **"Ingreso Promedio por Unidad"** con valores numéricos | ☐ |
| 4 | La tabla de ventas tiene la columna **"Categoría"** con datos traídos desde Catálogo_Productos | ☐ |
| 5 | Existe un gráfico de barras por región en la hoja "Ventas_2024" con etiquetas y texto alternativo | ☐ |
| 6 | La hoja **"Reporte_Semestre1"** existe y contiene los 4 elementos del escenario integrador | ☐ |
| 7 | La **Hoja de Respuestas** tiene todos los campos completados (prompts, análisis, reflexión) | ☐ |
| 8 | No hay errores `#N/A`, `#¡REF!` o `#¡DIV/0!` visibles en ninguna hoja | ☐ |

### Prueba funcional de fórmulas

Para verificar que las fórmulas generadas por Copilot funcionan correctamente:

1. Selecciona una celda de la columna **"Ingreso Promedio por Unidad"** y verifica en la barra de fórmulas que la fórmula hace referencia a las columnas correctas.
2. Cambia temporalmente el valor de **Unidades Vendidas** en una fila a `0` y verifica que la celda de Ingreso Promedio muestra `#¡DIV/0!` (esto confirma que la fórmula está activa). Presiona **Ctrl+Z** para deshacer el cambio.
3. Selecciona una celda de la columna **"Categoría"** y verifica en la barra de fórmulas que la función BUSCARV o XLOOKUP hace referencia a la hoja correcta.

### Guardado final

```
Ctrl + S   →   Guardar (el archivo se guarda en OneDrive automáticamente)
```

Confirma que el archivo se guardó correctamente: la barra de título no debe mostrar el asterisco (*) de cambios sin guardar.

---

## Solución de Problemas

### Problema 1: El botón de Copilot aparece atenuado o no está visible en la cinta

**Síntomas:**
- El botón de Copilot en la pestaña **Inicio** aparece gris y no responde al hacer clic.
- La pestaña **Copilot** no existe en la cinta de opciones.
- Al hacer clic en Copilot aparece el mensaje: *"Copilot no está disponible para tu cuenta"*.

**Causa probable:**
La cuenta de Microsoft 365 del estudiante no tiene asignada la licencia de **Microsoft 365 Copilot**, o la licencia fue asignada recientemente y aún no se ha propagado al cliente de Excel (puede tardar hasta 24 horas en activarse). También puede ocurrir si el archivo está guardado en disco local en lugar de OneDrive.

**Solución:**

```
Paso 1: Verifica que el archivo está en OneDrive (no en C:\Usuarios\...).
        Si está en disco local, guárdalo en OneDrive:
        Archivo → Guardar una copia → OneDrive → [tu carpeta de práctica]

Paso 2: Verifica la licencia en el portal:
        https://admin.microsoft365.com → Usuarios → [tu cuenta]
        → Licencias y aplicaciones → Confirma "Microsoft 365 Copilot" = Activado

Paso 3: Si la licencia está activa pero Copilot no aparece, cierra Excel
        completamente y vuelve a abrirlo. Si persiste, cierra sesión en
        Excel (Archivo → Cuenta → Cerrar sesión) y vuelve a iniciar sesión.

Paso 4: Si ninguno de los pasos anteriores funciona, notifica al instructor
        para acceder a la demostración proyectada como alternativa.
```

---

### Problema 2: Copilot genera una fórmula que devuelve errores `#N/A` o resultados incorrectos en todas las filas

**Síntomas:**
- La columna "Categoría" generada por Copilot muestra `#N/A` en todas o la mayoría de las filas.
- La fórmula BUSCARV o XLOOKUP parece correcta en el panel de Copilot pero falla al ejecutarse en la hoja.
- El gráfico generado por Copilot muestra datos incorrectos o no coincide con los datos de la tabla.

**Causa probable:**
Copilot puede referenciar incorrectamente el nombre de la hoja o tabla si hay espacios, caracteres especiales o tildes en los nombres. También puede ocurrir que el tipo de dato del campo de búsqueda no coincida entre las dos tablas (ej. ID_Producto es texto en una tabla y número en la otra), lo que impide la coincidencia exacta.

**Solución:**

```
Para errores #N/A en BUSCARV/XLOOKUP:

Paso 1: Verifica el tipo de dato del campo ID_Producto en ambas hojas.
        Selecciona una celda de ID_Producto en "Ventas_2024" → observa
        si el valor está alineado a la izquierda (texto) o derecha (número).
        Haz lo mismo en "Catálogo_Productos".

Paso 2: Si los tipos no coinciden, pide a Copilot que corrija la fórmula:
        "La fórmula BUSCARV devuelve #N/A. El ID_Producto en Ventas_2024
        es texto pero en Catálogo_Productos es número. ¿Cómo corrijo esto?"

Paso 3: Copilot sugerirá envolver el valor buscado con TEXTO() o VALOR()
        según corresponda. Aplica la corrección sugerida.

Paso 4: Si la referencia a la hoja falla, verifica el nombre exacto de la
        hoja (sin espacios extra ni caracteres especiales) y ajusta la
        fórmula manualmente en la barra de fórmulas.

Para gráficos con datos incorrectos:

Paso 1: Haz clic derecho sobre el gráfico → "Seleccionar datos".
Paso 2: Verifica que el rango de datos del gráfico apunta a la tabla
        correcta y al rango correcto de columnas.
Paso 3: Ajusta el rango manualmente si es necesario y haz clic en Aceptar.
```

---

## Limpieza del Entorno

Al finalizar la práctica, realiza los siguientes pasos de cierre ordenado:

1. **Guarda el archivo final** con Ctrl+S y confirma que se guardó en OneDrive.
2. **Cierra el panel de Copilot** haciendo clic en la X del panel lateral para liberar espacio en pantalla.
3. **Verifica el historial de versiones** (opcional pero recomendado):
   ```
   Archivo → Información → Historial de versiones
   ```
   Confirma que existe al menos una versión guardada automáticamente en OneDrive durante la práctica.
4. **Cierra el libro** sin cerrar Excel:
   ```
   Ctrl + W   →   Si pregunta guardar, haz clic en "Guardar"
   ```
5. **No elimines el archivo de OneDrive**; el instructor lo usará para la evaluación.
6. Si usaste pestañas del navegador para verificar la licencia, ciérralas.

---

## Resumen

En esta práctica de nivel **Crear**, el estudiante completó un flujo de trabajo completo de análisis de datos asistido por inteligencia artificial usando Microsoft Copilot en Excel. Los logros clave de la sesión fueron:

| Fase | Habilidad desarrollada | Herramienta de Copilot usada |
|---|---|---|
| Activación | Verificar acceso y explorar el panel de Copilot | Panel lateral — consulta descriptiva |
| Automatización | Formatear tabla, ordenar datos, agregar columna calculada | Prompts de acción directa |
| Análisis | Generar resúmenes, totales por región y tendencias temporales | Prompts analíticos — respuesta en texto y tabla |
| Visualización | Crear gráfico de barras comparativo con etiquetas | Prompt descriptivo de gráfico |
| Fórmulas | Generar y comprender BUSCARV / XLOOKUP para cruce de tablas | Prompt de generación de fórmula |
| Integración | Análisis autónomo: limpieza, análisis y mini-reporte | Ingeniería de prompts independiente |

### Conceptos clave reforzados

- **Copilot requiere tablas de Excel** (no rangos simples) para funcionar correctamente — principio central de la Lección 6.1.
- La **calidad del prompt** determina directamente la calidad del resultado: prompts específicos y contextuales producen mejores respuestas que instrucciones vagas.
- Copilot es un **amplificador de productividad**, no un reemplazo del conocimiento de Excel: el estudiante debe poder evaluar críticamente si el resultado generado es correcto.
- El **almacenamiento en OneDrive** es un requisito técnico no negociable para que Copilot acceda al contexto del archivo.
- La **accesibilidad** (texto alternativo en gráficos) es una responsabilidad profesional que aplica incluso cuando el contenido es generado por IA.

### Recursos adicionales

| Recurso | URL |
|---|---|
| Microsoft Learn — Introducción a Copilot en M365 | https://learn.microsoft.com/es-es/copilot/microsoft-365/microsoft-365-copilot-overview |
| Microsoft Support — Usar Copilot en Excel | https://support.microsoft.com/es-es/office/usar-copilot-en-excel-d7110502-0334-4b4f-a175-a73abdfc118a |
| Galería de prompts de Copilot para Excel | https://adoption.microsoft.com/es-es/copilot-scenario-library/ |
| Microsoft Tech Community — Copilot in Excel Blog | https://techcommunity.microsoft.com/t5/excel-blog/copilot-in-excel/ba-p/3800842 |

---

## Hoja de Respuestas del Estudiante

> 📋 **Instrucciones:** Completa esta sección durante la práctica. Puedes escribir tus respuestas directamente en una hoja nueva del libro llamada **"Mis_Respuestas"**, o en el documento complementario indicado por tu instructor.

### Paso 2 — Fórmula generada por Copilot para columna calculada

```
Fórmula anotada: _______________________________________________
```

### Paso 3 — Hallazgos del resumen analítico

```
Top 3 productos por ingresos:
  1. _______________________________________________
  2. _______________________________________________
  3. _______________________________________________

Región con mejor desempeño: _______________________________________________

Mes con mayores ingresos: _______________________________________________

Observación de verificación (¿los resultados parecen correctos?):
_______________________________________________
_______________________________________________
```

### Paso 4 — Evaluación del gráfico generado por Copilot

```
¿El título es descriptivo y profesional? (Sí / No / Parcialmente): _______
¿Los ejes tienen etiquetas claras? (Sí / No): _______
¿La leyenda es legible? (Sí / No): _______
Elementos que mejorarías manualmente: _______________________________________________
```

### Paso 5 — Comprensión de la fórmula de búsqueda

```
Fórmula generada por Copilot (BUSCARV o XLOOKUP):
_______________________________________________

Argumento = valor buscado: _______________________________________________
Argumento = dónde buscar: _______________________________________________
Argumento = qué devolver: _______________________________________________

Diferencia principal entre BUSCARV y XLOOKUP según Copilot:
_______________________________________________
_______________________________________________
```

### Paso 6 — Escenario integrador

**Problemas de calidad de datos identificados visualmente:**
```
1. _______________________________________________
2. _______________________________________________
3. _______________________________________________
```

**Prompts utilizados y evaluación de resultados:**

| # | Prompt escrito | Resultado satisfactorio (S/N) | Ajuste manual requerido |
|---|---|---|---|
| 1 | | | |
| 2 | | | |
| 3 | | | |
| 4 | | | |
| 5 | | | |

**Preguntas de reflexión:**
```
1. ¿En qué tarea fue más útil Copilot?
_______________________________________________
_______________________________________________

2. ¿Identificaste algún resultado incorrecto de Copilot? Descríbelo.
_______________________________________________
_______________________________________________

3. ¿Qué harías diferente en tus prompts si repitieras esta práctica?
_______________________________________________
_______________________________________________

4. ¿En qué situaciones profesionales reales aplicarías Copilot en Excel?
_______________________________________________
_______________________________________________
```

---

*Fin del Laboratorio 06-00-01 — Práctica 6: Uso de Microsoft Copilot en Excel*
