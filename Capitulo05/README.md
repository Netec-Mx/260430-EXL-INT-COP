# Creación y Gestión de Gráficos

## Metadatos

| Campo            | Detalle                                      |
|------------------|----------------------------------------------|
| **Duración**     | 25 minutos                                   |
| **Complejidad**  | Media                                        |
| **Nivel Bloom**  | Aplicar (Apply)                              |
| **Módulo**       | 5 — Creación y Gestión de Gráficos           |
| **Versión Excel**| Microsoft 365 (versión 2308 o superior)      |

---

## Descripción General

En esta práctica construirás y personalizarás tres tipos de gráficos —columnas agrupadas, líneas y circular— a partir de un conjunto de datos de ventas trimestrales por producto y región. Aprenderás a mover gráficos a hojas dedicadas, a agregar y modificar series de datos, a intercambiar filas y columnas para cambiar la perspectiva de análisis, y a aplicar formato profesional incluyendo títulos, etiquetas, leyendas y estilos corporativos. Finalizarás agregando texto alternativo a cada gráfico mediante el Panel de Accesibilidad, garantizando que las visualizaciones sean inclusivas y accesibles para todos los usuarios.

---

## Objetivos de Aprendizaje

Al completar esta práctica, serás capaz de:

- [ ] Crear gráficos de columnas agrupadas, líneas y circular a partir de rangos de datos seleccionados e insertar hojas de gráficos dedicadas.
- [ ] Agregar nuevas series de datos y modificar el rango de datos de series existentes usando el cuadro de diálogo **Seleccionar datos**.
- [ ] Intercambiar filas y columnas en un gráfico para cambiar la perspectiva de visualización y analizar cuál comunica mejor la información.
- [ ] Personalizar los elementos del gráfico: título, etiquetas de datos, leyenda, ejes y líneas de cuadrícula con formato específico.
- [ ] Aplicar estilos y esquemas de color predefinidos.

---

## Requisitos Previos

### Conocimientos Necesarios

- Haber completado las Prácticas 1 a 4 del curso, o tener conocimiento equivalente de gestión de datos y tablas en Excel.
- Saber seleccionar rangos de datos contiguos y no contiguos (usando la tecla **Ctrl**).
- Comprensión general de los tipos de gráficos y para qué tipo de datos es apropiado cada uno (comparación, tendencia, proporción).

### Acceso y Licencias Requeridos

- Cuenta de Microsoft 365 activa con acceso a Microsoft Excel 365.

- Archivo de práctica |  [Lab05_Ventas_Graficos (1).xlsx](./Lab05_Ventas_Graficos(1).xlsx)



## Entorno de Laboratorio

### Hardware Recomendado

| Componente         | Mínimo Requerido                              | Recomendado                     |
|--------------------|-----------------------------------------------|---------------------------------|
| Procesador         | Intel Core i5 8ª gen / AMD Ryzen 5 (64 bits)  | Intel Core i7 / AMD Ryzen 7     |
| Memoria RAM        | 8 GB                                          | 16 GB                           |
| Espacio en disco   | 10 GB disponibles                             | 20 GB disponibles               |
| Resolución         | 1366 × 768 px                                 | 1920 × 1080 px                  |
| Dispositivo entrada| Teclado con caracteres en español             | Ratón externo (recomendado)     |
| Conectividad       | 10 Mbps mínimo                                | 25 Mbps o superior              |

### Software Requerido

| Software                | Versión Mínima             | Notas                                              |
|-------------------------|----------------------------|----------------------------------------------------|
| Microsoft Excel         | Microsoft 365 (v2308+)     | Instalado en español o con fórmulas en español     |
| Microsoft OneDrive      | Integrado con M365         | Para acceder al archivo de práctica                |
| Navegador web           | Edge / Chrome 110+ / Firefox 110+ | Como alternativa si se usa Excel Online   |
| Sistema Operativo       | Windows 10 (21H2+) o Windows 11 | —                                           |

### Configuración Inicial del Entorno

Antes de comenzar los pasos de la práctica, realiza la siguiente configuración:

1. Abre **Microsoft OneDrive** y navega a la carpeta del curso.
2. Localiza el archivo **[Lab05_Ventas_Graficos(1).xlsx](./Lab03_Inventario_Productos(1).xlsx)** y ábrelo directamente en **Excel de escritorio** (no en Excel Online) haciendo clic en **Abrir en la aplicación de escritorio**.
3. Verifica que el archivo se abre correctamente y que puedes ver las siguientes hojas en la parte inferior:
   - `Datos_Ventas`
   - `Datos_Año_Anterior`
4. Asegúrate de que la **cinta de opciones** está completamente visible (no contraída). Si está contraída, haz doble clic en cualquier pestaña para expandirla.
5. Confirma que Excel está configurado en **idioma español** verificando que la pestaña **Insertar** aparece con ese nombre (no "Insert").

> **Nota sobre el idioma:** Todas las fórmulas y comandos en esta práctica están escritos en español. Si tu instalación de Excel está en inglés, los menús tendrán nombres equivalentes en inglés (por ejemplo, **Insert** en lugar de **Insertar**, **Chart Design** en lugar de **Diseño de gráfico**).

---

## Pasos del Laboratorio

---

### Paso 1: Explorar los Datos de Práctica

**Objetivo:** Familiarizarse con la estructura de los datos antes de crear cualquier gráfico, identificando los rangos que se utilizarán en cada visualización.

#### Instrucciones

1. Con el archivo **`Lab05_Ventas_Graficos.xlsx`** abierto, haz clic en la pestaña **`Datos_Ventas`** en la parte inferior del libro.

2. Observa la estructura de la tabla. Los datos están organizados de la siguiente manera:

   | Columna | Contenido                          |
   |---------|------------------------------------|
   | A       | Trimestre (T1, T2, T3, T4)         |
   | B       | Región Norte — Ventas (miles $)    |
   | C       | Región Sur — Ventas (miles $)      |
   | D       | Región Este — Ventas (miles $)     |
   | E       | Región Oeste — Ventas (miles $)    |
   | G       | Producto (Producto A, B, C, D)     |
   | H       | Porcentaje de participación (%)    |

3. Identifica los siguientes rangos clave que usarás durante la práctica:

   - **Rango principal:** `A1:E5` → Ventas trimestrales por región (4 trimestres × 4 regiones).
   - **Rango circular:** `G1:H5` → Distribución porcentual por producto.

4. Haz clic en la pestaña **`Datos_Año_Anterior`** y observa que contiene el mismo rango de columnas (A1:E5) pero con los datos del año anterior. Este rango se usará en el Paso 4.

5. Regresa a la hoja **`Datos_Ventas`**.

#### Resultado Esperado

Debes poder identificar claramente los dos rangos de datos y comprender que:
- Las **filas** representan trimestres (categorías del eje X).
- Las **columnas B a E** representan las cuatro regiones (series de datos).
- La columna G y H contienen datos independientes para el gráfico circular.

#### Verificación

- [ ] Puedes ver los datos de los 4 trimestres en el rango `A1:E5`.
- [ ] La hoja `Datos_Año_Anterior` contiene datos con la misma estructura.
- [ ] Identificaste el rango `G1:H5` para el gráfico circular.

---

### Paso 2: Crear un Gráfico de Columnas Agrupadas

**Objetivo:** Insertar un gráfico de columnas agrupadas a partir del rango de datos de ventas trimestrales por región, y moverlo a una hoja de gráfico dedicada.

#### Instrucciones

**Parte A — Insertar el gráfico:**

1. En la hoja **`Datos_Ventas`**, selecciona el rango **`A1:E5`** (incluyendo los encabezados de trimestres y regiones).

2. Ve a la pestaña **Insertar** en la cinta de opciones.

3. En el grupo **Gráficos**, haz clic en el ícono **Insertar gráfico de columnas o de barras** (ícono con barras verticales).

4. En el menú desplegable que aparece, selecciona **Columna agrupada** (primera opción en la sección "Columna 2D"). Es el ícono con barras verticales de diferentes colores agrupadas lado a lado.
![Imagen práctica](../images/imagen%20120.png)

5. Excel insertará el gráfico como un **objeto incrustado** sobre la hoja `Datos_Ventas`. Verás el gráfico flotando sobre las celdas, con cuatro grupos de barras (uno por trimestre) y cuatro colores (uno por región).

   > **Observación:** En este momento, las **filas** (trimestres T1–T4) son las categorías del eje X, y las **columnas** (regiones) son las series representadas por diferentes colores. Recuerda esta configuración; la modificarás en el Paso 5.

5.5  — Verifica en la leyenda que aparecen las regiones (Norte, Sur, Este, Oeste) y en el eje X los trimestres (T1–T4). Si en cambio la leyenda muestra T1, T2, T3, T4, ve a Diseño de gráfico → Cambiar fila/columna para corregirlo antes de continuar.

![Imagen práctica](../images/imagen%20120a.png)

**Parte B — Mover el gráfico a una hoja dedicada:**

6. Asegúrate de que el gráfico esté seleccionado (verás un borde azul alrededor y los controladores de tamaño en las esquinas).

7. En la cinta de opciones, aparecerá la pestaña contextual **Diseño de gráfico**. Haz clic en ella.

8. En el extremo derecho del grupo **Ubicación**, haz clic en el botón **Mover gráfico**.

9. En el cuadro de diálogo **Mover gráfico**, selecciona la opción **Hoja nueva**.

10. En el campo de texto, borra el nombre predeterminado y escribe: `Columnas_Regiones`

![Imagen práctica](../images/imagen%20121a.png)

11. Haz clic en **Aceptar**.

12. Verifica que el gráfico ahora aparece como una hoja completa. En la barra de pestañas inferior del libro debes ver la nueva pestaña llamada `Columnas_Regiones`.

#### Resultado Esperado

- El gráfico de columnas agrupadas ocupa la totalidad de la hoja `Columnas_Regiones`.
- El eje horizontal (X) muestra los cuatro trimestres: T1, T2, T3, T4.
- La leyenda muestra los nombres de las cuatro regiones con sus colores correspondientes.
- La hoja `Datos_Ventas` ya no muestra el gráfico incrustado.


#### Verificación

- [ ] La pestaña `Columnas_Regiones` aparece en la barra inferior del libro.
- [ ] El gráfico muestra 4 grupos de barras (uno por trimestre) con 4 barras por grupo (una por región).
- [ ] La hoja `Datos_Ventas` está limpia, sin el gráfico incrustado.

---

### Paso 3: Crear un Gráfico de Líneas en una Hoja Nueva

**Objetivo:** Crear un gráfico de líneas para visualizar la tendencia de ventas por región a lo largo de los trimestres, y moverlo a su propia hoja de gráfico.

#### Instrucciones

1. Haz clic en la pestaña **`Datos_Ventas`** para regresar a la hoja de datos.

2. Selecciona nuevamente el rango **`A1:E5`**.

3. Ve a **Insertar → Gráficos → Insertar gráfico de líneas o de áreas** (ícono con líneas ascendentes).

4. En el menú desplegable, selecciona **Línea con marcadores** (cuarta opción en la sección "Línea 2D", muestra líneas con puntos marcadores en cada valor).

![Imagen práctica](../images/imagen%20123.png)

5. Excel insertará el gráfico de líneas como objeto incrustado en la hoja `Datos_Ventas`.

6. Con el gráfico seleccionado, ve a **Diseño de gráfico → Mover gráfico**.

![Imagen práctica](../images/imagen%20124.png)

7. Selecciona **Hoja nueva**, escribe el nombre `Lineas_Tendencia` y haz clic en **Aceptar**.

8. Observa el gráfico resultante:
   - Hay **cuatro líneas**, una por cada región.
   - El eje X muestra los trimestres T1 a T4.
   - Los marcadores en cada punto permiten identificar el valor exacto de cada trimestre.

![Imagen práctica](../images/imagen%20125.png)

#### Resultado Esperado

- La hoja `Lineas_Tendencia` contiene un gráfico de líneas con cuatro series (una por región), con marcadores en cada punto de datos.
- La tendencia ascendente, descendente o estable de cada región es claramente visible.

#### Verificación

- [ ] La pestaña `Lineas_Tendencia` existe en la barra inferior del libro.
- [ ] El gráfico muestra 4 líneas diferenciadas por color con marcadores.
- [ ] El eje X refleja los 4 trimestres del año.

---

### Paso 4: Crear un Gráfico Circular

**Objetivo:** Crear un gráfico circular para representar la distribución porcentual de ventas por producto, insertarlo como objeto incrustado en la hoja de datos.

#### Instrucciones

1. Regresa a la hoja **`Datos_Ventas`**.

2. Selecciona el rango **`G1:H5`** (columna de Productos y columna de Porcentaje de participación, incluyendo encabezados).

3. Ve a **Insertar → Gráficos → Insertar gráfico circular o de anillos** (ícono con un círculo dividido en sectores).

4. Selecciona **Circular** (primera opción, gráfico de pastel estándar en 2D).

![Imagen práctica](../images/imagen%20126.png)

5. Excel insertará el gráfico circular como objeto incrustado en la hoja `Datos_Ventas`.

6. **En esta ocasión, NO moverás el gráfico a una hoja separada.** Reposiciona el gráfico incrustado para que no tape los datos:
   - Haz clic en el borde del gráfico (no en su interior) y mantén presionado el botón del ratón.
   - Arrastra el gráfico hacia el área vacía a la derecha de los datos (aproximadamente desde la columna J hacia adelante).
   - Suelta el botón del ratón.

7. Si necesitas ajustar el tamaño, arrastra los controladores de esquina (círculos pequeños en las esquinas del gráfico) mientras mantienes presionada la tecla **Alt** para que el gráfico se ajuste a los bordes de las celdas.

#### Resultado Esperado

![Imagen práctica](../images/imagen%20127.png)

- Un gráfico circular con cuatro sectores (uno por producto) aparece como objeto incrustado en la hoja `Datos_Ventas`, posicionado a la derecha de los datos sin taparlos.
- La leyenda muestra los nombres de los cuatro productos.

#### Verificación

- [ ] El gráfico circular está visible en la hoja `Datos_Ventas` sin tapar los datos de las columnas A a H.
- [ ] El gráfico muestra 4 sectores diferenciados por color.
- [ ] La leyenda identifica cada sector con el nombre del producto correspondiente.

---

### Paso 5: Agregar una Nueva Serie de Datos a un Gráfico Existente

**Objetivo:** Agregar los datos de ventas del año anterior al gráfico de columnas agrupadas para permitir una comparación directa entre años, usando el cuadro de diálogo **Seleccionar datos**.

#### Instrucciones

1. Haz clic en la pestaña **`Columnas_Regiones`** para ir al gráfico de columnas.

2. Haz clic en cualquier área del gráfico para asegurarte de que está seleccionado.

3. Ve a la pestaña contextual **Diseño de gráfico** en la cinta.

4. En el grupo **Datos**, haz clic en **Seleccionar datos**.

5. Se abrirá el cuadro de diálogo **Seleccionar origen de datos**. Observa que en el panel izquierdo (**Entradas de leyenda (Series)**) ya aparecen las cuatro series actuales: Norte, Sur, Este, Oeste.

![Imagen práctica](../images/imagen%20128a.png)

6. Para agregar la serie del año anterior, haz clic en el botón **Agregar** (en el panel izquierdo).

7. Se abrirá el cuadro de diálogo **Modificar serie**. Completa los campos:

   - **Nombre de la serie:** Haz clic en el campo y luego haz clic en la celda **`B1`** de la hoja `Datos_Año_Anterior`. Verás que el campo muestra la referencia `=Datos_Año_Anterior!$B$1`.

   ![Imagen práctica](../images/imagen%20129.png)

   - **Valores de la serie:** Borra el contenido del campo, luego selecciona el rango **`B2:B5`** de la hoja `Datos_Año_Anterior`. El campo mostrará `=Datos_Año_Anterior!$B$2:$B$5`.

8. Haz clic en **Aceptar** para cerrar el cuadro **Modificar serie**.

9. Verás que la nueva serie "Norte (Año Anterior)" aparece en el panel izquierdo del cuadro **Seleccionar origen de datos**.

![Imagen práctica](../images/imagen%20130a.png)

10. Haz clic en **Aceptar** para cerrar el cuadro de diálogo principal.

11. Observa el gráfico: ahora tiene **cinco grupos de barras** por trimestre (las cuatro regiones del año actual más la región Norte del año anterior). La nueva serie aparece con un color diferente.

![Imagen práctica](../images/imagen%20131a.png)

#### Resultado Esperado

- El gráfico de columnas ahora muestra 5 series en total.
- La nueva serie "Norte (Año Anterior)" aparece como una barra adicional en cada grupo de trimestre.
- La leyenda se actualizó automáticamente para incluir la nueva serie.

#### Verificación

- [ ] El cuadro de diálogo **Seleccionar origen de datos** muestra 5 entradas en el panel de series.
- [ ] El gráfico muestra 5 barras por grupo de trimestre.
- [ ] La leyenda incluye la nueva serie.

---

### Paso 6: Modificar el Rango de Datos de una Serie Existente

**Objetivo:** Editar el rango de datos de la serie "Norte" del año actual para practicar la modificación de series existentes en el cuadro de diálogo **Seleccionar datos**.

#### Instrucciones

1. Permanece en la hoja **`Columnas_Regiones`** con el gráfico visible.

2. Ve a **Diseño de gráfico → Seleccionar datos** para abrir nuevamente el cuadro de diálogo.

3. En el panel izquierdo (**Entradas de leyenda**), haz clic para seleccionar la serie **Norte** (la del año actual, no la del año anterior).

4. Haz clic en el botón **Modificar** (debajo del panel izquierdo).

5. Se abre el cuadro **Modificar serie**. Observa los campos actuales:
   - **Nombre de la serie:** referencia a la celda con el encabezado "Norte".
   - **Valores de la serie:** referencia al rango `=Datos_Ventas!$B$2:$B$5`.

6. Haz clic dentro del campo **Nombre de la serie** y verifica que la referencia apunta a `=Datos_Ventas!$B$1`. Si es correcto, no modifiques este campo.

7. Haz clic dentro del campo **Valores de la serie**, selecciona todo el contenido del campo (Ctrl+A) y escribe directamente la nueva referencia:
   ```
   =Datos_Ventas!$B$2:$B$5
   ```
   > **Nota:** En este ejercicio el rango es el mismo, pero estás practicando el flujo de modificación. En un escenario real, aquí cambiarías el rango para incluir nuevos períodos de datos.

![Imagen práctica](../images/imagen%20133.png)

8. Haz clic en **Aceptar** para cerrar **Modificar serie**.

9. Haz clic en **Aceptar** para cerrar **Seleccionar origen de datos**.

#### Resultado Esperado

- El gráfico se actualiza (puede no verse un cambio visual si el rango es idéntico, lo cual es correcto para este ejercicio de práctica).
- Comprendes el flujo completo para modificar el rango de datos de cualquier serie existente.

#### Verificación

- [ ] Pudiste abrir el cuadro **Modificar serie** para la serie "Norte".
- [ ] Identificaste los campos **Nombre de la serie** y **Valores de la serie**.
- [ ] El gráfico sigue mostrando todas las series sin errores.


### Paso 7: Personalizar los Elementos del Gráfico

**Objetivo:** Agregar y formatear el título del gráfico, habilitar etiquetas de datos con formato de número, reposicionar la leyenda, ajustar la escala del eje vertical y eliminar las líneas de cuadrícula secundarias.

#### Instrucciones

**Parte A — Agregar y formatear el título del gráfico:**

1. Permanece en la hoja **`Columnas_Regiones`** con el gráfico seleccionado.

2. Haz clic en el botón **Elementos de gráfico** (el ícono con el símbolo **+** que aparece en la esquina superior derecha del gráfico cuando está seleccionado).

3. En el menú flotante, verifica que la casilla **Título del gráfico** está marcada. Si no lo está, márcala.

4. Haz doble clic sobre el texto del título del gráfico (que puede decir "Título del gráfico" por defecto) para entrar en modo de edición.

5. Selecciona todo el texto con **Ctrl+A** y escribe:
   ```
   Ventas Trimestrales por Región 2024
   ```

![Imagen práctica](../images/imagen%20134.png)


6. Para formatear el título: con el texto del título aún seleccionado (o haciendo clic en el borde del cuadro de título para seleccionarlo como objeto), ve a la pestaña **Inicio** en la cinta y aplica:
   - **Tamaño de fuente:** 14
   - **Negrita:** activada (Ctrl+N)

7. Haz clic fuera del título para deseleccionarlo.

**Parte B — Habilitar etiquetas de datos:**

8. Haz clic en el gráfico para seleccionarlo.

9. Haz clic en el botón **+** (Elementos de gráfico).

10. Marca la casilla **Etiquetas de datos**. Aparecerán los valores numéricos sobre cada barra.

![Imagen práctica](../images/imagen%20135.png)

11. Para formatear las etiquetas: haz clic derecho sobre cualquier etiqueta de datos en el gráfico y selecciona **Dar formato a etiquetas de datos**.

12. En el panel **Dar formato a etiquetas de datos** que aparece a la derecha, busca la sección **Número** (puede que necesites desplazarte hacia abajo).

13. En el campo **Categoría**, selecciona **Número** y establece **Posiciones decimales: 0** (cero decimales, ya que los valores son enteros en miles).

![Imagen práctica](../images/imagen%20136.png)

14. Cierra el panel de formato haciendo clic en la **X** del panel.

**Parte C — Reposicionar la leyenda:**

15. Haz clic en el botón **+** (Elementos de gráfico).

16. Pasa el cursor sobre **Leyenda** (sin hacer clic) para ver la flecha de submenú, y haz clic en ella.

17. Selecciona **Abajo** para mover la leyenda a la parte inferior del gráfico.

18. Verifica que la leyenda ahora aparece debajo del área de trazado del gráfico.

**Parte D — Ajustar la escala del eje vertical:**

19. Haz doble clic sobre el eje vertical (eje Y, el que muestra los valores numéricos en el lado izquierdo del gráfico) para abrir el panel **Dar formato a eje**.

20. En el panel derecho, expande la sección **Opciones de eje** si no está ya expandida.

21. Localiza el campo **Mínimo** y cambia el valor de `Automático` a un valor fijo:
    - Haz clic en el botón de opción **Fijo** junto a "Mínimo".
    - Escribe `0` en el campo.

22. Localiza el campo **Unidad principal** y establece un valor fijo apropiado (por ejemplo, `50` si los valores están en el rango de 0–500, o el valor que el instructor indique según los datos del archivo).

![Imagen práctica](../images/imagen%20137.png)

23. Cierra el panel de formato.

**Parte E — Eliminar las líneas de cuadrícula secundarias:**

24. Haz clic en el botón **+** (Elementos de gráfico).

25. Pasa el cursor sobre **Líneas de cuadrícula** y haz clic en la flecha de submenú.

26. Verifica que solo esté marcada la opción **Horizontal principal** y desmarca cualquier otra opción marcada (como **Horizontal secundaria** o **Vertical principal**).

![Imagen práctica](../images/imagen%20138.png)

27. Haz clic fuera del menú para cerrarlo.

#### Resultado Esperado

- El gráfico muestra el título "Ventas Trimestrales por Región 2024" en negrita, tamaño 14.
- Cada barra tiene una etiqueta con el valor numérico sin decimales.
- La leyenda está posicionada en la parte inferior del gráfico.
- El eje Y comienza en 0 con una escala apropiada.
- Solo las líneas de cuadrícula horizontales principales son visibles.

#### Verificación

- [ ] El título del gráfico dice "Ventas Trimestrales por Región 2024" en negrita.
- [ ] Las etiquetas de datos aparecen sobre cada barra sin decimales.
- [ ] La leyenda está en la parte inferior del gráfico.
- [ ] El eje Y comienza en 0.
- [ ] No hay líneas de cuadrícula secundarias visibles.

---

### Paso 8: Aplicar Estilos y Esquema de Colores

**Objetivo:** Aplicar un estilo de gráfico predefinido del catálogo de Excel y cambiar el esquema de colores para alinearlo con una identidad corporativa simulada (colores azul y gris).

#### Instrucciones

**Parte A — Aplicar un estilo de gráfico:**

1. Permanece en la hoja **`Columnas_Regiones`** con el gráfico seleccionado.

2. Ve a la pestaña **Diseño de gráfico** en la cinta.

3. En el grupo **Estilos de gráfico**, verás una galería de estilos numerados. Haz clic en el botón **Más** (flecha con una línea horizontal, en la esquina inferior derecha de la galería) para expandir todos los estilos disponibles.

4. Pasa el cursor sobre diferentes estilos para previsualizar cómo quedaría el gráfico. Observa cómo cambian los fondos, las texturas y el formato de las barras.

5. Selecciona el **Estilo 8** (fondo oscuro con barras de colores brillantes) o el estilo que el instructor indique. Haz clic sobre él para aplicarlo.

![Imagen práctica](../images/imagen%20139.png)


   > **Alternativa rápida:** También puedes hacer clic en el ícono de **pincel** (Estilos de gráfico) que aparece a la derecha del gráfico cuando está seleccionado, y seleccionar el estilo desde ese panel flotante.

**Parte B — Cambiar el esquema de colores:**

6. Con el gráfico aún seleccionado, permanece en la pestaña **Diseño de gráfico**.

7. Haz clic en el botón **Cambiar colores** (en el grupo **Estilos de gráfico**, a la izquierda de la galería de estilos).

8. Se desplegará un panel con dos secciones:
   - **Colorido:** paletas con múltiples colores distintos.
   - **Monocromático:** paletas con variaciones de un solo color.

9. Para simular una identidad corporativa en tonos azules, selecciona la opción **Color 1** de la sección **Monocromático** (paleta de azules). Si prefieres mantener múltiples colores pero con una paleta más corporativa, selecciona **Colorido 2** de la sección Colorido.

10. Observa cómo el gráfico actualiza sus colores inmediatamente.

![Imagen práctica](../images/imagen%20140.png)

11. Si el resultado no es visualmente satisfactorio, prueba con otras opciones hasta encontrar una combinación que comunique los datos de manera clara y profesional.

#### Resultado Esperado

- El gráfico tiene aplicado un estilo predefinido diferente al estilo predeterminado inicial.
- Los colores del gráfico reflejan una paleta coherente (monocromática en azules o colorida corporativa).
- El gráfico mantiene todos los elementos personalizados del Paso 8 (título, etiquetas, leyenda).

#### Verificación

- [ ] El estilo del gráfico es diferente al estilo predeterminado.
- [ ] Los colores de las barras siguen una paleta coherente.
- [ ] El título, las etiquetas y la leyenda siguen siendo visibles y legibles con el nuevo estilo.


### Paso 9: Personalizar el Gráfico Circular

**Objetivo:** Aplicar personalizaciones específicas al gráfico circular: agregar etiquetas con porcentajes, aplicar un estilo y separar un sector para destacar el producto principal.

#### Instrucciones

1. En la hoja **`Datos_Ventas`**, haz clic sobre el gráfico circular para seleccionarlo.

2. **Agregar etiquetas de porcentaje:**
   - Haz clic en el botón **+** (Elementos de gráfico).
   - Marca la casilla **Etiquetas de datos**.
   - Haz clic derecho sobre cualquier etiqueta y selecciona **Dar formato a etiquetas de datos**.
   - En el panel derecho, en la sección **Opciones de etiqueta**, marca la casilla **Porcentaje** y desmarca la casilla **Valor** (si estaba marcada).
   - Cierra el panel de formato.

   ![Imagen práctica](../images/imagen%20141.png)

3. **Separar un sector (efecto "explosión"):**
   - Haz clic una vez sobre el gráfico circular para seleccionar toda la serie.
   - Haz clic una segunda vez (lentamente, no doble clic) sobre el sector del **Producto A** para seleccionar solo ese sector.
   - Con ese sector seleccionado, haz clic derecho y selecciona **Dar formato a punto de datos**.

 ![Imagen práctica](../images/imagen%20142.png)

   - En el panel derecho, en la sección **Opciones de serie**, busca el control deslizante **Sección de puntos** y arrástralo hasta aproximadamente **10%**.
   - Cierra el panel.
 ![Imagen práctica](../images/imagen%20143.png)

4. **Aplicar un estilo al gráfico circular:**
   - Haz clic en el borde del gráfico para seleccionar el gráfico completo (no solo el sector).
   - Ve a **Diseño de gráfico → Estilos de gráfico** y selecciona el **Estilo 3** 

   ![Imagen práctica](../images/imagen%20144.png)

5. **Agregar título al gráfico circular:**
   - Haz clic en el botón **+** y verifica que **Título del gráfico** está marcado.
   - Haz doble clic sobre el título y escribe:
     ```
     Distribución de Ventas por Producto 2024
     ```



#### Resultado Esperado

- El gráfico circular muestra el porcentaje de cada sector como etiqueta (en lugar de los valores absolutos).
- El sector del Producto A está ligeramente separado del resto del círculo para destacarlo.
- El gráfico tiene el título "Distribución de Ventas por Producto 2024".

![Imagen práctica](../images/imagen%20145.png)

#### Verificación

- [ ] Las etiquetas muestran porcentajes (%) en lugar de valores absolutos.
- [ ] El sector del Producto A está visualmente separado de los demás sectores.
- [ ] El título del gráfico circular es visible y correcto.

---

### Paso 10: Guardar el Libro de Trabajo

**Objetivo:** Guardar todos los cambios realizados durante la práctica asegurando que el archivo se guarda correctamente en OneDrive.

#### Instrucciones

1. Presiona **Ctrl+G** para guardar el archivo (o ve a **Archivo → Guardar**).

2. Dado que el archivo ya está en OneDrive, se guardará automáticamente en la misma ubicación. Verifica que en la barra de título aparece el nombre del archivo sin el indicador de cambios no guardados (el asterisco `*` o el punto antes del nombre).

3. Verifica en la barra de estado de Excel (parte inferior de la pantalla) que no aparece ningún mensaje de error.

4. Como verificación adicional, ve a **Archivo → Información** y confirma que la ruta del archivo muestra una ubicación de OneDrive (debe contener `OneDrive` en la ruta).

5. Presiona **Escape** o haz clic en la flecha de retroceso para volver al libro.

#### Resultado Esperado

- El archivo se guarda correctamente en OneDrive con todos los gráficos creados y personalizados.
- La barra de título muestra el nombre del archivo sin indicadores de cambios no guardados.

#### Verificación

- [ ] El archivo se guardó sin errores.
- [ ] La ruta del archivo confirma que está en OneDrive.
- [ ] Todos los gráficos y personalizaciones son visibles después de guardar.

---

## Validación y Pruebas Finales

Antes de considerar la práctica completada, realiza las siguientes verificaciones globales:

### Lista de Verificación Final

| # | Elemento a Verificar | Ubicación | ✓ |
|---|----------------------|-----------|---|
| 1 | Hoja de gráfico `Columnas_Regiones` existe con gráfico de columnas agrupadas | Pestaña inferior del libro | ☐ |
| 2 | Hoja de gráfico `Lineas_Tendencia` existe con gráfico de líneas con marcadores | Pestaña inferior del libro | ☐ |
| 3 | Gráfico circular incrustado en hoja `Datos_Ventas` sin tapar los datos | Hoja `Datos_Ventas` | ☐ |
| 4 | Gráfico de columnas tiene 5 series (4 regiones actuales + Norte año anterior) | Hoja `Columnas_Regiones` | ☐ |
| 5 | Gráfico de columnas muestra trimestres en eje X (no regiones) | Hoja `Columnas_Regiones` | ☐ |
| 6 | Título "Ventas Trimestrales por Región 2024" en negrita, tamaño 14 | Hoja `Columnas_Regiones` | ☐ |
| 7 | Etiquetas de datos sin decimales en gráfico de columnas | Hoja `Columnas_Regiones` | ☐ |
| 8 | Leyenda posicionada en la parte inferior del gráfico de columnas | Hoja `Columnas_Regiones` | ☐ |
| 9 | Eje Y comienza en 0 en el gráfico de columnas | Hoja `Columnas_Regiones` | ☐ |
| 10 | Solo líneas de cuadrícula horizontales principales visibles | Hoja `Columnas_Regiones` | ☐ |
| 11 | Estilo y esquema de colores corporativo aplicado al gráfico de columnas | Hoja `Columnas_Regiones` | ☐ |
| 12 | Gráfico circular muestra porcentajes como etiquetas | Hoja `Datos_Ventas` | ☐ |
| 13 | Sector del Producto A separado (explosión ~10%) | Hoja `Datos_Ventas` | ☐ |
| 14 | Texto alternativo descriptivo en los 3 gráficos | Todos los gráficos | ☐ |
| 15 | Archivo guardado en OneDrive sin errores | Barra de título / Archivo → Información | ☐ |

### Prueba de Integridad de Datos

Para verificar que los gráficos están correctamente vinculados a los datos:

1. Ve a la hoja **`Datos_Ventas`**.
2. Modifica temporalmente el valor de la celda **`B2`** (ventas Norte T1) cambiándolo a un número notablemente diferente (por ejemplo, multiplícalo por 2).
3. Ve a la hoja **`Columnas_Regiones`** y verifica que la barra correspondiente a "Norte" en el trimestre T1 cambió su altura.
4. Regresa a **`Datos_Ventas`** y presiona **Ctrl+Z** para deshacer el cambio.
5. Confirma que el gráfico regresó al valor original.

Este proceso confirma que los gráficos están dinámicamente vinculados a los datos fuente.

---

## Solución de Problemas

### Problema 1: El gráfico no se actualiza al cambiar los datos fuente

**Síntoma:** Después de modificar un valor en la hoja `Datos_Ventas`, el gráfico en la hoja `Columnas_Regiones` no refleja el cambio. Las barras siguen mostrando los valores anteriores.

**Causa probable:** El cálculo automático de Excel está desactivado, o el rango de datos del gráfico fue definido con valores estáticos (pegados como texto) en lugar de referencias de celda dinámicas. También puede ocurrir si el archivo se abrió en **Modo de compatibilidad** (visible en la barra de título como `[Modo de compatibilidad]`).

**Solución:**
1. Ve a **Archivo → Opciones → Fórmulas**.
2. En la sección **Opciones de cálculo**, verifica que **Cálculo del libro** esté configurado en **Automático**.
3. Haz clic en **Aceptar**.
4. Presiona **F9** para forzar el recálculo manual inmediato.
5. Si el problema persiste, verifica que el archivo no está en Modo de compatibilidad: ve a **Archivo → Información** y, si aparece el botón **Convertir**, haz clic en él para convertir el archivo al formato moderno `.xlsx`.
6. Si el rango del gráfico fue definido incorrectamente, abre **Diseño de gráfico → Seleccionar datos** y verifica que los valores de las series son referencias de celda (ejemplo: `=Datos_Ventas!$B$2:$B$5`) y no valores numéricos fijos.


---

## Limpieza del Entorno

Al finalizar la práctica, realiza los siguientes pasos para dejar el entorno en orden:

1. **Guardar el archivo final:** Presiona **Ctrl+G** una última vez para asegurarte de que todos los cambios están guardados en OneDrive.

2. **Cerrar paneles adicionales:** Si el Panel de Accesibilidad u otros paneles laterales siguen abiertos, ciérralos haciendo clic en la **X** de cada panel.

3. **No eliminar las hojas de gráfico:** Las hojas `Columnas_Regiones` y `Lineas_Tendencia` son parte del resultado final de la práctica y deben permanecer en el libro.

4. **Verificar el número de hojas:** El libro final debe contener exactamente las siguientes hojas:
   - `Datos_Ventas`
   - `Datos_Año_Anterior`
   - `Columnas_Regiones`
   - `Lineas_Tendencia`

5. **Cerrar Excel:** Ve a **Archivo → Cerrar** para cerrar el libro. Si Excel pregunta si deseas guardar, selecciona **Guardar**.

---

## Resumen

### Conceptos Clave Aplicados

En esta práctica aplicaste el flujo completo de creación y gestión de gráficos profesionales en Microsoft Excel 365:

| Habilidad Practicada | Herramienta/Función Utilizada | Paso |
|---|---|---|
| Crear gráfico de columnas agrupadas | Insertar → Gráficos → Columna agrupada | 2 |
| Mover gráfico a hoja dedicada | Diseño de gráfico → Mover gráfico → Hoja nueva | 2, 3 |
| Crear gráfico de líneas con marcadores | Insertar → Gráficos → Línea con marcadores | 3 |
| Crear gráfico circular | Insertar → Gráficos → Circular | 4 |
| Agregar nueva serie de datos | Diseño de gráfico → Seleccionar datos → Agregar | 5 |
| Modificar rango de serie existente | Diseño de gráfico → Seleccionar datos → Editar | 6 |
| Intercambiar filas y columnas | Diseño de gráfico → Cambiar fila/columna | 7 |
| Personalizar título, etiquetas, leyenda | Botón + (Elementos de gráfico) | 8 |
| Ajustar escala del eje Y | Doble clic en eje → Dar formato a eje | 8 |
| Aplicar estilos y esquemas de color | Diseño de gráfico → Estilos de gráfico / Cambiar colores | 9 |
| Separar sector en gráfico circular | Clic derecho → Dar formato a punto de datos → Explosión | 11 |


### Próximos Pasos

En la **Práctica 6**, integrarás **Microsoft Copilot en Excel** para automatizar tareas de análisis, generar fórmulas mediante lenguaje natural y obtener resúmenes inteligentes de tus datos. Recuerda que para esa práctica el archivo debe estar guardado en OneDrive (como ya lo tienes) y tu cuenta debe tener la licencia de Microsoft 365 Copilot activa.

---

### Recursos Adicionales

- [Microsoft Support: Crear un gráfico de principio a fin](https://support.microsoft.com/es-es/office/crear-un-gr%C3%A1fico-de-principio-a-fin-0baf399e-dd61-4e18-8a73-b3fd5d5680c2)
- [Microsoft Support: Agregar o quitar títulos en un gráfico](https://support.microsoft.com/es-es/office/agregar-o-quitar-t%C3%ADtulos-en-un-gr%C3%A1fico-d82d2d8a-3e62-4c3e-b5a2-7c8a0b7b3f4e)
- [Microsoft Support: Mover o cambiar el tamaño de un gráfico](https://support.microsoft.com/es-es/office/mover-o-cambiar-el-tama%C3%B1o-de-un-gr%C3%A1fico-5b9f6c5c-9de1-4f3c-8bb8-2b4f3c0bfb8d)
- [Microsoft Support: Tipos de gráficos disponibles en Office](https://support.microsoft.com/es-es/office/tipos-de-gr%C3%A1ficos-disponibles-en-office-a6187218-807e-4103-9e0a-27cdb19afb90)
- [Microsoft Support: Mejorar la accesibilidad con el Comprobador de accesibilidad](https://support.microsoft.com/es-es/office/mejorar-la-accesibilidad-con-el-comprobador-de-accesibilidad-a16f6de0-2f39-4a2b-8bd8-5ad801426c7f)
- [Chandoo.org: Excel Charts Guide – Learn to Create and Customize Charts](https://chandoo.org/wp/excel-charts/)

---
