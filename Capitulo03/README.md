# Gestión de Tablas y Datos Tabulares

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 20 minutos |
| **Complejidad** | Fácil |
| **Nivel de Bloom** | Aplicar |
| **Módulo / Capítulo** | Capítulo 3 — Gestión de Tablas y Datos Tabulares |
| **Práctica número** | 3 |

---

## 2. Descripción General

En esta práctica trabajarás con un conjunto de datos de inventario de productos que inicialmente se encuentra como un rango de celdas sin formato de tabla. Convertirás ese rango en una tabla estructurada de Excel, explorarás sus opciones de estilo y configuración, y practicarás las operaciones de filtrado y ordenación multinivel que son fundamentales en el análisis de datos profesional.

Al completar esta práctica habrás experimentado de forma directa la diferencia entre trabajar con rangos simples y con tablas formales de Excel, y comprenderás por qué la conversión a tabla es considerada una de las mejores prácticas en Excel intermedio.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica, serás capaz de:

- [ ] Convertir un rango de datos en una tabla estructurada de Excel utilizando la opción **Insertar → Tabla** y verificar la detección automática de encabezados.
- [ ] Aplicar y modificar estilos de tabla predefinidos, y configurar opciones de tabla como fila de encabezado, fila de totales y columnas con bandas.
- [ ] Insertar y configurar la fila de totales de una tabla utilizando funciones de resumen disponibles (SUMA, PROMEDIO, CONTEO, MÁX, MÍN).
- [ ] Filtrar datos de una tabla aplicando filtros por texto, por número y por color, y ordenar por múltiples criterios de columna con prioridades definidas.

---

## 4. Prerrequisitos

### Conocimiento previo

| Requisito | Descripción |
|---|---|
| Práctica 2 completada (o equivalente) | Conocimiento de manipulación de celdas y rangos en Excel |
| Selección de rangos | Saber seleccionar rangos contiguos y no contiguos con teclado y ratón |
| Encabezados de columna | Comprender qué es un encabezado y cómo se organiza un conjunto de datos tabular |
| Navegación básica por la cinta | Conocer las pestañas principales: Inicio, Insertar, Datos |

### Acceso y licencias

| Requisito | Detalle |
|---|---|
| Cuenta Microsoft 365 | Activa con acceso a Excel 365 (versión 2308 o superior recomendada) |
| Archivo de práctica | - Archivo [Lab03_Inventario_Productos (1).xlsx](./Lab03_Inventario_Productos(1).xlsx)|
| Acceso a OneDrive | Recomendado para guardar el archivo (no obligatorio en esta práctica) |

> **Nota para el instructor:** Esta práctica no requiere licencia de Copilot. Sin embargo, si algún estudiante tiene Excel en inglés, deberá usar los equivalentes en inglés para las funciones (SUM, AVERAGE, COUNT, MAX, MIN). Aclarar esto antes de iniciar.

---

## 5. Entorno de Laboratorio

### Hardware requerido

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 8ª gen / AMD Ryzen 5 (64 bits) | Intel Core i7 / AMD Ryzen 7 |
| Memoria RAM | 8 GB | 16 GB |
| Espacio en disco | 500 MB libres | 2 GB libres |
| Resolución de pantalla | 1366 × 768 px | 1920 × 1080 px |
| Dispositivo señalador | Ratón o trackpad | Ratón externo (recomendado para selección de rangos) |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | Microsoft 365 (v2308+) | Instalado en español |
| Sistema operativo | Windows 10 (21H2+) o Windows 11 | — |
| Navegador web | Edge, Chrome 110+ o Firefox 110+ | Para acceso a Microsoft 365 online si aplica |

### Preparación del entorno

Antes de comenzar los pasos de la práctica, realiza las siguientes acciones de configuración:

1. Abre **Microsoft Excel 365**.
2. Abre el archivo **`Lab03_Inventario_Productos.xlsx`** proporcionado por tu instructor.
3. Verifica que el archivo se abre en la hoja denominada **`Inventario`**.
4. Confirma que la hoja contiene datos en el rango **`A1:H51`** (1 fila de encabezados + 50 filas de datos).
5. Guarda una copia de seguridad del archivo antes de comenzar:
   - Presiona `F12` → guarda como **`Lab03_Inventario_MiNombre.xlsx`** en tu carpeta de trabajo.

> **Estructura esperada del archivo de práctica:**

| Columna | Encabezado | Tipo de dato |
|---|---|---|
| A | ID_Producto | Número entero |
| B | Nombre_Producto | Texto |
| C | Categoría | Texto |
| D | Precio_Unitario | Moneda (número decimal) |
| E | Stock_Actual | Número entero |
| F | Stock_Mínimo | Número entero |
| G | Proveedor | Texto |
| H | Estado_Stock | Texto (con color de celda aplicado) |

> **Nota sobre la columna H (Estado_Stock):** El archivo de práctica ya tiene aplicado formato condicional en esta columna con los siguientes colores: **verde** = "Suficiente", **amarillo** = "Revisar", **rojo** = "Crítico". Estos colores serán utilizados en el Paso 4 para practicar filtros por color.

---

## 6. Pasos del Laboratorio

---

### Paso 1 — Convertir el Rango en Tabla Estructurada

**Objetivo:** Convertir el rango de datos `A1:H51` en una tabla formal de Excel, verificar la detección automática de encabezados y asignar un nombre descriptivo a la tabla.

**Duración estimada:** 6 minutos

#### Instrucciones

1. Haz clic en la celda **`A1`** (encabezado "ID_Producto") para posicionarte dentro del rango de datos.

2. Ve a la pestaña **Insertar** en la cinta de opciones.

3. En el grupo **Tablas**, haz clic en el botón **Tabla**.

   > **Alternativa rápida:** Puedes usar el atajo de teclado `Ctrl + T` desde cualquier celda dentro del rango. El resultado es idéntico.

4. Se abre el cuadro de diálogo **Crear tabla**. Verifica los siguientes elementos:
   - **Campo de rango:** Debe mostrar `=$A$1:$H$51`. Si el rango detectado es diferente, corrígelo manualmente escribiendo `=$A$1:$H$51`.
   - **Casilla "La tabla tiene encabezados":** Debe estar **marcada** (activada). Si no lo está, actívala.

![Imagen práctica](../images/imagen%2073.png)

5. Haz clic en **Aceptar**.

6. Observa los cambios inmediatos en la hoja:
   - Aparecen **flechas de filtro** (▼) en cada celda de encabezado (fila 1).
   - Las filas de datos muestran **colores alternados** (bandas).
   - En la cinta de opciones aparece la pestaña contextual **Diseño de tabla** (visible solo cuando hay una celda de la tabla seleccionada).

7. **Renombra la tabla** (paso crítico):
   - Asegúrate de que alguna celda de la tabla esté seleccionada.
   - Ve a la pestaña **Diseño de tabla**.
   - En el extremo izquierdo de la cinta, localiza el campo **Nombre de la tabla** (mostrará `Tabla1` por defecto).
   - Haz clic dentro del campo, borra el texto existente y escribe: `TblInventario`
   - Presiona `Enter` para confirmar.

   ![Imagen práctica](../images/imagen%2074.png)

8. Verifica el nombre en el **Cuadro de nombres** (esquina superior izquierda de la hoja): al hacer clic en cualquier celda de la tabla, el cuadro debe mostrar la referencia de celda normalmente, pero si haces clic en la flecha del Cuadro de nombres, `TblInventario` debe aparecer en la lista de nombres definidos.

#### Resultado esperado

La hoja `Inventario` muestra el rango `A1:H51` convertido en una tabla con bandas de color, flechas de filtro en los encabezados y la pestaña **Diseño de tabla** visible en la cinta. El nombre `TblInventario` aparece en el campo Nombre de la tabla.

#### Verificación

- [ ] El rango `A1:H51` tiene formato de tabla (bandas de color visibles).
- [ ] Las flechas de filtro (▼) aparecen en todas las celdas de la fila 1.
- [ ] La pestaña **Diseño de tabla** aparece en la cinta al seleccionar cualquier celda de la tabla.
- [ ] El campo **Nombre de la tabla** muestra `TblInventario`.

---

### Paso 2 — Aplicar y Modificar Estilos de Tabla

**Objetivo:** Explorar el catálogo de estilos de tabla predefinidos, aplicar un estilo específico y modificar las opciones de estilo (bandas, primera columna, botones de filtro).

**Duración estimada:** 8 minutos

#### Instrucciones

**Parte A — Explorar y aplicar un estilo de tabla**

1. Haz clic en cualquier celda dentro de `TblInventario` para activar la pestaña **Diseño de tabla**.

2. Ve a la pestaña **Diseño de tabla** en la cinta.

3. En el grupo **Estilos de tabla**, observa la galería de estilos disponibles. Haz clic en el botón **Más** (el pequeño botón con flecha hacia abajo y línea horizontal, ubicado en la esquina inferior derecha de la galería) para expandir el catálogo completo.

4. El catálogo se organiza en tres secciones: **Claro**, **Medio** y **Oscuro**. Pasa el cursor sobre varios estilos para previsualizar cómo quedaría la tabla (sin hacer clic aún).

5. Aplica el estilo **"Estilo de tabla medio 16"** (azul, segunda fila de la sección Medio):
   - Ubica la sección **Medio** en el catálogo.
   - Haz clic en el estilo de color **azul** de la segunda fila (el nombre exacto aparece en el tooltip al pasar el cursor).
   
   > **Referencia visual:** Es el estilo con encabezado azul oscuro y bandas en gris claro/blanco alternadas.

![Imagen práctica](../images/imagen%2075.png)

6. Confirma que el estilo se ha aplicado observando el cambio de colores en la tabla.

**Parte B — Modificar opciones de estilo**

7. Con cualquier celda de la tabla seleccionada y en la pestaña **Diseño de tabla**, localiza el grupo **Opciones de estilo de tabla** (extremo izquierdo de la cinta, junto al campo de nombre).

8. Observa las casillas disponibles. El estado inicial típico es:
   - ☑ Fila de encabezado
   - ☐ Fila de totales
   - ☑ Filas con bandas
   - ☐ Primera columna
   - ☐ Última columna
   - ☑ Columnas con bandas (puede variar)
   - ☑ Botón de filtro

![Imagen práctica](../images/imagen%2076.png)

9. Realiza los siguientes cambios **uno por uno**, observando el efecto visual en la tabla después de cada cambio:

   **a)** Activa la casilla **"Primera columna"** → Observa cómo la columna A (ID_Producto) adquiere un formato resaltado (negrita).

   **b)** Desactiva la casilla **"Filas con bandas"** → Observa cómo desaparecen los colores alternados en las filas.

   **c)** Activa nuevamente la casilla **"Filas con bandas"** → Las bandas regresan.

   **d)** Desactiva la casilla **"Botón de filtro"** → Observa cómo desaparecen las flechas (▼) de los encabezados.

   **e)** Activa nuevamente la casilla **"Botón de filtro"** → Las flechas regresan.

   **f)** Desactiva la casilla **"Primera columna"** → La columna A vuelve a su formato normal.

10. Deja la configuración final de las opciones de estilo de la siguiente manera:
    - ☑ Fila de encabezado
    - ☐ Fila de totales (la activaremos en el Paso 3)
    - ☑ Filas con bandas
    - ☐ Primera columna
    - ☐ Última columna
    - ☐ Columnas con bandas
    - ☑ Botón de filtro

#### Resultado esperado

La tabla `TblInventario` muestra el estilo de tabla medio azul con bandas de filas activas, sin primera columna resaltada, y con los botones de filtro visibles en los encabezados.

![Imagen práctica](../images/imagen%2077.png)

#### Verificación

- [ ] El estilo aplicado muestra encabezados en azul oscuro con texto blanco.
- [ ] Las filas de datos alternan entre azul claro y blanco.
- [ ] Los botones de filtro (▼) son visibles en todos los encabezados.
- [ ] La primera columna NO tiene formato especial resaltado.

---

### Paso 3 — Insertar y Configurar la Fila de Totales

**Objetivo:** Activar la fila de totales de la tabla y configurar diferentes funciones de resumen para columnas específicas: CONTEO para ID_Producto, PROMEDIO para Precio_Unitario, SUMA para Stock_Actual, MÍN para Stock_Mínimo.

**Duración estimada:** 8 minutos

#### Instrucciones

**Parte A — Activar la fila de totales**

1. Haz clic en cualquier celda dentro de `TblInventario`.

2. Ve a la pestaña **Diseño de tabla**.

3. En el grupo **Opciones de estilo de tabla**, activa la casilla **"Fila de totales"**.

4. Observa que aparece una nueva fila al final de la tabla (fila 52) con la etiqueta **"Total"** en la columna A y un valor automático en la última columna (H).

   > **Nota:** Excel coloca automáticamente una función SUBTOTALES en la última columna. Este valor inicial puede no ser el que necesitamos; lo configuraremos en los siguientes pasos.

![Imagen práctica](../images/imagen%2078.png)

**Parte B — Configurar funciones de resumen por columna**

5. **Columna A — Columna ID_Producto (CONTEO):**
   - Haz clic en la celda de la fila de totales correspondiente a la columna **A** (celda `A52`).
   - Haz clic en la **flecha desplegable** (▼) que aparece en la celda.
   - Del menú desplegable, selecciona **Contar números** (equivale a la función CONTEO/COUNT).

![Imagen práctica](../images/imagen%2079.png)

   - Verifica que el resultado muestra **50** (el total de registros de producto).


6. **Columna D — Precio_Unitario (PROMEDIO):**
   - Haz clic en la celda de la fila de totales correspondiente a la columna **D** (celda `D52`).
   - Haz clic en la **flecha desplegable** (▼).
   - Selecciona **Promedio**.
   - Observa el valor calculado (precio promedio de todos los productos).

![Imagen práctica](../images/imagen%2080.png)

7. **Columna E — Stock_Actual (SUMA):**
   - Haz clic en la celda de la fila de totales correspondiente a la columna **E** (celda `E52`).
   - Haz clic en la **flecha desplegable** (▼).
   - Selecciona **Suma**.
   - Observa el valor calculado (total de unidades en stock).

   ![Imagen práctica](../images/imagen%2081.png)

8. **Columna F — Stock_Mínimo (MÍN):**
   - Haz clic en la celda de la fila de totales correspondiente a la columna **F** (celda `F52`).
   - Haz clic en la **flecha desplegable** (▼).
   - Selecciona **Mín**.
   - Observa el valor calculado (el stock mínimo más bajo de todos los productos).

![Imagen práctica](../images/imagen%2082.png)

9. **Columna H — Estado_Stock (CONTEO de texto):**
   - Haz clic en la celda de la fila de totales correspondiente a la columna **H** (celda `H52`).
   - Haz clic en la **flecha desplegable** (▼).
   - Selecciona **Recuento** (esta opción cuenta celdas no vacías, incluyendo texto).
   - Verifica que el resultado muestra **50**.

10. Haz clic en la celda `D52` (fila de totales, Precio_Unitario) y examina la barra de fórmulas. Observa que Excel usa la función `=SUBTOTALES(101,[Precio_Unitario])` en lugar de `=PROMEDIO(...)`. Esto es intencional: la función SUBTOTALES respeta los filtros activos, lo que significa que el promedio se recalculará automáticamente cuando apliques filtros en el Paso 4.

    > **Concepto clave — Referencias estructuradas:** Nota que la fórmula usa `'[Precio_Unitario]` en lugar de `D2:D51`. Esta es una **referencia estructurada**, una de las ventajas de trabajar con tablas formales. Hace que las fórmulas sean más legibles y se ajusten automáticamente cuando la tabla crece.

#### Resultado esperado

La fila 52 muestra la fila de totales con: CONTEO en columna A (valor: 50), PROMEDIO en columna D, SUMA en columna E, MÍN en columna F y CONTEO en columna H.

#### Verificación

- [ ] La fila de totales es visible en la fila 52 con la etiqueta "Total" en columna A.
- [ ] La celda `A52` muestra el valor **50** (CONTEO de registros).
- [ ] La celda `E52` muestra la SUMA del stock actual (un número mayor a 0).
- [ ] La barra de fórmulas de `D52` muestra una función `SUBTOTALES(...)` con referencia estructurada.

---

### Paso 4 — Filtrar Datos por Texto, Número y Color

**Objetivo:** Aplicar tres tipos de filtros diferentes sobre la tabla: filtro de texto por categoría, filtro numérico por umbral de stock y filtro por color de celda. Limpiar los filtros entre cada ejercicio.

**Duración estimada:** 10 minutos

#### Instrucciones

**Parte A — Filtro de texto por categoría**

1. Haz clic en la flecha de filtro (▼) del encabezado **Categoría** (columna C).

2. Se abre el panel de filtro. Observa que en la parte inferior aparece una lista con todas las categorías únicas presentes en los datos (por ejemplo: Electrónica, Cables, Papelería, Limpieza, etc.).

3. Haz clic en **Seleccionar todo** para desmarcar todas las categorías.

4. Marca únicamente la categoría **"Electrónica"**.

![Imagen práctica](../images/imagen%2083.png)

5. Haz clic en **Aceptar**.

6. Observa los resultados:
   - Solo se muestran las filas cuya categoría es "Electrónica".
   - Los números de fila (izquierda de la hoja) aparecen en **azul** y no son consecutivos, indicando que hay filas ocultas.
   - El ícono de la flecha en el encabezado Categoría cambia a un ícono de **embudo** (🔽 con filtro), indicando que hay un filtro activo.
   - La fila de totales se **recalcula automáticamente** mostrando solo los valores de los productos de Electrónica.

   ![Imagen práctica](../images/imagen%2084.png)

7. Anota mentalmente (o en papel) cuántos productos de la categoría Electrónica aparecen (observa el valor de CONTEO en `A52`).

8. **Limpia el filtro:** Haz clic nuevamente en la flecha/embudo del encabezado **Categoría** → selecciona **Borrar filtro de "Categoría"** → Haz clic en **Aceptar** (o simplemente haz clic en "Borrar filtro"). Todos los registros deben volver a ser visibles.

**Parte B — Filtro numérico por umbral de stock**

9. Haz clic en la flecha de filtro (▼) del encabezado **Stock_Actual** (columna E).

10. En el panel de filtro, pasa el cursor sobre la opción **Filtros de número** para expandir el submenú.

![Imagen práctica](../images/imagen%2085.png)

11. Selecciona **Menor que...** del submenú.

12. Se abre el cuadro de diálogo **Filtro personalizado**. En el campo junto a "es menor que", escribe el valor: `20`

![Imagen práctica](../images/imagen%2086.png)

13. Haz clic en **Aceptar**.

14. Observa los resultados:
    - Solo se muestran productos con Stock_Actual menor a 20 unidades (productos con posible riesgo de desabasto).
    - La fila de totales muestra la SUMA únicamente del stock filtrado.

15. **Limpia el filtro:** Haz clic en el embudo del encabezado **Stock_Actual** → selecciona **Borrar filtro de "Stock_Actual"** → Confirma. Todos los registros deben volver a ser visibles.

**Parte C — Filtro por color de celda**

> **Recordatorio:** La columna H (Estado_Stock) tiene celdas con colores de fondo aplicados mediante formato condicional: verde = "Suficiente", amarillo = "Revisar", rojo = "Crítico".

16. Haz clic en la flecha de filtro (▼) del encabezado **Estado_Stock** (columna H).

17. En el panel de filtro, pasa el cursor sobre la opción **Filtrar por color** para expandir el submenú.

18. En la sección **Filtrar por color de celda**, selecciona el color **rojo**.

![Imagen práctica](../images/imagen%2087.png)

19. Haz clic en **Aceptar** (o el filtro puede aplicarse directamente al hacer clic en el color).

20. Observa los resultados:
    - Solo se muestran los productos con Estado_Stock = "Crítico" (celdas rojas).
    - Estos son los productos que requieren atención inmediata de reabastecimiento.

![Imagen práctica](../images/imagen%2088.png)

21. **Limpia el filtro:** Haz clic en el embudo del encabezado **Estado_Stock** → selecciona **Borrar filtro de "Estado_Stock"**. Todos los registros deben volver a ser visibles.

22. Verifica que todos los 50 registros están visibles nuevamente (CONTEO en `A52` = 50).

#### Resultado esperado

Después de completar las tres partes y limpiar todos los filtros, la tabla muestra los 50 registros completos. Has comprobado que los tres tipos de filtro (texto, número, color) funcionan correctamente y que la fila de totales se recalcula dinámicamente con cada filtro aplicado.

#### Verificación

- [ ] El filtro de texto por "Electrónica" mostró únicamente productos de esa categoría.
- [ ] El filtro numérico "menor que 20" mostró solo productos con stock bajo.
- [ ] El filtro por color rojo mostró solo productos con estado "Crítico".
- [ ] La fila de totales se recalculó correctamente en cada filtro aplicado.
- [ ] Al limpiar todos los filtros, el CONTEO en `A52` muestra nuevamente **50**.

---

### Paso 5 — Ordenación Multinivel

**Objetivo:** Aplicar una ordenación con tres criterios de prioridad: primero por Categoría (A-Z), luego por Precio_Unitario (mayor a menor) y finalmente por Nombre_Producto (A-Z). Comprender el orden de prioridad de los criterios.

**Duración estimada:** 4 minutos

#### Instrucciones

1. Haz clic en cualquier celda dentro de `TblInventario` para asegurarte de que la tabla está activa.

2. Ve a la pestaña **Datos** en la cinta de opciones.

3. En el grupo **Ordenar y filtrar**, haz clic en el botón **Ordenar** (el ícono con flechas A↕Z y líneas horizontales). Se abre el cuadro de diálogo **Ordenar**.

![Imagen práctica](../images/imagen%2089.png)

   > **Nota:** No uses las flechas de filtro de los encabezados para esta tarea, ya que solo permiten ordenar por un criterio a la vez. El cuadro de diálogo **Ordenar** permite configurar múltiples niveles.

4. **Configura el Nivel 1 (criterio principal — Categoría):**
   - En la fila que ya aparece (o haz clic en **Agregar nivel** si el cuadro está vacío):
   - **Columna:** Selecciona `Categoría` del desplegable "Ordenar según".
   - **Ordenar según:** `Valores de celda`.
   - **Criterio:** `A a Z`.

![Imagen práctica](../images/imagen%2090.png)

5. **Agrega el Nivel 2 (criterio secundario — Precio_Unitario):**
   - Haz clic en el botón **Agregar nivel** (parte superior izquierda del cuadro de diálogo).
   - Aparece una nueva fila "Luego por".
   - **Columna:** Selecciona `Precio_Unitario`.
   - **Ordenar según:** `Valores de celda`.
   - **Criterio:** `De mayor a menor`.

   ![Imagen práctica](../images/imagen%2091.png)

6. **Agrega el Nivel 3 (criterio terciario — Nombre_Producto):**
   - Haz clic nuevamente en **Agregar nivel**.
   - Aparece otra fila "Luego por".
   - **Columna:** Selecciona `Nombre_Producto`.
   - **Ordenar según:** `Valores de celda`.
   - **Criterio:** `A a Z`.

   ![Imagen práctica](../images/imagen%2092.png)

7. Verifica que el cuadro de diálogo muestre los tres niveles en el orden correcto:
   - **Nivel 1:** Categoría — A a Z
   - **Nivel 2:** Precio_Unitario — De mayor a menor
   - **Nivel 3:** Nombre_Producto — A a Z

8. Haz clic en **Aceptar**.

9. Examina los resultados en la tabla:
   - Los datos deben estar agrupados visualmente por categoría (todas las filas de "Electrónica" juntas, todas las de "Almacenamiento" juntas, etc.).
   - Dentro de cada categoría, los productos están ordenados de precio más alto a precio más bajo.
   - Si hay productos con la misma categoría y el mismo precio, estarán ordenados alfabéticamente por nombre.

10. **Comprende la lógica de prioridad:** El criterio de Nivel 1 (Categoría) tiene la mayor prioridad. El Nivel 2 (Precio) solo se aplica para desempatar dentro de la misma categoría. El Nivel 3 (Nombre) solo actúa cuando dos productos de la misma categoría tienen exactamente el mismo precio.

#### Resultado esperado

La tabla muestra los 50 productos ordenados jerárquicamente: agrupados por categoría en orden alfabético, y dentro de cada categoría ordenados de precio mayor a menor. Los botones de filtro y la fila de totales permanecen intactos.

 ![Imagen práctica](../images/imagen%2093.png)

#### Verificación

- [ ] Los productos están visiblemente agrupados por categoría (todas las filas de una misma categoría aparecen juntas).
- [ ] Dentro de cada categoría, el primer producto listado tiene el precio más alto y el último tiene el precio más bajo.
- [ ] La fila de totales sigue mostrando los valores de resumen correctos (CONTEO = 50).
- [ ] Los botones de filtro (▼) siguen visibles en todos los encabezados.

---

## 7. Validación y Pruebas Finales

Una vez completados todos los pasos, realiza las siguientes verificaciones integrales para confirmar que la práctica se completó correctamente:

### Lista de verificación final

| # | Verificación | Cómo comprobarla | Resultado esperado |
|---|---|---|---|
| 1 | La tabla existe con el nombre correcto | Pestaña **Diseño de tabla** → campo Nombre de la tabla | Muestra `TblInventario` |
| 2 | El rango de la tabla es correcto | Selecciona toda la tabla con `Ctrl + A` | Selecciona `A1:H52` (incluyendo fila de totales) |
| 3 | El estilo de tabla es el correcto | Observar colores de la tabla | Encabezados azul oscuro con texto blanco, bandas azul claro/blanco |
| 4 | La fila de totales está activa | Observar fila 52 | Muestra "Total" en A52 y valores de resumen en D52, E52, F52, H52 |
| 5 | No hay filtros activos | Observar los íconos de los encabezados | Todos muestran flechas (▼) sin embudo de filtro activo |
| 6 | Los datos están ordenados correctamente | Revisar las primeras y últimas filas de cada categoría | Agrupados por categoría A-Z, luego por precio mayor a menor |
| 7 | La fila de totales muestra 50 registros | Observar celda `A52` | Valor = 50 |

### Prueba de expansión dinámica (opcional — 2 minutos adicionales)

Esta prueba demuestra una de las ventajas más importantes de las tablas:

1. Haz clic en la celda **`A53`** (la fila inmediatamente debajo de la fila de totales).

   > **Nota:** Si la fila de totales está activa, deberás ir a la celda inmediatamente después. Excel puede pedirte que confirmes si deseas agregar datos a la tabla.

2. Escribe el valor `51` y presiona `Tab`.
3. Escribe `Producto de Prueba` y presiona `Tab`.
4. Escribe `Electrónica` y presiona `Tab`.
5. Escribe `99.99` y presiona `Tab`.
6. Escribe `15` y presiona `Tab`.
7. Escribe `5` y presiona `Enter`.

8. Observa que Excel **expandió automáticamente** la tabla para incluir la nueva fila, aplicando el estilo de bandas y manteniendo las referencias estructuradas.

9. Verifica que el CONTEO en la fila de totales ahora muestra **51**.

10. **Elimina la fila de prueba:** Haz clic derecho sobre el número de la fila recién agregada → selecciona **Eliminar filas de la tabla**. El CONTEO debe volver a **50**.

---

## 8. Resolución de Problemas

### Problema 1: El cuadro de diálogo "Crear tabla" detecta un rango incorrecto

**Síntoma:** Al presionar `Ctrl + T` o al usar **Insertar → Tabla**, el campo de rango en el cuadro de diálogo muestra un rango diferente a `=$A$1:$H$51` (por ejemplo, solo `=$A$1:$A$1` o un rango parcial).

**Causa probable:** Hay celdas vacías dentro del rango de datos (filas o columnas completamente vacías) que interrumpen la detección automática del límite del rango. También puede ocurrir si la celda activa estaba fuera del rango de datos al momento de ejecutar el comando.

**Solución:**
1. Haz clic en **Cancelar** para cerrar el cuadro de diálogo sin crear la tabla.
2. Selecciona manualmente el rango completo `A1:H51` (haz clic en `A1`, mantén `Shift` y haz clic en `H51`).
3. Revisa visualmente si hay filas o columnas vacías dentro del rango seleccionado y elimínalas si las hay.
4. Con el rango `A1:H51` seleccionado, presiona `Ctrl + T` nuevamente.
5. En el cuadro de diálogo, el campo de rango ahora debe mostrar `=$A$1:$H$51`. Verifica que la casilla "La tabla tiene encabezados" esté marcada y haz clic en **Aceptar**.

---

### Problema 2: La opción "Filtrar por color" no aparece o aparece deshabilitada en el menú de filtro

**Síntoma:** Al hacer clic en la flecha de filtro (▼) del encabezado Estado_Stock y pasar el cursor sobre **Filtrar por color**, el submenú aparece vacío (sin colores para seleccionar) o la opción está en gris y no es seleccionable.

**Causa probable:** Las celdas de la columna H no tienen colores de fondo aplicados, o el formato condicional no se cargó correctamente desde el archivo de práctica. Esto puede ocurrir si el archivo fue guardado en un formato diferente (como `.csv`) que no preserva el formato, o si el archivo de práctica fue modificado accidentalmente antes de la práctica.

**Solución:**
1. Verifica que las celdas de la columna H (Estado_Stock) tienen colores visibles. Si todas las celdas aparecen con fondo blanco, el formato condicional no está presente.
2. Solicita al instructor el archivo de práctica original (`Lab03_Inventario_Productos.xlsx`) y ábrelo nuevamente.
3. Si necesitas aplicar los colores manualmente como alternativa temporal:
   - Selecciona todas las celdas de la columna H que contengan el valor "Crítico".
   - Ve a **Inicio → Color de relleno** → selecciona rojo.
   - Repite para "Revisar" (amarillo) y "Suficiente" (verde).
4. Una vez que las celdas tienen color de fondo, regresa al filtro de la columna H → **Filtrar por color** → el submenú ahora mostrará los colores disponibles.

---

## 9. Limpieza del Entorno

Al finalizar la práctica, realiza los siguientes pasos para dejar el entorno en orden:

1. **Verifica que no haya filtros activos** en la tabla: Si algún encabezado muestra el ícono de embudo (filtro activo), haz clic en ese encabezado → **Borrar filtro**. Alternativamente, ve a **Datos → Borrar** (en el grupo Ordenar y filtrar) para eliminar todos los filtros de una vez.

2. **Guarda el archivo de trabajo** con los cambios realizados:
   - Presiona `Ctrl + S`.
   - Si el archivo está en formato `.xlsx`, confirma guardar en el mismo formato.

3. **Guarda una copia final con nombre descriptivo** (opcional pero recomendado):
   - Presiona `F12`.
   - Guarda como `Lab03_Inventario_Completado_[TuNombre].xlsx` en tu carpeta de entregas.

4. **No elimines la tabla ni reviertas los cambios:** El archivo con la tabla `TblInventario` configurada será el punto de partida para prácticas posteriores del módulo.

5. Si abriste otros libros de Excel durante la práctica, ciérralos para liberar memoria: **Archivo → Cerrar** en cada libro adicional.

---

## 10. Resumen

### Lo que aprendiste en esta práctica

En esta práctica aplicaste cuatro operaciones fundamentales de gestión de tablas en Excel 365:

| Operación | Método utilizado | Beneficio clave |
|---|---|---|
| **Crear tabla desde rango** | `Ctrl + T` / Insertar → Tabla | Activa filtros, expansión dinámica y referencias estructuradas automáticamente |
| **Aplicar y modificar estilos** | Pestaña Diseño de tabla → Galería de estilos | Comunicación visual profesional sin formato manual |
| **Fila de totales con funciones** | Opciones de estilo → Fila de totales | Resúmenes dinámicos que respetan los filtros activos (función SUBTOTALES) |
| **Filtros y ordenación multinivel** | Flechas de filtro / Datos → Ordenar | Análisis rápido de subconjuntos de datos con múltiples criterios |

### Conceptos clave para recordar

- **Renombrar la tabla** inmediatamente después de crearla (`TblInventario`) es una práctica esencial para el trabajo con fórmulas y referencias en hojas complejas.
- **La fila de totales usa SUBTOTALES**, no SUMA directa, lo que permite que los valores se recalculen automáticamente cuando se aplican filtros.
- **Las referencias estructuradas** (`TblInventario[Precio_Unitario]`) son más legibles y robustas que las referencias de celda tradicionales (`D2:D51`).
- **El orden de los niveles en la ordenación multinivel importa:** el Nivel 1 tiene la mayor prioridad; los niveles siguientes solo actúan como desempate.
- **Los filtros por color** solo funcionan cuando las celdas tienen colores de fondo aplicados (ya sea manualmente o mediante formato condicional).

### Conexión con el resto del curso

Las tablas estructuradas que creaste en esta práctica son la base para:
- **Capítulo 4 (Tablas Dinámicas):** Las tablas de Excel son la fuente de datos ideal para tablas dinámicas.
- **Capítulo 5 (Gráficos):** Los gráficos basados en tablas se actualizan automáticamente cuando se agregan nuevos datos.
- **Capítulo 6 (Copilot en Excel):** Copilot trabaja de forma más efectiva con datos organizados en tablas estructuradas con nombres descriptivos.

### Recursos adicionales

- [Documentación oficial de Microsoft: Crear una tabla en Excel](https://support.microsoft.com/es-es/office/crear-una-tabla-en-excel-e81aa349-b006-4f8a-9806-5af9df0ac664)
- [Microsoft Support: Usar referencias estructuradas con tablas de Excel](https://support.microsoft.com/es-es/office/usar-referencias-estructuradas-con-tablas-de-excel-f5ed2452-2337-4f71-bed3-c8ae6d2b276e)
- [Microsoft Support: Filtrar datos en un rango o tabla](https://support.microsoft.com/es-es/office/filtrar-datos-en-un-rango-o-tabla-01832226-31b5-4568-8806-38c37dcc180e)
- [Microsoft Support: Ordenar datos en una tabla o rango](https://support.microsoft.com/es-es/office/ordenar-datos-en-un-rango-o-tabla-62d0b95d-2a90-4610-a6ae-2e545c4a4654)

---
