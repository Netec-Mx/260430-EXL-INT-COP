# Manipulación de celdas y rangos de datos

## Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 25 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (*Apply*) |
| **Módulo** | 2 — Manipulación de celdas y rangos de datos |
| **Versión de Excel requerida** | Microsoft 365 (versión 2308 o superior) |
| **Archivo de práctica** | [Lab02-VentasRegionales (1).xlsx](./Lab02-VentasRegionales%20(1).xlsx)

---

## Descripción General

En esta práctica trabajarás con un libro de Excel que contiene datos de ventas mensuales por región. Aplicarás pegado especial para transferir únicamente valores, formatos o transponer tablas; usarás autorrelleno y relleno rápido para completar series de fechas, números y texto; generarás datos de prueba con las funciones `ALEATORIO.ENTRE()` y `SECUENCIA()`; y finalizarás configurando estilos, rangos nombrados, minigráficos y reglas de formato condicional. Al terminar, habrás construido una hoja de análisis de ventas visualmente efectiva y con datos estructurados de manera profesional.

> **Nota sobre idioma de Excel:** Todas las fórmulas en esta práctica están escritas en español (`SUMA`, `ALEATORIO.ENTRE`, `SECUENCIA`, etc.), que es el idioma predeterminado cuando Excel está instalado en español. Si tu instalación está en inglés, utiliza los equivalentes: `SUM`, `RANDBETWEEN`, `SEQUENCE`. Consulta a tu instructor si tienes dudas.

> **Nota sobre `ALEATORIO.ENTRE()`:** Esta función genera valores diferentes cada vez que la hoja se recalcula (al presionar **F9** o guardar). Tus resultados numéricos serán distintos a los de tus compañeros. Esto es normal y esperado; lo importante es aplicar correctamente la función.

---

## Objetivos de Aprendizaje

Al finalizar esta práctica, serás capaz de:

- [ ] Aplicar opciones avanzadas de pegado especial (valores, formatos, transponer) para reorganizar datos sin alterar la fuente original.
- [ ] Utilizar autorrelleno y relleno rápido (`Ctrl+E`) para completar series de fechas, secuencias numéricas y patrones de texto de forma eficiente.
- [ ] Generar tablas de datos de prueba usando `ALEATORIO.ENTRE()` y `SECUENCIA()`, e insertar o eliminar filas y columnas desplazando datos correctamente.
- [ ] Definir rangos nombrados y referenciarlos en fórmulas para mejorar la legibilidad de la hoja.
- [ ] Insertar minigráficos de línea y aplicar reglas de formato condicional con escalas de color y barras de datos para comunicar tendencias visualmente.

---

## Prerrequisitos

### Conocimientos previos
- Haber completado la Práctica 1 o tener conocimiento equivalente de gestión básica de libros y hojas de Excel.
- Saber seleccionar rangos de celdas, ingresar datos y aplicar formato básico de números y texto.
- Comprender qué es una fórmula en Excel y cómo ingresarla en una celda (uso del signo `=`).

### Acceso y recursos
- Cuenta de Microsoft 365 activa con acceso a Excel en escritorio (no versión web para esta práctica).
- Archivo [Lab02-VentasRegionales (1).xlsx](./Lab02-VentasRegionales%20(1).xlsx) 
- Conexión a Internet estable (mínimo 10 Mbps) para sincronización con OneDrive.

---

## Entorno de Laboratorio

### Requisitos de hardware

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 8ª gen / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| Memoria RAM | 8 GB | 16 GB |
| Espacio en disco | 10 GB disponibles | 20 GB disponibles |
| Resolución de pantalla | 1366 × 768 px | 1920 × 1080 px |
| Dispositivo señalador | Ratón externo recomendado | Ratón externo |

### Requisitos de software

| Software | Versión requerida |
|---|---|
| Sistema operativo | Windows 10 (21H2+) o Windows 11 |
| Microsoft Excel | Microsoft 365, versión 2308 o superior |
| Navegador web | Microsoft Edge, Chrome 110+ o Firefox 110+ |
| OneDrive | Integrado con Microsoft 365 (sincronizado) |

### Configuración inicial del entorno

Antes de comenzar los pasos de la práctica, realiza estas verificaciones:

1. **Abre OneDrive** y navega a `Curso Excel Intermedio > Módulo 2`. Confirma que el archivo `Lab02-VentasRegionales.xlsx` está disponible.
2. **Descarga y abre el archivo** haciendo doble clic sobre él. Excel debe abrirse en la aplicación de escritorio (no en el navegador). Si se abre en el navegador, haz clic en **Abrir en la aplicación de escritorio**.
3. **Verifica la versión de Excel:** Ve a `Archivo → Cuenta → Acerca de Excel`. Confirma que la versión es 2308 o superior.
4. **Habilita el guardado automático:** En la barra de título, activa el interruptor **Autoguardado** (debe aparecer en color verde/activo). Esto garantiza que el archivo se guarda en OneDrive continuamente.
5. **Confirma el idioma de las fórmulas:** En cualquier celda vacía, escribe `=SUMA(1,2)` y presiona **Enter**. Si obtienes `3`, Excel está en español. Elimina el contenido de esa celda antes de continuar.

---

## Pasos del Laboratorio

---

### Paso 1 — Explorar el archivo y aplicar Pegado Especial de Valores

**Objetivo:** Familiarizarse con la estructura del libro y practicar el pegado de solo valores para eliminar dependencias de fórmulas en una copia de los datos.

**Duración estimada:** 8 minutos

#### Instrucciones

1. Con el archivo [Lab02-VentasRegionales (1).xlsx](./Lab02-VentasRegionales%20(1).xlsx)  abierto, observa las hojas disponibles en la parte inferior. Deberías ver al menos las hojas: **`Ventas_Q1`**, **`Ventas_Q2`** y **`Resumen`**.

![Imagen práctica](../images/imagen%2037.png)

2. Haz clic en la hoja **`Ventas_Q1`**. Observa su estructura:
   - **Columna A:** Nombres de regiones (Norte, Sur, Este, Oeste, Centro).
   - **Columnas B a D:** Ventas de los meses Enero, Febrero y Marzo (con fórmulas que calculan los totales).
   - **Columna E:** Total trimestral calculado con `=SUMA(B2:D2)`.

3. Selecciona el rango **`B3:E7`** (ventas y totales de las 5 regiones).

4. Copia el rango con **Ctrl+C**. Verás el borde parpadeante alrededor del rango copiado.

![Imagen práctica](../images/imagen%2038.png)

5. Haz clic en la celda **`B10`** para establecer el destino del pegado.

6. Abre el cuadro de diálogo de Pegado Especial presionando **Ctrl+Alt+V**.

7. En el cuadro de diálogo, selecciona la opción **Valores** y haz clic en **Aceptar**.

![Imagen práctica](../images/imagen%2039.png)

8. Haz clic en cualquier celda del rango **`B10:E14`** y observa la barra de fórmulas. Confirma que muestra solo un número (no una fórmula).



9. Compara con una celda del rango original (por ejemplo, `E3`): la barra de fórmulas debe mostrar `=SUMA(B2:D2)`. La copia en `E10` debe mostrar solo el valor numérico equivalente.

![Imagen práctica](../images/imagen%2040.png)

10. Presiona **Escape** para limpiar el portapapeles (el borde parpadeante desaparecerá).

> **¿Por qué es importante esto?** Al compartir reportes con personas externas, pegar solo valores protege la lógica interna de tus cálculos y evita que fórmulas se rompan al moverlas fuera de su contexto original.

#### Resultado esperado

El rango `B10:E14` contiene los mismos números que `B3:E7`, pero sin fórmulas. Al hacer clic en cualquier celda de `B10:E14`, la barra de fórmulas muestra únicamente el valor numérico.

#### Verificación

- Haz clic en la celda `E10`. La barra de fórmulas debe mostrar un número (por ejemplo, `145700`), **no** una fórmula.
- Haz clic en la celda `E3`. La barra de fórmulas debe mostrar `=SUMA(B3:D3)`.
- Ambas celdas deben mostrar el mismo número en la hoja.

---

### Paso 2 — Aplicar Pegado Especial de Formatos y Transponer datos

**Objetivo:** Usar pegado especial para copiar únicamente el formato visual de un rango y transponer una tabla de filas a columnas.

**Duración estimada:** 7 minutos

#### Instrucciones

**Parte A — Pegado de solo formatos:**

1. En la hoja **`Ventas_Q1`**, selecciona el rango **`A2:E2`** (fila de encabezados con formato de color, negrita y bordes).

2. Copia con **Ctrl+C**.

3. Haz clic en la celda **`A9`** (que actualmente está vacía y servirá como encabezado de la tabla de valores pegados en el Paso 1).

4. Abre Pegado Especial con **Ctrl+Alt+V**.

5. Selecciona **Formatos** y haz clic en **Aceptar**.

![Imagen práctica](../images/imagen%2041.png)

6. Escribe manualmente los encabezados en las celdas `A9` a `E9`:
   - `A9`: `Región`
   - `B9`: `Enero`
   - `C9`: `Febrero`
   - `D9`: `Marzo`
   - `E9`: `TotalQ1`

7. Observa que las celdas `A9:E9` ahora tienen el mismo formato visual (color de fondo, negrita, bordes) que la fila 1, pero con el texto que tú escribiste.

![Imagen práctica](../images/imagen%2042.png)

8. Presiona **Escape** para limpiar el portapapeles.

**Parte B — Transponer una tabla:**

1. Selecciona el rango **`A2:E7`** (tabla completa con encabezados y datos).

2. Copia con **Ctrl+C**.

3. Haz clic en la celda **`A18`** (deja espacio suficiente debajo de los datos anteriores).

4. Abre Pegado Especial con **Ctrl+Alt+V**.

5. En el cuadro de diálogo, marca la casilla **Transponer** (en la esquina inferior derecha del cuadro) y selecciona **Valores** en la sección "Pegar". Haz clic en **Aceptar**.


6. Observa el resultado: lo que antes eran 5 filas de regiones ahora son 5 columnas, y lo que eran 4 columnas (Enero, Feb, Mar, Total) ahora son 4 filas.

![Imagen práctica](../images/imagen%2043.png)

7. Presiona **Escape**.

#### Resultado esperado

- El rango `A9:E9` tiene el mismo formato visual que `A1:E1` con los encabezados escritos manualmente.
- A partir de la celda `A18`, la tabla original aparece transpuesta: los meses están en filas y las regiones en columnas.

#### Verificación

- Haz clic en `B9`. La barra de fórmulas debe mostrar el texto `Enero` (no una fórmula) y la celda debe tener el mismo formato que `B1`.
- En la tabla transpuesta, la celda `A19` debe contener `Enero` (el primer mes, que ahora es una fila).
- La celda `B18` debe contener el nombre de la primera región (por ejemplo, `Norte`).

---

### Paso 3 — Autorrelleno de series de fechas y texto

**Objetivo:** Usar el controlador de relleno para completar automáticamente series de fechas, números y texto con patrones personalizados.

**Duración estimada:** 7 minutos

#### Instrucciones

1. Haz clic en la hoja **`Resumen`** en la parte inferior de la pantalla.

2. **Serie de fechas semanales:**
   - En la celda **`A2`**, escribe la fecha: `03/01/2025` y presiona **Enter**.
   - En la celda **`A3`**, escribe: `10/01/2025` (exactamente 7 días después) y presiona **Enter**.
   - Selecciona el rango **`A2:A3`**.
   - Posiciona el cursor sobre el **controlador de relleno** (pequeño cuadrado verde en la esquina inferior derecha del rango seleccionado). El cursor cambiará a una cruz negra delgada (`+`).
   - Arrastra el controlador de relleno hacia abajo hasta la celda **`A13`** (12 semanas en total).
   - Verifica que Excel completó la serie con incrementos de 7 días (fechas de los lunes).

   ![Imagen práctica](../images/imagen%2044.png)

3. **Serie de identificadores de texto:**
   - En la celda **`B2`**, escribe: `Semana-01` y presiona **Enter**.
   - En la celda **`B3`**, escribe: `Semana-02` y presiona **Enter**.
   - Selecciona **`B2:B3`**.
   - Arrastra el controlador de relleno hacia abajo hasta **`B13`**.
   - Excel debe completar la serie: `Semana-03`, `Semana-04`… hasta `Semana-12`.

   ![Imagen práctica](../images/imagen%2045.png)

4. **Serie numérica con incremento personalizado:**
   - En la celda **`C2`**, escribe: `100`.
   - En la celda **`C3`**, escribe: `115`.
   - Selecciona **`C2:C3`**.
   - Arrastra el controlador de relleno hacia abajo hasta **`C13`**.
   - Excel debe completar la serie: `130`, `145`, `160`… (incremento de 15).

   ![Imagen práctica](../images/imagen%2046.png)

5. **Autorrelleno de meses:**
   - En la celda **`D2`**, escribe: `Enero`.
   - Selecciona solo **`D2`** (una sola celda).
   - Arrastra el controlador de relleno hacia abajo hasta **`D13`**.
   - Excel debe completar automáticamente: `Febrero`, `Marzo`… hasta `Diciembre` 

6. Escribe el encabezado en la fila 1 para cada columna:
   - `A1`: `Fecha`
   - `B1`: `Semana`
   - `C1`: `Meta`
   - `D1`: `Mes`

 ![Imagen práctica](../images/imagen%2047.png)

#### Resultado esperado

Las columnas A a D contienen 12 filas de datos (filas 2 a 13) con series correctamente completadas: fechas semanales, identificadores `Semana-01` a `Semana-12`, valores numéricos de 100 a 265 (incremento de 15) y los 12 meses del año.

#### Verificación

- La celda `A13` debe contener la fecha `19/03/2025` (12ª semana desde el 03/01/2025).
- La celda `B13` debe contener `Semana-12`.
- La celda `C13` debe contener `265`.
- La celda `D13` debe contener `Diciembre`.

---

### Paso 4 — Generar datos de prueba con `ALEATORIO.ENTRE()` y `SECUENCIA()`

**Objetivo:** Crear una tabla de datos de ventas simuladas usando funciones dinámicas de Excel 365 para poblar rangos automáticamente.

**Duración estimada:** 8 minutos

> **Nota:** `SECUENCIA()` y `ALEATORIO.ENTRE()` son funciones de Microsoft 365. Si tu versión de Excel no las reconoce, notifica a tu instructor.

#### Instrucciones

1. En la hoja **`Resumen`**, desplázate hacia la derecha y haz clic en la celda **`F1`**.

2. **Crear encabezados de la tabla de datos de prueba:**
   - `F1`: `ID_Venta`
   - `G1`: `Región`
   - `H1`: `Ventas_Ene`
   - `I1`: `Ventas_Feb`
   - `J1`: `Ventas_Mar`

3. **Generar IDs automáticos con `SECUENCIA()`:**
   - Haz clic en la celda **`F2`**.
   - Escribe la siguiente fórmula y presiona **Enter**:
     ```
     =SECUENCIA(10,1,1001,1)
     ```
   - Esta fórmula genera 10 números en 1 columna, comenzando en 1001 con incremento de 1 (1001, 1002… 1010).
   - Observa que Excel rellena automáticamente las celdas `F2:F11` con el resultado (función dinámica de desbordamiento).

![Imagen práctica](../images/imagen%2048.png)

4. **Ingresar nombres de región manualmente:**
   - En las celdas `G2:G11`, escribe los siguientes valores (2 por región):
     - `G2` y `G3`: `Norte`
     - `G4` y `G5`: `Sur`
     - `G6` y `G7`: `Este`
     - `G8` y `G9`: `Oeste`
     - `G10` y `G11`: `Centro`

5. **Generar ventas simuladas de Enero con `ALEATORIO.ENTRE()`:**
   - Haz clic en la celda **`H2`**.
   - Escribe la siguiente fórmula y presiona **Enter**:
     ```
     =ALEATORIO.ENTRE(20000,150000)
     ```
   - Copia la celda `H2` (Ctrl+C).
   - Selecciona el rango **`H2:J11`** (ventas de los 3 meses para los 10 registros).
   - Pega con **Ctrl+V**.
   - Cada celda del rango tendrá un valor aleatorio entre 20,000 y 150,000.

   ![Imagen práctica](../images/imagen%2049.png)

6. **Observar el comportamiento dinámico:**
   - Presiona **F9** (recalcular). Observa cómo todos los valores de `ALEATORIO.ENTRE()` cambian.
   - Esto es normal. Tus valores serán diferentes a los de tus compañeros.

7. **Convertir los valores aleatorios a valores fijos** (para evitar que cambien al guardar):
   - Selecciona el rango **`H2:J11`**.
   - Copia con **Ctrl+C**.
   - Con el rango aún seleccionado, abre Pegado Especial con **Ctrl+Alt+V**.
   - Selecciona **Valores** y haz clic en **Aceptar**.
   - Ahora los valores son fijos. Presiona **F9** para confirmar que ya no cambian.

![Imagen práctica](../images/imagen%2050.png)

8. **Aplicar formato de número a las ventas:**
   - Selecciona el rango **`H2:J11`**.
   - Ve a la pestaña **Inicio** → grupo **Número** → haz clic en el botón **Formato de número de contabilidad** (símbolo `$`) o selecciona **Número** con separador de miles desde el menú desplegable de formatos.

   ![Imagen práctica](../images/imagen%2051.png)

#### Resultado esperado

La tabla en `F1:J11` contiene: IDs del 1001 al 1010 generados por `SECUENCIA()`, nombres de región en pares, y valores de ventas fijos (ya no aleatorios) con formato numérico con separador de miles en las columnas H, I y J.

#### Verificación

- Haz clic en la celda `F2`. La barra de fórmulas debe mostrar `=SECUENCIA(10,1,1001,1)`.
- Haz clic en la celda `H2`. La barra de fórmulas debe mostrar un número fijo (no la fórmula `=ALEATORIO.ENTRE(...)`).
- Presiona **F9**: los valores en `H2:J11` **no deben cambiar** (ya son valores fijos).
- El rango `F2:F11` debe mostrar los números del 1001 al 1010.

---

### Paso 5 — Insertar y eliminar filas y columnas

**Objetivo:** Practicar la inserción y eliminación de filas y columnas desplazando datos correctamente sin romper la estructura de la hoja.

**Duración estimada:** 5 minutos

#### Instrucciones

**Insertar una fila:**

1. En la hoja **`Resumen`**, haz clic en el **número de fila `6`** (encabezado de fila, en el margen izquierdo) para seleccionar la fila completa.

2. Haz clic derecho sobre la selección y elige **Insertar** del menú contextual. Se insertará una fila en blanco encima de la fila 6, y todos los datos debajo se desplazarán una fila hacia abajo.

![Imagen práctica](../images/imagen%2052.png)

3. En la nueva fila 6 (que quedó en blanco), escribe en la celda `G6`: `Este` (para completar el par de la región Este que quedó separada).

![Imagen práctica](../images/imagen%2053.png)

   > Nota: Si los datos de región quedaron desalineados por la inserción, ajusta los valores en la columna G para que el par de cada región esté en filas consecutivas.

**Insertar una columna:**

4. Haz clic en la **letra de columna `G`** para seleccionar toda la columna G.

5. Haz clic derecho y elige **Insertar**. Se insertará una columna en blanco a la izquierda de G, y las columnas de datos se desplazarán a la derecha.

6. En la celda **`G1`** (nueva columna vacía), escribe: `Vendedor`.

![Imagen práctica](../images/imagen%2054.png)

7. En las celdas `G2:G12`, escribe nombres de vendedor (puedes inventarlos, por ejemplo: `Ana López`, `Carlos Ruiz`, etc., uno por fila).

**Eliminar una fila:**

8. Haz clic en el número de la **última fila de datos** que corresponde a la región Centro (aproximadamente fila 12 o 13, dependiendo de los desplazamientos anteriores). Selecciona esa fila completa.

![Imagen práctica](../images/imagen%2055.png)

9. Haz clic derecho y elige **Eliminar**. La fila se elimina y los datos de abajo suben.

10. Verifica que la tabla de datos sigue teniendo exactamente **10 registros** (filas 2 a 11) después de los ajustes.

11. Agrega valores aleatorios a los huecos del registro que quedaba en blanco y copialos como valores

![Imagen práctica](../images/imagen%2056.png)

#### Resultado esperado

La tabla de datos en la hoja `Resumen` tiene una columna `Vendedor` insertada entre `ID_Venta` y `Región`, y mantiene 10 filas de datos sin filas en blanco intermedias.

#### Verificación

- Cuenta las filas de datos: deben ser exactamente 10 (de la fila 2 a la fila 11).
- La columna `Vendedor` debe estar entre `ID_Venta` y `Región`.
- No deben existir filas completamente en blanco dentro del rango de datos.

---

### Paso 6 — Aplicar estilos de celda, alineación y agrupar hojas

**Objetivo:** Mejorar la presentación visual de los datos usando estilos predefinidos, configurar alineación y aplicar cambios simultáneos en múltiples hojas mediante agrupación.

**Duración estimada:** 6 minutos

#### Instrucciones

**Aplicar estilos de celda:**

1. En la hoja **`Resumen`**, selecciona la fila de encabezados **`F1:K1`** (ajusta según las columnas reales de tu tabla).

2. Ve a la pestaña **Inicio** → grupo **Estilos** → haz clic en **Estilos de celda**.

3. En la galería que aparece, selecciona el estilo **"Encabezado 1"** (o el estilo de encabezado que prefieras de la sección "Títulos y encabezados").

![Imagen práctica](../images/imagen%2057.png)

4. Selecciona el rango de datos **`F2:K11`** (todas las filas de datos).

5. En la misma galería de **Estilos de celda**, selecciona el estilo **"20% - Énfasis 1"** (un color de fondo suave) para dar un tono diferente a los datos.

![Imagen práctica](../images/imagen%2058.png)

**Configurar alineación:**

6. Selecciona nuevamente los encabezados **`F1:K1`**.

7. En la pestaña **Inicio** → grupo **Alineación**, haz clic en:
   - **Centrar** (alineación horizontal).
   - **Centrar en vertical** (alineación vertical).
   - **Ajustar texto** (para que los encabezados largos no se corten).

![Imagen práctica](../images/imagen%2059.png)

8. Selecciona las celdas de la columna de IDs (`F2:F11`) y aplica **alineación centrada**.

![Imagen práctica](../images/imagen%2060.png)

**Agrupar hojas para aplicar encabezados simultáneamente:**

9. Haz clic en la pestaña de la hoja **`Ventas_Q1`**.

10. Mantén presionada la tecla **Ctrl** y haz clic en la pestaña **`Ventas_Q2`**. Ambas hojas quedarán seleccionadas (agrupadas). La barra de título mostrará `[Grupo]`.

11. Con las hojas agrupadas, haz clic en la celda **`A1`** de la hoja activa.

12. Escribe: `Reporte de Ventas Trimestrales` y presiona **Enter**.

13. Aplica el estilo **"Título"** desde la galería de Estilos de celda.

![Imagen práctica](../images/imagen%2061.png)

14. Para desagrupar las hojas, haz clic derecho sobre cualquier pestaña de hoja y selecciona **Desagrupar hojas**, o simplemente haz clic en una pestaña que no esté en el grupo.

15. Verifica que tanto `Ventas_Q1` como `Ventas_Q2` tienen el mismo encabezado en `A1`.

#### Resultado esperado

- La tabla en `Resumen` tiene encabezados con estilo "Encabezado 1" y datos con fondo de color suave.
- Las hojas `Ventas_Q1` y `Ventas_Q2` tienen el texto `Reporte de Ventas Trimestrales` en `A1` con el mismo formato, aplicado simultáneamente gracias a la agrupación.

#### Verificación

- Haz clic en la hoja `Ventas_Q2` y verifica que la celda `A1` contiene `Reporte de Ventas Trimestrales` con el estilo "Título" aplicado.
- En la hoja `Resumen`, los encabezados de la tabla deben tener fondo de color y texto en negrita (estilo "Encabezado 1").

---

### Paso 7 — Crear y usar rangos nombrados

**Objetivo:** Definir rangos nombrados para las áreas principales de datos y referenciarlos en fórmulas para mejorar la legibilidad.

**Duración estimada:** 5 minutos

#### Instrucciones

1. Ve a la hoja **`Ventas_Q1`** y selecciona el rango **`B2:D6`** (ventas de Enero, Febrero y Marzo para las 5 regiones, sin encabezados ni totales).

2. Haz clic en el **Cuadro de nombres** (el campo que muestra la referencia de celda, ubicado a la izquierda de la barra de fórmulas, donde normalmente dice algo como `B2`).

3. Escribe el nombre: `Ventas_Q1_Datos` y presiona **Enter**. El rango queda nombrado.

![Imagen práctica](../images/imagen%2062.png)

4. Selecciona el rango **`E3:E7`** (columna de totales trimestrales).

5. En el **Cuadro de nombres**, escribe: `Totales_Q1` y presiona **Enter**.


6. **Verificar los rangos nombrados en el Administrador de nombres:**
   - Ve a la pestaña **Fórmulas** → grupo **Nombres definidos** → haz clic en **Administrador de nombres**.
   - Confirma que aparecen `Ventas_Q1_Datos` y `Totales_Q1` en la lista.
   - Cierra el Administrador de nombres.

   ![Imagen práctica](../images/imagen%2063.png)

7. **Usar los rangos nombrados en fórmulas:**
   - Haz clic en la celda **`G2`** de la hoja `Ventas_Q1` (o cualquier celda vacía disponible).
   - Escribe la siguiente fórmula y presiona **Enter**:
     ```
     =SUMA(Totales_Q1)
     ```
   - Esta fórmula suma todas las ventas del trimestre usando el nombre del rango en lugar de la referencia de celdas.

   ![Imagen práctica](../images/imagen%2064.png)

8. En la celda **`G3`**, escribe:
   ```
   =PROMEDIO(Totales_Q1)
   ```
   Esta fórmula calcula el promedio de los totales por región.

   ![Imagen práctica](../images/imagen%2065.png)



#### Resultado esperado

El Administrador de nombres muestra los rangos `Ventas_Q1_Datos` y `Totales_Q1` correctamente definidos. Las celdas `G2` y `G3` muestran el total general de ventas y el promedio por región, calculados usando los nombres de rango en lugar de referencias de celda directas.

#### Verificación

- Haz clic en `G2`. La barra de fórmulas debe mostrar `=SUMA(Totales_Q1)`.
- Haz clic en `G3`. La barra de fórmulas debe mostrar `=PROMEDIO(Totales_Q1)`.
- Abre el Administrador de nombres (`Fórmulas → Administrador de nombres`) y confirma que ambos rangos están listados con sus referencias correctas.

---

### Paso 8 — Insertar y configurar Minigráficos (Sparklines)

**Objetivo:** Insertar minigráficos de línea para visualizar la tendencia de ventas mensual por región directamente dentro de las celdas.

**Duración estimada:** 4 minutos

#### Instrucciones

1. En la hoja **`Ventas_Q1`**, haz clic en la celda **`F3`** (debe estar vacía; si no lo está, usa la siguiente columna disponible después de los datos).

2. Ve a la pestaña **Insertar** → grupo **Minigráficos** → haz clic en **Línea**.

3. En el cuadro de diálogo **Crear minigráficos**:
   - **Rango de datos:** Selecciona `B3:D3` (ventas de Enero, Febrero, Marzo para la primera región).
   - **Rango de ubicación:** Confirma que dice `$F$3`.
   - Haz clic en **Aceptar**.

   ![Imagen práctica](../images/imagen%2066.png)

4. Haz clic en la celda `F3` que ahora contiene el minigráfico.

5. Arrastra el **controlador de relleno** de `F3` hacia abajo hasta `F7` para crear minigráficos para las 5 regiones. Excel ajustará automáticamente el rango de datos de cada minigráfico.

6. **Personalizar los minigráficos:**
   - Con los minigráficos `F3:F7` seleccionados, aparecerá la pestaña contextual **Minigráfico** en la cinta de opciones.
   - En el grupo **Mostrar**, activa las casillas: **Punto alto** y **Punto bajo**.
   - En el grupo **Estilo**, selecciona un estilo de color que contraste bien (por ejemplo, el estilo de color naranja o azul oscuro).

   ![Imagen práctica](../images/imagen%2067.png)

7. **Agregar un encabezado:**
   - En la celda `F2`, escribe: `Tendencia`.


#### Resultado esperado

Las columnas F de la hoja `Ventas_Q1` contiene minigráficos de línea. Los minigráficos de línea muestran marcadores en los puntos más alto y más bajo.

![Imagen práctica](../images/imagen%2068.png)

#### Verificación

- Haz clic en la celda `F3`. La pestaña contextual **Minigráfico** debe aparecer en la cinta de opciones.
- Los puntos alto y bajo deben estar marcados con un color diferente en los minigráficos de línea.

---

### Paso 9 — Aplicar Formato Condicional

**Objetivo:** Configurar reglas de formato condicional con escalas de color, barras de datos y conjuntos de iconos para identificar visualmente los valores más altos y más bajos en la tabla de ventas.

**Duración estimada:** 5 minutos

#### Instrucciones

**Escala de color en ventas de Enero:**

1. En la hoja **`Ventas_Q1`**, selecciona el rango **`B3:B7`** (ventas de Enero para las 5 regiones).

2. Ve a la pestaña **Inicio** → grupo **Estilos** → haz clic en **Formato condicional**.

3. En el menú desplegable, selecciona **Escalas de color** y elige la escala **Verde - Amarillo - Rojo** (la primera opción de la galería). Los valores más altos aparecerán en verde y los más bajos en rojo.

![Imagen práctica](../images/imagen%2069.png)

**Barras de datos en ventas de Febrero:**

4. Selecciona el rango **`C3:C7`** (ventas de Febrero).

5. Ve a **Formato condicional → Barras de datos** y selecciona **Relleno degradado: Azul** (o el color que prefieras).

6. Las celdas mostrarán barras proporcionales al valor de cada celda.

![Imagen práctica](../images/imagen%2070.png)

**Conjunto de iconos en los totales:**

7. Selecciona el rango **`E3:E7`** (totales trimestrales).

8. Ve a **Formato condicional → Conjuntos de iconos** y selecciona el conjunto de **3 flechas (de colores)** (flecha verde hacia arriba, amarilla horizontal, roja hacia abajo).

![Imagen práctica](../images/imagen%2071.png)


**Verificar y administrar las reglas:**

9. Con el rango `E2:E6` aún seleccionado, ve a **Formato condicional → Administrar reglas**.

10. Confirma que la regla de iconos aparece en la lista. Haz clic en **Editar regla** para explorar las opciones de configuración (umbrales porcentuales). No es necesario modificarlas; solo observa la configuración. Haz clic en **Cancelar** para cerrar sin cambios.

![Imagen práctica](../images/imagen%2072.png)

11. Haz clic en **Cerrar** para cerrar el Administrador de reglas.

#### Resultado esperado

- La columna B (Enero) muestra un gradiente de color verde a rojo según el valor de cada celda.
- La columna C (Febrero) muestra barras de datos azules proporcionales a cada valor.
- La columna E (Totales) muestra iconos de flechas de colores: verde para los valores más altos, rojo para los más bajos.

#### Verificación

- La región con el mayor total en la columna E debe mostrar una flecha verde apuntando hacia arriba.
- La región con el menor total en la columna E debe mostrar una flecha roja apuntando hacia abajo.
- Al pasar el cursor sobre cualquier celda de `B3:B7`, debe ser evidente cuál tiene el color más verde (valor más alto) y cuál tiene el color más rojo (valor más bajo).

---

## Validación y Pruebas Finales

Antes de dar por completada la práctica, realiza las siguientes verificaciones globales:

### Lista de verificación final

| # | Elemento a verificar | Hoja | ¿Correcto? |
|---|---|---|---|
| 1 | El rango `B10:E14` en `Ventas_Q1` contiene solo valores (sin fórmulas) | `Ventas_Q1` | ☐ |
| 2 | La tabla transpuesta comienza en `A18` con regiones en columnas y meses en filas | `Ventas_Q1` | ☐ |
| 3 | Las series de fechas, semanas, metas y meses en columnas A-D están completas (12 filas) | `Resumen` | ☐ |
| 4 | La tabla de datos de prueba (`F1:K11`) tiene IDs del 1001-1010 y valores fijos de ventas | `Resumen` | ☐ |
| 5 | Los rangos nombrados `Ventas_Q1_Datos` y `Totales_Q1` existen en el Administrador de nombres | `Ventas_Q1` | ☐ |
| 6 | Las celdas `G2` y `G3` usan fórmulas con nombres de rango (`SUMA(Ventas_Q1_Datos)`, etc.) | `Ventas_Q1` | ☐ |
| 7 | Las columnas F y G tienen minigráficos de línea y columna para las 5 regiones | `Ventas_Q1` | ☐ |
| 8 | Las columnas B, C y E tienen reglas de formato condicional aplicadas | `Ventas_Q1` | ☐ |
| 9 | Las hojas `Ventas_Q1` y `Ventas_Q2` tienen el mismo encabezado en `A1` | Ambas | ☐ |
| 10 | El autoguardado está activo y el archivo está guardado en OneDrive | Todas | ☐ |

### Prueba de integridad de rangos nombrados

1. Presiona **Ctrl+G** (o **F5**) para abrir el cuadro de diálogo **Ir a**.
2. Escribe `Ventas_Q1_Datos` en el campo de referencia y presiona **Enter**.
3. Excel debe seleccionar automáticamente el rango `B2:D6` en la hoja `Ventas_Q1`.
4. Repite el proceso con `Totales_Q1` y verifica que selecciona `E2:E6`.

### Prueba de formato condicional

1. En la hoja `Ventas_Q1`, modifica manualmente el valor de la celda `B2` a un número muy alto (por ejemplo, `999999`).
2. Verifica que la celda `B2` cambia inmediatamente a color verde intenso (el valor más alto de la columna).
3. Presiona **Ctrl+Z** para deshacer el cambio.

---

## Resolución de Problemas

### Problema 1: `SECUENCIA()` muestra `#¿NOMBRE?` o `#NOMBRE?` en la celda

**Síntoma:** Al ingresar `=SECUENCIA(10,1,1001,1)` en la celda `F2`, Excel muestra el error `#¿NOMBRE?` en lugar del rango de números esperado.

**Causa:** Este error indica que Excel no reconoce la función `SECUENCIA()`. Las causas más comunes son:
1. La versión de Excel instalada es anterior a Microsoft 365 (por ejemplo, Excel 2019 o 2016), donde esta función no existe.
2. Excel está en inglés y se está usando el nombre en español (`SECUENCIA` en lugar de `SEQUENCE`).

**Solución:**
- **Si Excel está en inglés:** Reemplaza la fórmula por su equivalente en inglés:
  ```
  =SEQUENCE(10,1,1001,1)
  ```
- **Si la versión es anterior a Microsoft 365:** Usa esta alternativa con funciones tradicionales para generar los IDs en el rango `F2:F11`:
  - En `F2`: `=1001`
  - En `F3`: `=F2+1`
  - Luego copia `F3` hasta `F11`.
  - Notifica a tu instructor para que registre la incompatibilidad de versión.
- **Verificación de versión:** Ve a `Archivo → Cuenta → Acerca de Excel` y confirma el número de versión. Debe ser 2308 o superior para Microsoft 365.

---

### Problema 2: Los minigráficos no aparecen o muestran una línea plana sin variación

**Síntoma:** Después de insertar los minigráficos en el rango `F2:F6`, las celdas aparecen vacías o muestran una línea completamente horizontal sin variación visible.

**Causa:** Existen dos causas frecuentes:
1. **Rango de datos incorrecto:** El rango de datos especificado al crear el minigráfico no coincide con las celdas que contienen los valores de ventas. Esto ocurre cuando las columnas se desplazaron al insertar la columna de `Vendedor` en el Paso 5 y el rango de datos del minigráfico no se actualizó.
2. **Todos los valores son idénticos:** Si los datos de ventas de los tres meses son exactamente iguales para una región, la línea aparecerá plana (esto es técnicamente correcto, no un error).

**Solución:**
- **Para verificar y corregir el rango de datos:**
  1. Haz clic en la celda con el minigráfico problemático (por ejemplo, `F2`).
  2. Ve a la pestaña contextual **Minigráfico** → grupo **Minigráfico** → haz clic en **Editar datos**.
  3. En el cuadro de diálogo, verifica que el **Rango de datos** apunta a las celdas correctas de ventas (por ejemplo, `B2:D2` para la primera región). Si el rango es incorrecto, corrígelo manualmente.
  4. Haz clic en **Aceptar**.
- **Para verificar que los datos no son idénticos:**
  - Revisa las celdas de ventas de la región afectada. Si los tres valores son iguales, modifica uno ligeramente para confirmar que el minigráfico responde. Luego restaura el valor original.
- **Si las celdas del minigráfico aparecen completamente vacías (sin ninguna línea):**
  - Elimina el minigráfico haciendo clic derecho → **Minigráfico → Borrar minigráficos seleccionados**.
  - Repite el proceso de inserción desde el Paso 8, verificando cuidadosamente el rango de datos en el cuadro de diálogo.

---

## Limpieza del Entorno

Al finalizar la práctica, realiza los siguientes pasos para dejar el entorno ordenado:

1. **Guardar el archivo final:**
   - Presiona **Ctrl+S** para guardar manualmente (además del autoguardado).
   - Confirma que el archivo se guardó en OneDrive verificando que el indicador de autoguardado muestra "Guardado en OneDrive".

2. **Eliminar datos temporales de prueba en la hoja `Ventas_Q1`:**
   - Selecciona el rango `B10:E14` (tabla de valores pegados del Paso 1) y presiona **Supr** para limpiar el contenido (mantén el rango para referencia futura si tu instructor lo indica).
   - Selecciona el rango `A18:E22` (tabla transpuesta del Paso 2) y presiona **Supr**.

   > **Nota:** Consulta con tu instructor si debes mantener estos rangos para la evaluación. En algunos casos, el instructor puede querer revisar el trabajo completo.

3. **Cerrar el Administrador de nombres** (si quedó abierto): Ve a `Fórmulas → Administrador de nombres` y ciérralo si está visible.

4. **Desagrupar hojas** (si quedaron agrupadas): Haz clic derecho en cualquier pestaña de hoja → **Desagrupar hojas**.

5. **Cerrar Excel** si no continuarás con la siguiente práctica: `Archivo → Cerrar`. El archivo ya está guardado en OneDrive.

---

## Resumen

En esta práctica aplicaste un conjunto completo de técnicas intermedias de manipulación de celdas y rangos en Microsoft Excel 365:

| Técnica | Herramienta utilizada | Resultado obtenido |
|---|---|---|
| Pegado selectivo | Ctrl+Alt+V → Valores / Formatos / Transponer | Copias limpias sin fórmulas; tabla rotada |
| Series automáticas | Controlador de relleno (autorrelleno) | Fechas semanales, identificadores y meses completados |
| Datos dinámicos | `SECUENCIA()` y `ALEATORIO.ENTRE()` | Tabla de 10 registros con IDs y ventas simuladas |
| Estructura de datos | Insertar/eliminar filas y columnas | Tabla reorganizada con columna `Vendedor` |
| Presentación visual | Estilos de celda y agrupación de hojas | Encabezados uniformes en múltiples hojas simultáneamente |
| Legibilidad de fórmulas | Rangos nombrados + Administrador de nombres | Fórmulas `=SUMA(Totales_Q1)` más legibles |
| Visualización de tendencias | Minigráfico de línea | Tendencias visibles por región dentro de celdas |
| Análisis visual | Formato condicional (escalas, barras, iconos) | Identificación inmediata de valores altos y bajos |

### Conceptos clave para recordar

- **Pegado especial** (Ctrl+Alt+V) es tu herramienta principal para controlar exactamente *qué* se transfiere al pegar: valores, formatos, fórmulas o la estructura transpuesta de una tabla.
- **Autorrelleno** detecta patrones; necesitas al menos **dos celdas** para que Excel identifique el incremento en series numéricas. Para listas predefinidas (meses, días), basta con **una celda**.
- **`SECUENCIA()`** y **`ALEATORIO.ENTRE()`** son funciones exclusivas de Microsoft 365; siempre convierte los resultados de `ALEATORIO.ENTRE()` a valores fijos con Pegado Especial → Valores cuando necesites datos estables.
- Los **rangos nombrados** no solo mejoran la legibilidad; también hacen que las fórmulas sean más resistentes a cambios de estructura en la hoja.
- Los **minigráficos** y el **formato condicional** son herramientas de comunicación visual: úsalos para guiar la atención del lector hacia la información más relevante.

### Recursos adicionales

- [Documentación oficial: Pegado especial en Excel](https://support.microsoft.com/es-es/office/pegar-especial-en-excel-f9f2b9a0-4a62-4b1e-b4b8-573a74f49b98)
- [Documentación oficial: Autorrelleno en hojas de cálculo](https://support.microsoft.com/es-es/office/rellenar-datos-autom%C3%A1ticamente-en-celdas-de-hojas-de-c%C3%A1lculo-74e31bdd-d993-45da-aa82-35a236c5b5db)
- [Documentación oficial: Función SECUENCIA](https://support.microsoft.com/es-es/office/funci%C3%B3n-secuencia-57467a98-57e0-4817-9f14-2eb78519ca90)
- [Documentación oficial: Minigráficos en Excel](https://support.microsoft.com/es-es/office/usar-minigr%C3%A1ficos-para-mostrar-tendencias-de-datos-1474e169-008c-4783-926b-5c60e620f5ca)
- [Documentación oficial: Formato condicional](https://support.microsoft.com/es-es/office/usar-formato-condicional-para-resaltar-informaci%C3%B3n-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
- [Microsoft Learn: Ruta de aprendizaje de Excel](https://learn.microsoft.com/es-es/training/paths/get-started-with-excel/)

---

> **Próxima práctica:** En la **Práctica 3** aplicarás referencias absolutas y mixtas en fórmulas, y comenzarás a trabajar con funciones de texto (`IZQUIERDA`, `DERECHA`, `EXTRAE`, `CONCATENAR`) y funciones de conteo condicional (`CONTAR.SI`, `SUMAR.SI`). Las técnicas de rangos nombrados que aprendiste hoy serán la base para escribir esas fórmulas de manera más legible y mantenible.
