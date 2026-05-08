# Gestión de libros y hojas

## Metadatos

| Campo            | Detalle                                      |
|------------------|----------------------------------------------|
| **Duración**     | 25 minutos                                   |
| **Complejidad**  | Media                                        |
| **Nivel Bloom**  | Aplicar                                      |
| **Módulo**       | 1 — Gestión de libros, hojas y datos externos |
| **Versión Excel**| Microsoft 365 (versión 2308 o superior)      |

---

## Descripción General

En esta práctica trabajarás con un libro de Excel que simula un entorno de gestión administrativa de recursos humanos. Comenzarás importando datos de empleados desde un archivo de texto plano (.txt) usando Power Query, y luego conectarás el libro a una fuente de datos online de ejemplo. A continuación, practicarás la navegación eficiente mediante el cuadro de nombres y rangos con nombre, insertarás hipervínculos funcionales, configurarás vistas personalizadas e inmovilización de paneles, y finalizarás ajustando las opciones de impresión y las propiedades del documento.

---

## Objetivos de Aprendizaje

Al completar esta práctica, serás capaz de:

- [ ] Importar un archivo `.txt` delimitado por comas usando Power Query desde la pestaña **Datos**, seleccionando correctamente el delimitador y la codificación de caracteres.
- [ ] Conectar Excel a una fuente de datos online y configurar la actualización de la consulta.
- [ ] Insertar hipervínculos funcionales hacia celdas de otra hoja, hacia un archivo externo y hacia una URL dentro de un libro de Excel.
- [ ] Inmovilizar paneles y agregar encabezados y pies de página con información del documento.
- [ ] Establecer el área de impresión, los títulos de impresión repetidos y las propiedades del libro (autor y palabras clave).

---

## Prerrequisitos

### Conocimientos previos

- Conocimiento básico de la interfaz de Excel: cinta de opciones, hojas, celdas y barra de fórmulas.
- Capacidad para abrir, guardar y cerrar libros de Excel.
- Familiaridad con el sistema de archivos de Windows para localizar archivos en carpetas.

### Acceso y recursos necesarios

- Cuenta de Microsoft 365 activa con acceso a Excel de escritorio (versión 2308 o superior).
- Acceso a Internet (mínimo 10 Mbps) para la importación de datos desde fuente online.
- Archivo de práctica **`empleados_datos.txt`** descargado o creado según las instrucciones del **Paso 0** de esta guía.
- Permisos de escritura en la carpeta de trabajo local (por ejemplo, `C:\LabExcel\Lab01\`).

---

## Entorno de Laboratorio

### Hardware recomendado

| Componente        | Mínimo                              | Recomendado                        |
|-------------------|-------------------------------------|------------------------------------|
| Procesador        | Intel Core i5 8ª gen / AMD Ryzen 5  | Intel Core i7 / AMD Ryzen 7        |
| RAM               | 8 GB                                | 16 GB                              |
| Almacenamiento    | 10 GB disponibles                   | SSD con 20 GB disponibles          |
| Resolución        | 1366 × 768 px                       | 1920 × 1080 px                     |
| Conexión          | 10 Mbps                             | 25 Mbps o superior                 |

### Software requerido

| Software                    | Versión mínima                  | Notas                                                   |
|-----------------------------|---------------------------------|---------------------------------------------------------|
| Microsoft Excel             | Microsoft 365 (2308+)           | Instalación de escritorio; no Excel Online              |
| Windows                     | Windows 10 21H2 / Windows 11    | —                                                       |
| Navegador web               | Edge / Chrome 110+ / Firefox 110+ | Para verificar URLs en los hipervínculos               |
| Bloc de notas (Notepad)     | Cualquier versión               | Para crear el archivo `.txt` de práctica                |

### Preparación del entorno (Paso 0)

> ⏱ **Tiempo estimado:** 3 minutos  
> Realiza estos pasos **antes** de iniciar la práctica cronometrada.

**1. Crear la carpeta de trabajo:**

```
C:\LabExcel\Lab01\
```

Abre el Explorador de archivos, navega a `C:\`, crea la carpeta `LabExcel` y dentro de ella la subcarpeta `Lab01`.

**2. Crear el archivo de práctica `empleados_datos.txt`:**

Abre el Bloc de notas (Notepad) y copia exactamente el siguiente contenido:

```
ID,Nombre,Apellido,Departamento,Cargo,Salario,FechaIngreso
1001,Ana,García,Recursos Humanos,Coordinadora,42000,15/03/2019
1002,Carlos,Mendoza,Tecnología,Desarrollador,58000,02/07/2020
1003,Laura,Pérez,Finanzas,Analista,47500,10/11/2018
1004,Miguel,Torres,Ventas,Ejecutivo,39000,23/01/2021
1005,Sofía,Ramírez,Tecnología,Arquitecta,72000,05/09/2017
1006,Andrés,López,Marketing,Director,85000,14/06/2016
1007,Valentina,Castillo,Finanzas,Contadora,44000,30/08/2022
1008,Roberto,Núñez,Ventas,Supervisor,52000,18/04/2019
1009,Diana,Vargas,Recursos Humanos,Analista,41000,07/12/2020
1010,Felipe,Morales,Marketing,Diseñador,46000,22/02/2021
```

Guarda el archivo como:
- **Nombre:** `empleados_datos.txt`
- **Tipo:** Todos los archivos (`*.*`)
- **Codificación:** UTF-8
- **Ubicación:** `C:\LabExcel\Lab01\`

![Imagen práctica](../images/imagen%2001.png)


**3. Abrir Excel y crear un nuevo libro:**

Abre Microsoft Excel 365. Crea un libro en blanco. Guárdalo inmediatamente como:
- **Nombre:** `Lab01_GestionLibros.xlsx`
- **Ubicación:** `C:\LabExcel\Lab01\`

> 💡 **Nota sobre el idioma de Excel:** Todas las fórmulas y comandos de esta práctica están escritos en español. Si tu Excel está en inglés, los nombres de los comandos en la cinta de opciones serán equivalentes en inglés. Consulta con tu instructor si tienes dudas.

---

## Procedimiento Paso a Paso

---

### Paso 1 — Importar datos desde el archivo `empleados_datos.txt` usando Power Query

> ⏱ **Tiempo estimado:** 8 minutos

**Objetivo:** Importar correctamente los datos del archivo de texto delimitado por comas hacia una hoja de Excel usando Power Query, garantizando que los tipos de datos y la codificación sean correctos.

#### Instrucciones

1. Con el libro `Lab01_GestionLibros.xlsx` abierto, haz clic en la pestaña **Datos** en la cinta de opciones.

![Imagen práctica](../images/imagen%2002.png)

2. En el grupo **Obtener y transformar datos**, haz clic en el botón **Obtener datos**.

3. En el menú desplegable, selecciona **Desde archivo** → **Desde texto/CSV**.

![Imagen práctica](../images/imagen%2004.png)

   > 📌 **¿Por qué este camino y no abrir el archivo directamente?** Si abrieras el archivo `.txt` con doble clic, Excel podría interpretar mal los formatos de fecha o los ceros iniciales en los IDs. Power Query te da control total sobre el proceso.

4. En el cuadro de diálogo **Importar datos**, navega hasta `C:\LabExcel\Lab01\`, selecciona `empleados_datos.txt` y haz clic en **Transformar datos**.

5. Se abre la ventana de vista previa de Power Query. Verifica y configura los siguientes campos:


   | Campo                  | Valor correcto a seleccionar         |
   |------------------------|--------------------------------------|
   | **Origen del archivo** | `65001: Unicode (UTF-8)`             |
   | **Delimitador**        | `Coma`                               |
   | **Detección de tipo de datos** | `Basado en las primeras 200 filas` |

![Imagen práctica](../images/imagen%2003.png)

6. Revisa la vista previa de los datos. Deberías ver 10 filas de empleados con las columnas: `ID`, `Nombre`, `Apellido`, `Departamento`, `Cargo`, `Salario`, `FechaIngreso`. Verifica que los nombres con tildes (García, Pérez, Ramírez, etc.) se muestren correctamente.

   > ⚠️ **Problema de codificación:** Si ves caracteres como `GarcÃ­a` en lugar de `García`, cambia el campo **Origen del archivo** a `65001: Unicode (UTF-8)`. Si el problema persiste, prueba con `1252: Europa occidental (Windows)`.

7. Haz clic en **Transformar datos** (no en "Cargar") para abrir el Editor de Power Query.

8. En el Editor de Power Query, verifica los tipos de datos asignados automáticamente a cada columna haciendo clic en el ícono de tipo de datos (ABC, 123, etc.) en el encabezado de cada columna:

   | Columna         | Tipo esperado     | Acción si es incorrecto                    |
   |-----------------|-------------------|--------------------------------------------|
   | `ID`            | Número entero (123) | Clic en el ícono → seleccionar **Número entero** |
   | `Nombre`        | Texto (ABC)       | Debería ser correcto automáticamente       |
   | `Apellido`      | Texto (ABC)       | Debería ser correcto automáticamente       |
   | `Departamento`  | Texto (ABC)       | Debería ser correcto automáticamente       |
   | `Cargo`         | Texto (ABC)       | Debería ser correcto automáticamente       |
   | `Salario`       | Número decimal (1.2) | Clic en el ícono → seleccionar **Número entero** |
   | `FechaIngreso`  | Fecha             | Clic en el ícono → seleccionar **Fecha**   |

9. Para cambiar el tipo de la columna `FechaIngreso`: haz clic en el ícono de tipo de datos en el encabezado de esa columna → selecciona **Fecha** 

![Imagen práctica](../images/imagen%2005.png)

→ en el cuadro de diálogo que aparece, selecciona **Sustituir la actual**.

![Imagen práctica](../images/imagen%2006.png)


10. Haz clic en **Cerrar y cargar** (botón en la esquina superior izquierda del Editor de Power Query) → selecciona **Cerrar y cargar en...** 

![Imagen práctica](../images/imagen%2007.png)

→ en el cuadro de diálogo, elige **Tabla** y selecciona **Hoja de cálculo existente** → celda `$A$1` → **Aceptar**.

![Imagen práctica](../images/imagen%2008.png)

11. Los datos se cargan en la **Hoja1** del libro como una tabla de Excel formateada. Haz doble clic en la pestaña de la hoja y renómbrala como **`Empleados`**.

#### Resultado esperado

La hoja `Empleados` debe mostrar una tabla con:
- Encabezados en la fila 1: `ID`, `Nombre`, `Apellido`, `Departamento`, `Cargo`, `Salario`, `FechaIngreso`
- 10 filas de datos (filas 2 a 11)
- Formato de tabla de Excel aplicado (con filtros automáticos en los encabezados)
- Nombres con tildes y ñ correctamente mostrados
- La columna `FechaIngreso` mostrando fechas (no texto)

#### Verificación

- [ ] La tabla tiene exactamente 11 filas de datos y 7 columnas.
- [ ] Los nombres con tildes (García, Pérez, Ramírez) se muestran correctamente, sin caracteres extraños.
- [ ] La columna `FechaIngreso` muestra valores de tipo fecha (alineados a la derecha o con formato de fecha visible).
- [ ] En el panel **Consultas y conexiones** (derecha de la pantalla) aparece la consulta `empleados_datos`.

![Imagen práctica](../images/imagen%2009.png)
---

### Paso 2 — Conectar a una fuente de datos online

> ⏱ **Tiempo estimado:** 6 minutos

**Objetivo:** Importar datos desde una URL pública que contiene una tabla HTML, y configurar la consulta para que pueda actualizarse.

#### Instrucciones

1. En el libro, crea una nueva hoja: haz clic en el botón **+** (Nueva hoja) en la barra de pestañas. Renómbrala como **`DatosOnline`**.

2. Asegúrate de estar en la hoja `DatosOnline`. Haz clic en la pestaña **Datos** → **Obtener datos** → **Desde otras fuentes** → **Desde la Web**.
![Imagen práctica](../images/imagen%2010.png)

3. En el cuadro de diálogo **Desde la Web**, selecciona la opción **Básico** e introduce la siguiente URL (una tabla de ejemplo de países y capitales publicada por W3Schools):

   ```
   https://www.w3schools.com/html/html_tables.asp
   ```

4. Haz clic en **Aceptar**. Si aparece Acceder a contenido web, selecciona anónimo y después da clic a conectar.

![Imagen práctica](../images/imagen%2011.png)

Excel se conecta a la página. Puede tardar entre 10 y 30 segundos dependiendo de la velocidad de conexión.

5. Se abre el **Navegador** de Power Query. En el panel izquierdo verás una lista de tablas detectadas en la página (pueden aparecer como `Table 1`, `Table 2`, etc.).

6. Haz clic en cada tabla para previsualizar su contenido en el panel derecho. Selecciona la tabla que contenga datos de ejemplo con columnas reconocibles (nombres de empresas, países o cualquier tabla con datos tabulares claros).

   > 💡 **Nota:** El contenido exacto de las tablas detectadas puede variar según la versión actual de la página web. El objetivo del ejercicio es practicar el proceso de selección y carga, no el contenido específico de los datos.

7. Con la tabla seleccionada, haz clic en **Cargar** (no en "Transformar datos" en este paso).

![Imagen práctica](../images/imagen%2012.png)

8. Los datos se cargan en la hoja `DatosOnline` como una tabla de Excel.

9. Ahora configura la actualización automática: haz clic derecho sobre la consulta en el panel lateral → **Propiedades**.

![Imagen práctica](../images/imagen%2013.png)


10. En la pestaña **Uso** del cuadro de diálogo **Propiedades de la consulta**, activa la casilla **Actualizar cada** y establece el valor en **`60`** minutos. Haz clic en **Aceptar**.

![Imagen práctica](../images/imagen%2014.png)

11. Para probar la actualización manual: pestaña **Datos** → botón **Actualizar todo** → observa el indicador de actualización en la barra de estado.

#### Resultado esperado

- La hoja `DatosOnline` contiene una tabla con datos importados desde la URL especificada.
- En el panel **Consultas y conexiones** aparece una segunda consulta (además de `empleados_datos`).
- La tabla tiene configurada una actualización automática cada 60 minutos.

#### Verificación

- [ ] La hoja `DatosOnline` contiene datos tabulares importados desde la web (al menos 2 columnas y 3 filas).
- [ ] En el panel **Consultas y conexiones** se listan dos consultas activas.
- [ ] Al hacer clic en **Actualizar todo**, Excel no muestra error de conexión (puede mostrar "Sin cambios" si los datos no han variado).

---

### Paso 3 — Navegación con el Cuadro de Nombres y rangos con nombre

> ⏱ **Tiempo estimado:** 4 minutos

**Objetivo:** Definir rangos con nombre en la hoja `Empleados` y practicar la navegación eficiente usando el cuadro de nombres.

#### Instrucciones

1. Haz clic en la pestaña de la hoja **`Empleados`**.

2. Selecciona el rango **`A1:G11`** (toda la tabla de empleados incluyendo encabezados).

3. Haz clic en el **Cuadro de nombres** (el campo que muestra la referencia de celda, ubicado a la izquierda de la barra de fórmulas). El contenido actual se selecciona automáticamente.

4. Escribe el nombre **`TablaEmpleados`** y presiona **Enter**. Has definido un rango con nombre.

![Imagen práctica](../images/imagen%2015.png)


5. Ahora define un segundo rango con nombre para la columna de salarios:
   - Selecciona el rango **`F2:F11`** (los valores de salario, sin el encabezado).
   - Haz clic en el **Cuadro de nombres**, escribe **`Salarios`** y presiona **Enter**.

![Imagen práctica](../images/imagen%2016.png)

6. Verifica los rangos definidos usando el Administrador de nombres: pestaña **Fórmulas** → grupo **Nombres definidos** → **Administrador de nombres**. Deberías ver `TablaEmpleados` y `Salarios` en la lista. Haz clic en **Cerrar**.

![Imagen práctica](../images/imagen%2017.png)


7. Practica la navegación: haz clic en cualquier celda aleatoria del libro. Luego haz clic en la flecha desplegable del **Cuadro de nombres** → selecciona **`TablaEmpleados`**. Excel selecciona inmediatamente el rango completo de la tabla.

8. Repite el paso anterior seleccionando **`Salarios`** desde el cuadro de nombres.

#### Resultado esperado

- El Administrador de nombres muestra dos entradas: `TablaEmpleados` (referencia `=Empleados!$A$1:$G$11`) y `Salarios` (referencia `=Empleados!$F$2:$F$11`).
- Al seleccionar un nombre desde el cuadro de nombres, Excel navega y selecciona el rango correspondiente instantáneamente.

#### Verificación

- [ ] El Administrador de nombres muestra exactamente 2 rangos con nombre definidos.
- [ ] Al seleccionar `TablaEmpleados` desde el cuadro de nombres, se selecciona el rango `A1:G11` en la hoja `Empleados`.
- [ ] Al seleccionar `Salarios` desde el cuadro de nombres, se selecciona el rango `F2:F11`.

---

### Paso 4 — Insertar hipervínculos

> ⏱ **Tiempo estimado:** 5 minutos

**Objetivo:** Insertar tres tipos de hipervínculos en una hoja de índice: uno hacia otra hoja del libro, uno hacia un archivo externo y uno hacia una URL web.

#### Instrucciones

1. Crea una nueva hoja: haz clic en **+** (Nueva hoja) → renómbrala como **`Índice`**. Arrastra esta pestaña para que quede como la **primera hoja** (más a la izquierda) del libro.

2. En la hoja `Índice`, escribe los siguientes textos en las celdas indicadas:

   | Celda | Texto a escribir                        |
   |-------|-----------------------------------------|
   | `A1`  | `ÍNDICE DE NAVEGACIÓN DEL LIBRO`        |
   | `A3`  | `Ver datos de empleados`                |
   | `A4`  | `Ver datos importados de la web`        |
   | `A5`  | `Abrir archivo de origen de datos`      |
   | `A6`  | `Visitar documentación de Power Query`  |

![Imagen práctica](../images/imagen%2018.png)

3. **Hipervínculo 1 — Hacia otra hoja del libro:**
   - Haz clic derecho en la celda **`A3`** → **Vínculo** (o **Hipervínculo**) → se abre el cuadro de diálogo **Insertar hipervínculo**.
   - En el panel izquierdo, selecciona **Lugar de este documento**.
   - En la lista de hojas, selecciona **`Empleados`**.
   - En el campo **Referencia de celda**, escribe **`A1`**.
   - Haz clic en **Aceptar**.

   ![Imagen práctica](../images/imagen%2019.png)


4. **Hipervínculo 2 — Hacia la hoja DatosOnline:**
   - Haz clic derecho en la celda **`A4`** → **Vínculo** → **Lugar de este documento**.
   - Selecciona la hoja **`DatosOnline`**, referencia de celda **`A1`**.
   - Haz clic en **Aceptar**.

   ![Imagen práctica](../images/imagen%2020.png)

5. **Hipervínculo 3 — Hacia un archivo externo:**
   - Haz clic derecho en la celda **`A5`** → **Vínculo** → en el panel izquierdo selecciona **Archivo o página web existente**.
   - Navega hasta `C:\LabExcel\Lab01\` y selecciona el archivo **`empleados_datos.txt`**.
   - Haz clic en **Aceptar**.

   ![Imagen práctica](../images/imagen%2021.png)

6. **Hipervínculo 4 — Hacia una URL web:**
   - Haz clic derecho en la celda **`A6`** → **Vínculo** → **Archivo o página web existente**.
   - En el campo **Dirección**, escribe:
     ```
     https://learn.microsoft.com/es-es/power-query/power-query-what-is-power-query
     ```
   - Haz clic en **Aceptar**.

      ![Imagen práctica](../images/imagen%2022.png)

7. Prueba cada hipervínculo haciendo **Ctrl + clic** sobre cada celda:
   - `A3` debe navegar a la hoja `Empleados`, celda `A1`.
   - `A4` debe navegar a la hoja `DatosOnline`, celda `A1`.
   - `A5` debe intentar abrir el archivo `empleados_datos.txt` en el Bloc de notas.
   - `A6` debe abrir el navegador web con la página de Microsoft Learn.

#### Resultado esperado

- Las celdas `A3`, `A4`, `A5` y `A6` muestran el texto en color azul subrayado (formato estándar de hipervínculo en Excel).
- Cada hipervínculo navega o abre el destino correcto al hacer Ctrl + clic.

#### Verificación

- [ ] Las 4 celdas muestran formato de hipervínculo (texto azul subrayado).
- [ ] Ctrl + clic en `A3` navega a la hoja `Empleados`.
- [ ] Ctrl + clic en `A4` navega a la hoja `DatosOnline`.
- [ ] Ctrl + clic en `A6` abre el navegador web en la URL de Microsoft Learn.

---

### Paso 5 — Inmovilizar paneles y configurar vistas personalizadas

> ⏱ **Tiempo estimado:** 5 minutos

**Objetivo:** Configurar la inmovilización de filas y columnas de encabezado en la hoja `Empleados`

#### Instrucciones

**Inmovilizar paneles:**

1. Haz clic en la pestaña de la hoja **`Empleados`**.

2. Haz clic en la celda **`B2`** (la celda inmediatamente debajo y a la derecha de los encabezados que deseas inmovilizar: la fila 1 con los encabezados de columna y la columna A con los IDs).


3. Ve a la pestaña **Vista** → grupo **Ventana** → **Inmovilizar paneles** → **Inmovilizar paneles** (la primera opción del submenú).

   > 💡 Aparecerá una línea horizontal debajo de la fila 1 y una línea vertical a la derecha de la columna A, indicando que esas filas/columnas están inmovilizadas.

![Imagen práctica](../images/imagen%2023.png)


4. Prueba el efecto: si la tabla fuera más grande, al desplazarte hacia abajo o hacia la derecha, la fila 1 y la columna A permanecerían visibles. Para simular esto, expande artificialmente la vista: mantén presionada la flecha hacia abajo durante unos segundos para desplazarte por debajo de los datos.


#### Verificación

- [ ] Se observan líneas de inmovilización en la hoja `Empleados` (línea horizontal bajo la fila 1, línea vertical tras la columna A).

---

### Paso 6 — Agregar encabezados y pies de página

> ⏱ **Tiempo estimado:** 4 minutos

**Objetivo:** Configurar encabezados y pies de página en la hoja `Empleados` con información profesional del documento.

#### Instrucciones

1. Con la hoja **`Empleados`** activa, ve a la pestaña **Insertar** → grupo **Texto** → **Encabezado y pie de página**.

![Imagen práctica](../images/imagen%2024.png)


   > Excel cambia a la vista **Diseño de página** y muestra tres secciones editables en el encabezado (izquierda, centro, derecha).

2. **Sección izquierda del encabezado:** haz clic en la sección izquierda del encabezado y escribe:
   ```
   Gestión de Empleados
   ```

3. **Sección central del encabezado:** haz clic en la sección central. En la pestaña **Diseño de encabezado y pie de página** que aparece en la cinta, haz clic en **Nombre de archivo** (inserta el código `&[Archivo]`).

4. **Sección derecha del encabezado:** haz clic en la sección derecha. En la cinta, haz clic en **Fecha actual** (inserta el código `&[Fecha]`).

![Imagen práctica](../images/imagen%2025.png)

5. Ahora configura el pie de página: haz clic en el botón **Ir al pie de página** en la cinta de opciones (pestaña **Diseño**).

6. **Sección izquierda del pie de página:** escribe:
   ```
   Confidencial - Uso interno
   ```

![Imagen práctica](../images/imagen%2026.png)

7. **Sección central del pie de página:** haz clic en **Número de página** (`&[Página]`), luego escribe ` de ` y haz clic en **Número de páginas** (`&[Páginas]`). El resultado será: `&[Página] de &[Páginas]`.

![Imagen práctica](../images/imagen%2027.png)

8. **Sección derecha del pie de página:** escribe:
   ```
   Lab01 - Excel 365
   ```
![Imagen práctica](../images/imagen%2028.png)

9. Haz clic en cualquier celda de la hoja para salir del modo de edición de encabezado/pie de página.

10. Para verificar el resultado visual: pestaña **Vista** → **Diseño de página** (si no estás ya en esa vista). Deberías ver el encabezado y pie de página configurados en la parte superior e inferior de la hoja.

11. Regresa a la vista normal: pestaña **Vista** → **Normal**.

![Imagen práctica](../images/imagen%2029.png)

#### Resultado esperado

- En la vista **Diseño de página**, el encabezado muestra: izquierda "Gestión de Empleados", centro el nombre del archivo, derecha la fecha actual.
- El pie de página muestra: izquierda "Confidencial - Uso interno", centro "1 de 1" (o el número de páginas correspondiente), derecha "Lab01 - Excel 365".

#### Verificación

- [ ] En la vista Diseño de página, el encabezado es visible con las tres secciones configuradas.
- [ ] El pie de página muestra numeración de páginas en el centro.
- [ ] Al volver a la vista Normal, la hoja se muestra correctamente sin el encabezado/pie en pantalla.

---

### Paso 7 — Configurar área de impresión y títulos de impresión

> ⏱ **Tiempo estimado:** 4 minutos

**Objetivo:** Definir el área de impresión de la hoja `Empleados`, configurar títulos de impresión repetidos y ajustar las opciones de página.

#### Instrucciones

1. En la hoja **`Empleados`**, selecciona el rango **`A1:G11`** (toda la tabla).

2. Ve a la pestaña **Diseño de página** → grupo **Configurar página** → **Área de impresión** → **Establecer área de impresión**.

![Imagen práctica](../images/imagen%2030.png)

   > Una línea discontinua aparece alrededor del rango seleccionado, indicando el área de impresión.

3. Ahora configura los títulos de impresión (para que los encabezados de columna se repitan en cada página impresa si la tabla crece):
   - Pestaña **Disposición de página** → **Imprimir Titulos** (abre el cuadro de diálogo **Configurar página**, pestaña **Hoja**).
   - En el campo **Repetir filas en el extremo superior**, haz clic en el botón de selección de rango (ícono de flecha roja) y selecciona la fila **1** (la fila de encabezados). El campo debe mostrar `$1:$1`.
   - Haz clic en **Aceptar**.

![Imagen práctica](../images/imagen%2031.png)

4. Ajusta la orientación de la página: pestaña **Disposición de página** → **Orientación** → **Horizontal**.

![Imagen práctica](../images/imagen%2032.png)

5. Ajusta los márgenes: pestaña **Disposición de página** → **Márgenes** → **Imprmir** → **Estrecho**.

![Imagen práctica](../images/imagen%2033.png)

6. Configura el ajuste de escala: pestaña **Disposición de página** → grupo **Ajuste de escala** → en el campo **Ancho**, selecciona **1 página**; en **Alto**, selecciona **Automático**.

![Imagen práctica](../images/imagen%2034.png)

7. Verifica la configuración en la vista previa: pestaña **Archivo** → **Imprimir** (o **Ctrl + P**). Revisa que la vista previa muestre la tabla completa en una sola página con los encabezados visibles. **No imprimas**; solo verifica la vista previa.

8. Presiona **Escape** para cerrar la vista de impresión sin imprimir.

#### Resultado esperado

- La hoja `Empleados` tiene un área de impresión definida (visible como línea discontinua).
- La vista previa de impresión muestra todos los datos en una sola página horizontal con los encabezados de columna en la parte superior.

![Imagen práctica](../images/imagen%2035.png)

#### Verificación

- [ ] El área de impresión está definida (línea discontinua visible alrededor de `A1:G11`).
- [ ] La vista previa de impresión muestra la tabla completa en una sola página.
- [ ] La orientación es horizontal (apaisada) y los márgenes son estrechos.

---

### Paso 8 — Configurar propiedades del libro e insertar comentarios

> ⏱ **Tiempo estimado:** 3 minutos

**Objetivo:** Establecer las propiedades del documento (autor, palabras clave, descripción) e insertar comentarios en celdas específicas de la hoja `Empleados`.

#### Instrucciones

**Propiedades del libro:**

1. Ve a la pestaña **Archivo** → **Información**.

2. En el panel derecho, localiza la sección **Propiedades** (puede estar en la columna derecha de la pantalla).

3. Haz clic en **Mostrar todas las propiedades** si no están todas visibles.

4. Completa los siguientes campos:

   | Propiedad        | Valor a introducir                                   |
   |------------------|------------------------------------------------------|
   | **Título**       | `Gestión de Empleados - Lab 01`                      |
   | **Etiquetas**    | `empleados; RRHH; Excel 365; importación; Power Query` |
   | **Comentarios**  | `Libro de práctica para el Lab 01 del curso de Excel intermedio` |
   | **Compañía**     | `Curso Excel 365`                                    |

5. Haz clic en la flecha **← Atrás** para volver al libro.


#### Resultado esperado

![Imagen práctica](../images/imagen%2036.png)

- Las propiedades del libro están completadas (visible desde **Archivo → Información**).


#### Verificación

- [ ] Desde **Archivo → Información**, el campo "Etiquetas" muestra las palabras clave configuradas.
---

## Validación y Pruebas Finales

Antes de dar por completada la práctica, realiza las siguientes verificaciones globales del libro:

### Lista de comprobación final

| # | Verificación | Resultado esperado | ✓ |
|---|---|---|---|
| 1 | **Estructura del libro** | El libro tiene exactamente 3 hojas: `Índice`, `Empleados`, `DatosOnline` (en ese orden) | ☐ |
| 2 | **Datos importados** | La hoja `Empleados` contiene 11 filas de datos con 7 columnas correctamente tipadas | ☐ |
| 3 | **Conexiones activas** | **Datos → Consultas y conexiones** muestra 2 consultas activas | ☐ |
| 4 | **Rangos con nombre** | **Fórmulas → Administrador de nombres** muestra `TablaEmpleados` y `Salarios` | ☐ |
| 5 | **Hipervínculos** | Las 4 celdas de la hoja `Índice` tienen hipervínculos funcionales | ☐ |
| 6 | **Inmovilización** | La hoja `Empleados` tiene paneles inmovilizados (fila 1 y columna A) | ☐ |
| 7 | **Encabezado/pie de página** | La hoja `Empleados` tiene encabezado y pie de página configurados | ☐ |
| 8 | **Área de impresión** | La vista previa de impresión muestra la tabla en una sola página horizontal | ☐ |
| 9 | **Propiedades del libro** | **Archivo → Información** muestra título y etiquetas configuradas | ☐ |
| 10 | **Archivo guardado** | El libro está guardado como `Lab01_GestionLibros.xlsx` en `C:\LabExcel\Lab01\` | ☐ |

### Prueba de integridad de la conexión

Para confirmar que la conexión online sigue activa:

1. Ve a la hoja `DatosOnline`.
2. Pestaña **Datos** → **Actualizar todo**.
3. Observa la barra de estado en la parte inferior de Excel. Debe mostrar brevemente el mensaje de actualización y luego completarse sin errores.

---

## Resolución de Problemas

### Problema 1: Los caracteres con tilde aparecen como símbolos extraños al importar el archivo `.txt`

**Síntoma:** Al cargar el archivo `empleados_datos.txt`, los nombres como "García" aparecen como "GarcÃ­a", "GarcÃa" o con cuadros negros en lugar de las letras con tilde o ñ.

**Causa:** El archivo `.txt` fue guardado con una codificación de caracteres diferente a la que Power Query está intentando usar para leerlo. El conflicto más común es entre UTF-8 (estándar moderno) y Windows-1252 (ANSI, usado por sistemas Windows más antiguos). Si el Bloc de notas guardó el archivo en ANSI pero Power Query intenta leerlo como UTF-8 (o viceversa), los caracteres especiales del español se corrompen.

**Solución:**
1. En la ventana de vista previa de Power Query (antes de cargar los datos), localiza el campo **Origen del archivo** en la parte superior.
2. Haz clic en el menú desplegable y prueba las siguientes opciones en este orden:
   - `65001: Unicode (UTF-8)` → verifica si los nombres se muestran correctamente.
   - `1252: Europa occidental (Windows)` → si UTF-8 no funcionó.
   - `65001: Unicode (UTF-8 con BOM)` → como última alternativa.
3. Una vez que los nombres se muestren correctamente en la vista previa, continúa con el proceso de importación.
4. **Prevención futura:** Al guardar archivos `.txt` desde el Bloc de notas, selecciona siempre la codificación **UTF-8** en el cuadro de diálogo "Guardar como". Esto garantiza compatibilidad con Power Query y la mayoría de sistemas modernos.

---

### Problema 2: El hipervínculo hacia el archivo `.txt` no abre el archivo o muestra un error de seguridad

**Síntoma:** Al hacer Ctrl + clic en la celda `A5` de la hoja `Índice`, Excel muestra el mensaje "Microsoft Office ha identificado un posible problema de seguridad" o "No se puede abrir el archivo especificado", y el archivo no se abre.

**Causa:** Excel aplica restricciones de seguridad a los hipervínculos que apuntan a archivos locales o de red, especialmente si el libro de Excel no está en la misma ubicación de confianza que el archivo de destino. Este comportamiento es una medida de seguridad de Microsoft Office para prevenir la apertura automática de archivos potencialmente maliciosos. También puede ocurrir si la ruta del archivo contiene espacios o caracteres especiales que no fueron codificados correctamente al insertar el hipervínculo.

**Solución:**
1. **Verificar la ruta:** Haz clic derecho en la celda `A5` → **Editar hipervínculo**. Confirma que la ruta mostrada es exactamente `C:\LabExcel\Lab01\empleados_datos.txt` sin caracteres adicionales o espacios al inicio/final.
2. **Agregar la carpeta a ubicaciones de confianza:** Ve a **Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza → Ubicaciones de confianza → Agregar nueva ubicación** → introduce `C:\LabExcel\Lab01\` y activa la casilla "Las subcarpetas de esta ubicación también son de confianza" → **Aceptar** en todos los cuadros de diálogo. Cierra y vuelve a abrir el libro, luego prueba el hipervínculo nuevamente.
3. **Alternativa:** Si el problema persiste, modifica el hipervínculo para que apunte a una URL en lugar de una ruta local, o usa la opción **Lugar de este documento** para reemplazarlo por un hipervínculo interno al libro (que no tiene restricciones de seguridad).

---

## Limpieza del Entorno

Al finalizar la práctica, realiza los siguientes pasos para dejar el entorno ordenado:

1. **Guardar el libro final:**
   - Presiona **Ctrl + S** para guardar todos los cambios.
   - Confirma que el archivo `Lab01_GestionLibros.xlsx` está guardado en `C:\LabExcel\Lab01\`.

2. **Cerrar las conexiones activas (opcional según instrucción del instructor):**
   - Si el instructor indica que se deben cerrar las conexiones: pestaña **Datos** → **Consultas y conexiones** → clic derecho sobre cada consulta → **Eliminar**. Confirma la eliminación.
   - **Nota:** Eliminar las consultas no borra los datos ya cargados en las hojas; solo elimina la conexión a la fuente original.

3. **Restaurar la vista normal:**
   - Asegúrate de que la vista activa sea **Normal** (pestaña **Vista** → **Normal**).
   - Aplica la vista personalizada **`Vista_RRHH`** para que todos los datos sean visibles al entregar el archivo.

4. **Verificar el archivo entregable:**
   - El archivo `Lab01_GestionLibros.xlsx` debe contener:
     - 3 hojas: `Índice`, `Empleados`, `DatosOnline`
     - Todos los elementos configurados durante la práctica
   - Comparte o entrega el archivo según las instrucciones de tu instructor.

5. **Archivos de práctica:** El archivo `empleados_datos.txt` puede conservarse en `C:\LabExcel\Lab01\` para referencia futura; no es necesario eliminarlo.

---

## Resumen

En esta práctica aplicaste un conjunto completo de habilidades de gestión de libros y datos externos en Microsoft Excel 365:

| Habilidad practicada | Herramienta/Función utilizada |
|---|---|
| Importación de archivos de texto | Power Query → Desde texto/CSV |
| Conexión a datos online | Power Query → Desde la Web |
| Gestión de tipos de datos en importación | Editor de Power Query |
| Navegación eficiente | Cuadro de nombres + Administrador de nombres |
| Hipervínculos internos y externos | Insertar hipervínculo |
| Inmovilización de paneles | Vista → Inmovilizar paneles |
| Vistas personalizadas | Vista → Vistas personalizadas |
| Encabezados y pies de página | Insertar → Encabezado y pie de página |
| Configuración de impresión | Diseño de página → Área/Títulos de impresión |
| Propiedades del documento | Archivo → Información |
| Comentarios y notas en celdas | Revisar → Comentarios / Notas |

### Conceptos clave para recordar

- **Power Query es preferible a la apertura directa** de archivos `.txt` o `.csv` porque permite controlar el delimitador, la codificación y los tipos de datos, y guarda el proceso como una consulta reutilizable.
- **Los rangos con nombre** aceleran la navegación y hacen las fórmulas más legibles; se gestionan desde el Administrador de nombres en la pestaña Fórmulas.
- **Las vistas personalizadas** permiten que diferentes usuarios del mismo libro vean solo la información relevante para su rol sin modificar los datos subyacentes.
- **La codificación UTF-8** es el estándar recomendado para archivos de texto en español, ya que soporta tildes, ñ y otros caracteres especiales sin corrupción.

### Recursos adicionales

- [Importar o exportar archivos de texto (.txt o .csv) en Excel — Microsoft Support](https://support.microsoft.com/es-es/office/importar-o-exportar-archivos-de-texto-txt-o-csv-5250ac4c-663c-47ce-937b-339e391393ba)
- [Conectarse a una página web con Power Query — Microsoft Support](https://support.microsoft.com/es-es/office/conectarse-a-una-p%C3%A1gina-web-power-query-b2725852-6b0e-4b1d-bad2-4c8f6d3b0f8e)
- [Introducción a Power Query — Microsoft Learn](https://learn.microsoft.com/es-es/power-query/power-query-what-is-power-query)
- [Administrar consultas en Excel para Windows — Microsoft Support](https://support.microsoft.com/es-es/office/administrar-consultas-en-excel-para-windows-35f8f6f6-5e58-4e4e-9b8e-09c7a5c07b4e)
- [Definir y usar nombres en fórmulas — Microsoft Support](https://support.microsoft.com/es-es/office/definir-y-usar-nombres-en-f%C3%B3rmulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64)

---
*Lab 01-00-01 | Curso Excel 365 Intermedio | Duración: 25 minutos*
