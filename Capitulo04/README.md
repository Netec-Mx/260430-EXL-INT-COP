# Práctica 4: Operaciones con Fórmulas y Funciones

## 1. Metadatos

| Atributo         | Detalle                        |
|------------------|--------------------------------|
| **Duración**     | 60 minutos                     |
| **Complejidad**  | Alta                           |
| **Nivel Bloom**  | Aplicar                        |
| **Módulo**       | 4 — Fórmulas y Funciones       |
| **Versión Excel**| Microsoft 365 (versión 2308+)  |

---

## 2. Descripción General

En esta práctica construirás un libro de trabajo con tres hojas interconectadas que representan los datos de una empresa ficticia: empleados, ventas mensuales y productos. A lo largo de seis módulos progresivos aplicarás referencias relativas, absolutas y mixtas; referencias estructuradas en tablas; funciones aritméticas y de conteo; lógica condicional con `SI()`; funciones dinámicas `UNICOS()` y `ORDENAR()`; y un conjunto completo de funciones de texto para transformar y combinar datos. Al finalizar, habrás construido un modelo funcional que integra todas las técnicas intermedias de fórmulas cubiertas en el curso.

> **Nota sobre idioma de fórmulas:** Todas las fórmulas de esta práctica están escritas en español, que es el idioma predeterminado de Excel instalado en español. Si tu Excel está en inglés, consulta la tabla de equivalencias al final de esta guía.

---

## 3. Objetivos de Aprendizaje

Al completar esta práctica serás capaz de:

- [ ] Distinguir y aplicar referencias relativas, absolutas (`$`) y mixtas al copiar fórmulas a través de rangos, incluyendo una tabla de multiplicar con referencias mixtas.
- [ ] Utilizar referencias estructuradas (`[@Columna]`, `[Columna]`) para operar sobre tablas de Excel sin depender de rangos de celdas tradicionales.
- [ ] Implementar `SUMA()`, `PROMEDIO()`, `MAX()` y `MIN()` para generar un panel de resumen de ventas, y aplicar `CONTAR()`, `CONTARA()` y `CONTAR.BLANCO()` para auditar la completitud de datos.
- [ ] Construir fórmulas `SI()` simples y anidadas para clasificar empleados en categorías de desempeño, y generar listas únicas y ordenadas con `UNICOS()` y `ORDENAR()`.
- [ ] Manipular cadenas de texto con `IZQUIERDA()`, `DERECHA()`, `EXTRAE()`, `LARGO()`, `MAYUSC()`, `MINUSC()`, `CONCAT()` y `UNIRCADENAS()` para transformar nombres y construir correos electrónicos corporativos.

---

## 4. Prerrequisitos

### Conocimientos previos
- Haber completado las Prácticas 1, 2 y 3 del curso, o tener conocimiento equivalente de Excel intermedio.
- Saber ingresar, editar y copiar fórmulas básicas en Excel.
- Conocer la diferencia entre valores numéricos y texto en celdas.
- Estar familiarizado con el concepto de tablas estructuradas de Excel (insertar tabla, nombrar tabla, encabezados de columna).

### Acceso y licencias
- Cuenta de Microsoft 365 activa con acceso a Excel 365 (versión 2308 o superior).
- Las funciones `UNICOS()` y `ORDENAR()` son **exclusivas de Microsoft 365**; no están disponibles en Excel 2019 o 2016. Si usas una versión anterior, consulta con tu instructor las alternativas tradicionales.
- Archivo de práctica `Lab04_Formulas_Funciones_INICIO.xlsx` proporcionado por el instructor o disponible en la carpeta del curso en OneDrive.

---

## 5. Entorno de Laboratorio

### Hardware recomendado

| Componente          | Mínimo                              | Recomendado                  |
|---------------------|-------------------------------------|------------------------------|
| Procesador          | Intel Core i5 8ª gen / AMD Ryzen 5  | Intel Core i7 / AMD Ryzen 7  |
| Memoria RAM         | 8 GB                                | 16 GB                        |
| Espacio en disco    | 10 GB disponibles                   | 20 GB disponibles            |
| Resolución pantalla | 1366 × 768 px                       | 1920 × 1080 px               |

### Software requerido

| Software              | Versión mínima         | Notas                                      |
|-----------------------|------------------------|--------------------------------------------|
| Microsoft Excel       | 365 (versión 2308+)    | Instalado en español                       |
| Sistema Operativo     | Windows 10 (21H2+)     | o Windows 11                               |
| Navegador web         | Edge / Chrome 110+     | Para descargar archivos de OneDrive        |

### Preparación del entorno

1. Abre **Microsoft Excel 365**.
2. Descarga el archivo `Lab04_Formulas_Funciones_INICIO.xlsx` desde la carpeta del curso en OneDrive.
3. Ábrelo en Excel. Verifica que el libro contiene exactamente **cinco hojas** con las siguientes pestañas:
   - `Empleados`
   - `Ventas`
   - `Productos`
   - `Tabla_Multiplicar` *(puede estar en blanco)*
   - `Texto_Empleados` *(puede estar en blanco)*
4. Guarda una copia de trabajo con el nombre `Lab04_TuNombre.xlsx` usando **Archivo → Guardar una copia**.

> **Importante:** Si el archivo de inicio no está disponible, el instructor te indicará cómo crear los datos manualmente. Las instrucciones de creación de datos están incluidas al inicio de cada módulo como paso opcional.

---

## 6. Instrucciones Paso a Paso

---

### Módulo A — Referencias Relativas y Absolutas: Cálculo de Comisiones

**Objetivo del módulo:** Aplicar correctamente referencias relativas y absolutas en una fórmula de comisiones que se copia a través de un rango de celdas.

---

#### Paso A1: Verificar la estructura de datos en la hoja Ventas

**Instrucciones:**

1. Haz clic en la pestaña **`Ventas`**.
2. Verifica que la hoja contiene la siguiente estructura (o créala si el archivo de inicio no la incluye):

   | Celda | Contenido                        |
   |-------|----------------------------------|
   | `A1`  | `Tasa de Comisión`               |
   | `B1`  | `0.08` *(8%)*                    |
   | `A3`  | `Vendedor`                       |
   | `B3`  | `Ventas Mensuales`               |
   | `C3`  | `Comisión`                       |
   | `A4`  | `Ana López`                      |
   | `A5`  | `Carlos Méndez`                  |
   | `A6`  | `Sofía Ramírez`                  |
   | `A7`  | `Diego Torres`                   |
   | `A8`  | `Valeria Núñez`                  |
   | `B4`  | `125000`                         |
   | `B5`  | `98000`                          |
   | `B6`  | `143000`                         |
   | `B7`  | `87500`                          |
   | `B8`  | `162000`                         |

3. Formatea la celda `B1` como porcentaje: selecciona `B1`, presiona **Ctrl+1**, elige **Porcentaje** con 0 decimales, haz clic en **Aceptar**.

**Resultado esperado:** La celda `B1` muestra `8%`. Las celdas `B4:B8` contienen los montos de ventas en formato numérico.

**Verificación:** Confirma que `B1` muestra `8%` y que `B4` contiene el valor `125000`.

---

#### Paso A2: Escribir la fórmula de comisión con referencia absoluta

**Instrucciones:**

1. Haz clic en la celda **`C4`**.
2. Escribe la siguiente fórmula y presiona **Enter**:
   ```
   =B4*$B$1
   ```
3. Observa el resultado: debería mostrar `10000` (125,000 × 8%).
4. Vuelve a seleccionar `C4`.
5. Coloca el cursor sobre la esquina inferior derecha de la celda hasta que aparezca el cursor de **cruz negra (+)**.
6. Arrastra hacia abajo hasta la celda **`C8`** para copiar la fórmula.

**Resultado esperado:**

| Celda | Fórmula resultante | Valor    |
|-------|--------------------|----------|
| `C4`  | `=B4*$B$1`         | `10,000` |
| `C5`  | `=B5*$B$1`         | `7,840`  |
| `C6`  | `=B6*$B$1`         | `11,440` |
| `C7`  | `=B7*$B$1`         | `7,000`  |
| `C8`  | `=B8*$B$1`         | `12,960` |

**Verificación:**
- Haz clic en `C5` y confirma en la barra de fórmulas que dice `=B5*$B$1` (la referencia a `B5` cambió, pero `$B$1` permanece fija).
- Haz clic en `C8` y confirma `=B8*$B$1`.
- Cambia temporalmente `B1` a `0.10` (10%) y verifica que todas las comisiones se recalculan. Devuelve `B1` a `0.08` al terminar.

> **Concepto clave:** `B4` es una referencia **relativa** que se ajusta al copiar. `$B$1` es una referencia **absoluta** que siempre apunta a la tasa de comisión, sin importar a qué fila se copie la fórmula.

---

### Módulo B — Referencias Mixtas: Tabla de Multiplicar

**Objetivo del módulo:** Construir una tabla de multiplicación bidimensional usando referencias mixtas que fijan solo la columna o solo la fila según corresponda.

---

#### Paso B1: Preparar la estructura de la tabla de multiplicar

**Instrucciones:**

1. Haz clic en la pestaña **`Tabla_Multiplicar`**.
2. En la celda **`A1`** escribe el título: `Tabla de Multiplicar (1–10)`.
3. En el rango **`B2:K2`** (fila de encabezados de columna), escribe los números del 1 al 10:
   - `B2` = `1`, `C2` = `2`, `D2` = `3`, ..., `K2` = `10`
   - Truco rápido: escribe `1` en `B2`, `2` en `C2`, selecciona `B2:C2` y arrastra el controlador de relleno hasta `K2`.
4. En el rango **`A3:A12`** (columna de encabezados de fila), escribe los números del 1 al 10:
   - `A3` = `1`, `A4` = `2`, ..., `A12` = `10`
   - Usa el mismo método de relleno automático.

**Resultado esperado:** La fila 2 (columnas B a K) contiene los números 1–10. La columna A (filas 3 a 12) contiene los números 1–10.

**Verificación:** Confirma que `K2` = `10` y `A12` = `10`.

---

#### Paso B2: Ingresar la fórmula con referencias mixtas

**Instrucciones:**

1. Haz clic en la celda **`B3`**.
2. Escribe la siguiente fórmula:
   ```
   =$A3*B$2
   ```
   - `$A3`: La columna A está **fija** (siempre tomará el multiplicador de la columna A), pero la fila se ajusta al copiar hacia abajo.
   - `B$2`: La fila 2 está **fija** (siempre tomará el multiplicador de la fila 2), pero la columna se ajusta al copiar hacia la derecha.
3. Presiona **Enter** y vuelve a seleccionar `B3`.
4. Copia la fórmula hacia la derecha hasta **`K3`**: selecciona `B3`, luego arrastra el controlador de relleno hasta `K3`.
5. Selecciona el rango **`B3:K3`** completo.
6. Arrastra el controlador de relleno del rango hacia abajo hasta la fila **12** (`B12:K12`).

**Resultado esperado:** La tabla completa (rango `B3:K12`) contiene los productos de multiplicación correctos. Ejemplos:

| Celda | Fórmula generada | Valor esperado |
|-------|-----------------|----------------|
| `B3`  | `=$A3*B$2`      | `1`            |
| `K3`  | `=$A3*K$2`      | `10`           |
| `B12` | `=$A12*B$2`     | `10`           |
| `K12` | `=$A12*K$2`     | `100`          |
| `E7`  | `=$A7*E$2`      | `25`           |

**Verificación:**
- Haz clic en `E7` y confirma en la barra de fórmulas: `=$A7*E$2`.
- Verifica que `K12` = `100` (10 × 10).
- Haz clic en `C5` y confirma `=$A5*C$2` = `6` (2 × 3).

> **Concepto clave:** Sin referencias mixtas, copiar la fórmula desplazaría ambas referencias y la tabla sería incorrecta. `$A3` garantiza que siempre se lea la columna A (los multiplicandos de fila), y `B$2` garantiza que siempre se lea la fila 2 (los multiplicandos de columna).

---

### Módulo C — Referencias Estructuradas y Funciones Aritméticas

**Objetivo del módulo:** Usar referencias estructuradas para operar sobre una tabla de ventas y construir un panel de resumen con `SUMA()`, `PROMEDIO()`, `MAX()` y `MIN()`.

---

#### Paso C1: Verificar y nombrar la tabla de ventas

**Instrucciones:**

1. Haz clic en la pestaña **`Ventas`**.
2. Desplázate hacia abajo. Debajo de la tabla de comisiones, a partir de la fila **12**, debes encontrar (o crear) la siguiente tabla de ventas mensuales:

   | Vendedor       | Ene    | Feb    | Mar    | Abr    | May    | Jun    | Total Vendedor |
   |----------------|--------|--------|--------|--------|--------|--------|----------------|
   | Ana López      | 42000  | 38000  | 45000  | 51000  | 39000  | 47000  |                |
   | Carlos Méndez  | 31000  | 29000  | 35000  | 28000  | 33000  | 37000  |                |
   | Sofía Ramírez  | 48000  | 52000  | 44000  | 55000  | 49000  | 53000  |                |
   | Diego Torres   | 27000  | 31000  | 29000  | 33000  | 28000  | 30000  |                |
   | Valeria Núñez  | 55000  | 58000  | 62000  | 59000  | 61000  | 64000  |                |

   Los encabezados deben estar en la fila 12, comenzando en `A12`.

3. Selecciona cualquier celda dentro de esta tabla (por ejemplo, `B13`).
4. Ve a **Insertar → Tabla** (o presiona **Ctrl+T**).
5. Confirma que el rango es correcto y que la opción **"La tabla tiene encabezados"** está marcada. Haz clic en **Aceptar**.
6. Con la tabla seleccionada, ve a la pestaña **Diseño de tabla** (o **Herramientas de tabla → Diseño**).
7. En el campo **Nombre de tabla** (extremo izquierdo de la cinta), escribe: `VentasMensuales` y presiona **Enter**.

**Resultado esperado:** La tabla está formateada con el estilo predeterminado de Excel y se llama `VentasMensuales`.

**Verificación:** Haz clic en cualquier celda de la tabla y confirma que en el campo **Nombre de tabla** aparece `VentasMensuales`.

---

#### Paso C2: Calcular el total por vendedor usando referencias estructuradas

**Instrucciones:**

1. Haz clic en la primera celda de la columna **`Total Vendedor`** (debería ser `H13`, correspondiente a Ana López).
2. Escribe la siguiente fórmula usando referencia estructurada:
   ```
   =SUMA(VentasMensuales[@[Ene]:[Jun]])
   ```
   Esta fórmula suma todas las columnas desde `Ene` hasta `Jun` para la fila actual (`@` significa "esta fila").
3. Presiona **Enter**. Excel puede completar automáticamente la fórmula para el resto de las filas de la tabla.
4. Si no se completó automáticamente, copia la fórmula hacia abajo hasta la última fila de datos.

**Resultado esperado:**

| Vendedor       | Total Vendedor |
|----------------|----------------|
| Ana López      | `262,000`      |
| Carlos Méndez  | `193,000`      |
| Sofía Ramírez  | `301,000`      |
| Diego Torres   | `178,000`      |
| Valeria Núñez  | `359,000`      |

**Verificación:** Confirma que `H13` = `262,000` y que la barra de fórmulas muestra `=SUMA(VentasMensuales[@[Ene]:[Jun]])`.

---

#### Paso C3: Construir el panel de resumen de ventas

**Instrucciones:**

1. En la hoja `Ventas`, ubícate en una zona libre, por ejemplo a partir de la celda **`J12`**.
2. Crea los siguientes encabezados y fórmulas:

   | Celda  | Contenido a escribir              | Fórmula                                          |
   |--------|-----------------------------------|--------------------------------------------------|
   | `J12`  | `PANEL DE RESUMEN`                | *(solo texto)*                                   |
   | `J13`  | `Total General de Ventas`         | *(solo texto)*                                   |
   | `K13`  | *(fórmula)*                       | `=SUMA(VentasMensuales[Total Vendedor])`          |
   | `J14`  | `Promedio por Vendedor`           | *(solo texto)*                                   |
   | `K14`  | *(fórmula)*                       | `=PROMEDIO(VentasMensuales[Total Vendedor])`      |
   | `J15`  | `Vendedor con Mayor Venta`        | *(solo texto)*                                   |
   | `K15`  | *(fórmula)*                       | `=MAX(VentasMensuales[Total Vendedor])`           |
   | `J16`  | `Vendedor con Menor Venta`        | *(solo texto)*                                   |
   | `K16`  | *(fórmula)*                       | `=MIN(VentasMensuales[Total Vendedor])`           |

3. Ingresa cada fórmula en la columna K presionando **Enter** después de cada una.

**Resultado esperado:**

| Indicador                  | Valor        |
|----------------------------|--------------|
| Total General de Ventas    | `1,293,000`  |
| Promedio por Vendedor      | `258,600`    |
| Vendedor con Mayor Venta   | `359,000`    |
| Vendedor con Menor Venta   | `178,000`    |

**Verificación:** Confirma que `K13` = `1,293,000`. Modifica temporalmente un valor en la tabla (por ejemplo, cambia el `42000` de Ana en Ene a `50000`) y verifica que `K13` se actualiza automáticamente. Deshaz el cambio con **Ctrl+Z**.

---

### Módulo D — Funciones de Conteo y Auditoría de Datos

**Objetivo del módulo:** Aplicar `CONTAR()`, `CONTARA()` y `CONTAR.BLANCO()` para auditar la completitud de los datos de empleados.

---

#### Paso D1: Verificar la estructura de datos de empleados

**Instrucciones:**

1. Haz clic en la pestaña **`Empleados`**.
2. Verifica que la hoja contiene una tabla con las siguientes columnas (o créala a partir de `A1`):

   | ID  | Nombre Completo    | Departamento  | Fecha Ingreso | Salario | % Cumplimiento | Teléfono   | Email |
   |-----|--------------------|---------------|---------------|---------|----------------|------------|-------|
   | 101 | Ana López          | Ventas        | 15/03/2019    | 28000   | 95             | 5512345678 |       |
   | 102 | Carlos Méndez      | Ventas        | 22/07/2020    | 25000   | 78             |            |       |
   | 103 | Sofía Ramírez      | Marketing     | 08/01/2018    | 32000   | 88             | 5598765432 |       |
   | 104 | Diego Torres       | Operaciones   | 30/11/2021    | 22000   | 65             | 5567891234 |       |
   | 105 | Valeria Núñez      | Ventas        | 14/05/2017    | 35000   | 92             |            |       |
   | 106 | Roberto Sánchez    | Marketing     | 03/09/2022    | 24000   | 71             | 5534567890 |       |
   | 107 | Lucía Fernández    | RRHH          | 17/02/2020    |         | 84             | 5523456789 |       |
   | 108 | Martín Guzmán      | Operaciones   | 25/06/2019    | 26000   | 90             | 5545678901 |       |
   | 109 | Isabella Morales   | RRHH          | 11/10/2021    | 23000   |                |            |       |
   | 110 | Alejandro Ríos     | Marketing     | 28/04/2023    | 21000   | 76             | 5556789012 |       |

   > Nota: Las celdas vacías en Teléfono (filas 103, 106), Salario (fila 108) y % Cumplimiento / Email (fila 110) son **intencionales** para el ejercicio de auditoría.

3. Convierte este rango en tabla: selecciona `A1:H11`, presiona **Ctrl+T**, confirma encabezados y nómbrala `TablaEmpleados`.

**Resultado esperado:** Tabla `TablaEmpleados` con 10 empleados y 8 columnas, con algunas celdas vacías intencionales.

---

#### Paso D2: Aplicar funciones de conteo para auditar los datos

**Instrucciones:**

1. En la misma hoja `Empleados`, ubícate en la celda **`J1`** y crea el siguiente panel de auditoría:

   | Celda  | Texto / Fórmula                                                    |
   |--------|--------------------------------------------------------------------|
   | `J1`   | `AUDITORÍA DE COMPLETITUD DE DATOS`                               |
   | `J2`   | `Total de registros (numéricos en ID):`                           |
   | `K2`   | `=CONTAR(TablaEmpleados[ID])`                                      |
   | `J3`   | `Total de nombres registrados (texto):`                           |
   | `K3`   | `=CONTARA(TablaEmpleados[Nombre Completo])`                        |
   | `J4`   | `Teléfonos registrados:`                                          |
   | `K4`   | `=CONTARA(TablaEmpleados[Teléfono])`                               |
   | `J5`   | `Teléfonos faltantes:`                                            |
   | `K5`   | `=CONTAR.BLANCO(TablaEmpleados[Teléfono])`                         |
   | `J6`   | `Salarios registrados:`                                           |
   | `K6`   | `=CONTAR(TablaEmpleados[Salario])`                                 |
   | `J7`   | `Salarios faltantes:`                                             |
   | `K7`   | `=CONTAR.BLANCO(TablaEmpleados[Salario])`                          |
   | `J8`   | `% Cumplimiento registrados:`                                     |
   | `K8`   | `=CONTAR(TablaEmpleados[% Cumplimiento])`                          |
   | `J9`   | `% Cumplimiento faltantes:`                                       |
   | `K9`   | `=CONTAR.BLANCO(TablaEmpleados[% Cumplimiento])`                   |

2. Ingresa cada fórmula en la columna K.

**Resultado esperado:**

| Indicador                           | Valor esperado |
|-------------------------------------|----------------|
| Total de registros (ID numéricos)   | `10`           |
| Total de nombres registrados        | `10`           |
| Teléfonos registrados               | `8`            |
| Teléfonos faltantes                 | `2`            |
| Salarios registrados                | `9`            |
| Salarios faltantes                  | `1`            |
| % Cumplimiento registrados          | `9`            |
| % Cumplimiento faltantes            | `1`            |

**Verificación:**
- Confirma que `K2` = `10` (CONTAR cuenta solo valores numéricos en la columna ID).
- Confirma que `K5` = `2` (dos teléfonos vacíos: Carlos Méndez y Valeria Núñez).
- Confirma que `K3` = `10` (CONTARA cuenta texto; todos los nombres están registrados).

> **Concepto clave:** `CONTAR()` solo cuenta celdas con **números**. `CONTARA()` cuenta celdas con **cualquier contenido** (números, texto, fechas). `CONTAR.BLANCO()` cuenta celdas **vacías**. La suma de `CONTARA()` + `CONTAR.BLANCO()` debe ser igual al total de filas de la tabla.

---

### Módulo E — Función SI() Simple y Anidada: Clasificación de Desempeño

**Objetivo del módulo:** Construir fórmulas `SI()` para clasificar empleados en categorías de desempeño usando condiciones simples y anidadas.

---

#### Paso E1: Agregar columna de clasificación con SI() simple

**Instrucciones:**

1. En la hoja **`Empleados`**, haz clic en la celda de encabezado de la primera columna vacía a la derecha de la tabla `TablaEmpleados`. Si la tabla termina en la columna H (`Email`), haz clic en `I1`.
2. Escribe el encabezado: `Clasificación` y presiona **Enter**. Excel extenderá automáticamente la tabla.
3. En la primera celda de datos de esta columna (por ejemplo, `I2`), escribe la siguiente fórmula de `SI()` simple:
   ```
   =SI([@[% Cumplimiento]]>=80,"Cumple","No Cumple")
   ```
4. Presiona **Enter**. Excel completará automáticamente la fórmula para todos los empleados.

**Resultado esperado:**

| Empleado           | % Cumplimiento | Clasificación |
|--------------------|----------------|---------------|
| Ana López          | 95             | Cumple        |
| Carlos Méndez      | 78             | No Cumple     |
| Sofía Ramírez      | 88             | Cumple        |
| Diego Torres       | 65             | No Cumple     |
| Valeria Núñez      | 92             | Cumple        |
| Roberto Sánchez    | 71             | No Cumple     |
| Lucía Fernández    | 84             | Cumple        |
| Martín Guzmán      | 90             | Cumple        |
| Isabella Morales   | *(vacío)*      | No Cumple     |
| Alejandro Ríos     | 76             | No Cumple     |

**Verificación:** Confirma que Ana López muestra "Cumple" y Carlos Méndez muestra "No Cumple".

---

#### Paso E2: Agregar columna de nivel de desempeño con SI() anidado

**Instrucciones:**

1. Haz clic en la celda de encabezado de la siguiente columna vacía de la tabla (por ejemplo, `J1` si la tabla se extendió).
2. Escribe el encabezado: `Nivel Desempeño` y presiona **Enter**.
3. En la primera celda de datos (por ejemplo, `J2`), escribe la siguiente fórmula con `SI()` anidado de tres categorías:
   ```
   =SI([@[% Cumplimiento]]>=90,"Alto",SI([@[% Cumplimiento]]>=75,"Medio","Bajo"))
   ```
   Esta fórmula evalúa:
   - Si `% Cumplimiento` ≥ 90 → `"Alto"`
   - Si `% Cumplimiento` ≥ 75 (pero < 90) → `"Medio"`
   - En cualquier otro caso (< 75 o vacío) → `"Bajo"`
4. Presiona **Enter**.

**Resultado esperado:**

| Empleado           | % Cumplimiento | Nivel Desempeño |
|--------------------|----------------|-----------------|
| Ana López          | 95             | Alto            |
| Carlos Méndez      | 78             | Medio           |
| Sofía Ramírez      | 88             | Medio           |
| Diego Torres       | 65             | Bajo            |
| Valeria Núñez      | 92             | Alto            |
| Roberto Sánchez    | 71             | Bajo            |
| Lucía Fernández    | 84             | Medio           |
| Martín Guzmán      | 90             | Alto            |
| Isabella Morales   | *(vacío)*      | Bajo            |
| Alejandro Ríos     | 76             | Medio           |

**Verificación:**
- Confirma que Ana López = "Alto" (95 ≥ 90).
- Confirma que Sofía Ramírez = "Medio" (88 ≥ 75 pero < 90).
- Confirma que Diego Torres = "Bajo" (65 < 75).

---

#### Paso E3: Extraer lista de departamentos únicos y ordenados

**Instrucciones:**

1. Ubícate en una zona libre de la hoja `Empleados`, por ejemplo la celda **`L1`**.
2. Escribe el encabezado: `Departamentos (únicos)`.
3. En la celda **`L2`**, escribe la siguiente fórmula dinámica:
   ```
   =ORDENAR(UNICOS(TablaEmpleados[Departamento]))
   ```
   Esta fórmula:
   - `UNICOS()`: extrae los valores únicos (sin repetición) de la columna Departamento.
   - `ORDENAR()`: ordena el resultado alfabéticamente (orden ascendente por defecto).
4. Presiona **Enter**. La fórmula derramará los resultados hacia abajo automáticamente.

**Resultado esperado:** A partir de `L2`, aparecerá la lista:
```
Marketing
Operaciones
RRHH
Ventas
```
*(en orden alfabético, sin repeticiones)*

**Verificación:** Confirma que aparecen exactamente 4 departamentos únicos en orden alfabético. Haz clic en `L2` y verifica que la barra de fórmulas muestra `=ORDENAR(UNICOS(TablaEmpleados[Departamento]))`.

> **Nota:** Si tu Excel no es versión 365, `UNICOS()` y `ORDENAR()` no estarán disponibles. Consulta con tu instructor las alternativas para versiones anteriores.

---

### Módulo F — Manipulación de Texto

**Objetivo del módulo:** Aplicar funciones de texto para separar nombres, normalizar mayúsculas y construir correos electrónicos corporativos.

---

#### Paso F1: Preparar la hoja de trabajo de texto

**Instrucciones:**

1. Haz clic en la pestaña **`Texto_Empleados`**.
2. Verifica o crea la siguiente estructura de datos a partir de `A1`:

   | A (Nombre Completo)   | B (Código Depto) |
   |-----------------------|------------------|
   | ANA LOPEZ             | VEN              |
   | CARLOS MENDEZ         | VEN              |
   | SOFIA RAMIREZ         | MKT              |
   | DIEGO TORRES          | OPS              |
   | VALERIA NUNEZ         | VEN              |
   | ROBERTO SANCHEZ       | MKT              |
   | LUCIA FERNANDEZ       | RRH              |
   | MARTIN GUZMAN         | OPS              |
   | ISABELLA MORALES      | RRH              |
   | ALEJANDRO RIOS        | MKT              |

   > Los nombres están en MAYÚSCULAS y sin acentos para facilitar el ejercicio de normalización.

3. Agrega los siguientes encabezados en la fila 1 (si los datos empiezan en fila 2):
   - `A1`: `Nombre Completo`
   - `B1`: `Código Depto`
   - `C1`: `Primer Nombre`
   - `D1`: `Apellido`
   - `E1`: `Nombre Normalizado`
   - `F1`: `Apellido Normalizado`
   - `G1`: `Largo Nombre`
   - `H1`: `Correo Corporativo`
   - `I1`: `Correo con UNIRCADENAS`

**Resultado esperado:** La hoja tiene encabezados en la fila 1 y datos de 10 empleados en las filas 2–11.

---

#### Paso F2: Extraer el primer nombre con IZQUIERDA() y HALLAR()

**Instrucciones:**

1. Haz clic en la celda **`C2`**.
2. Escribe la siguiente fórmula para extraer el primer nombre (todo lo que está antes del espacio):
   ```
   =IZQUIERDA(A2,HALLAR(" ",A2)-1)
   ```
   - `HALLAR(" ",A2)` encuentra la posición del primer espacio en el nombre completo.
   - `-1` excluye el espacio del resultado.
   - `IZQUIERDA()` extrae los caracteres desde la izquierda hasta esa posición.
3. Presiona **Enter** y copia la fórmula hacia abajo hasta `C11`.

**Resultado esperado:**
- `C2` = `ANA`
- `C3` = `CARLOS`
- `C5` = `VALERIA`

**Verificación:** Confirma que `C2` = `ANA` y `C3` = `CARLOS`.

---

#### Paso F3: Extraer el apellido con DERECHA() y LARGO()

**Instrucciones:**

1. Haz clic en la celda **`D2`**.
2. Escribe la siguiente fórmula para extraer el apellido (todo lo que está después del espacio):
   ```
   =DERECHA(A2,LARGO(A2)-HALLAR(" ",A2))
   ```
   - `LARGO(A2)` devuelve el número total de caracteres del nombre completo.
   - `HALLAR(" ",A2)` devuelve la posición del espacio.
   - La diferencia indica cuántos caracteres hay después del espacio.
   - `DERECHA()` extrae esa cantidad de caracteres desde la derecha.
3. Presiona **Enter** y copia la fórmula hacia abajo hasta `D11`.

**Resultado esperado:**
- `D2` = `LOPEZ`
- `D3` = `MENDEZ`
- `D9` = `MORALES`

**Verificación:** Confirma que `D2` = `LOPEZ` y `D9` = `MORALES`.

---

#### Paso F4: Normalizar mayúsculas con MAYUSC() y MINUSC()

**Instrucciones:**

1. Haz clic en la celda **`E2`**.
2. Escribe la fórmula para convertir el primer nombre a formato de título (primera letra mayúscula, resto minúsculas). Usaremos `CONCAT()` con `MAYUSC()` e `IZQUIERDA()` y `MINUSC()` con `EXTRAE()`:
   ```
   =CONCAT(MAYUSC(IZQUIERDA(C2,1)),MINUSC(EXTRAE(C2,2,LARGO(C2)-1)))
   ```
   - `MAYUSC(IZQUIERDA(C2,1))`: convierte la primera letra a mayúscula.
   - `MINUSC(EXTRAE(C2,2,LARGO(C2)-1))`: convierte el resto del nombre a minúsculas.
   - `CONCAT()`: une ambas partes.
3. Presiona **Enter** y copia la fórmula hasta `E11`.
4. Haz clic en la celda **`F2`** y aplica la misma lógica para el apellido:
   ```
   =CONCAT(MAYUSC(IZQUIERDA(D2,1)),MINUSC(EXTRAE(D2,2,LARGO(D2)-1)))
   ```
5. Copia la fórmula hasta `F11`.

**Resultado esperado:**
- `E2` = `Ana`, `F2` = `Lopez`
- `E3` = `Carlos`, `F3` = `Mendez`
- `E9` = `Isabella`, `F9` = `Morales`

**Verificación:** Confirma que `E2` = `Ana` (no `ANA` ni `ana`).

---

#### Paso F5: Calcular el largo del nombre completo con LARGO()

**Instrucciones:**

1. Haz clic en la celda **`G2`**.
2. Escribe la fórmula:
   ```
   =LARGO(A2)
   ```
3. Presiona **Enter** y copia la fórmula hasta `G11`.

**Resultado esperado:**
- `G2` = `9` (ANA LOPEZ tiene 9 caracteres incluyendo el espacio)
- `G3` = `13` (CARLOS MENDEZ tiene 13 caracteres)
- `G10` = `14` (ALEJANDRO RIOS tiene 14 caracteres)

**Verificación:** Cuenta manualmente los caracteres de "ANA LOPEZ" (A-N-A-espacio-L-O-P-E-Z = 9) y confirma que `G2` = `9`.

---

#### Paso F6: Construir correos electrónicos con CONCAT()

**Instrucciones:**

1. Haz clic en la celda **`H2`**.
2. Escribe la siguiente fórmula para construir el correo electrónico corporativo con el formato `nombre.apellido@empresa.com`:
   ```
   =CONCAT(MINUSC(C2),".",MINUSC(D2),"@empresa.com")
   ```
   - `MINUSC(C2)`: primer nombre en minúsculas.
   - `"."`: separador punto.
   - `MINUSC(D2)`: apellido en minúsculas.
   - `"@empresa.com"`: dominio corporativo fijo.
3. Presiona **Enter** y copia la fórmula hasta `H11`.

**Resultado esperado:**
- `H2` = `ana.lopez@empresa.com`
- `H3` = `carlos.mendez@empresa.com`
- `H5` = `valeria.nunez@empresa.com`

**Verificación:** Confirma que `H2` = `ana.lopez@empresa.com` (todo en minúsculas, con punto separador y dominio correcto).

---

#### Paso F7: Construir correos alternativos con UNIRCADENAS()

**Instrucciones:**

1. Haz clic en la celda **`I2`**.
2. Escribe la siguiente fórmula usando `UNIRCADENAS()` para lograr el mismo resultado de forma alternativa:
   ```
   =UNIRCADENAS(".",VERDADERO,MINUSC(C2),MINUSC(D2))&"@empresa.com"
   ```
   - `UNIRCADENAS(".", VERDADERO, ...)`: une los elementos con "." como separador, ignorando celdas vacías (`VERDADERO`).
   - `&"@empresa.com"`: concatena el dominio al final usando el operador `&`.
3. Presiona **Enter** y copia la fórmula hasta `I11`.

**Resultado esperado:** Los valores en la columna I deben ser idénticos a los de la columna H:
- `I2` = `ana.lopez@empresa.com`
- `I3` = `carlos.mendez@empresa.com`

**Verificación:**
- Confirma que `H2` e `I2` contienen exactamente el mismo texto.
- En una celda vacía, escribe `=H2=I2` y verifica que devuelve `VERDADERO`.

> **Concepto clave:** `CONCAT()` une elementos en el orden especificado sin separador automático (tú defines cada separador). `UNIRCADENAS()` permite definir un separador único que se aplica entre todos los elementos, lo que resulta más eficiente cuando tienes muchos campos a unir.

---

## 7. Validación y Pruebas

Al finalizar todos los módulos, realiza las siguientes verificaciones globales para confirmar que el libro está correcto:

### Lista de verificación final

| # | Verificación | Hoja | Resultado esperado |
|---|-------------|------|--------------------|
| 1 | `C4` en hoja Ventas contiene `=B4*$B$1` | Ventas | `10,000` |
| 2 | Al cambiar `B1` a `0.10`, todas las comisiones en C4:C8 se actualizan | Ventas | Proporcional |
| 3 | `K12` en Tabla_Multiplicar = 100 | Tabla_Multiplicar | `100` |
| 4 | `E7` en Tabla_Multiplicar contiene `=$A7*E$2` | Tabla_Multiplicar | `25` |
| 5 | `K13` en panel de resumen = 1,293,000 | Ventas | `1,293,000` |
| 6 | `K5` (teléfonos faltantes) = 2 | Empleados | `2` |
| 7 | Isabella Morales tiene "Bajo" en Nivel Desempeño | Empleados | `Bajo` |
| 8 | La lista de departamentos únicos contiene exactamente 4 valores | Empleados | `4` |
| 9 | `H2` en Texto_Empleados = `ana.lopez@empresa.com` | Texto_Empleados | Correcto |
| 10 | `H2` = `I2` (CONCAT y UNIRCADENAS producen el mismo resultado) | Texto_Empleados | `VERDADERO` |

### Prueba de integridad de referencias

1. En la hoja **Ventas**, cambia la tasa de comisión en `B1` a `0.12` (12%).
2. Verifica que **todas** las celdas en `C4:C8` se actualizan automáticamente.
3. Devuelve `B1` a `0.08`.

### Prueba de referencias estructuradas

1. Agrega un nuevo vendedor en la tabla `VentasMensuales` (última fila): `Luis Herrera` con ventas de `40000, 38000, 42000, 39000, 41000, 43000`.
2. Verifica que la columna `Total Vendedor` calcula automáticamente el total del nuevo registro.
3. Verifica que el panel de resumen (Total General, Promedio, MAX, MIN) se actualiza automáticamente.
4. Elimina el registro de Luis Herrera con **Ctrl+Z** al terminar.

---

## 8. Solución de Problemas

### Problema 1: La fórmula de comisión produce `#¡VALOR!` o resultados incorrectos al copiar

**Síntoma:** Al copiar la fórmula `=B4*$B$1` hacia abajo, algunas celdas muestran `#¡VALOR!`, `0` o un resultado inesperadamente diferente.

**Causa probable:** La celda `B1` contiene el valor `8%` como **texto** en lugar de como número, o la referencia absoluta no se escribió correctamente. Si `B1` contiene el texto `"8%"` en lugar del número `0.08`, Excel no puede multiplicarlo. Alternativamente, si se escribió `B$1` en lugar de `$B$1`, la columna no está fija y Excel buscará la tasa en columnas incorrectas al copiar hacia la derecha.

**Solución:**
1. Haz clic en `B1` y verifica en la barra de fórmulas: debe mostrar `0.08`, no `8%` como texto.
2. Si `B1` contiene texto, bórralo, escribe `0.08` y aplica formato porcentaje desde la cinta (**Inicio → Número → %**).
3. Haz clic en `C4` y verifica en la barra de fórmulas que la referencia dice exactamente `$B$1` con el signo `$` antes de la letra B **y** antes del número 1.
4. Si falta algún `$`, edita la celda (F2), selecciona `B1` en la fórmula y presiona **F4** hasta obtener `$B$1`.
5. Copia nuevamente la fórmula hacia abajo.

---

### Problema 2: `UNICOS()` o `ORDENAR()` devuelven `#¿NOMBRE?`

**Síntoma:** Al escribir `=ORDENAR(UNICOS(TablaEmpleados[Departamento]))` en la hoja Empleados, la celda muestra el error `#¿NOMBRE?` en lugar de la lista de departamentos.

**Causa probable:** La versión de Excel instalada **no es Microsoft 365** (puede ser Excel 2019, 2016 o una versión anterior), por lo que las funciones `UNICOS()` y `ORDENAR()` no existen en esa instalación. El error `#¿NOMBRE?` en Excel siempre indica que la función escrita no es reconocida por la versión instalada.

**Solución:**
1. Verifica tu versión de Excel: ve a **Archivo → Cuenta → Acerca de Excel**. Si no dice "Microsoft 365", estas funciones no están disponibles.
2. **Alternativa para versiones anteriores:** Para obtener valores únicos sin `UNICOS()`, puedes usar la funcionalidad de **Filtro avanzado** (pestaña Datos → Avanzadas → Copiar a otro lugar → Sólo registros únicos) o crear una tabla dinámica con el campo Departamento.
3. **Si tienes Microsoft 365 pero el error persiste:** Verifica que el nombre de la tabla sea exactamente `TablaEmpleados` (sin espacios ni caracteres adicionales). Ve a **Fórmulas → Administrador de nombres** para confirmar el nombre correcto y corrígelo en la fórmula si es necesario.
4. Consulta con tu instructor si necesitas la alternativa con funciones tradicionales.

---

## 9. Limpieza

Al finalizar la práctica, realiza los siguientes pasos para cerrar correctamente el trabajo:

1. **Guarda el archivo final:**
   - Presiona **Ctrl+S** o ve a **Archivo → Guardar**.
   - Confirma que el archivo se guarda con el nombre `Lab04_TuNombre.xlsx`.

2. **Verifica que todas las hojas estén completas:**
   - Revisa que las cinco pestañas (`Empleados`, `Ventas`, `Productos`, `Tabla_Multiplicar`, `Texto_Empleados`) existen y contienen datos.

3. **Elimina fórmulas de prueba temporales:**
   - Si creaste fórmulas de prueba como `=H2=I2` en celdas sueltas, elimínalas antes de entregar.

4. **Devuelve los valores modificados durante las pruebas:**
   - Confirma que la tasa de comisión en `B1` de la hoja Ventas es `0.08` (8%).
   - Confirma que no hay registros de prueba adicionales en las tablas.

5. **Entrega el archivo:**
   - Sube el archivo `Lab04_TuNombre.xlsx` a la carpeta designada en OneDrive según las instrucciones de tu instructor.

---

## 10. Resumen

En esta práctica construiste un libro de trabajo completo que integra las técnicas fundamentales de fórmulas y funciones intermedias de Excel 365:

| Módulo | Técnica aplicada | Concepto clave |
|--------|-----------------|----------------|
| **A** | Referencias relativas y absolutas | `$B$1` fija la tasa; `B4` se ajusta al copiar |
| **B** | Referencias mixtas | `$A3*B$2` para tabla bidimensional |
| **C** | Referencias estructuradas + SUMA, PROMEDIO, MAX, MIN | `[@[Ene]:[Jun]]` opera en la fila actual |
| **D** | CONTAR, CONTARA, CONTAR.BLANCO | Auditoría de completitud de datos |
| **E** | SI() simple y anidado + UNICOS + ORDENAR | Clasificación en 3 niveles; lista sin duplicados |
| **F** | IZQUIERDA, DERECHA, EXTRAE, LARGO, MAYUSC, MINUSC, CONCAT, UNIRCADENAS | Transformación y combinación de texto |

### Puntos clave para recordar

- **Referencia absoluta (`$B$1`):** Fija la celda completamente. Esencial para parámetros constantes (tasas, porcentajes, factores de conversión) que deben permanecer iguales al copiar la fórmula.
- **Referencia mixta (`$A3` o `B$2`):** Fija solo una dimensión. Indispensable para tablas bidimensionales donde cada celda combina un valor de su fila con un valor de su columna.
- **Referencias estructuradas:** Hacen las fórmulas más legibles y robustas. `[@Ventas]` siempre apunta a la columna correcta, incluso si se insertan columnas nuevas en la tabla.
- **CONTAR vs CONTARA:** `CONTAR()` solo cuenta números; `CONTARA()` cuenta cualquier contenido. Usa ambas estratégicamente para auditar datos.
- **SI() anidado:** Permite múltiples categorías evaluando condiciones en cascada. Evalúa de mayor a menor (o menor a mayor) para evitar solapamientos.
- **UNICOS() + ORDENAR():** Combinación poderosa para generar catálogos dinámicos que se actualizan automáticamente cuando los datos fuente cambian.
- **CONCAT vs UNIRCADENAS:** `CONCAT()` une elementos con separadores definidos manualmente; `UNIRCADENAS()` aplica un separador uniforme entre todos los elementos de forma más eficiente.

### Tabla de equivalencias de funciones (Español → Inglés)

| Función en Español    | Función en Inglés  |
|-----------------------|--------------------|
| `SUMA()`              | `SUM()`            |
| `PROMEDIO()`          | `AVERAGE()`        |
| `MAX()`               | `MAX()`            |
| `MIN()`               | `MIN()`            |
| `CONTAR()`            | `COUNT()`          |
| `CONTARA()`           | `COUNTA()`         |
| `CONTAR.BLANCO()`     | `COUNTBLANK()`     |
| `SI()`                | `IF()`             |
| `UNICOS()`            | `UNIQUE()`         |
| `ORDENAR()`           | `SORT()`           |
| `IZQUIERDA()`         | `LEFT()`           |
| `DERECHA()`           | `RIGHT()`          |
| `EXTRAE()`            | `MID()`            |
| `LARGO()`             | `LEN()`            |
| `MAYUSC()`            | `UPPER()`          |
| `MINUSC()`            | `LOWER()`          |
| `CONCAT()`            | `CONCAT()`         |
| `UNIRCADENAS()`       | `TEXTJOIN()`       |
| `HALLAR()`            | `FIND()`           |
| `VERDADERO`           | `TRUE`             |

### Recursos adicionales

- [Microsoft Support: Cambiar entre referencias relativas, absolutas y mixtas](https://support.microsoft.com/es-es/office/cambiar-entre-referencias-relativas-absolutas-y-mixtas-dfec08cd-ae65-4f56-839e-5f0d8d0baca9)
- [Microsoft Support: Función SI](https://support.microsoft.com/es-es/office/funci%C3%B3n-si-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2)
- [Microsoft Support: Función UNIRCADENAS](https://support.microsoft.com/es-es/office/funci%C3%B3n-unircadenas-357b449a-ec91-49d0-80c3-0e8fc845691c)
- [Microsoft Support: Uso de referencias estructuradas con tablas de Excel](https://support.microsoft.com/es-es/office/uso-de-referencias-estructuradas-con-tablas-de-excel-f5ed2452-2337-4f71-bed3-c8ae6d2b276e)
- [ExcelJet: UNIQUE function](https://exceljet.net/functions/unique-function)

---
