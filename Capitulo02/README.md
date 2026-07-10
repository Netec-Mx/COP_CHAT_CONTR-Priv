# Laboratorio 2: Diseño de flujos de trabajo para cruzar cuentas contables, identificar discrepancias automáticamente y generar reportes de excepción utilizando IA

## 1. Metadatos del Laboratorio

| Atributo | Detalle |
| :--- | :--- |
| **Duración** | 90 minutos |
| **Complejidad** | Básica-Media |
| **Audiencia** | Profesionales de Finanzas, Contabilidad y Auditoría |
| **Tecnologías** | Microsoft Excel (Escritorio o Web) y Microsoft Copilot Chat (M365) |
| **Enfoque** | Auditoría Financiera, Cruce de Datos y Fórmulas Automatizadas sin Macros |

---

## 2. Descripción 

En este laboratorio práctico de 90 minutos, aprenderás a utilizar **Copilot Chat** para diseñar un modelo de conciliación contable de forma autosuficiente directamente en la cuadrícula de Microsoft Excel. Guiado por la IA, estructurarás datos de prueba e implementarás fórmulas de verificación cruzada y formatos condicionales para detectar anomalías al instante, sin necesidad de usar código ni macros.

---

## 3. Objetivos

Al completar este laboratorio, serás capaz de:
* **Generar escenarios financieros de prueba** usando estructuras Markdown reconocibles por Excel.
* **Construir lógica de auditoría cruzada** mediante el uso de fórmulas lógicas simplificadas (`SI` y `CONTAR.SI`).
* **Implementar alertas visuales automáticas** para aislar discrepancias de montos o registros faltantes.
* **Iterar soluciones con IA** para adaptar los modelos a las reglas del negocio contable.

---

## 4. Prerrequisitos

* Cuenta activa de **Microsoft 365** con acceso a **Copilot Chat**.
* Aplicación de Microsoft Excel (Escritorio o Web).
* Conocimientos básicos de cómo copiar, pegar y arrastrar una fórmula en Excel.

---

## 5. Entorno de Laboratorio

* Aplicación de Microsoft Excel (Escritorio o versión Web).
* Panel de Copilot Chat abierto a un costado del libro de trabajo para interactuar en un solo hilo conversacional.
* Un libro de Excel nuevo y limpio guardado con el nombre `Laboratorio_Conciliacion.xlsx`.

---

## 6. Procedimiento Paso a Paso

### Sección A: Generación Autosuficiente de Datos Contables de Prueba 

El primer paso consiste en poblar nuestro libro de Excel con un escenario financiero simulado de manera ágil. Le pediremos a Copilot Chat que actúe como un diseñador de datos contables para generar dos estructuras de tabla independientes en la misma hoja de cálculo, inyectando discrepancias intencionales que auditaremos más adelante.

1. Abra su archivo `Laboratorio_Conciliacion.xlsx`.
2. Renombre la pestaña actual con el nombre **"Panel_Conciliacion"**.
3. Abra el panel lateral de **Copilot Chat**.
4. Copie, pegue y ejecute el siguiente prompt estructurado en la caja de chat:

```
Actúa como un experto en auditoría financiera y análisis de datos en Excel. Necesito generar un escenario de simulación de conciliación contable dentro de una sola hoja de Excel. 

Para garantizar que los datos se peguen correctamente en columnas y filas separadas en Excel, por favor proporcióname la información estrictamente en formato de TABLAS MARKDOWN (utilizando barras verticales '|' y guiones para los encabezados). Genérame dos tablas independientes separadas:

Tabla 1 (Para iniciar en la celda B3):
| ID_Transaccion | Fecha | Concepto_Empresa | Monto_Empresa |
|---|---|---|---|
| 101 | 2026-06-01 | Pago proveedor A | 1250.00 |
| 102 | 2026-06-02 | Compra insumos | 450.75 |
| 103 | 2026-06-03 | Venta producto X | 3100.00 |
| 104 | 2026-06-04 | Pago servicios | 890.00 |
| 105 | 2026-06-05 | Cobro cliente Y | 1500.00 |

Tabla 2 (Para iniciar en la celda G3):
| ID_Banco | Fecha_Banco | Detalle_Banco | Monto_Banco |
|---|---|---|---|
| 101 | 2026-06-01 | Pago proveedor A | 1250.00 |
| 102 | 2026-06-02 | Compra insumos | 450.75 |
| 103 | 2026-06-03 | Venta producto X | 3100.00 |
| 104 | 2026-06-04 | Pago servicios | 89.00 |
| 106 | 2026-06-06 | Depósito cliente Z | 2000.00 |

Muestra las tablas directamente en la pantalla del chat (no uses bloques de código ni texto plano). Necesito que se rendericen como tablas visuales para que, al seleccionarlas con el mouse y arrastrar, Excel reconozca las filas y columnas nativamente al presionar Ctrl+C y Ctrl+V.
```

5. Seleccione con el cursor del mouse la primera tabla renderizada visualmente en la respuesta de Copilot (la de la Empresa).
6. Copie la selección (Ctrl+C), sitúese en la celda **B3** de la pestaña "Panel_Conciliacion" y péguelo (Ctrl+V). Compruebe que cada dato quede de inmediato en su propia columna (Rango B3:E8).
7. Repita el proceso de selección con el mouse para la segunda tabla renderizada (la del Banco).
8. Copie la selección (Ctrl+C), sitúese en la celda **G3** de Excel y péguelo (Ctrl+V). Compruebe que se separe perfectamente en las celdas (Rango G3:J8), dejando la columna F vacía en el medio.

---

### Sección B: Construcción del Motor de Conciliación mediante Fórmulas Simples

Una vez estructurados los datos directamente en las celdas, utilizaremos una fórmula lógica directa en español para cruzar la información sin riesgo de errores. Usaremos el punto y coma (`;`) como separador para garantizar que sea aceptada por Excel de inmediato.

1. Asegúrese de estar en el mismo hilo de conversación con Copilot Chat para mantener el contexto.
2. Copie, pegue y ejecute el siguiente prompt estructurado:

```
Actúa como un analista de datos senior experto en Excel financiero. Basándote en las dos tablas que acabamos de colocar en la hoja (Tabla 1 de Empresa en el rango B3:E8 y Tabla 2 de Banco en el rango G3:J8), necesito crear una columna de validación al final de cada una. 

Por favor, proporcióname las instrucciones y las dos fórmulas exactas en español, configuradas estrictamente con PUNTO Y COMA (;) como separador de argumentos:

1. Para la Tabla 1 de la Empresa: Añadiremos el encabezado "Estado_Conciliacion" en la celda F3. Necesito una fórmula para la celda F4 (para arrastrar hasta F8) que busque el ID_Transaccion (B4) en la columna de IDs del Banco ($G$4:$G$8). Si no lo encuentra, debe mostrar "No Encontrado en Banco". Si sí lo encuentra, debe evaluar si el Monto_Empresa (E4) es igual al Monto_Banco (J4). Si son iguales, mostrará "Conciliado", y si son diferentes, mostrará "Diferencia de Monto".

2. Para la Tabla 2 del Banco: Añadiremos el encabezado "Verificacion_Banco" en la celda K3. Dame la fórmula espejo para colocar en K4 (para arrastrar hasta K8) que busque el ID_Banco (G4) en los IDs de la Empresa ($B$4:$B$8) y aplique la misma lógica comparando el Monto_Banco (J4) con el Monto_Empresa (E4).
```

3. Aplique las fórmulas simplificadas y directas en su hoja de cálculo:
   - Haga clic en la celda **F4**, copie y pegue la siguiente fórmula exacta:
     `=SI(CONTAR.SI($G$4:$G$8;B4)=0;"No Encontrado en Banco";SI(E4=J4;"Conciliado";"Diferencia de Monto"))`
   - Arrastre el controlador de relleno desde F4 hacia abajo hasta la celda F8.

   - Haga clic en la celda **K4**, copie y pegue la siguiente fórmula espejo exacta:
     `=SI(CONTAR.SI($B$4:$B$8;G4)=0;"No Encontrado en Empresa";SI(J4=E4;"Conciliado";"Diferencia de Monto"))`
   - Arrastre el controlador de relleno desde K4 hacia abajo hasta la celda K8.

4. Valide que los estados se calculen instantáneamente: las filas 101 a 103 mostrarán "Conciliado", la fila 104 mostrará "Diferencia de Monto" y las filas 105 y 106 marcarán sus respectivos estados de no encontrado sin generar ninguna alerta de error.

---

### Sección C: Formato Condicional y Reto de Adaptación 

Para facilitar la auditoría visual, utilizaremos las herramientas de Excel para destacar los errores. Posteriormente, resolverás un reto práctico.

1. En el chat de Copilot, ejecute el siguiente prompt para recibir la guía visual:

```
Actúa como diseñador de dashboards en Excel. Necesito resaltar visualmente las columnas de estado que acabo de crear (Columnas E e I). Dime los pasos exactos en los menús de Excel para aplicar tres reglas de Formato Condicional basadas en el texto de la celda:
1. Si el texto es "Conciliado", aplicar relleno verde claro.
2. Si el texto es "Error de Monto", aplicar relleno rojo claro.
3. Si el texto es "No Encontrado", aplicar relleno amarillo claro.
```

2. Siga los pasos indicados por Copilot sobre los rangos E4:E8 e I4:I8 para colorear su conciliación.

---

### RECUADRO DEL RETO DE ADAPTACIÓN 

Para medir tu comprensión en este capítulo, ejecuta el siguiente desafío rápido directamente con tu asistente de IA.

3. Copie, adapte y ejecute este prompt en su chat para añadir una métrica clave de control en su hoja:

```
Actúa como un analista financiero. Ya tengo mi modelo de conciliación operando con estados y colores en las columnas F y K. Ahora, necesito agregar un indicador rápido en la parte inferior de la hoja para presentárselo a mi supervisor.

Por favor, indícame cómo lograr lo siguiente:
1. En la celda B11 escribiré el texto: "Total Transacciones Conciliadas".
2. Proporcióname una fórmula muy sencilla utilizando CONTAR.SI para colocar en la celda C11, que cuente cuántas veces aparece la palabra "Conciliado" en la columna de la Empresa (Rango F4:F8).

Dame la fórmula en español lista para copiar y pegar utilizando punto y coma (;) como separador.
```

4. Analice la propuesta de Copilot, implemente la fórmula en la celda **C11** y verifique que el sistema devuelva el número exacto de transacciones exitosas de forma autosuficiente.

---

## 7. Conceptos Clave para Recordar

* **Uso del Separador Regional:** En sistemas operativos configurados en español, el punto y coma (`;`) es el estándar obligatorio para separar argumentos en las fórmulas. Conocer esto evita el 90% de los errores al interactuar con Copilot.
* **Lógica Avanzada Simplificada:** Combinar funciones básicas como `SI` y `CONTAR.SI` permite estructurar auditorías cruzadas potentes, seguras y altamente legibles sin complicar el mantenimiento del archivo.
* **Modelos Autosuficientes:** Al diseñar hojas estructuradas con fórmulas nativas dinámicas, el analista financiero asegura que el archivo funcione en cualquier dispositivo (Excel Escritorio, Web, Teams) de forma automática.

---

## 8. Resultado Esperado

Al finalizar, tu archivo `Laboratorio_Conciliacion.xlsx` mostrará en la pestaña **"Panel_Conciliacion"** un sistema automatizado de control donde:
1. Los registros duplicados idénticos se validan en color **Verde** ("Conciliado").
2. Las diferencias de digitación numérica (ID 104) saltan inmediatamente en color **Rojo** ("Error de Monto").
3. Las partidas pendientes de cobro o depósitos en tránsito (IDs 105 y 106) se aíslan en color **Amarillo** ("No Encontrado").
