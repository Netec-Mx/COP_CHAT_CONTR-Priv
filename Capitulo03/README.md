# Laboratorio 3: Estrategias para integrar datos extraídos de Oracle con funciones de Copilot en Excel. Limpieza, transformación y estructuración de grandes volúmenes de datos financieros.

## 1. Metadatos del Laboratorio

| Atributo | Detalle |
| :--- | :--- |
| **Duración** | 90 minutos |
| **Complejidad** | Avanzada (Enfoque Estratégico y de Arquitectura de Datos) |
| **Audiencia** | Líderes de Proyectos BI, Auditores, Contadores y Analistas Financieros Senior |
| **Tecnologías** | Microsoft Copilot Chat (M365) y Microsoft Excel (Escritorio o Web) |
| **Enfoque** | Integración de ERPs, Data Cleansing asistido por IA y Gobierno de Datos |

---

## 2. Descripción 

Este laboratorio entrena a los profesionales en las estrategias clave para integrar extracciones masivas e incompatibles de **Oracle** dentro de entornos analíticos utilizando **Microsoft Copilot**. El diseño del ejercicio elimina por completo la necesidad de escribir fórmulas manuales en Excel. El estudiante asume el rol de Orquestador de Datos, forzando a la IA a actuar como un motor ETL (Extracción, Transformación y Carga) dentro del chat. El resultado de cada fase se inyecta directamente en Excel mediante copiado y pegado directo, culminando con una auditoría analítica del flujo de caja.

---

## 3. Objetivos

Al finalizar este laboratorio, el estudiante será capaz de:
* **Aplicar estrategias de integración** para procesar volcados de datos planos e impuros de Oracle ERP sin depender de macros ni herramientas externas de desarrollo.
* **Orquestar procesos de Data Cleansing en la interfaz de la IA** para estandarizar formatos monetarios bloqueados, limpiar texto descriptivo e indexar identificadores.
* **Estructurar un modelo de datos financiero limpio en Excel** utilizando exclusivamente flujos de copiado directo de tablas Markdown nativas renderizadas por el chat.
* **Generar reportes de gobierno contable de nivel ejecutivo**, utilizando a Copilot para auditar saldos consolidados y detectar riesgos en la cadena de suministro.

---

## 4. Prerrequisitos

* Cuenta activa de **Microsoft 365** con acceso a **Copilot Chat**.
* Aplicación de Microsoft Excel abierta con un libro nuevo guardado como `Estrategia_Oracle_Copilot.xlsx`.

---

## 5. Procedimiento Paso a Paso 

### Fase A: Inyección del Reporte Contable Crudo de Oracle 

El primer paso es simular una extracción real, masiva y descuidada de un libro mayor de Oracle ERP. El texto vendrá contaminado con códigos de sistema e importes con texto de divisa ("USD") que bloquean los cálculos ordinarios.

1. En su libro de Excel, renombre la primera pestaña como **"1. Reporte_Crudo_Oracle"**.
2. Abra su panel lateral de **Copilot Chat** e introduzca el siguiente prompt estratégico para generar la base del caso de estudio:

```
Actúa como un Administrador de Base de Datos y Líder de Sistemas Financieros en Oracle ERP. Necesito simular un volcado masivo en bruto de asientos contables correspondientes a las operaciones de la organización del mes de junio de 2026. 

Por favor, proporcióname una TABLA MARKDOWN única (usando barras '|' y guiones) que contenga exactamente 25 filas de datos. Es un requisito crítico de interoperabilidad que la tabla se renderice como una estructura visual nativa directamente en la pantalla de este chat (no uses bloques de código gris ni texto plano) para que pueda seleccionarla con el mouse y pegarla en Excel de inmediato. 

Usa esta estructura exacta y varía los datos fila por fila manteniendo la coherencia financiera:
| ORACLE_REF_ID | ISO_DATE | TRANSACTION_DESC | RAW_AMOUNT |
|---|---|---|---|
| ID_ERP_GL_1001_2026 | 2026-06-01 | [PROV_NORT] Compra Suministros #991 | USD 14500.50 |
| ID_ERP_GL_1002_2026 | 2026-06-01 | [CORP_SUD] Reembolso Viáticos Mtz | USD -320.15 |

Condiciones para las 25 filas:
1. Incrementa el ID secuencialmente desde 1001 hasta 1025.
2. Intercala las fechas entre el 01/06/2026 y el 10/06/2026.
3. Distribuye equitativamente las etiquetas de control: [PROV_NORT], [CORP_SUD], [REV_INT] (Ingresos por intereses) y [PROV_ANOM] (Cuentas anómalas por auditar).
4. Todos los importes en 'RAW_AMOUNT' deben llevar el prefijo 'USD ' y alternar números positivos altos y flujos negativos.
```

3. Seleccione con el mouse toda la tabla visual que Copilot desplegó en el chat.
4. Cópiela (Ctrl+C), sitúese en la celda **B3** de la pestaña "1. Reporte_Crudo_Oracle" y péguela (Ctrl+V).
5. Valide que los datos se hayan distribuido correctamente en sus celdas. Esta es la base "sucia" del sistema legado.

---

### Fase B: Transformación Masiva de Datos 

En un flujo tradicional, el analista pasaría horas escribiendo fórmulas de texto en Excel para limpiar las impurezas. En este laboratorio de IA, obligaremos a Copilot a realizar el proceso de ETL (Extracción, Transformación y Carga) dentro de su propia memoria y a entregarnos la matriz 100% limpia.

1. Cree una segunda pestaña en su libro de Excel y llámela **"2. Maestro_Datos_Limpio"**.
2. Vaya al chat de Copilot y ejecute el siguiente prompt de ingeniería de datos para procesar el texto:

```
Actúa como un Ingeniero de Datos Financieros y especialista en Business Intelligence. Necesito que proceses la tabla de 25 filas que acabas de generar en el turno anterior y apliques una estrategia de transformación masiva (Data Cleansing masivo).

Requiero que realices las siguientes transformaciones sobre las 25 filas en tu memoria:
1. ID_Limpio: Remueve los prefijos y sufijos 'ID_ERP_GL_' y '_2026' dejando solo el número limpio (ej. 1001).
2. Fecha_Estandar: Convierte la fecha ISO al formato estándar de visualización DD/MM/AAAA.
3. Categoria: Extrae únicamente el texto que se encuentra encerrado dentro de los corchetes (ej. PROV_NORT).
4. Descripcion_Operacion: Quita el bloque de corchetes inicial y deja solo el texto descriptivo de la transacción.
5. Importe_Numerico: Elimina de raíz la palabra 'USD ' de la columna de montos. Asegúrate de conservar el signo negativo si el flujo es una salida de dinero y entrega el número limpio con sus decimales (ej. 14500.50).

Por favor, devuélveme las 25 filas transformadas estructuradas en una NUEVA TABLA MARKDOWN VISUAL NATIVA en la pantalla del chat con las columnas corregidas: | ID_Limpio | Fecha_Estandar | Categoria | Descripcion_Operacion | Importe_Numerico |. No uses bloques de código plano, necesito la tabla visual directa para copiar y pegar en mi hoja.
```

3. Seleccione con el cursor del mouse la nueva tabla limpia generada por Copilot en la pantalla del chat.
4. Cópiela (Ctrl+C), vaya a la pestaña **"2. Maestro_Datos_Limpio"** de Excel, haga clic en la celda **B3** y péguela (Ctrl+V).
5. Seleccione la columna F (`Importe_Numerico`) y aplíquele el formato de **Moneda ($)** desde la barra de inicio de Excel. Ahora tiene una base de datos maestra perfecta, sin haber escrito una sola fórmula manual.

---

### Fase C: Agregación Contable y Dashboard Ejecutivo 

Procederemos a construir el cuadro de mando gerencial. En lugar de usar complejas funciones condicionales en la hoja, le ordenaremos a Copilot que realice los cálculos matemáticos del consolidado basándose en la información que ya conoce.

1. Cree una tercera pestaña en su libro de Excel y llámela **"3. Dashboard_Financiero"**.
2. Vaya al chat de Copilot e introduzca el siguiente prompt analítico:

```
Actúa como un Arquitecto de Business Intelligence y Director Financiero. Basándote en la tabla limpia de 25 filas que acabamos de estructurar, necesito que realices un procesamiento matemático de agregación de datos para construir un resumen ejecutivo.

Calcula dentro del chat las siguientes métricas exactas utilizando los valores de la columna 'Importe_Numerico':
1. Total de Entradas de Efectivo: La suma de todos los montos que son mayores a cero.
2. Total de Salidas de Efectivo (Egresos): La suma de todos los montos que son menores a cero.
3. Balance Neto Operativo: El resultado neto total del mes (Entradas menos Salidas).
4. Cantidad de transacciones registradas bajo la categoría de riesgo [PROV_ANOM].

Preséntame estos 4 resultados consolidados estructurados en una pequeña TABLA MARKDOWN VISUAL NATIVA con dos columnas: | Indicador Financiero | Saldo Consolidado |. Asegúrate de incluir el cálculo real basado en los datos de las 25 filas anteriores.
```

3. Copie la tabla de indicadores devuelta por Copilot Chat.
4. Vaya a la pestaña **"3. Dashboard_Financiero"** en Excel, sitúese en la celda **B3** y pegue (Ctrl+V).
5. Aplique un formato visual profesional a las celdas (negritas a los encabezados y formato de moneda a los resultados). El cuadro de mando gerencial está completo utilizando la IA como motor de cálculo.

---

### Fase D: Auditoría Forense y Reporte de Gobierno 

La última fase del laboratorio consolida el perfil estratégico del estudiante. Utilizaremos la capacidad analítica avanzada de Copilot para interpretar el impacto del balance neto y redactar un reporte de riesgos corporativos para la junta directiva.

1. Ejecute el siguiente prompt avanzado en Copilot Chat para activar la fase de consultoría:

```
Actúa como un Auditor de Control Interno y Consultor Estratégico Senior. Basándote en el Balance Neto Operativo y en el volumen de operaciones detectadas para la categoría [PROV_ANOM] que calculaste en el paso anterior, necesito que redactes un informe técnico de auditoría dirigido al Comité de Gobierno Corporativo.

El informe debe constar de 3 párrafos formales con la siguiente estructura:
1. Diagnóstico de Liquidez: Evaluación de la salud financiera de la empresa según el balance neto del periodo (si es positivo o si hay déficit operativo).
2. Alerta de Mitigación de Riesgos: Un análisis profundo sobre el peligro de mantener transacciones bajo el rubro [PROV_ANOM] (Cuentas Anómalas) sin conciliación fiscal y su impacto en auditorías externas.
3. Plan de Acción de Gobierno: Recomendaciones de control interno específicas (ej. congelamiento de cuentas, auditoría retroactiva al trimestre o implementación de políticas de cumplimiento).

Usa un lenguaje corporativo de alta dirección, sofisticado y preciso.
```

2. Analice detenidamente el texto estructurado por la IA.
3. Copie el informe de auditoría generado por Copilot, regrese a la pestaña **"3. Dashboard_Financiero"** de Excel y péguelo directamente en un cuadro de texto o en la celda **B10**. 
4. El entregable final está listo: un artefacto financiero que integra la extracción, la transformación masiva (ETL), la analítica de indicadores y el gobierno corporativo en un flujo puro de IA.

---

### Fase E: Reto de Aplicación Autónoma – Auditoría de Brechas Presupuestarias 

**Instrucciones para el estudiante:** Hasta este punto del laboratorio, has actuado de forma guiada. Ahora te enfrentarás a un reto de negocio real de manera independiente. La junta directiva te acaba de entregar la tabla oficial de "Presupuesto Máximo Autorizado" para egresos (salidas de dinero) del mes de junio de 2026:

| Categoria | Presupuesto_Maximo_Egresos |
| :--- | :--- |
| PROV_NORT | $ 30,000.00 |
| CORP_SUD | $ 15,000.00 |
| PROV_ANOM | $ 0.00 |

#### El Desafío Financiero:
Debes interactuar de forma autónoma con Copilot para obligarlo a realizar un **análisis de brechas (Gap Analysis)**. La IA debe consolidar los egresos reales de tus 25 filas por cada categoría, cruzarlos contra esta nueva tabla de presupuestos autorizados, identificar si alguna categoría sufrió un desborde presupuestario (overbudget) y calcular el monto exacto de la desviación. 

**Regla de Oro:** No puedes calcular nada a mano ni escribir fórmulas en Excel. Excel solo recibirá los bloques que logres extraer de la IA mediante copiado y pegado directo.

#### Pistas de Ingeniería de Prompts para el Éxito:
Para resolver este reto, debes estructurar un prompt avanzado en el chat utilizando las siguientes estrategias:
* **Asignación de Rol:** Inicia tu mensaje forzando a la IA a actuar como un *Auditor Forense y Controller de Presupuesto Corporativo*.
* **Inyección de Contexto:** Declara explícitamente en tu prompt la tabla de presupuestos autorizados que te acabamos de dar arriba.
* **Instrucción de Cruce Relacional:** Pídele que recuerde las 25 filas de datos limpios de la Fase B, que agrupe y sume únicamente los valores negativos (egresos) de cada categoría, y que cree una nueva tabla comparativa.
* **Restricción de Formato:** Exígele que el resultado final sea una nueva **Tabla Markdown Visual Nativa** que incluya las columnas: `| Categoria | Total_Egresos_Reales | Presupuesto_Autorizado | Desviación (Diferencia) | Estado (Aprobado / Alerta de Desvío) |`.

#### Pasos para cerrar el laboratorio una vez que Copilot te entregue el resultado:
1. Copia la tabla de desvíos presupuestarios calculada por la IA y pégala en la celda **B18** de la pestaña **"3. Dashboard_Financiero"** (debajo del informe de la Fase D).
2. Inserta un nuevo prompt pidiéndole a Copilot que redacte un *Memorando Ejecutivo de Hallazgos Críticos* basado en las desviaciones encontradas (especialmente analizando el comportamiento de `PROV_ANOM`).
3. Copia ese texto de alta dirección y pégalo en la celda **B26** de tu dashboard.

---

## 6. Conceptos Clave para Recordar

* **Estrategia ETL en el Chat:** La capacidad de procesar transformaciones complejas (limpieza de cadenas de texto y conversiones numéricas) dentro de la interfaz del LLM optimiza los tiempos de desarrollo y mitiga los errores humanos asociados al ingreso de fórmulas manuales en las hojas de cálculo.
* **Interoperabilidad sin Fricciones:** El renderizado nativo de tablas Markdown en pantallas de chat permite un puente limpio de datos hacia Excel, asegurando que las columnas y filas se mantengan alineadas de forma predecible sin importar la configuración regional del sistema operativo del usuario.
* **Análisis Inverso (Data to Insights):** Los entornos analíticos modernos combinan la organización de datos numéricos en hojas de cálculo con la capacidad de reinyectar esos resúmenes en motores de IA para generar narrativas de riesgos, planes de mitigación y memorandos ejecutivos para la toma de decisiones estratégicas.

---

## 7. Resultado Esperado

Al concluir la sesión práctica, el estudiante habrá completado y guardado el archivo `Estrategia_Oracle_Copilot.xlsx`, el cual deberá estructurarse de la siguiente manera:

1. **Pestaña "1. Reporte_Crudo_Oracle":** Contendrá la base semilla original de 25 filas con los textos e importes sucios extraídos directamente de los servidores de Oracle (Rango B3:E28).
2. **Pestaña "2. Maestro_Datos_Limpio":** Mostrará la matriz de datos maestros completamente depurada por el motor ETL de Copilot (Rango B3:F28). Allá las columnas serán legibles, sin prefijos y los importes serán numéricos operables.
3. **Pestaña "3. Dashboard_Financiero":** Presentará:
   * **Rango B3:C7:** El cuadro macro consolidado generado en la Fase C (Entradas, Salidas, Balance Neto).
   * **Rango B10:B15:** El informe ejecutivo original de auditoría e interpretación de la Fase D.
   * **Rango B18:F22:** La tabla del **Reto de Aplicación Autónoma (Fase E)** mostrando el análisis comparativo de egresos vs presupuesto y el estatus de alerta por categoría.
   * **A partir de la celda B26:** El memorando de hallazgos críticos final redactado por la IA sobre las categorías sobrepasadas y el plan de acción de gobierno corporativo.
