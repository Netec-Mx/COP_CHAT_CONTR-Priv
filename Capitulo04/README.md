# Laboratorio 4: Uso de Copilot para la revisión de evidencias, detección de anomalías en registros contables y generación de resúmenes ejecutivos para la dirección financiera.

## 1. Metadatos del Laboratorio

| Atributo | Detalle |
| :--- | :--- |
| **Duración** | 60 minutos  |
| **Complejidad** | Avanzada - Forense |
| **Audiencia** | Líderes de Proyectos BI, Auditores, Contadores y Analistas Financieros Senior |
| **Tecnologías** | Microsoft Copilot Chat (M365) y Microsoft Excel (Escritorio o Web) |
| **Enfoque** | Uso de Copilot para la revisión de evidencias, detección de anomalías en registros contables y generación de resúmenes ejecutivos para la dirección financiera. |

---

## 2. Descripción 

Este laboratorio práctico entrena a los profesionales en el despliegue de prompts avanzados para automatizar la auditoría de registros contables y la validación de sus evidencias documentales asociadas. Utilizando las capacidades de razonamiento conceptual de Microsoft Copilot, los estudiantes estructurarán un entorno de control sin necesidad de programar macros ni escribir fórmulas complejas en Excel. El ejercicio se centra en identificar discrepancias normativas, alertar sobre comportamientos sospechosos en transacciones financieras y compilar resúmenes de hallazgos listos para ser presentados ante la Alta Dirección.

---

## 3. Objetivos

Al finalizar este laboratorio, el estudiante será capaz de:
* **Evaluar registros contables complejos** utilizando la IA como un motor de revisión de evidencias basado en reglas de control interno.
* **Detectar anomalías financieras y patrones atípicos** (duplicidades, saltos de secuencia, inconsistencias en descripciones) mediante ingeniería de prompts.
* **Sintetizar hallazgos contables críticos**, delegando en Copilot la redacción técnica de resúmenes ejecutivos con impacto financiero directo.
* **Tomar decisiones de gobierno analítico**, consolidando pistas de auditoría e informes de cumplimiento en una estación de trabajo limpia en Excel.

---

## 4. Prerrequisitos

* Cuenta activa de **Microsoft 365** con acceso a **Copilot Chat**.
* Aplicación de Microsoft Excel abierta con un libro nuevo guardado como `Auditoria_Evidencias_Copilot.xlsx`.

---

## 5. Procedimiento Paso a Paso (Flujo Basado en Respuestas de IA)

### Fase A: Inyección del Libro Diario y Evidencias de Control
El laboratorio inicia construyendo un escenario simulado donde se cruzan las transacciones del libro diario con el estado de su documentación de soporte (facturas, aprobaciones, firmas digitales). Forzaremos a Copilot a renderizar una matriz con inconsistencias intencionales para su posterior análisis.

1. En su libro de Excel, renombre la primera pestaña como **"1. Registro_y_Evidencias"**.
2. Vaya al chat de **Copilot** e introduzca el siguiente prompt corporativo:

```
Actúa como un Auditor Contable Forense. Necesito generar una tabla de control que cruce un libro diario financiero con el estado de sus evidencias documentales del mes de junio de 2026.

Por favor, proporcióname una TABLA MARKDOWN única (usando barras '|' y guiones) que contenga exactamente 15 filas de datos. Es un requisito obligatorio que la tabla se muestre renderizada visualmente en la pantalla del chat (no uses bloques de código gris) para poder copiarla y pegarla directamente en Excel. 

Usa esta estructura de columnas exacta y varía los datos fila por fila:
| REGISTRO_ID | FECHA | CUENTA_CONTABLE | IMPORTE | ESTADO_EVIDENCIA | RESPONSABLE_APROBACIÓN |
|---|---|---|---|---|---|
| REG-2026-001 | 05/06/2026 | 520595 (Diversos) | $ 8,500.00 | Factura digital validada y cargada | Ing. Carlos Mendoza |
| REG-2026-002 | 05/06/2026 | 110505 (Caja General) | $ 12,000.00 | Recibo provisional físico (Sin XML) | Sin Firma Autorizada |

Condiciones críticas para las 15 filas:
1. Incrementa el REGISTRO_ID secuencialmente de 001 a 015.
2. Intercala cuentas de Gastos (52), Activos (11) y Pasivos (21).
3. Asegúrate de incluir de forma oculta 3 filas con anomalías severas: una con un importe inusualmente alto bajo la descripción 'Pendiente por legalizar', otra con el estado de evidencia 'Documento borroso / No legible', y una tercera fila que repita exactamente el mismo importe y cuenta en una fecha consecutiva (duplicado sospechoso).
```

3. Seleccione con el cursor la tabla markdown renderizada por Copilot.
4. Cópiela (`Ctrl+C`), sitúese en la celda **B3** de la pestaña "1. Registro_y_Evidencias" en Excel y péguela (`Ctrl+V`).
5. Dé un formato rápido de cuadrícula en Excel para asegurar la visualización limpia de los 15 registros.

---

### Fase B: Procesamiento Forense y Detección Automática de Anomalías
En esta fase, delegaremos la auditoría analítica en la memoria de Copilot. El objetivo es que la IA actúe como un filtro de cumplimiento inteligente, escaneando la matriz, aislando los riesgos y categorizándolos sin requerir condicionales manuales en Excel.

1. Cree una segunda pestaña en su archivo de Excel y llámela **"2. Hallazgos_Auditoria"**.
2. Regrese al chat de Copilot e introduzca el siguiente prompt estructurado de detección:

```
Actúa como un Auditor Interno de Control de Riesgos Financieros. Necesito que analices minuciosamente la tabla de 15 filas que generamos en el paso anterior. Tu objetivo es detectar de forma automática todas las anomalías, inconsistencias normativas y fallas en la revisión de evidencias documentales.

Por favor, devuélveme una NUEVA TABLA MARKDOWN VISUAL NATIVA en el chat que extraiga únicamente las filas que presenten riesgos potenciales (mínimo 3 o 4 registros detectados). La tabla debe estructurarse con las siguientes columnas:
| ID_Anomalía | Cuenta_Afectada | Importe_Riesgo | Descripción_del_Hallazgo | Nivel_de_Riesgo (Alto / Medio) | Acción_Inmediata_Sugerida |

Reglas analíticas:
- Clasifica como 'Riesgo Alto' la ausencia de firmas autorizadas, los documentos no legibles o los montos sin XML válido.
- Clasifica como 'Riesgo Medio' las duplicidades de montos en cuentas operativas que requieran conciliación complementaria.
No uses bloques de texto plano, genera la tabla visual de manera directa.
```

3. Con el mouse, seleccione la tabla de anomalías generada por Copilot en el chat.
4. Cópiela (`Ctrl+C`), vaya a la pestaña **"2. Hallazgos_Auditoria"** en Excel, haga clic en la celda **B3** y péguela (`Ctrl+V`).
5. Aplique color de relleno condicional básico o formato manual en Excel para resaltar en rojo las celdas etiquetadas como "Riesgo Alto".

---

### Fase C: Generación de Resúmenes Ejecutivos para la Dirección Financiera
El valor de un auditor radica en su capacidad para sintetizar problemas operativos complejos en reportes estratégicos legibles para la alta gerencia. Usaremos a Copilot para procesar los hallazgos aislados en la fase anterior y redactar un informe ejecutivo formal.

1. En el mismo hilo de conversación con Copilot, ejecute el siguiente prompt de redacción corporativa:

```
Actúa como el Director de Control Interno y Auditoría Forense de la compañía. Con base en la tabla de anomalías y riesgos documentales detectados en la Fase B, necesito presentar un informe ejecutivo de alto nivel dirigido al Director Financiero (CFO).

Redacta un reporte estructurado utilizando el siguiente esquema formal (utiliza títulos claros en negrita):
1. **Resumen Ejecutivo**: Un párrafo conciso que explique la situación general del libro diario analizado del mes de junio de 2026 y el estado general del cumplimiento.
2. **Impacto Financiero y Operativo Estimado**: Un análisis de cómo la falta de XMLs válidos, los documentos no legibles y las transacciones duplicadas exponen a la empresa a sanciones fiscales o fraudes internos.
3. **Recomendaciones de Mitigación**: Tres puntos de acción concretos y obligatorios para robustecer el proceso de carga de evidencias y firmas electrónicas de manera inmediata.

Escribe el texto con un tono directivo, sofisticado, directo y riguroso.
```

2. Lea y valide la estructura del texto redactado por Copilot.
3. Copie el resumen ejecutivo completo desde el chat, regrese a la pestaña **"2. Hallazgos_Auditoria"** de Excel y péguelo a partir de la celda **B12** (dejando un espacio limpio debajo de la tabla de anomalías). El reporte de dirección queda consolidado en un único entorno de trabajo.

---

### Fase D: Reto de Aplicación Autónoma – Matriz de Severidad Fiscal 

**Instrucciones:** Hasta este punto has seguido un proceso guiado de extracción y análisis. Ahora, debes demostrar criterio analítico y habilidades de orquestación autónoma aplicando ingeniería de prompts independiente.

#### El Desafío:
La Dirección Financiera ha solicitado mapear los hallazgos encontrados contra una nueva normativa de **Severidad Fiscal y Riesgo Tributario**. Debes interactuar con Copilot por tu cuenta para que procese las anomalías detectadas en la Fase B y construya una **Matriz de Impacto Impositivo**.

#### Pistas de Ingeniería de Prompts para el Éxito:
Para resolver este reto de forma autónoma, construye un prompt avanzado en el chat considerando las siguientes pautas:
* **Establecimiento de Rol Avanzado:** Fuerza a Copilot a actuar como un *Consultor Tributario Senior y Experto en Cumplimiento Fiscal de la DIAN / Entidades de Control*.
* **Definición de Nuevas Reglas de Negocio:** Explícale a la IA que bajo la normativa vigente:
  1. Cualquier transacción sin Factura XML Válida genera una "Inadmisión de Costos y Deducciones del 100%".
  2. Documentos no legibles o provisionales se consideran "Pasivos no Soportados" con riesgo de sanción por inexactitud.
  3. Los registros duplicados representan "Riesgo de Doble Contabilización Errónea".
* **Solicitud de Salida Estructurada:** Exígele a Copilot que evalúe tus anomalías bajo estos tres criterios y te devuelva una nueva **Tabla Markdown Visual Nativa** que incluya: `| ID_Anomalía | Cuenta | Riesgo_Fiscal | Impacto_Sancionatorio_Estimado | Prioridad_de_Corrección (Crítica / Alta) |`.

#### Cierre del Laboratorio:
1. Copie la tabla impositiva autónoma generada por Copilot y péguela en una nueva pestaña de Excel llamada **"3. Matriz_Riesgo_Fiscal"**, iniciando en la celda **B3**.
2. Póngase a prueba pidiéndole a la IA que concluya con una frase de cierre de 2 líneas sobre la urgencia de regularizar los pasivos detectados, y péguela en la celda **B10** para finalizar su entregable.

---

## 6. Conceptos Clave para Recordar

* **Auditoría Asistida por IA:** Delegar el análisis cualitativo de textos (como las descripciones de las evidencias o las notas de control) en un LLM acelera la detección de patrones de fraude o errores contables que las fórmulas matemáticas tradicionales de Excel suelen pasar por alto.
* **Gobierno Documental Completo:** El ciclo de Business Intelligence financiero no se limita a sumar columnas numéricas; requiere validar la integridad legal del soporte transaccional (entornos XML, flujos de firmas y legibilidad), asegurando el blindaje fiscal de la organización.
* **Productividad Cero Fricción:** Utilizar la interfaz de la IA como un motor analítico intermedio permite que Excel se transforme en una consola limpia de visualización corporativa, maximizando el tiempo invertido en la toma de decisiones estratégicas en lugar de la limpieza manual de filas.

---

## 7. Resultado Esperado

Al finalizar el laboratorio práctico, el estudiante entregará el archivo `Auditoria_Evidencias_Copilot.xlsx`, el cual deberá cumplir estrictamente la siguiente arquitectura:

1. **Pestaña "1. Registro_y_Evidencias":** Contendrá la base de datos contable cruda de 15 filas (Rango B3:G18) generada originalmente por la IA, sirviendo como el histórico de auditoría.
2. **Pestaña "2. Hallazgos_Auditoria":** Mostrará la consolidación analítica guiada:
   * **Rango B3:G7 (aprox):** La tabla de anomalías clave con sus niveles de riesgo asignados por Copilot.
   * **A partir de la celda B12:** El Resumen Ejecutivo e informe técnico formal dirigido a la Dirección Financiera (CFO).
3. **Pestaña "3. Matriz_Riesgo_Fiscal":** Presentará el resultado del **Reto de Aplicación Autónoma (Fase E)**: la matriz de impacto impositivo mapeada según las reglas tributarias de deducibilidad, junto con la conclusión estratégica de urgencia de regularización.
