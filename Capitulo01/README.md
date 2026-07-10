# Laboratorio 1: Uso de Copilot para escribir, depurar y optimizar código VBA en Excel. Creación de automatizaciones personalizadas para tareas recurrentes sin conocimientos previos de programación.

## 1. Metadatos del Laboratorio

| Atributo | Detalle |
| :--- | :--- |
| **Duración** | 60 minutos |
| **Complejidad** | Media |
| **Audiencia** | Profesionales, Analistas e Ingenieros |
| **Tecnologías** | Microsoft Excel y Microsoft Copilot Chat (M365) |
| **Enfoque** | Ingeniería de Prompts aplicada a VBA |

---

## 2. Descripción 

En este laboratorio práctico, los participantes aprenderán a diseñar, refinar y desplegar macros de VBA en Microsoft Excel utilizando instrucciones en lenguaje natural mediante Copilot Chat. El ejercicio adopta un enfoque incremental y estructurado para construir una solución automatizada sin escribir código de forma manual.

---

## 3. Objetivos 

Al completar este laboratorio, serás capaz de:
* Configurar e interactuar con el entorno de desarrollo de VBA en Excel.
* Aplicar técnicas de ingeniería de prompts para la generación de código.
* Construir de forma incremental una macro compleja a través de iteraciones.
* Desplegar y probar la solución directamente en un libro habilitado para macros (.xlsm).

---

## 4. Prerrequisitos de Conocimiento y Licencias

* Manejo básico-intermedio de Microsoft Excel.
* No se requieren conocimientos previos de programación en VBA.
* Licencia activa de Microsoft 365 con acceso a Copilot Chat.
* Aplicación de escritorio de Microsoft Excel instalada (Windows o Mac).

---

## 5. Entorno de Laboratorio

* Microsoft Excel de escritorio (las macros no se ejecutan en la versión web).
* Pestaña de Programador / Desarrollador habilitada.
* Panel de Copilot Chat abierto y listo para interactuar.

---

## 6. Procedimiento Paso a Paso

### Sección A: Configuración Inicial del Entorno 

1. Abra una aplicación de Microsoft Excel en blanco.
2. Guarde el archivo inmediatamente con el nombre `Laboratorio_Capitulo1.xlsm` (Libro de Excel habilitado para macros).
3. Habilite la pestaña Programador: Vaya a Archivo > Opciones > Personalizar cinta de opciones, marque la casilla "Programador" o "Desarrollador" y haga clic en Aceptar.
4. En la pestaña Programador, haga clic en Visual Basic (o presione Alt + F11).
5. En el menú del editor de VBA, vaya a Insertar > Módulo.
---
### Sección B: Generación de la Macro Base - Formato y Estructura 

1. Abra su chat de Copilot.
2. Copie, pegue y ejecute el siguiente prompt:
   
```
Actúa como un experto en automatización de procesos en Excel y desarrollo de VBA. Necesito una macro llamada 'InicializarTablaCalidad' que prepare la Hoja1 de mi libro. La macro debe realizar las siguientes tareas estrictamente:
1. Limpiar todo el contenido y formatos existentes en la 'Hoja1'.
2. Crear un encabezado en la fila 3 (desde la columna B hasta la columna F) con los siguientes títulos: 'ID Incidente', 'Proceso Impactado', 'Causa Raíz Presunta', 'Impacto Operativo', 'Fecha de Registro'.
3. Aplicar formato profesional al encabezado: fondo azul oscuro, texto en negrita y color blanco, y alineación centrada.
4. Ajustar automáticamente el ancho de las columnas (AutoFit) de la B a la F.
Por favor, proporciona únicamente el código VBA limpio dentro de un bloque de código, incluyendo comentarios breves por cada sección.
```

3. Copie el código generado por Copilot.
4. Vaya al editor de VBA en Excel, haga doble clic en Módulo1 y pegue el código.
5. Regrese a Excel, vaya a Programador > Macros, seleccione "InicializarTablaCalidad" y haga clic en Ejecutar.

---

### Sección C: Incremento de Complejidad - Lógica de Negocio y Validación 

1. En el mismo chat de Copilot, copie y ejecute el siguiente prompt:

```
Basándote en la macro anterior, necesito crear un procedimiento llamado 'RegistrarNuevoIncidente'. Esta macro debe permitir al usuario ingresar datos fila por fila de forma interactiva:
1. Debe encontrar de forma automática la primera fila vacía disponible justo debajo de los encabezados creados anteriormente.
2. Debe solicitar al usuario el 'Proceso Impactado' mediante un InputBox.
3. Si el usuario deja el cuadro vacío o cancela, la macro debe mostrar un mensaje de aviso que diga 'Operación cancelada por datos incompletos' y detener la ejecución (Exit Sub).
4. Si el proceso es ingresado, debe generar automáticamente un 'ID Incidente' secuencial en la columna B.
5. Debe solicitar la 'Causa Raíz Presunta' mediante otro InputBox.
6. En la columna 'Fecha de Registro' (Columna F), la macro debe insertar automáticamente la fecha y hora actuales del sistema (Now).
7. Al finalizar, debe enviar un mensaje emergente (MsgBox) confirmando: 'Incidente registrado exitosamente'.
Proporcióname el código VBA optimizado listo para ser agregado al módulo existente.
```

2. Copie el nuevo código.
3. En el editor de VBA, vaya al final del código anterior y pegue esta nueva macro.
4. Pruebe la macro ejecutando "RegistrarNuevoIncidente" desde Excel.

### Sección D: Desafío de Optimización - Formato Condicional 

1. En el mismo chat de Copilot, ingrese el prompt del desafío final:

```
Necesito una macro llamada 'AplicarFormatoAlertas' que analice toda la columna E ('Impacto Operativo') desde la fila 4 en adelante.
La macro debe recorrer cada fila escrita y evaluar:
- Si el texto ingresado en la columna E contiene o es igual a la palabra 'Alto', debe pintar el fondo de esa celda específica de color rojo claro y la fuente de color rojo oscuro.
- La búsqueda no debe discriminar entre mayúsculas y minúsculas.
- Aplica bordes delgados grises a todo el rango de datos que tenga registros.
Entrégame el código VBA correspondiente bien estructurado.
```

2. Copie el código generado y péguelo al final de su módulo de VBA.
3. Registre un incidente con impacto "Alto" para probar la regla y ejecute la macro "AplicarFormatoAlertas".

---

## 7. Conceptos Clave para Recordar

* Mantenimiento de Contexto: Construir en un solo hilo permite que la IA recuerde nombres de variables y hojas anteriores.
* Estructura del Prompt: Indicar siempre un Rol, Objetivo, Restricciones claras y el Formato de salida deseado.
* Modularidad: Es mejor crear funciones separadas y específicas que intentar hacer una única macro masiva.

---

## 8. Resultado Esperado

Al concluir, dispondrás de un archivo `.xlsm` funcional con una tabla estructurada de forma automática, un sistema interactivo de captura de datos con validación ante campos vacíos, y un formateo condicional inteligente ejecutado mediante lenguaje natural.

![diagrama1](/1.jpg)
