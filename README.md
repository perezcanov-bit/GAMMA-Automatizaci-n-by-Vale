# GAMMA Automatización de Inscripciones

Sistema de automatización para el proceso de inscripción de GAMMA usando **Google Forms, Google Sheets y Google Apps Script**.

Este repositorio está pensado para que futuros oficiales puedan repetir el proceso cada año sin tener que reconstruir la automatización desde cero.

## Qué hace el sistema

Cuando una persona completa el formulario de inscripción, el sistema:

- Lee la nueva respuesta del Google Form.
- Crea una fila organizada en la hoja `MASTER`.
- Asigna automáticamente Periodo 1 y Periodo 2 según preferencias, cupo y reglas de Kinder.
- Mantiene las asignaciones finales en columnas separadas para permitir cambios manuales.
- Actualiza automáticamente los archivos de:
  - listas por clase,
  - grupos de salida,
  - información médica,
  - asistencia.

## Archivos principales

### `Code.gs`

Contiene todo el Apps Script de la automatización.

Antes de usarlo en un nuevo año, hay que actualizar la configuración correspondiente, especialmente los IDs de los cuatro archivos de salida.

Busca esta sección:

```javascript
const OUTPUT_SHEETS = {
  CLASS_ROSTERS: "PASTE_CLASS_ROSTERS_SPREADSHEET_ID_HERE",
  EXIT_GROUPS: "PASTE_EXIT_GROUPS_SPREADSHEET_ID_HERE",
  MEDICAL: "PASTE_MEDICAL_SPREADSHEET_ID_HERE",
  ATTENDANCE: "PASTE_ATTENDANCE_SPREADSHEET_ID_HERE"
};
```

Cada ID se obtiene desde la URL de una hoja de cálculo:

```text
https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit
```

Copia únicamente el texto entre `/d/` y `/edit`.

## Recursos que se deben crear cada año

Se recomienda crear estos seis recursos nuevos:

1. Google Form de inscripción.
2. Google Sheet central enlazado al Form.
3. Google Sheet de listas por clase.
4. Google Sheet de grupos de salida.
5. Google Sheet de información médica.
6. Google Sheet de asistencia.

## Hoja central

El archivo principal debe contener una pestaña llamada:

```text
MASTER
```

El sistema está diseñado para trabajar con estas columnas:

| Columna | Encabezado |
|---|---|
| A | Fecha de inscripción |
| B | Nombre completo |
| C | Edad |
| D | Talla de camiseta |
| E | Adulto responsable 1 |
| F | Teléfono 1 |
| G | Correo 1 |
| H | Adulto responsable 2 |
| I | Teléfono 2 |
| J | Correo 2 |
| K | Adulto responsable 3 |
| L | Teléfono 3 |
| M | Adulto responsable 4 |
| N | Teléfono 4 |
| O | Enfermedades |
| P | Alergias |
| Q | Otra condición |
| R | Seguro médico |
| S | Opción A |
| T | Opción B |
| U | Alternativa C |
| V | Asignación automática P1 |
| W | Asignación automática P2 |
| X | Asignación final P1 |
| Y | Asignación final P2 |
| Z | Estado |
| AA | Alerta |

Conviene congelar la fila 1 y activar filtros.

## Configuración anual que se debe revisar

Antes de abrir inscripciones, revisar:

- Nombre de la pestaña de respuestas del Form.
- Capacidad máxima por clase.
- Lista de clases disponibles.
- Texto exacto de la opción de Kinder.
- IDs de los cuatro archivos de salida.
- Mapeo de nombres de clases a pestañas de asistencia.
- Rangos de filas destinados a Periodo 1 y Periodo 2.
- Títulos exactos de las preguntas del Form.

La configuración actual usa una capacidad máxima de 20 estudiantes por clase.

## Triggers necesarios

En Apps Script se deben crear dos triggers instalables.

### 1. Registro de nueva inscripción

Función:

```text
onRegistrationSubmit
```

Configuración:

```text
Event source: From spreadsheet
Event type: On form submit
```

### 2. Cambio manual de asignación final

Función:

```text
onFinalAssignmentEdit
```

Configuración:

```text
Event source: From spreadsheet
Event type: On edit
```

Estos triggers son importantes porque el script abre y actualiza otros Google Sheets.

## Lógica general de asignación

Para estudiantes fuera de Kinder, el sistema intenta combinaciones de preferencias y alternativas respetando el cupo disponible.

También puede marcar casos como:

```text
AUTO
PARTIAL CHOICE
OUTSIDE CHOICES
MANUAL REVIEW
```

Las columnas `X` y `Y` contienen la asignación final y pueden editarse manualmente por un oficial.

Las listas de clase y asistencia utilizan estas asignaciones finales.

## Reglas de Kinder

La automatización incluye reglas específicas según edad y selección de Kinder.

Estas reglas deben revisarse cada año por si cambia la política de inscripción.

## Asistencia

El archivo de asistencia conserva la posición de estudiantes ya existentes en su clase.

Los nuevos estudiantes se agregan en el primer espacio disponible.

Importante: el código actual está diseñado suponiendo que la zona de estudiante/asistencia se encuentra dentro de `B:Z` y deja intactas las columnas `AA` en adelante.

## Antes de abrir inscripciones

Haz al menos una prueba completa:

1. Envía una inscripción de prueba desde el Form.
2. Confirma que aparece en `Form Responses 1`.
3. Confirma que aparece correctamente en `MASTER`.
4. Revisa la asignación automática.
5. Confirma que se actualizaron los cuatro archivos externos.
6. Cambia manualmente una asignación en `X` o `Y`.
7. Verifica que listas de clase y asistencia se actualicen.
8. Borra los datos de prueba antes de abrir el Form oficialmente.

## Importante sobre `buildMasterFromResponses()`

Esta función reconstruye el contenido de `MASTER` a partir de las respuestas del Form.

**No debe ejecutarse durante operación normal después de haber realizado asignaciones o cambios manuales**, porque puede borrar información existente en `MASTER`.

Lo ideal es terminar toda la configuración antes de abrir el proceso de inscripción.

## Documentación completa

La guía anual paso a paso se encuentra en la carpeta:

```text
docs/
```

Archivo recomendado:

```text
Gamma_Inscripciones_Automatizacion_Guia_Anual.docx
```

Esa guía contiene el proceso completo para reconstruir el sistema en un nuevo ciclo de inscripciones.

## Recomendación para futuros oficiales

Antes de modificar el código:

1. Haz una copia del Google Sheet central.
2. Haz una copia del Apps Script.
3. Prueba los cambios con registros ficticios.
4. No cambies nombres de hojas, preguntas o clases sin revisar las referencias correspondientes en `Code.gs`.

## Tecnología

- Google Forms
- Google Sheets
- Google Apps Script / JavaScript

## Propósito del repositorio

Este repositorio funciona como respaldo técnico y manual de continuidad para que la automatización de inscripciones de GAMMA pueda mantenerse y reutilizarse en futuros años.
