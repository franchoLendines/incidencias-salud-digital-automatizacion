# Automatización de gestión de incidencias médicas (Google Apps Script)

Script para Google Apps Script que automatiza la gestión de incidencias de un servicio de salud digital: recibe correos de incidencias, determina automáticamente a qué proveedor médico corresponde cada caso, se lo reenvía, detecta cuándo el proveedor confirma el cierre y traslada la resolución al solicitante original — todo dentro del mismo hilo de Gmail.

## Qué hace

1. **Recepción**: detecta correos entrantes de incidencias (por asunto y remitente) que aún no han sido procesados.
2. **Clasificación**: analiza el texto del correo para distinguir incidencias **técnicas** (login, fallos de app, cobertura, etc.) de **asistenciales** (médico no se presentó, demoras, cambios de cita, etc.), y detecta la especialidad/servicio implicado.
3. **Enrutamiento automático**: decide a qué proveedor médico enviar el caso, combinando el servicio/especialidad solicitado con el listado de médicos de cada proveedor.
4. **Reenvío**: reenvía la incidencia al proveedor correspondiente en el mismo hilo de correo, con un resumen estructurado del caso.
5. **Seguimiento de respuestas**: cuando el proveedor responde, comprueba si el mensaje contiene la confirmación de cierre. Si la hay, reenvía la resolución al solicitante original; si no, genera una alerta interna para revisión manual.
6. **Registro**: cada incidencia se registra en Google Sheets (uno por proveedor / grupo de proveedores), con deduplicación por número de caso.
7. **Recordatorios de SLA**: si un caso lleva abierto 3, 5, 7, 14... días, envía recordatorios automáticos al proveedor con tono creciente de urgencia.
8. **Resumen semanal** (opcional): envía por correo un resumen con el número de casos, resueltos, pendientes y tiempo medio de resolución por proveedor.

Todas las alertas internas (incidencias técnicas, casos sin proveedor claro, respuestas que no confirman el cierre) se centralizan en un único buzón con el asunto prefijado `[ACCIÓN MANUAL - ...]`.

## Estructura del código

Un único archivo, `Codigo.gs`, organizado en secciones:

- `CONFIG`: todos los parámetros configurables (IDs de Sheets, correos de contacto, textos, plazos de SLA).
- Clasificación de incidencias por palabras clave (técnica vs. asistencial, con subcategorías).
- Enrutamiento por servicio/especialidad → proveedor.
- Envío y reenvío de correos (con historial completo del hilo incluido en HTML).
- Registro y cierre de casos en Google Sheets.
- Recordatorios de SLA con control de duplicados.
- Utilidades (normalización de texto, detección de auto-respuestas, etc.).

## Requisitos previos

- Una cuenta de Google Workspace o Gmail desde la que se reciban las incidencias.
- Google Sheets para registrar los casos (uno para cada proveedor o grupo de proveedores — ver `CONFIG`).
- Un Google Sheet con el listado de médicos por proveedor (columnas: `PROVEEDOR | NOMBRE | APELLIDO1 | APELLIDO2 | ESPECIALIDAD`).

## Instalación

1. Crea un proyecto nuevo en [script.google.com](https://script.google.com).
2. Copia el contenido de `Codigo.gs` en el editor.
3. Rellena la sección `CONFIG` al principio del archivo con tus datos:
   - IDs de tus Google Sheets (se sacan de la URL: `https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit`).
   - Tu correo de entrada (`CORREO_ENTRADA`) y el asunto que filtra las incidencias entrantes (`ASUNTO_FILTRO`).
   - El correo de contacto (`to`/`cc`) de cada proveedor.
4. En cada Google Sheet de proveedores, crea las pestañas y cabeceras necesarias. Puedes usar la función auxiliar `crearCabecerasOtrosProveedores()` (ejecútala una vez manualmente desde el editor) para generarlas automáticamente en el Sheet de "otros proveedores".
5. La primera vez que ejecutes cualquier función, Google te pedirá autorizar permisos de Gmail, Sheets y Drive.

## Activadores (ejecución automática)

En el editor de Apps Script, ve a **Activadores** (icono de reloj) y crea:

| Función | Frecuencia recomendada |
|---|---|
| `procesarIncidencias` | Cada 15 minutos |
| `procesarRespuestasProveedores` | Cada 15 minutos |
| `gestionarRecordatorios` | Diario, a una hora fija (p. ej. 09:00) |
| `resumenSemanal` | Semanal, un día fijo (opcional) |

También puedes ejecutar cualquier función manualmente desde el editor con el botón ▶️ para probarla.

## Notas

- El código incluido en este repositorio usa **valores de ejemplo** (`TU_SHEET_ID...`, `tu_correo@tudominio.com`, etc.) en lugar de los IDs y correos reales del entorno donde se desarrolló originalmente. Sustitúyelos por los tuyos antes de ejecutar.
- Los proveedores están identificados de forma genérica como `PROVEEDOR1`...`PROVEEDOR7` (en `CONFIG`, `NOMBRES_PROVEEDORES`, `MAPA_SERVICIOS`, etc.). `PROVEEDOR1` y `PROVEEDOR2` tienen su propio Google Sheet de seguimiento; `PROVEEDOR3` a `PROVEEDOR7` comparten un mismo Sheet (uno por pestaña). Renombra estos identificadores o ajusta las palabras clave de `MAPA_SERVICIOS` según tus proveedores reales.
- El script usa `LockService` para evitar ejecuciones simultáneas y etiqueta cada hilo de Gmail nada más procesarlo, para no reprocesar nunca el mismo caso.
- Si tu listado de médicos está en un archivo `.xlsx` de Drive (en vez de un Google Sheet nativo), necesitas habilitar el servicio avanzado **Drive API** en el editor de Apps Script (Servicios → "+" → Drive API).
