# Capa de aplicación: casos de uso y puertos

Los casos de uso describen **qué** hace el sistema con las reglas del dominio, sin atarse a una herramienta. Cada uno indica los **puertos** (capacidades) que necesita. La tabla del final dice qué adaptador implementa cada puerto en la Fase 1.

## Puertos

| Puerto | Operaciones que necesitan los casos de uso | Adaptador en Fase 1 |
|---|---|---|
| **BuzonEntrada** | Recibir un correo con asunto, remitente, fecha y adjuntos | Office 365 Outlook, desencadenador de LEG-01 |
| **RegistroCasos** | Crear un caso, actualizarlo, buscar casos previos por remitente y fecha, detectar cambios de estado | Lista SharePoint `Registro_Legalizaciones_TC` (Crear, Actualizar y Obtener elementos; desencadenador de LEG-02) |
| **AlmacenSoportes** | Guardar un archivo en la carpeta de su caso, marcarlo con el ID, listar los archivos de un caso y moverlos | Biblioteca SharePoint `Legalizaciones_TC` (Crear archivo, Actualizar propiedades, Obtener archivos, Mover archivo) |
| **Notificador** | Enviar un correo al colaborador, a Contabilidad o a TI | Office 365 Outlook, envío desde el buzón compartido |
| **LectorComprobantes** | Leer las líneas de gasto del Excel quincenal | Power Query (`fnLeerComprobante` sobre la tabla `tblGastos`) |
| **Analisis** | Consolidar, controlar la calidad y preparar la carga en el ERP | Libro de Excel con Power Query |
| **Reloj** | Fecha actual y conversión de UTC a la hora local | `utcNow()` / `convertFromUtc()` en Power Automate; `fnUtcALocal` en Power Query |

---

## CU-01 · Recibir y registrar una legalización

- **Actor:** colaborador (indirecto, por correo). **Disparador:** llega un correo al buzón.
- **Pasos del proceso:** 1 a 4.
- **Precondición:** el asunto cumple R-01.
- **Flujo principal:**
  1. Obtener asunto, remitente, fecha local y colaborador (R-02, R-10).
  2. Crear el caso en **RegistroCasos** con estado inicial `Pendiente` y un ID temporal.
  3. Calcular el ID definitivo con el número interno del caso (R-09).
  4. Guardar cada adjunto válido en **AlmacenSoportes**, en la carpeta del caso (R-08) y con el nombre `ID_original`, y marcarlo con el ID.
  5. Actualizar el caso con el ID, los nombres de archivo y las Observaciones (resultado de CU-02 y CU-03).
- **Postcondición:** el caso existe, tiene un ID único y sus soportes están guardados y enlazados.
- **Implementación:** LEG-01, desencadenador y bloques A, D, F y G.

## CU-02 · Validar la documentación

- **Paso del proceso:** 5.
- **Entrada:** los adjuntos del correo, sin las imágenes incrustadas (R-05).
- **Reglas:** R-03 (tipos por extensión), R-04 (completa si hay factura y comprobante), R-06 (archivos no válidos) y R-02 (asunto sin nombre).
- **Salida:** documentación completa (sí o no) y lista de hallazgos para el colaborador.
- **Implementación:** LEG-01, bloques B y C. Referencia: `Test-Documentacion`.

## CU-03 · Detectar un posible duplicado

- **Paso del proceso:** 5.
- **Entrada:** remitente, fecha y nombres de los PDF y XML del correo.
- **Regla:** R-07. Es heurístico y no revisa el contenido.
- **Salida:** los IDs de los casos que coinciden y el aviso para Observaciones y para Contabilidad.
- **Implementación:** LEG-01, bloque E. Referencia: `Find-PosiblesDuplicados`.

## CU-04 · Notificar la recepción

- **Paso del proceso:** 6.
- **Flujo:**
  - documentación completa → correo "Legalización recibida" al colaborador, con notas si hubo hallazgos menores;
  - incompleta → correo "Legalización incompleta" con lo que falta y cómo reenviar;
  - posible duplicado → aviso **interno** a Contabilidad, nunca al colaborador.
- **Puerto:** Notificador.
- **Implementación:** LEG-01, bloque H.

## CU-05 · Revisar y cerrar un caso

- **Actor:** Contabilidad. **Paso del proceso:** 7.
- **Flujo principal:**
  1. Contabilidad abre los casos pendientes (vista "Pendientes de revisión", los más antiguos primero), revisa los soportes y, si quiere, marca el caso "En revisión".
  2. Decide **Aprobada** o **Rechazada** y escribe el motivo en Observaciones. Esta decisión es **manual por diseño**.
  3. El sistema mueve los archivos a la carpeta del estado final (R-08), registra la Fecha Cierre y avisa al colaborador (configurable).
- **Reglas:** R-11. Aprobada y Rechazada son finales.
- **Implementación:** SharePoint (decisión manual) y LEG-02 (pasos automáticos).

## CU-06 · Consolidar para Operam

- **Actor:** Contabilidad. **Paso del proceso:** 8.
- **Flujo:**
  1. Actualizar el libro de análisis.
  2. Tomar de *Consolidado_Operam* las líneas de gasto de los casos **Aprobados** que aún no se cargaron.
  3. Cargarlas **manualmente** en Operam.
  4. Anotar cada ID en *Control_Operam*, para que no se vuelva a cargar.
- **Puertos:** LectorComprobantes y Analisis.
- **Implementación:** consultas `Gastos`, `Consolidado_Operam` y `Control_Operam`.

## CU-07 · Controlar la calidad

- **Actor:** Contabilidad o TI.
- **Controles:**
  - casos abiertos hace más de 5 días (R-12);
  - posibles duplicados sin resolver;
  - casos cerrados sin Fecha Cierre (LEG-02 falló);
  - casos reabiertos;
  - casos aprobados sin comprobante legible;
  - comprobantes que no usan la plantilla;
  - totales que no cuadran (Subtotal + IVA ≠ Total);
  - casos sin archivos.
- **Implementación:** consulta `Alertas` del libro de análisis.

---

## Trazabilidad con el proceso acordado

| Paso acordado | Caso de uso | Implementación |
|---|---|---|
| 1. El colaborador envía el correo | CU-01 | Outlook |
| 2. Power Automate detecta el correo | CU-01 | LEG-01, desencadenador |
| 3. Guarda cada adjunto en su subcarpeta | CU-01 | LEG-01, bloque F |
| 4. Crea el registro con estado Pendiente | CU-01 | LEG-01, bloques D y G |
| 5. Validaciones automáticas | CU-02, CU-03 | LEG-01, bloques B, C y E |
| 6. Notificaciones | CU-04 | LEG-01, bloque H |
| 7. Contabilidad revisa y aprueba o rechaza | CU-05 | SharePoint + LEG-02 |
| 8. Consolidación en Excel antes de Operam | CU-06, CU-07 | Libro de análisis (Power Query) |
