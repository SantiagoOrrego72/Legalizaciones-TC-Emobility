# LEG-01 · Recepción de legalizaciones

Flujo de nube de Power Automate que implementa los **pasos 2 a 6** del proceso (casos de uso CU-01 a CU-04 de [`src/2-aplicacion/CASOS-DE-USO.md`](../../2-aplicacion/CASOS-DE-USO.md)):

1. Detecta el correo de legalización en el buzón.
2. Crea el caso en `Registro_Legalizaciones_TC` con estado **Pendiente** y le asigna su ID (`LEG-AAAAMM-NNNNN`).
3. Guarda cada adjunto válido en `Legalizaciones_TC/Pendientes/AAAA/MM/Colaborador`, con nombre `ID_nombre-original`.
4. Valida la documentación: factura PDF, comprobante Excel, formatos y posible duplicado.
5. Avisa al colaborador (documentación completa o incompleta) y, si hay posible duplicado, a Contabilidad.
6. Si algo falla, marca el caso con "ERROR TÉCNICO", avisa a TI y deja la ejecución como fallida.

Solo usa conectores **estándar** incluidos en Microsoft 365 (Office 365 Outlook y SharePoint), sin licencias premium.

---

## Antes de empezar

| Requisito | Cómo se cumple |
|---|---|
| SharePoint aprovisionado y verificado | `01-Aprovisionar-SharePoint.ps1` y `02-Verificar-SharePoint.ps1` con **FALLA: 0** |
| Cuenta de servicio con licencia de Microsoft 365 | Es la dueña del flujo y de sus conexiones. El script 01 le da nivel Colaborar en el sitio (`Permisos.CuentaServicioFlujo`) |
| Buzón de legalizaciones | Recomendado: buzón compartido (`Flujos.BuzonCompartido = $true`). Un administrador de Exchange le da a la cuenta de servicio **Acceso total** y **Enviar como** sobre ese buzón |
| Configuración de los flujos | `Generar-ConfiguracionFlujos.ps1` sin `-Borrador` → `salida/Configuracion-Flujos.json` |
| Plantillas de correo | Carpeta [`plantillas-correo/`](plantillas-correo/) |

Inicie sesión en <https://make.powerautomate.com> **con la cuenta de servicio**. Si el entorno no es el predeterminado, elija el correcto arriba a la derecha.

## Convenciones de esta guía

- **Nombre de la acción:** después de agregar cada acción, cámbiele el nombre (menú `…` > *Cambiar nombre*) exactamente como se indica, **sin tildes y con guiones bajos**. Las expresiones usan esos nombres.
- **Expresión:** en el campo, pulse *Insertar expresión* (**fx**), pegue el texto del bloque y pulse *Agregar* / *Aceptar*. No escriba una expresión como texto normal.
- **Filtrar matriz (Filter array):** pulse *Editar en modo avanzado* y pegue la condición completa, que empieza con `@`.
- **Seleccionar (Select):** pulse el botón *Cambiar al modo de texto* (icono **T**) junto a *Asignar* y pegue la expresión con **fx**.
- **Condición:** en el primer cuadro inserte la expresión indicada, elija el operador y escriba el valor del segundo cuadro. Si el valor es `true`, insértelo como expresión.
- Los nombres de las acciones están en español y, entre paréntesis, en inglés.

## Vista general

```text
LEG-01 Recepción de legalizaciones
├─ Desencadenador: correo nuevo en el buzón (con condición del asunto, simultaneidad 1)
├─ Configuracion · SaltoLinea · Plantilla_completa · Plantilla_incompleta · Plantilla_duplicado · Plantilla_error
├─ Inicializar varIdElemento · varIdLegalizacion · varHallazgos
├─ Intentar (Ámbito)
│  ├─ A. Datos del correo:   Asunto, Remitente, Fecha_local, Anio, Mes, Nombre_en_asunto, Colaborador,
│  │                         Carpeta_colaborador_base, Carpeta_colaborador, Ruta_carpeta
│  ├─ B. Adjuntos:           Adjuntos_reales, Facturas, Comprobantes, Xmls, No_soportados, Soportados,
│  │                         Nombres_guardados, Nombres_no_soportados
│  ├─ C. Validaciones:       Si_falta_factura, Si_falta_comprobante, Si_hay_no_soportados, Si_asunto_sin_nombre,
│  │                         Documentacion_completa
│  ├─ D. Registro:           Crear_registro, Guardar_id_elemento, Id_legalizacion, Guardar_id_legalizacion
│  ├─ E. Duplicados:         Casos_previos_remitente, Adjuntos_comparables, Nombres_comparables, Casos_duplicados,
│  │                         Ids_duplicados, Es_posible_duplicado, Aviso_duplicado, Observaciones
│  ├─ F. Archivos:           Guardar_adjuntos → Nombre_archivo_destino, Crear_archivo, Marcar_archivo_con_id
│  ├─ G. Actualizar_registro
│  └─ H. Notificaciones:     Lista_archivos_html, Hallazgos_html, Si_documentacion_completa, Si_posible_duplicado
├─ Capturar_error (Ámbito; solo si Intentar falla)
│  └─ Acciones_fallidas, Detalle_error, Url_ejecucion, Si_se_creo_registro, Cuerpo_error, Avisar_error_TI, Terminar_con_error
└─ (Opcional) Mover_correo_procesado
```

---

## Paso 1 · Crear el flujo

1. *Crear* > **Flujo de nube automatizado** (Automated cloud flow).
2. Nombre: `LEG-01 Recepción de legalizaciones`.
3. Desencadenador: vea el paso 2 y pulse *Crear*.

## Paso 2 · Desencadenador

Use **una** de las dos opciones, según `Flujos.BuzonCompartido`.

**Opción A: buzón compartido** (recomendada)

- Conector **Office 365 Outlook** › **Cuando llega un nuevo correo electrónico a un buzón compartido (V2)** (When a new email arrives in a shared mailbox (V2)).
- *Dirección del buzón original* (Original Mailbox Address): el valor de `Flujos.BuzonLegalizaciones`, por ejemplo `legalizaciones@emobility.com`.
- *Carpeta*: `Inbox` (Bandeja de entrada).
- *Solo con datos adjuntos*: **No**. Así se registran y avisan también los correos que llegan sin adjuntos.
- *Incluir datos adjuntos*: **Sí**.
- *Filtro de asunto*: vacío. El filtro se hace con la condición del desencadenador, que tolera tildes y mayúsculas.

**Opción B: buzón propio de la cuenta de servicio**

- **Office 365 Outlook** › **Cuando llega un nuevo correo electrónico (V3)** (When a new email arrives (V3)), con *Carpeta* `Inbox`, *Solo con datos adjuntos* **No** e *Incluir datos adjuntos* **Sí**.

**Configuración del desencadenador** (menú `…` > *Configuración*):

| Opción | Valor | Por qué |
|---|---|---|
| División (Split On) | **Activada** (viene activada) | Una ejecución por correo |
| Control de simultaneidad | **Activado**, grado de paralelismo **1** | Los correos se procesan de a uno. Evita que dos correos del mismo día se crucen al buscar duplicados. Una vez activado no se puede desactivar sin volver a crear el desencadenador |
| Condiciones del desencadenador | la expresión de abajo | Solo se procesan los correos cuyo asunto contiene "Legalización Tarjeta Crédito", con o sin tildes y en cualquier combinación de mayúsculas |

Condición del desencadenador (pegue el texto tal cual, **con** la `@` inicial):

```expresion
@contains(replace(replace(replace(replace(replace(replace(replace(toLower(coalesce(triggerBody()?['subject'], '')), 'á', 'a'), 'é', 'e'), 'í', 'i'), 'ó', 'o'), 'ú', 'u'), '  ', ' '), '  ', ' '), 'legalizacion tarjeta credito')
```

> Regla del dominio: `correo.textoClaveAsunto`. Ejemplos verificados en `pruebas/unitarias/Dominio.Tests.ps1` (CU-01).

## Paso 3 · Acciones de configuración (antes de cualquier ámbito)

Agréguelas en este orden directamente debajo del desencadenador.

#### P1 · `Configuracion`
- Acción: **Redactar** (Compose).
- *Entradas*: pegue el **contenido completo** de `salida/Configuracion-Flujos.json` (texto, no expresión).

#### P2 · `SaltoLinea`
- Acción: **Redactar**. *Entradas*, como expresión:

```expresion
decodeUriComponent('%0A')
```

#### P3 · `Plantilla_completa`
- Acción: **Redactar**. *Entradas*: pegue el contenido de `plantillas-correo/recepcion-completa.html` como texto.

#### P4 · `Plantilla_incompleta`
- Acción: **Redactar**. *Entradas*: contenido de `plantillas-correo/recepcion-incompleta.html`.

#### P5 · `Plantilla_duplicado`
- Acción: **Redactar**. *Entradas*: contenido de `plantillas-correo/aviso-duplicado.html`.

#### P6 · `Plantilla_error`
- Acción: **Redactar**. *Entradas*: contenido de `plantillas-correo/error-tecnico.html`.

#### P7 · Variables
Tres acciones **Inicializar variable** (Initialize variable):

| Nombre de la acción | Nombre de la variable | Tipo | Valor |
|---|---|---|---|
| `Inicializar_varIdElemento` | `varIdElemento` | Entero (Integer) | `0` |
| `Inicializar_varIdLegalizacion` | `varIdLegalizacion` | Cadena (String) | (vacío) |
| `Inicializar_varHallazgos` | `varHallazgos` | Matriz (Array) | `[]` |

## Paso 4 · Ámbito `Intentar`

Agregue una acción **Ámbito** (Scope) y llámela `Intentar`. Todas las acciones de los bloques A a H van **dentro** de este ámbito, en el orden indicado.

### Bloque A · Datos del correo

#### A1 · `Asunto`
Redactar:
```expresion
trim(coalesce(triggerBody()?['subject'], ''))
```

#### A2 · `Remitente`
Redactar:
```expresion
toLower(trim(coalesce(triggerBody()?['from'], '')))
```

#### A3 · `Fecha_local`
Redactar. Fecha y hora de recepción en la hora del país, para los correos:
```expresion
convertFromUtc(triggerBody()?['receivedDateTime'], outputs('Configuracion')?['zonaHorariaWindows'], 'dd/MM/yyyy HH:mm')
```

#### A4 · `Anio`
Redactar. Año **local** de recepción. Un correo del 31 de agosto a las 21:30 de México va a la carpeta de agosto aunque en UTC ya sea septiembre:
```expresion
convertFromUtc(triggerBody()?['receivedDateTime'], outputs('Configuracion')?['zonaHorariaWindows'], 'yyyy')
```

#### A5 · `Mes`
Redactar:
```expresion
convertFromUtc(triggerBody()?['receivedDateTime'], outputs('Configuracion')?['zonaHorariaWindows'], 'MM')
```

#### A6 · `Nombre_en_asunto`
Redactar. Texto entre el primer `" - "` y el siguiente. Si no hay `" - "`, queda vacío:
```expresion
trim(coalesce(split(outputs('Asunto'), ' - ')?[1], ''))
```
Ejemplo: `RV: Legalización Tarjeta Crédito - Ana Pérez` → `Ana Pérez`.

#### A7 · `Colaborador`
Redactar. Si el asunto no trae el nombre, se usa la parte del correo antes de la `@`:
```expresion
if(empty(outputs('Nombre_en_asunto')), first(split(outputs('Remitente'), '@')), outputs('Nombre_en_asunto'))
```

#### A8 · `Carpeta_colaborador_base`
Redactar. Quita los caracteres que SharePoint no admite en carpetas (`" * : < > ? / \ | # % .`) y limita el largo:
```expresion
trim(take(trim(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(outputs('Colaborador'), '"', ''), '*', ''), ':', ''), '<', ''), '>', ''), '?', ''), '/', ''), '\', ''), '|', ''), '#', ''), '%', ''), '.', '')), int(outputs('Configuracion')?['largoMaximoCarpeta'])))
```

#### A9 · `Carpeta_colaborador`
Redactar:
```expresion
if(empty(outputs('Carpeta_colaborador_base')), outputs('Configuracion')?['nombreSinDato'], outputs('Carpeta_colaborador_base'))
```

#### A10 · `Ruta_carpeta`
Redactar. Resultado: `/Legalizaciones_TC/Pendientes/2026/09/Ana Pérez`.
```expresion
concat(outputs('Configuracion')?['rutaBiblioteca'], '/', outputs('Configuracion')?['carpetaInicial'], '/', outputs('Anio'), '/', outputs('Mes'), '/', outputs('Carpeta_colaborador'))
```

### Bloque B · Adjuntos

#### B1 · `Adjuntos_reales`
**Filtrar matriz** (Filter array). Descarta las imágenes incrustadas en el cuerpo (firmas, logos).
- *Desde*, como expresión:
```expresion
coalesce(triggerBody()?['attachments'], json('[]'))
```
- Condición (modo avanzado):
```expresion
@not(equals(item()?['isInline'], true))
```

#### B2 · `Facturas`
Filtrar matriz. *Desde*: `body('Adjuntos_reales')` (expresión). Condición:
```expresion
@contains(outputs('Configuracion')?['extensiones']?['factura'], concat('.', toLower(last(split(item()?['name'], '.')))))
```

#### B3 · `Comprobantes`
Filtrar matriz. *Desde*: `body('Adjuntos_reales')`. Condición:
```expresion
@contains(outputs('Configuracion')?['extensiones']?['comprobante'], concat('.', toLower(last(split(item()?['name'], '.')))))
```

#### B4 · `Xmls`
Filtrar matriz. *Desde*: `body('Adjuntos_reales')`. Condición:
```expresion
@contains(outputs('Configuracion')?['extensiones']?['xml'], concat('.', toLower(last(split(item()?['name'], '.')))))
```

#### B5 · `No_soportados`
Filtrar matriz. *Desde*: `body('Adjuntos_reales')`. Condición:
```expresion
@not(contains(outputs('Configuracion')?['extensionesPermitidas'], concat('.', toLower(last(split(item()?['name'], '.'))))))
```

#### B6 · `Soportados`
Redactar:
```expresion
union(body('Facturas'), body('Comprobantes'), body('Xmls'))
```

#### B7 · `Nombres_guardados`
**Seleccionar** (Select). *Desde*: `outputs('Soportados')` (expresión). *Asignar* en modo texto:
```expresion
item()?['name']
```

#### B8 · `Nombres_no_soportados`
Seleccionar. *Desde*: `body('No_soportados')`. *Asignar* en modo texto:
```expresion
item()?['name']
```

### Bloque C · Validaciones (paso 5 del proceso)

Cada condición agrega un mensaje a `varHallazgos` con **Anexar a variable de matriz** (Append to array variable) en la rama **Sí** (True). La rama **No** queda vacía. Los textos salen de `Configuracion.mensajes`, es decir, del dominio.

#### C1 · `Si_falta_factura`
Condición: `length(body('Facturas'))` **es igual a** `0`.
- Sí › `Agregar_falta_factura` (Anexar a variable de matriz `varHallazgos`), valor:
```expresion
outputs('Configuracion')?['mensajes']?['faltaFactura']
```

#### C2 · `Si_falta_comprobante`
Condición: `length(body('Comprobantes'))` **es igual a** `0`.
- Sí › `Agregar_falta_comprobante` (varHallazgos):
```expresion
outputs('Configuracion')?['mensajes']?['faltaComprobante']
```

#### C3 · `Si_hay_no_soportados`
Condición: `length(body('No_soportados'))` **es mayor que** `0`.
- Sí › `Agregar_no_soportados` (varHallazgos):
```expresion
replace(replace(outputs('Configuracion')?['mensajes']?['archivosNoSoportados'], '{archivos}', join(body('Nombres_no_soportados'), ', ')), '{formatos}', join(outputs('Configuracion')?['extensionesPermitidas'], ', '))
```

#### C4 · `Si_asunto_sin_nombre`
Condición: `empty(outputs('Nombre_en_asunto'))` **es igual a** `true`.
- Sí › `Agregar_asunto_sin_nombre` (varHallazgos):
```expresion
outputs('Configuracion')?['mensajes']?['asuntoSinNombre']
```

#### C5 · `Documentacion_completa`
Redactar. Hay factura **y** comprobante (el XML es opcional):
```expresion
and(greater(length(body('Facturas')), 0), greater(length(body('Comprobantes')), 0))
```

### Bloque D · Registro del caso (paso 4 del proceso)

#### D1 · `Crear_registro`
**SharePoint › Crear elemento** (Create item).
- *Dirección del sitio*: elija **Legalizaciones TC** de la lista.
- *Nombre de la lista*: elija **Registro_Legalizaciones_TC**.
- Campos, todos como expresión:

| Campo | Expresión |
|---|---|
| ID Legalización | `concat('TMP-', guid())` |
| Fecha Recepción | `triggerBody()?['receivedDateTime']` |
| Colaborador | `take(outputs('Colaborador'), 255)` |
| Correo | `outputs('Remitente')` |
| Asunto | `take(outputs('Asunto'), 255)` |
| Estado Value | `outputs('Configuracion')?['estadoInicial']` (*Escribir un valor personalizado* > fx) |

El ID temporal es necesario porque el definitivo usa el número interno que SharePoint asigna al crear el elemento. La columna no admite valores repetidos, así que dos correos simultáneos nunca comparten ID.

#### D2 · `Guardar_id_elemento`
**Establecer variable** (Set variable) `varIdElemento`:
```expresion
outputs('Crear_registro')?['body/ID']
```

#### D3 · `Id_legalizacion`
Redactar. Resultado: `LEG-202609-00123`.
```expresion
concat(outputs('Configuracion')?['prefijoId'], '-', outputs('Anio'), outputs('Mes'), '-', formatNumber(outputs('Crear_registro')?['body/ID'], outputs('Configuracion')?['formatoConsecutivo']))
```

#### D4 · `Guardar_id_legalizacion`
Establecer variable `varIdLegalizacion`:
```expresion
outputs('Id_legalizacion')
```

### Bloque E · Posible duplicado (paso 5, heurístico)

Regla: es un **posible duplicado** si otro caso del **mismo remitente**, recibido en los últimos **45 días** y **no rechazado**, tiene **el mismo nombre de archivo PDF o XML**. El Excel no cuenta porque suele llamarse igual en cada envío. No se revisa el contenido.

#### E1 · `Casos_previos_remitente`
**SharePoint › Obtener elementos** (Get items). Sitio **Legalizaciones TC**, lista **Registro_Legalizaciones_TC**.
- *Consulta de filtro* (Filter Query), como una sola expresión:
```expresion
concat('Correo eq ''', replace(outputs('Remitente'), '''', ''''''), ''' and FechaRecepcion ge ''', addDays(triggerBody()?['receivedDateTime'], mul(-1, int(outputs('Configuracion')?['diasVentanaDuplicado'])), 'yyyy-MM-ddTHH:mm:ssZ'), ''' and ID ne ', string(variables('varIdElemento')), ' and Estado ne ''', join(outputs('Configuracion')?['estadosExcluidosDuplicado'], ''' and Estado ne '''), '''')
```
- *Recuento superior* (Top Count): `100`.

Resultado de ejemplo: `Correo eq 'ana@emobility.com' and FechaRecepcion ge '2026-08-06T15:00:00Z' and ID ne 123 and Estado ne 'Rechazada'`.

#### E2 · `Adjuntos_comparables`
Filtrar matriz. *Desde*: `outputs('Soportados')`. Condición:
```expresion
@contains(outputs('Configuracion')?['extensionesComparanDuplicados'], concat('.', toLower(last(split(item()?['name'], '.')))))
```

#### E3 · `Nombres_comparables`
Seleccionar. *Desde*: `body('Adjuntos_comparables')`. *Asignar* en modo texto:
```expresion
toLower(item()?['name'])
```

#### E4 · `Casos_duplicados`
Filtrar matriz. *Desde*: `outputs('Casos_previos_remitente')?['body/value']`. Condición:
```expresion
@greater(length(intersection(split(replace(toLower(coalesce(item()?['NombreArchivo'], '')), decodeUriComponent('%0D'), ''), decodeUriComponent('%0A')), body('Nombres_comparables'))), 0)
```

#### E5 · `Ids_duplicados`
Seleccionar. *Desde*: `body('Casos_duplicados')`. *Asignar* en modo texto:
```expresion
item()?['Title']
```

#### E6 · `Es_posible_duplicado`
Redactar:
```expresion
greater(length(body('Casos_duplicados')), 0)
```

#### E7 · `Aviso_duplicado`
Redactar:
```expresion
replace(replace(outputs('Configuracion')?['mensajes']?['posibleDuplicado'], '{casos}', join(body('Ids_duplicados'), ', ')), '{dias}', string(outputs('Configuracion')?['diasVentanaDuplicado']))
```

#### E8 · `Observaciones`
Redactar. Todos los hallazgos, uno por línea, más el aviso de duplicado:
```expresion
join(if(outputs('Es_posible_duplicado'), union(variables('varHallazgos'), createArray(outputs('Aviso_duplicado'))), variables('varHallazgos')), outputs('SaltoLinea'))
```

### Bloque F · Guardar los adjuntos (paso 3 del proceso)

#### F1 · `Guardar_adjuntos`
**Aplicar a cada uno** (Apply to each). *Seleccionar una salida de los pasos anteriores*: `outputs('Soportados')`.
Dentro, en orden:

#### F1.1 · `Nombre_archivo_destino`
Redactar. Resultado: `LEG-202609-00123_factura_A001.pdf`.
```expresion
concat(variables('varIdLegalizacion'), '_', replace(replace(replace(replace(replace(replace(replace(replace(replace(items('Guardar_adjuntos')?['name'], '"', ''), '*', ''), ':', ''), '<', ''), '>', ''), '?', ''), '/', ''), '\', ''), '|', ''))
```

#### F1.2 · `Crear_archivo`
**SharePoint › Crear archivo** (Create file). Si la carpeta del colaborador o del mes no existe, la acción la crea.
- *Dirección del sitio*: **Legalizaciones TC**.
- *Ruta de acceso de la carpeta* (Folder Path):
```expresion
outputs('Ruta_carpeta')
```
- *Nombre de archivo*:
```expresion
outputs('Nombre_archivo_destino')
```
- *Contenido del archivo*:
```expresion
base64ToBinary(items('Guardar_adjuntos')?['contentBytes'])
```

#### F1.3 · `Marcar_archivo_con_id`
**SharePoint › Actualizar propiedades de archivo** (Update file properties). Sitio **Legalizaciones TC**, biblioteca **Legalizaciones_TC**.
- *Id*:
```expresion
outputs('Crear_archivo')?['body/ItemId']
```
- *ID Legalización*: `variables('varIdLegalizacion')` (expresión).

### Bloque G · Completar el registro

#### G1 · `Actualizar_registro`
**SharePoint › Actualizar elemento** (Update item). Sitio **Legalizaciones TC**, lista **Registro_Legalizaciones_TC**. Los campos obligatorios se vuelven a enviar porque la acción los exige.

| Campo | Expresión |
|---|---|
| Id | `variables('varIdElemento')` |
| ID Legalización | `variables('varIdLegalizacion')` |
| Fecha Recepción | `triggerBody()?['receivedDateTime']` |
| Colaborador | `take(outputs('Colaborador'), 255)` |
| Correo | `outputs('Remitente')` |
| Asunto | `take(outputs('Asunto'), 255)` |
| Nombre Archivo | `join(body('Nombres_guardados'), outputs('SaltoLinea'))` |
| Estado Value | `outputs('Configuracion')?['estadoInicial']` |
| Observaciones | `outputs('Observaciones')` |

### Bloque H · Notificaciones (paso 6 del proceso)

#### H1 · `Lista_archivos_html`
Redactar:
```expresion
if(empty(body('Nombres_guardados')), '<p>(ninguno)</p>', concat('<ul><li>', join(body('Nombres_guardados'), '</li><li>'), '</li></ul>'))
```

#### H2 · `Hallazgos_html`
Redactar:
```expresion
if(empty(variables('varHallazgos')), '', concat('<ul><li>', join(variables('varHallazgos'), '</li><li>'), '</li></ul>'))
```

#### H3 · `Si_documentacion_completa`
Condición: `outputs('Documentacion_completa')` **es igual a** `true`.

**Rama Sí**

#### H3a · `Cuerpo_completa`
Redactar:
```expresion
replace(replace(replace(replace(replace(outputs('Plantilla_completa'), '{{COLABORADOR}}', outputs('Colaborador')), '{{ID}}', variables('varIdLegalizacion')), '{{FECHA}}', outputs('Fecha_local')), '{{ARCHIVOS}}', outputs('Lista_archivos_html')), '{{NOTAS}}', if(empty(variables('varHallazgos')), '', concat('<p><b>Notas:</b></p>', outputs('Hallazgos_html'))))
```

#### H3b · `Enviar_correo_completa`
**Office 365 Outlook › Enviar un correo electrónico desde un buzón compartido (V2)** (Send an email from a shared mailbox (V2)). Con la opción B del desencadenador use **Enviar un correo electrónico (V2)**, que no pide el buzón.
- *Dirección del buzón original*: `outputs('Configuracion')?['buzonLegalizaciones']`
- *Para*: `outputs('Remitente')`
- *Asunto*:
```expresion
concat('Legalización recibida - ', variables('varIdLegalizacion'))
```
- *Cuerpo*: cambie el editor a vista de código (**</>**) e inserte la expresión `outputs('Cuerpo_completa')`.

**Rama No**

#### H3c · `Cuerpo_incompleta`
Redactar:
```expresion
replace(replace(replace(replace(replace(outputs('Plantilla_incompleta'), '{{COLABORADOR}}', outputs('Colaborador')), '{{ID}}', variables('varIdLegalizacion')), '{{FECHA}}', outputs('Fecha_local')), '{{HALLAZGOS}}', outputs('Hallazgos_html')), '{{ARCHIVOS}}', outputs('Lista_archivos_html'))
```

#### H3d · `Enviar_correo_incompleta`
Misma acción de envío que H3b.
- *Dirección del buzón original*: `outputs('Configuracion')?['buzonLegalizaciones']`
- *Para*: `outputs('Remitente')`
- *Asunto*:
```expresion
concat('Legalización incompleta - ', variables('varIdLegalizacion'))
```
- *Cuerpo* (vista de código): `outputs('Cuerpo_incompleta')`.

#### H4 · `Si_posible_duplicado`
Después de H3, fuera de sus ramas. Condición: `outputs('Es_posible_duplicado')` **es igual a** `true`. La rama **No** queda vacía.

#### H4a · `Cuerpo_duplicado`
Rama Sí. Redactar:
```expresion
replace(replace(replace(replace(replace(replace(outputs('Plantilla_duplicado'), '{{ID}}', variables('varIdLegalizacion')), '{{COLABORADOR}}', outputs('Colaborador')), '{{CORREO}}', outputs('Remitente')), '{{FECHA}}', outputs('Fecha_local')), '{{AVISO}}', outputs('Aviso_duplicado')), '{{ENLACE}}', concat(outputs('Configuracion')?['urlElemento'], string(variables('varIdElemento'))))
```

#### H4b · `Avisar_duplicado`
Misma acción de envío. Es un aviso **interno**: no va al colaborador.
- *Dirección del buzón original*: `outputs('Configuracion')?['buzonLegalizaciones']`
- *Para*: `outputs('Configuracion')?['correoContabilidad']`
- *Asunto*:
```expresion
concat('[Legalizaciones] Posible duplicado - ', variables('varIdLegalizacion'))
```
- *Cuerpo* (vista de código): `outputs('Cuerpo_duplicado')`
- *Importancia*: **Alta**.

## Paso 5 · Ámbito `Capturar_error`

Agregue otro **Ámbito** debajo de `Intentar` (no dentro) y llámelo `Capturar_error`. En su menú `…` > *Configurar ejecución posterior* (Configure run after) marque **tiene errores** y **agotó el tiempo de espera** (has failed, has timed out) y desmarque **es correcto**. Así solo se ejecuta si algo de `Intentar` falla.

Dentro, en orden:

#### J1 · `Acciones_fallidas`
Filtrar matriz. *Desde*:
```expresion
result('Intentar')
```
Condición:
```expresion
@equals(item()?['status'], 'Failed')
```

#### J2 · `Detalle_error`
Redactar:
```expresion
take(string(first(body('Acciones_fallidas'))), 3000)
```

#### J3 · `Url_ejecucion`
Redactar. Enlace directo a esta ejecución:
```expresion
concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'], '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])
```

#### J4 · `Si_se_creo_registro`
Condición: `variables('varIdElemento')` **es mayor que** `0`.

#### J4a · `Marcar_registro_con_error`
Rama Sí. **SharePoint › Actualizar elemento**, sitio y lista del registro:

| Campo | Expresión |
|---|---|
| Id | `variables('varIdElemento')` |
| ID Legalización | `if(empty(variables('varIdLegalizacion')), concat('ERR-', string(variables('varIdElemento'))), variables('varIdLegalizacion'))` |
| Fecha Recepción | `triggerBody()?['receivedDateTime']` |
| Colaborador | `take(outputs('Colaborador'), 255)` |
| Correo | `outputs('Remitente')` |
| Asunto | `take(outputs('Asunto'), 255)` |
| Estado Value | `outputs('Configuracion')?['estadoInicial']` |
| Observaciones | `concat(outputs('Configuracion')?['prefijoNotaInterna'], ' ERROR TÉCNICO al procesar el correo; TI fue notificado. Ejecución: ', outputs('Url_ejecucion'))` |

El texto empieza con `[Interno]`, así que LEG-02 no lo incluye nunca en el correo al colaborador.

#### J5 · `Cuerpo_error`
Redactar, **después** de J4. En *Configurar ejecución posterior* de esta acción marque **es correcto**, **tiene errores** y **se omite**, para que el aviso salga aunque falle el marcado del registro.
```expresion
replace(replace(replace(replace(replace(replace(outputs('Plantilla_error'), '{{FLUJO}}', 'LEG-01 Recepción de legalizaciones'), '{{ASUNTO}}', coalesce(triggerBody()?['subject'], '(sin asunto)')), '{{REMITENTE}}', coalesce(triggerBody()?['from'], '(desconocido)')), '{{FECHA}}', convertFromUtc(utcNow(), outputs('Configuracion')?['zonaHorariaWindows'], 'dd/MM/yyyy HH:mm')), '{{DETALLE}}', outputs('Detalle_error')), '{{ENLACE_EJECUCION}}', outputs('Url_ejecucion'))
```

#### J6 · `Avisar_error_TI`
Acción de envío (como H3b).
- *Dirección del buzón original*: `outputs('Configuracion')?['buzonLegalizaciones']`
- *Para*: `outputs('Configuracion')?['correoSoporteTI']`
- *Asunto*:
```expresion
concat('[Legalizaciones] Error técnico en LEG-01 - ', coalesce(triggerBody()?['subject'], '(sin asunto)'))
```
- *Cuerpo* (vista de código): `outputs('Cuerpo_error')`
- *Importancia*: **Alta**.

#### J7 · `Terminar_con_error`
**Terminar** (Terminate). *Configurar ejecución posterior*: **es correcto** y **tiene errores** de J6.
- *Estado*: **Con errores** (Failed).
- *Código*: `LEG01_ERROR`.
- *Mensaje*:
```expresion
take(outputs('Detalle_error'), 1000)
```

## Paso 6 · (Opcional) Mover el correo procesado

Sirve para ver en el buzón qué correos ya se procesaron. Antes, cree en el buzón la carpeta **Procesadas**.

#### I1 · `Mover_correo_procesado`
Debajo de `Capturar_error`, **fuera** de los ámbitos. *Configurar ejecución posterior*: marque solo **es correcto** de `Intentar`, desmarque lo demás.
- **Office 365 Outlook › Mover correo electrónico (V2)** (Move email (V2)).
- *Id. de mensaje*:
```expresion
triggerBody()?['id']
```
- *Carpeta*: **Procesadas**.
- *Dirección del buzón original* (solo opción A): `outputs('Configuracion')?['buzonLegalizaciones']`.

Va fuera de `Intentar` a propósito: si mover el correo falla, el caso ya quedó bien registrado y no debe marcarse con error.

## Paso 7 · Guardar y probar

1. **Guardar**. El diseñador marca en rojo cualquier expresión mal pegada; corríjala antes de seguir.
2. Envíe **un** correo de prueba (escenario E01 de [`pruebas/PLAN-DE-PRUEBAS.md`](../../../pruebas/PLAN-DE-PRUEBAS.md)) y abra la ejecución en *Historial de ejecuciones*.
3. Compruebe:
   - un elemento nuevo en el registro con ID `LEG-AAAAMM-NNNNN` y estado Pendiente;
   - los archivos en `Pendientes/AAAA/MM/Colaborador` con el prefijo del ID y la columna ID Legalización llena;
   - el correo "Legalización recibida" en el buzón del remitente.
4. Siga con el plan de pruebas completo.

## Solución de problemas

| Síntoma | Causa probable | Qué hacer |
|---|---|---|
| El flujo no se dispara | El asunto no contiene "Legalización Tarjeta Crédito" o el correo llegó a otra carpeta | Revise el asunto y la carpeta del desencadenador. Pruebe la condición con un asunto de los ejemplos |
| `InvalidTemplate` al guardar | Expresión pegada como texto o con comillas cambiadas por el editor | Vuelva a insertarla con **fx**. Use comillas rectas `'`, no tipográficas `’` |
| `outputs('Configuracion')?['…']` devuelve vacío | En *Configuracion* no se pegó el JSON completo, o no es JSON válido | Vuelva a pegar el contenido de `salida/Configuracion-Flujos.json` |
| Error 403 en *Crear archivo* o *Crear elemento* | La cuenta de la conexión no tiene permiso en el sitio | Agregue la cuenta en `Permisos.CuentaServicioFlujo` y ejecute otra vez `01-Aprovisionar-SharePoint.ps1` |
| No envía desde el buzón compartido | Falta el permiso **Enviar como** | Pida al administrador de Exchange **Acceso total** y **Enviar como** para la cuenta de servicio |
| Error en *Obtener elementos* con la consulta de filtro | Falta la columna o el índice | Ejecute `02-Verificar-SharePoint.ps1`: Correo y FechaRecepcion deben existir e indexarse |
| Valores únicos: "Ya existe un elemento con ese valor" | Ejecución duplicada muy rara, o ID escrito a mano | Revise el caso existente. El flujo no reutiliza IDs |

## Trazabilidad

| Bloque | Paso del proceso | Caso de uso | Reglas del dominio (`dominio.json`) |
|---|---|---|---|
| Desencadenador | 1 y 2 | CU-01 | `correo.textoClaveAsunto` |
| A | 2 | CU-01 | `correo.separadorNombre`, `carpetas.*` |
| B | 5 (formatos) | CU-02 | `documentos.*.extensiones` |
| C | 5 (existencia) | CU-02 | `documentos.*.obligatorio`, `mensajes.*` |
| D | 4 | CU-01 | `caso.prefijoId`, `caso.digitosConsecutivo`, `caso.estadoInicial` |
| E | 5 (duplicado) | CU-03 | `duplicados.*`, `documentos.*.comparaDuplicados` |
| F | 3 | CU-01 | `carpetas.raizPorEstado` |
| G | 4 | CU-01 | — |
| H | 6 | CU-04 | `mensajes.*` |
| Capturar_error | — | Manejo de errores | — |
