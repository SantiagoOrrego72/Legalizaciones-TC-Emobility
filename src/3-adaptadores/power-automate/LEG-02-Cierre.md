# LEG-02 · Cierre de casos

Flujo de nube que completa el **paso 7** del proceso (caso de uso CU-05). Cuando Contabilidad cambia el Estado de un caso a **Aprobada** o **Rechazada** en SharePoint:

1. Mueve los archivos del caso de `Pendientes/…` a `Aprobadas/…` o `Rechazadas/…` (misma estructura `AAAA/MM/Colaborador`).
2. Registra la **Fecha Cierre**.
3. Avisa el resultado al colaborador (configurable con `Flujos.NotificarCierre`), con las Observaciones como motivo.

La decisión de aprobar o rechazar sigue siendo manual. Este flujo solo ejecuta lo que antes había que hacer a mano: mover archivos y fechar el cierre.

Por qué es seguro:
- Solo se dispara cuando el caso está en Aprobada o Rechazada **y** todavía no tiene Fecha Cierre (condición del desencadenador). Cuando el flujo escribe la fecha, la nueva modificación ya no cumple la condición, así que no se repite.
- Si mover un archivo falla, **no** registra la Fecha Cierre. El caso aparece en la alerta "Caso cerrado sin Fecha Cierre" del libro de Excel. Cuando TI corrige la causa y alguien guarda de nuevo el elemento, el flujo se vuelve a disparar.

Requisitos: los mismos de LEG-01 (misma cuenta de servicio, mismas conexiones y el mismo `salida/Configuracion-Flujos.json`).

## Vista general

```text
LEG-02 Cierre de casos
├─ Desencadenador: se crea o modifica un elemento del registro (condición: estado final y sin Fecha Cierre)
├─ Configuracion · Plantilla_aprobada · Plantilla_rechazada · Plantilla_error
├─ Intentar (Ámbito)
│  ├─ Estado_cierre, Id_legalizacion, Anio, Mes, Carpeta_colaborador_base, Carpeta_colaborador, Carpeta_destino
│  ├─ Crear_carpeta_destino
│  ├─ Archivos_del_caso, Archivos_por_mover, Mover_archivos → Mover_archivo
│  ├─ Registrar_fecha_cierre
│  └─ Si_notificar_cierre → Observaciones_para_colaborador, Cuerpo_cierre, Enviar_resultado
└─ Capturar_error (Ámbito; solo si Intentar falla)
   └─ Acciones_fallidas, Detalle_error, Url_ejecucion, Cuerpo_error, Avisar_error_TI, Terminar_con_error
```

Se aplican las mismas convenciones de [LEG-01](LEG-01-Recepcion.md#convenciones-de-esta-guía): nombres exactos, expresiones con **fx** y *Filtrar matriz* en modo avanzado.

## Paso 1 · Crear el flujo y el desencadenador

1. *Crear* > **Flujo de nube automatizado**. Nombre: `LEG-02 Cierre de casos`.
2. Desencadenador: **SharePoint › Cuando se crea o modifica un elemento** (When an item is created or modified).
   - *Dirección del sitio*: **Legalizaciones TC**.
   - *Nombre de la lista*: **Registro_Legalizaciones_TC**.
3. Configuración del desencadenador (`…` > *Configuración*):
   - Control de simultaneidad: **Activado**, grado de paralelismo **1**.
   - Condición del desencadenador (con la `@` inicial):

```expresion
@and(or(equals(triggerBody()?['Estado']?['Value'], 'Aprobada'), equals(triggerBody()?['Estado']?['Value'], 'Rechazada')), equals(triggerBody()?['FechaCierre'], null))
```

> Los estados finales vienen del dominio: `caso.estadoAprobado` y `caso.estadoRechazado`. La condición del desencadenador no puede leer la acción *Configuracion*, por eso lleva los nombres escritos. Si el dominio cambia, hay que actualizarla (la prueba de arquitectura lo detecta).

## Paso 2 · Acciones de configuración

#### P1 · `Configuracion`
Redactar. Pegue el contenido completo de `salida/Configuracion-Flujos.json` (el mismo de LEG-01).

#### P2 · `Plantilla_aprobada`
Redactar. Contenido de `plantillas-correo/cierre-aprobada.html`.

#### P3 · `Plantilla_rechazada`
Redactar. Contenido de `plantillas-correo/cierre-rechazada.html`.

#### P4 · `Plantilla_error`
Redactar. Contenido de `plantillas-correo/error-tecnico.html`.

## Paso 3 · Ámbito `Intentar`

Agregue un **Ámbito** llamado `Intentar`. Dentro, en orden:

#### K1 · `Estado_cierre`
Redactar:
```expresion
triggerBody()?['Estado']?['Value']
```

#### K2 · `Id_legalizacion`
Redactar:
```expresion
triggerBody()?['Title']
```

#### K3 · `Anio`
Redactar. Año de recepción, leído del ID (`LEG-202609-00123` → `2026`):
```expresion
substring(outputs('Id_legalizacion'), add(length(outputs('Configuracion')?['prefijoId']), 1), 4)
```

#### K4 · `Mes`
Redactar (`LEG-202609-00123` → `09`):
```expresion
substring(outputs('Id_legalizacion'), add(length(outputs('Configuracion')?['prefijoId']), 5), 2)
```

#### K5 · `Carpeta_colaborador_base`
Redactar. Misma limpieza que LEG-01 (A8), aplicada al Colaborador guardado en el registro:
```expresion
trim(take(trim(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(coalesce(triggerBody()?['Colaborador'], ''), '"', ''), '*', ''), ':', ''), '<', ''), '>', ''), '?', ''), '/', ''), '\', ''), '|', ''), '#', ''), '%', ''), '.', '')), int(outputs('Configuracion')?['largoMaximoCarpeta'])))
```

#### K6 · `Carpeta_colaborador`
Redactar:
```expresion
if(empty(outputs('Carpeta_colaborador_base')), outputs('Configuracion')?['nombreSinDato'], outputs('Carpeta_colaborador_base'))
```

#### K7 · `Carpeta_destino`
Redactar. Resultado: `Aprobadas/2026/09/Ana Pérez`, relativo a la biblioteca.
```expresion
concat(if(equals(outputs('Estado_cierre'), outputs('Configuracion')?['estadoAprobado']), outputs('Configuracion')?['carpetaAprobadas'], outputs('Configuracion')?['carpetaRechazadas']), '/', outputs('Anio'), '/', outputs('Mes'), '/', outputs('Carpeta_colaborador'))
```

#### K8 · `Crear_carpeta_destino`
**SharePoint › Crear nueva carpeta** (Create new folder). Sitio **Legalizaciones TC**, *Lista o biblioteca* **Legalizaciones_TC**.
- *Ruta de la carpeta* (Folder Path):
```expresion
outputs('Carpeta_destino')
```

#### K9 · `Archivos_del_caso`
**SharePoint › Obtener archivos (solo propiedades)** (Get files (properties only)). Sitio **Legalizaciones TC**, biblioteca **Legalizaciones_TC**.
- *Consulta de filtro*:
```expresion
concat('IdLegalizacion eq ''', outputs('Id_legalizacion'), '''')
```
- *Configurar ejecución posterior*: marque **es correcto** y **tiene errores** de `Crear_carpeta_destino`. Si la carpeta ya existía, esa acción falla y el flujo debe seguir igual.

#### K10 · `Archivos_por_mover`
Filtrar matriz. Solo los archivos que siguen en la carpeta inicial (Pendientes).
- *Desde*:
```expresion
body('Archivos_del_caso')?['value']
```
- Condición:
```expresion
@contains(item()?['{Path}'], concat('/', outputs('Configuracion')?['carpetaInicial'], '/'))
```

#### K11 · `Mover_archivos`
**Aplicar a cada uno** sobre `body('Archivos_por_mover')`. Dentro:

#### K11.1 · `Mover_archivo`
**SharePoint › Mover archivo** (Move file).
- *Dirección del sitio actual* y *Dirección del sitio de destino*: **Legalizaciones TC**.
- *Archivo que se va a mover* (File to Move):
```expresion
items('Mover_archivos')?['{Identifier}']
```
- *Carpeta de destino*:
```expresion
concat(outputs('Configuracion')?['rutaBiblioteca'], '/', outputs('Carpeta_destino'))
```
- *Si ya hay otro archivo*: **Error en esta acción** (Fail this action). Los nombres llevan el ID del caso, así que un choque de nombres indica un problema que TI debe revisar.

#### K12 · `Registrar_fecha_cierre`
**SharePoint › Actualizar elemento**, sitio y lista del registro. Nombre Archivo y Observaciones se dejan **vacíos** para no cambiarlos.

| Campo | Expresión |
|---|---|
| Id | `triggerBody()?['ID']` |
| ID Legalización | `triggerBody()?['Title']` |
| Fecha Recepción | `triggerBody()?['FechaRecepcion']` |
| Colaborador | `triggerBody()?['Colaborador']` |
| Correo | `triggerBody()?['Correo']` |
| Asunto | `triggerBody()?['Asunto']` |
| Estado Value | `outputs('Estado_cierre')` |
| Fecha Cierre | `utcNow()` |

#### K13 · `Si_notificar_cierre`
Condición: `outputs('Configuracion')?['notificarCierre']` **es igual a** `true`. Rama **Sí**:

#### K13.1 · `Observaciones_para_colaborador`
Filtrar matriz. Deja solo las líneas de Observaciones que el colaborador puede leer. Quita las líneas vacías y las **notas internas**: las que empiezan con `[Interno]`, como el aviso de posible duplicado o de error técnico, o las que Contabilidad escriba así.
- *Desde*:
```expresion
split(coalesce(triggerBody()?['Observaciones'], ''), decodeUriComponent('%0A'))
```
- Condición:
```expresion
@and(not(empty(trim(item()))), not(startsWith(trim(item()), outputs('Configuracion')?['prefijoNotaInterna'])))
```

#### K13.2 · `Cuerpo_cierre`
Redactar:
```expresion
replace(replace(replace(replace(if(equals(outputs('Estado_cierre'), outputs('Configuracion')?['estadoAprobado']), outputs('Plantilla_aprobada'), outputs('Plantilla_rechazada')), '{{COLABORADOR}}', coalesce(triggerBody()?['Colaborador'], '')), '{{ID}}', outputs('Id_legalizacion')), '{{OBSERVACIONES}}', if(empty(body('Observaciones_para_colaborador')), '(sin observaciones)', join(body('Observaciones_para_colaborador'), '<br>'))), '{{FECHA_CIERRE}}', convertFromUtc(utcNow(), outputs('Configuracion')?['zonaHorariaWindows'], 'dd/MM/yyyy HH:mm'))
```

#### K13.3 · `Enviar_resultado`
**Enviar un correo electrónico desde un buzón compartido (V2)**, o **Enviar un correo electrónico (V2)** si no hay buzón compartido.
- *Dirección del buzón original*: `outputs('Configuracion')?['buzonLegalizaciones']`
- *Para*: `triggerBody()?['Correo']`
- *Asunto*:
```expresion
concat('Legalización ', toLower(outputs('Estado_cierre')), ' - ', outputs('Id_legalizacion'))
```
- *Cuerpo* (vista de código): `outputs('Cuerpo_cierre')`

## Paso 4 · Ámbito `Capturar_error`

Debajo de `Intentar`, con *Configurar ejecución posterior* en **tiene errores** y **agotó el tiempo de espera**. Dentro:

#### L1 · `Acciones_fallidas`
Filtrar matriz. *Desde*:
```expresion
result('Intentar')
```
Condición:
```expresion
@equals(item()?['status'], 'Failed')
```

#### L2 · `Detalle_error`
Redactar:
```expresion
take(string(first(body('Acciones_fallidas'))), 3000)
```

#### L3 · `Url_ejecucion`
Redactar:
```expresion
concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'], '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])
```

#### L4 · `Cuerpo_error`
Redactar:
```expresion
replace(replace(replace(replace(replace(replace(outputs('Plantilla_error'), '{{FLUJO}}', 'LEG-02 Cierre de casos'), '{{ASUNTO}}', coalesce(triggerBody()?['Title'], '(sin ID)')), '{{REMITENTE}}', coalesce(triggerBody()?['Correo'], '(desconocido)')), '{{FECHA}}', convertFromUtc(utcNow(), outputs('Configuracion')?['zonaHorariaWindows'], 'dd/MM/yyyy HH:mm')), '{{DETALLE}}', outputs('Detalle_error')), '{{ENLACE_EJECUCION}}', outputs('Url_ejecucion'))
```

#### L5 · `Avisar_error_TI`
Acción de envío.
- *Dirección del buzón original*: `outputs('Configuracion')?['buzonLegalizaciones']`
- *Para*: `outputs('Configuracion')?['correoSoporteTI']`
- *Asunto*:
```expresion
concat('[Legalizaciones] Error técnico en LEG-02 - ', coalesce(triggerBody()?['Title'], '(sin ID)'))
```
- *Cuerpo* (vista de código): `outputs('Cuerpo_error')`
- *Importancia*: **Alta**.

#### L6 · `Terminar_con_error`
**Terminar**. *Configurar ejecución posterior*: **es correcto** y **tiene errores** de L5.
- *Estado*: **Con errores**.
- *Código*: `LEG02_ERROR`.
- *Mensaje*:
```expresion
take(outputs('Detalle_error'), 1000)
```

## Paso 5 · Guardar y probar

1. Guardar.
2. En SharePoint, cambie a **Aprobada** el caso de la prueba E01 y guarde (escenario E12 del plan de pruebas).
3. En uno o dos minutos compruebe:
   - Fecha Cierre registrada;
   - archivos en `Aprobadas/AAAA/MM/Colaborador`;
   - correo "Legalización aprobada" al remitente.
4. Repita con **Rechazada** y un motivo en Observaciones (escenario E13).

## Solución de problemas

| Síntoma | Causa probable | Qué hacer |
|---|---|---|
| No se dispara al aprobar | El caso ya tenía Fecha Cierre, o el estado no es exactamente "Aprobada" / "Rechazada" | Vacíe Fecha Cierre y guarde de nuevo. Revise la condición del desencadenador |
| Falla `Anio` o `Mes` | El ID no tiene el formato `LEG-AAAAMM-NNNNN` (por ejemplo `TMP-…` o `ERR-…` de una recepción con error) | Corrija el caso con TI antes de cerrarlo |
| `Archivos_del_caso` vacío | Los archivos no tienen la columna ID Legalización (se subieron a mano) | Complete la columna en la biblioteca y vuelva a guardar el caso |
| Falla `Mover_archivo` por nombre repetido | Ya existe un archivo igual en el destino | TI revisa ambos archivos; luego guarde de nuevo el caso |
