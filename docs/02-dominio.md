# Dominio: reglas de negocio

Fuente única: [`src/1-dominio/dominio.json`](../src/1-dominio/dominio.json). Implementación de referencia: [`Legalizaciones.Dominio.psm1`](../src/1-dominio/Legalizaciones.Dominio.psm1). Ejemplos verificados: [`pruebas/unitarias/Dominio.Tests.ps1`](../pruebas/unitarias/Dominio.Tests.ps1).

## Lenguaje común

| Término | Significado |
|---|---|
| **Legalización (caso)** | Un correo de un colaborador con los soportes de sus gastos de tarjeta. Es un elemento del registro con un ID `LEG-AAAAMM-NNNNN` |
| **Colaborador** | Quien envía la legalización. Su nombre viene del asunto |
| **Factura** | Soporte fiscal en **PDF**, obligatorio. En Fase 1 no se lee su contenido |
| **Comprobante quincenal** | Excel (**.xlsx**) con una fila por gasto, hecho con la plantilla. Obligatorio. Excel lo lee con Power Query |
| **XML (CFDI)** | XML de la factura. Opcional en Fase 1; se guarda para validaciones fiscales futuras |
| **Hallazgo** | Algo que el colaborador debe corregir: documento faltante, formato no válido o asunto sin nombre |
| **Posible duplicado** | Sospecha heurística de que el caso repite otro. Es un aviso interno para Contabilidad |
| **Cierre** | Paso a Aprobada o Rechazada. Mueve los archivos y registra la Fecha Cierre |

## Entidad: Caso de legalización

| Atributo | Columna de SharePoint (interna) | Regla |
|---|---|---|
| ID Legalización | `Title` | R-09. Único |
| Fecha Recepción | `FechaRecepcion` | Fecha del correo; se guarda en UTC y se muestra en la hora local (R-10) |
| Colaborador | `Colaborador` | R-02 |
| Correo | `Correo` | Remitente, en minúsculas |
| Asunto | `Asunto` | — |
| Nombre Archivo | `NombreArchivo` | Nombres originales de los adjuntos guardados, uno por línea |
| Estado | `Estado` | R-11 |
| Observaciones | `Observaciones` | Hallazgos, aviso de duplicado o motivo de rechazo |
| Fecha Cierre | `FechaCierre` | La registra LEG-02 al cerrar |

## Estados y transiciones

```text
                 ┌──────────────┐
  correo nuevo ─►│  Pendiente   │◄──────┐
                 └──────┬───────┘       │ (volver a pendiente,
                        │               │  p. ej. esperando datos)
                        ▼               │
                 ┌──────────────┐       │
                 │ En revisión  ├───────┘
                 └──┬────────┬──┘
                    │        │      (también se puede cerrar directo desde Pendiente)
                    ▼        ▼
            ┌──────────┐  ┌───────────┐
            │ Aprobada │  │ Rechazada │   estados FINALES: no se reabren
            └──────────┘  └───────────┘
```

SharePoint no puede bloquear transiciones sin código. Las reglas se cumplen por proceso y con controles: LEG-02 solo actúa al pasar a un estado final, y la alerta "Caso reabierto" del Excel detecta un caso con Fecha Cierre que volvió a un estado abierto.

## Reglas

| Id | Regla | Clave en `dominio.json` | Dónde se aplica | Referencia |
|---|---|---|---|---|
| R-01 | Es una legalización si el asunto contiene "legalizacion tarjeta credito", sin importar mayúsculas, tildes en á é í ó ú ni espacios dobles | `correo.textoClaveAsunto` | Condición del desencadenador de LEG-01 (E-01) | `Test-AsuntoLegalizacion` |
| R-02 | El colaborador es el texto entre el primer `" - "` y el siguiente. Si no hay, se usa la parte del correo antes de la @ y se agrega un hallazgo | `correo.separadorNombre` | LEG-01 A6, A7, C4 (E-02, E-03) | `Get-NombreColaborador` |
| R-03 | Tipos de documento por extensión: factura `.pdf`, comprobante `.xlsx`, XML `.xml`. Todo lo demás no es válido y no se guarda | `documentos.*.extensiones` | LEG-01 B2 a B5 (E-06) | `Get-TipoDocumento` |
| R-04 | La documentación está completa si hay al menos una factura y un comprobante. El XML es opcional | `documentos.*.obligatorio` | LEG-01 C5 (E-07) | `Test-Documentacion` |
| R-05 | Las imágenes incrustadas en el cuerpo del correo (firmas) no son adjuntos | — | LEG-01 B1 | — |
| R-06 | Un archivo no válido de más no impide que la documentación esté completa, pero se informa | `mensajes.archivosNoSoportados` | LEG-01 C3 | `Test-Documentacion` |
| R-07 | Posible duplicado: otro caso del **mismo remitente**, de los **últimos 45 días**, **no rechazado**, con el **mismo nombre de archivo PDF o XML** (el Excel no cuenta). Es un aviso solo para Contabilidad | `duplicados.*`, `documentos.*.comparaDuplicados` | LEG-01 bloque E (E-10, E-11) | `Find-PosiblesDuplicados` |
| R-08 | Carpeta del caso: `<carpeta del estado>/AAAA/MM/<colaborador>`, con el mes **local** de recepción y el nombre sin `" * : < > ? / \ \| # % .`, de hasta 60 caracteres | `carpetas.*` | LEG-01 A4 a A10 y LEG-02 K3 a K7 (E-04, E-05) | `Get-RutaCarpetaCaso` |
| R-09 | ID: `LEG-` + AAAAMM local + `-` + número interno de SharePoint con 5 dígitos. Es único | `caso.prefijoId`, `caso.digitosConsecutivo` | LEG-01 D3 (E-08); valores únicos en SharePoint | `New-IdLegalizacion` |
| R-10 | Las fechas se guardan en UTC y se muestran en la hora del país (`Config.Pais`) | — | LEG-01 A3; Excel `fnUtcALocal` | `ConvertTo-FechaLocal` |
| R-11 | Estados y transiciones del diagrama; Aprobada y Rechazada son finales | `caso.estados`, `caso.transiciones` | SharePoint (opciones), LEG-02 (E-12), Excel (alertas) | `Test-TransicionEstado` |
| R-12 | Un caso abierto hace más de 5 días genera una alerta | `analisis.diasAlertaCasoAbierto` | Excel, hoja Alertas | — |
| R-13 | Las líneas de Observaciones que empiezan con `[Interno]` son notas internas: nunca se envían al colaborador. El aviso de duplicado y el de error técnico lo son | `observaciones.prefijoNotaInterna` | LEG-02 K13.1 (E-13) | `Get-ObservacionesParaColaborador` |

## Mensajes

Los textos que recibe el colaborador están en `dominio.json` → `mensajes`, con marcadores `{archivos}`, `{formatos}`, `{casos}` y `{dias}`. Cambiarlos no exige tocar los flujos: basta con volver a generar `salida/Configuracion-Flujos.json` y pegarlo en *Configuracion*.

## Cómo cambiar una regla

1. Edite `src/1-dominio/dominio.json`.
2. Si la regla cambia su comportamiento, actualice los ejemplos de `Dominio.Tests.ps1` y la función de `Legalizaciones.Dominio.psm1`.
3. Ejecute `pruebas/Invoke-Pruebas.ps1`. Las pruebas de arquitectura indican qué adaptador quedó desalineado.
4. Vuelva a generar lo derivado:
   - `Generar-ConfiguracionFlujos.ps1`, y pegar el JSON en ambos flujos;
   - `Generar-LibroAnalisis.ps1`;
   - `01-Aprovisionar-SharePoint.ps1`, si cambió algo de SharePoint.
5. Si cambian los nombres de los estados, además hay que actualizar a mano las opciones de la columna Estado y las condiciones de los desencadenadores (E-01 y E-12).
