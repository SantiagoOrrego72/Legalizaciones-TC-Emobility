# Catálogo de expresiones de reglas de negocio

Las reglas de negocio se implementan dos veces:
- en **Power Automate**, que es lo que corre en producción, con las expresiones de este catálogo;
- en **`src/1-dominio/Legalizaciones.Dominio.psm1`**, la implementación de referencia que prueba `pruebas/unitarias/Dominio.Tests.ps1`.

Cada expresión indica su regla del dominio, la función equivalente y ejemplos **verificados** por las pruebas automáticas. Sirven para comprobar el flujo: si un correo con el asunto del ejemplo no da el resultado del ejemplo, la expresión está mal copiada.

La prueba `pruebas/unitarias/Arquitectura.Tests.ps1` comprueba que cada expresión de este catálogo aparece idéntica en LEG-01 o LEG-02, y que todos los paréntesis y comillas de todas las expresiones están balanceados.

---

## E-01 · ¿Es un correo de legalización?

- **Dónde:** condición del desencadenador de LEG-01.
- **Regla:** `correo.textoClaveAsunto` = `legalizacion tarjeta credito`. Se pasa a minúsculas, se quitan las tildes de á é í ó ú y se unen los espacios dobles.
- **Referencia:** `Test-AsuntoLegalizacion`.

```expresion
@contains(replace(replace(replace(replace(replace(replace(replace(toLower(coalesce(triggerBody()?['subject'], '')), 'á', 'a'), 'é', 'e'), 'í', 'i'), 'ó', 'o'), 'ú', 'u'), '  ', ' '), '  ', ' '), 'legalizacion tarjeta credito')
```

| Asunto | Resultado |
|---|---|
| `Legalización Tarjeta Crédito - Ana Pérez` | se procesa |
| `LEGALIZACIÓN TARJETA CRÉDITO - JUAN DÍAZ` | se procesa |
| `Legalizacion Tarjeta Credito - Luis Gómez` | se procesa |
| `RV: Legalización Tarjeta Crédito - Ana Pérez` | se procesa |
| `Legalización  Tarjeta Crédito - Ana` (doble espacio) | se procesa |
| `Legalización de gastos de septiembre` | no se procesa |
| `Tarjeta Crédito - legalización` | no se procesa |

## E-02 · Nombre del colaborador en el asunto

- **Dónde:** LEG-01, A6 `Nombre_en_asunto`.
- **Regla:** `correo.separadorNombre` = `" - "`. Se toma el texto entre el primer separador y el siguiente.
- **Referencia:** `Get-NombreEnAsunto`.

```expresion
trim(coalesce(split(outputs('Asunto'), ' - ')?[1], ''))
```

| Asunto | Resultado |
|---|---|
| `Legalización Tarjeta Crédito - Ana Pérez` | `Ana Pérez` |
| `RV: Legalización Tarjeta Crédito - Ana María Pérez-Gómez` | `Ana María Pérez-Gómez` |
| `Legalización Tarjeta Crédito - Ana Pérez - Quincena 1` | `Ana Pérez` |
| `Legalización Tarjeta Crédito` | vacío |
| `Legalización Tarjeta Crédito-Ana Pérez` (sin espacios) | vacío |

## E-03 · Colaborador del caso

- **Dónde:** LEG-01, A7 `Colaborador`.
- **Regla:** si el asunto no trae el nombre, se usa la parte del correo antes de la `@`.
- **Referencia:** `Get-NombreColaborador`. Ejemplo: asunto sin nombre y remitente `Ana.Perez@Emobility.com` → `ana.perez`.

```expresion
if(empty(outputs('Nombre_en_asunto')), first(split(outputs('Remitente'), '@')), outputs('Nombre_en_asunto'))
```

## E-04 · Nombre de carpeta seguro

- **Dónde:** LEG-01 A8/A9 y LEG-02 K5/K6.
- **Regla:** `carpetas.caracteresNoPermitidos` (`" * : < > ? / \ | # % .`), `carpetas.largoMaximoNombre` (60) y `carpetas.nombreSinDato`.
- **Referencia:** `ConvertTo-NombreCarpeta`.

```expresion
trim(take(trim(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(outputs('Colaborador'), '"', ''), '*', ''), ':', ''), '<', ''), '>', ''), '?', ''), '/', ''), '\', ''), '|', ''), '#', ''), '%', ''), '.', '')), int(outputs('Configuracion')?['largoMaximoCarpeta'])))
```

```expresion
if(empty(outputs('Carpeta_colaborador_base')), outputs('Configuracion')?['nombreSinDato'], outputs('Carpeta_colaborador_base'))
```

| Colaborador | Carpeta |
|---|---|
| `Ana Pérez` | `Ana Pérez` |
| `Ana P.` | `Ana P` |
| `Juan "El Rápido" Díaz` | `Juan El Rápido Díaz` |
| `María/José: #1` | `MaríaJosé 1` |
| `???` | `Sin_nombre` |

## E-05 · Mes local de recepción

- **Dónde:** LEG-01 A4/A5 (y A3 para mostrar la fecha).
- **Regla:** la carpeta AAAA/MM y el ID usan la fecha **local** del país (`Config.Pais` → `zonaHorariaWindows`).
- **Referencia:** `ConvertTo-FechaLocal`. Ejemplo: 1/sep 03:30 UTC → 31/ago 21:30 en México → carpeta `2026/08`.

```expresion
convertFromUtc(triggerBody()?['receivedDateTime'], outputs('Configuracion')?['zonaHorariaWindows'], 'MM')
```

## E-06 · Clasificación de adjuntos por extensión

- **Dónde:** LEG-01 B2 a B5.
- **Regla:** `documentos.*.extensiones`. La extensión es lo que va después del último punto, en minúsculas.
- **Referencia:** `Get-TipoDocumento`.

```expresion
@contains(outputs('Configuracion')?['extensiones']?['factura'], concat('.', toLower(last(split(item()?['name'], '.')))))
```

```expresion
@not(contains(outputs('Configuracion')?['extensionesPermitidas'], concat('.', toLower(last(split(item()?['name'], '.'))))))
```

| Archivo | Tipo |
|---|---|
| `Factura_123.PDF` | factura |
| `Gastos quincena.xlsx` | comprobante |
| `cfdi.XML` | xml |
| `Gastos.xls`, `Gastos.xlsm`, `foto_factura.jpg`, `archivo`, `factura.pdf.exe` | no soportado |

## E-07 · Documentación completa

- **Dónde:** LEG-01 C5.
- **Regla:** `documentos.factura.obligatorio` y `documentos.comprobante.obligatorio` son verdaderos; el XML no es obligatorio.
- **Referencia:** `Test-Documentacion`.

```expresion
and(greater(length(body('Facturas')), 0), greater(length(body('Comprobantes')), 0))
```

## E-08 · Identificador del caso

- **Dónde:** LEG-01 D3.
- **Regla:** `caso.prefijoId` + `-` + AAAAMM local + `-` + ID interno de SharePoint con `caso.digitosConsecutivo` dígitos.
- **Referencia:** `New-IdLegalizacion`. Ejemplos: `LEG-202609-00123`, `LEG-202601-00007`, `LEG-202609-123456`.

```expresion
concat(outputs('Configuracion')?['prefijoId'], '-', outputs('Anio'), outputs('Mes'), '-', formatNumber(outputs('Crear_registro')?['body/ID'], outputs('Configuracion')?['formatoConsecutivo']))
```

## E-09 · Año y mes desde el ID (cierre)

- **Dónde:** LEG-02 K3/K4.
- **Referencia:** `Get-PeriodoDesdeId`. Ejemplo: `LEG-202608-00031` → año `2026`, mes `08`.

```expresion
substring(outputs('Id_legalizacion'), add(length(outputs('Configuracion')?['prefijoId']), 1), 4)
```

```expresion
substring(outputs('Id_legalizacion'), add(length(outputs('Configuracion')?['prefijoId']), 5), 2)
```

## E-10 · Casos previos del remitente (duplicados)

- **Dónde:** LEG-01 E1 (consulta de filtro).
- **Regla:** `duplicados.diasVentana` (45) y `duplicados.estadosExcluidos` (Rechazada).

```expresion
concat('Correo eq ''', replace(outputs('Remitente'), '''', ''''''), ''' and FechaRecepcion ge ''', addDays(triggerBody()?['receivedDateTime'], mul(-1, int(outputs('Configuracion')?['diasVentanaDuplicado'])), 'yyyy-MM-ddTHH:mm:ssZ'), ''' and ID ne ', string(variables('varIdElemento')), ' and Estado ne ''', join(outputs('Configuracion')?['estadosExcluidosDuplicado'], ''' and Estado ne '''), '''')
```

## E-11 · ¿Comparte un archivo PDF o XML?

- **Dónde:** LEG-01 E4.
- **Regla:** `documentos.*.comparaDuplicados`. Las comparaciones no distinguen mayúsculas.
- **Referencia:** `Find-PosiblesDuplicados`. Con `Factura_001.pdf` y `gastos.xlsx`, de 7 casos previos solo 3 son posibles duplicados (ver la prueba "CU-03").

```expresion
@greater(length(intersection(split(replace(toLower(coalesce(item()?['NombreArchivo'], '')), decodeUriComponent('%0D'), ''), decodeUriComponent('%0A')), body('Nombres_comparables'))), 0)
```

## E-12 · Estados finales (cierre)

- **Dónde:** condición del desencadenador de LEG-02.
- **Regla:** `caso.estadoAprobado` y `caso.estadoRechazado`; se ejecuta una sola vez, porque exige Fecha Cierre vacía.
- **Referencia:** `Test-TransicionEstado`. Aprobada y Rechazada son finales.

```expresion
@and(or(equals(triggerBody()?['Estado']?['Value'], 'Aprobada'), equals(triggerBody()?['Estado']?['Value'], 'Rechazada')), equals(triggerBody()?['FechaCierre'], null))
```

## E-13 · Notas internas fuera del correo de cierre

- **Dónde:** LEG-02, K13.1 `Observaciones_para_colaborador`.
- **Regla:** `observaciones.prefijoNotaInterna` = `[Interno]`. El aviso de duplicado y el de error técnico empiezan así, y Contabilidad también puede escribir notas internas con ese prefijo.
- **Referencia:** `Get-ObservacionesParaColaborador`.

```expresion
@and(not(empty(trim(item()))), not(startsWith(trim(item()), outputs('Configuracion')?['prefijoNotaInterna'])))
```

| Observaciones del caso | El colaborador ve |
|---|---|
| `Falta el comprobante…` / `[Interno] Posible duplicado de LEG-…` / `Factura ilegible, envíe una copia legible.` | `Falta el comprobante…` y `Factura ilegible, envíe una copia legible.` |
| `[Interno] ERROR TÉCNICO…` | (sin observaciones) |
