# Aprovisionamiento de SharePoint

Scripts que dejan listo SharePoint Online para la automatización. Todos se ejecutan **desde la raíz del proyecto, en PowerShell 7** (use `Abrir-PowerShell7.cmd`).

## Qué crea

| Elemento | Nombre | Detalle |
|---|---|---|
| Sitio | `/sites/Legalizaciones-TC` | Sitio de grupo **sin** Grupo de Microsoft 365, en español, con zona horaria y formato regional del país (`Pais`) |
| Biblioteca | `Legalizaciones_TC` | Carpetas `Pendientes`, `Aprobadas` y `Rechazadas`; columna **ID Legalización** (indexada); historial de versiones; sin desprotección obligatoria |
| Lista | `Registro_Legalizaciones_TC` | Los 9 campos acordados; historial de 100 versiones; sin adjuntos; índices |
| Vistas | `Todos los elementos`, `Pendientes de revisión`, `Cerradas` | La predeterminada va sin filtro (la usa Power Query) |
| Permisos | — | TI: Control total. Contabilidad y la cuenta del flujo: Colaborar. Nadie más |
| Uso compartido | — | Sin uso compartido externo; solo los propietarios comparten |

Las subcarpetas `AAAA/MM/Colaborador` las crea el flujo LEG-01 al guardar cada adjunto.

## Archivos

| Archivo | Para qué |
|---|---|
| `config/Config.Legalizaciones.psd1` | Datos del entorno (**se edita**) |
| `src/3-adaptadores/sharepoint/Esquema.Legalizaciones.psd1` | Contrato de la lista y la biblioteca: columnas, vistas y carpetas |
| `00-Registrar-AppEntraID.ps1` | Registra la aplicación de Entra ID con la que PnP inicia sesión (una vez por tenant) |
| `01-Aprovisionar-SharePoint.ps1` | Crea o corrige todo; se puede repetir sin riesgo |
| `02-Verificar-SharePoint.ps1` | Revisa todo sin cambiar nada; termina con código 1 si hay fallas |
| `Comun.Legalizaciones.ps1` | Funciones compartidas |

## Roles necesarios

| Paso | Rol |
|---|---|
| `00-Registrar-AppEntraID.ps1` | Administrador de aplicaciones en la nube o Administrador global (Entra ID), para el consentimiento |
| `01` con `Sitio.Crear = $true` o `DeshabilitarComparticionExterna = $true` | Administrador de SharePoint |
| `01` sobre un sitio existente sin esas opciones, y `02` | Propietario del sitio |

## Paso a paso

1. **Registrar la aplicación** (una vez por tenant; si la empresa ya tiene una aplicación de PnP PowerShell, use su Id. y salte este paso). Complete `Tenant.Nombre` y `Tenant.DominioEntra` en la configuración y ejecute:
   ```powershell
   .\src\4-infraestructura\sharepoint\00-Registrar-AppEntraID.ps1
   ```
   Acepte el consentimiento en el navegador y copie el Id. de aplicación en `Autenticacion.ClientId`.
2. **Completar la configuración:** todos los valores `< >` de las secciones `Tenant`, `Autenticacion`, `Sitio` y `Permisos`, y confirmar `Pais`. La sección `Flujos` se puede completar después.
3. **Aprovisionar:**
   ```powershell
   .\src\4-infraestructura\sharepoint\01-Aprovisionar-SharePoint.ps1
   ```
   Puede pedir iniciar sesión dos veces: una para el centro de administración y otra para el sitio. Cada paso muestra `CREADO`, `ACTUALIZADO`, `SIN CAMBIOS`, `AVISO` o `ERROR`. El registro queda en `logs/`.
4. **Verificar:**
   ```powershell
   .\src\4-infraestructura\sharepoint\02-Verificar-SharePoint.ps1
   ```
   Debe terminar con **FALLA: 0**. El detalle queda en `logs/verificacion_*.csv`.

## Contrato de datos

Estos **nombres internos** son los que usan Power Automate y Power Query. No se deben cambiar una vez en producción.

| Campo acordado | Nombre visible | Nombre interno | Tipo | Obligatorio | Índice |
|---|---|---|---|---|---|
| ID | ID Legalización | `Title` | Texto (255), **valores únicos** | Sí | Sí |
| Fecha Recepción | Fecha Recepción | `FechaRecepcion` | Fecha y hora | Sí | Sí |
| Colaborador | Colaborador | `Colaborador` | Texto (255) | Sí | No |
| Correo | Correo | `Correo` | Texto (255) | Sí | Sí |
| Asunto | Asunto | `Asunto` | Texto (255) | Sí | No |
| Nombre Archivo | Nombre Archivo | `NombreArchivo` | Varias líneas, texto sin formato | No | No |
| Estado | Estado | `Estado` | Opción: Pendiente, En revisión, Rechazada, Aprobada (predeterminada Pendiente) | Sí | Sí |
| Observaciones | Observaciones | `Observaciones` | Varias líneas, texto sin formato | No | No |
| Fecha Cierre | Fecha Cierre | `FechaCierre` | Fecha y hora | No | No |
| *(biblioteca)* | ID Legalización | `IdLegalizacion` | Texto | No | Sí |

Las razones de estas decisiones están en `docs/decisiones/` (ADR-003 y ADR-004).

## Volver a ejecutar y hacer cambios

- **Se puede ejecutar `01` las veces que haga falta.** Lo que ya está bien no se toca, lo que difiere se corrige y nunca borra datos.
- **No cambia el tipo de una columna existente:** se detiene y avisa.
- **No cambia las opciones de Estado** si ya existen distintas: solo avisa.
- **Se detiene si el esquema contradice al dominio**, por ejemplo si los estados del esquema no son los de `dominio.json`.
- **Agregar personas:** súmelas en la configuración y ejecute `01`. **Quitar personas:** a mano en *Configuración del sitio › Permisos del sitio*.
- **Columnas recomendadas** (IdMensajeOutlook, TieneFactura, TieneComprobante, TieneXML, PosibleDuplicado, RutaCarpeta, RevisadoPor): quite el `#` de las que quiera en el esquema y ejecute `01` y `02`.
- **Entorno de pruebas:** copie la configuración a `config/Config.Pruebas.psd1`, cambie `Sitio.RutaRelativa` y pase `-RutaConfig .\config\Config.Pruebas.psd1`.

## Solución de problemas

| Mensaje | Causa | Qué hacer |
|---|---|---|
| `Complete estos valores...` | Quedan marcadores `< >` | Completar los valores listados |
| `La configuración tiene errores...` | Un valor tiene un formato inválido o el esquema contradice al dominio | El mensaje dice cuál |
| `AADSTS700016` | `ClientId` incorrecto o de otro tenant | Revisar `Autenticacion.ClientId` |
| `AADSTS65001` | Falta el consentimiento de la aplicación | Un administrador: Entra ID › Aplicaciones empresariales › la aplicación › Permisos › Conceder consentimiento |
| `Access denied` o `403` al crear el sitio o en uso compartido | No es Administrador de SharePoint | Que lo ejecute TI, o crear el sitio a mano con `Sitio.Crear = $false` y `DeshabilitarComparticionExterna = $false` |
| El sitio ya existe o está en la papelera | Hubo un sitio con esa URL | Restaurarlo o eliminarlo desde el centro de administración, o usar otra `RutaRelativa` |
| `La zona horaria ... coincide con N zonas` | Una excepción mal escrita en `Sitio.ZonaHoraria` | Borre la excepción y use solo `Pais` |
| `No se pudo agregar ... al grupo` | El correo o el Id. del grupo no existen | Verificar en Entra ID |
| `La columna 'X' ya existe con tipo 'Y'` | Alguien la creó a mano | Corregirla a mano y volver a ejecutar |
| `running scripts is disabled` / `not digitally signed` | Política de ejecución o archivos bloqueados | `Get-ChildItem -Recurse *.ps1 \| Unblock-File` |
