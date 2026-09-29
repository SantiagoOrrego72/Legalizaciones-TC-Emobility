# Instalación en la empresa: qué va en cada lugar

## La idea clave: el sistema no corre en ningún computador

La solución vive en la **nube de Microsoft 365**:

| Pieza | Dónde corre | ¿Depende de un computador encendido? |
|---|---|---|
| Recepción y cierre (flujos LEG-01 y LEG-02) | Power Automate (nube de Microsoft) | **No.** Funcionan 24/7 con todos los equipos apagados |
| Registro de casos y archivos | SharePoint Online | No |
| Correos de legalización | Exchange Online (buzón compartido) | No |
| Análisis | Libro de Excel guardado en SharePoint | Solo cuando Contabilidad lo actualiza |

Por eso **no se instala ningún servidor, servicio de Windows, tarea programada ni base de datos** en ningún computador. Esto es lo que lo hace de nivel empresarial: no hay un equipo que, si se apaga o se daña, detenga el proceso. Microsoft opera la infraestructura.

El computador de la empresa se usa **solo para instalar y mantener** la solución: ejecutar los scripts de SharePoint, generar el libro de Excel y la configuración de los flujos. Una vez instalada, puede apagarse sin afectar nada.

## Qué se instala en cada equipo

### 1. Equipo de administración de TI (uno solo)

El computador de la empresa desde el que TI instala y mantiene la solución.

| Componente | ¿Obligatorio? | Para qué | Cómo lo instala TI |
|---|---|---|---|
| **Windows 10 u 11** | Sí | El generador del libro de Excel usa Excel de escritorio | — |
| **PowerShell 7.4 o superior** (versión LTS recomendada) | Sí | Ejecutar los scripts de SharePoint (PnP) | `winget install --id Microsoft.PowerShell` como administrador, o por Intune/SCCM. Sin permisos: `Instalar-Prerrequisitos.cmd` instala la versión portátil |
| **Módulo PnP.PowerShell 3.x** | Sí | Crear, verificar y mantener SharePoint | `Instalar-Prerrequisitos.cmd -Perfil Administracion`, o `Install-Module PnP.PowerShell` desde la PowerShell Gallery (o el repositorio interno de módulos de la empresa) |
| **Microsoft 365 Apps (Excel de escritorio)** | Sí | Generar el libro de análisis con Power Query | Normalmente ya está en los equipos de la empresa |
| **Navegador** (Edge o Chrome) | Sí | Construir los flujos en Power Automate y administrar SharePoint | Ya instalado |
| Pester 5 o superior | No | Ejecutar las pruebas automáticas antes de cambiar algo | `Instalar-Prerrequisitos.cmd` (perfil Completo) |
| Python 3 + openpyxl | No | Regenerar la plantilla del comprobante o los datos de prueba. La plantilla ya está hecha en `salida/` | python.org + `pip install openpyxl` |
| Outlook clásico | No | Enviar automáticamente los correos de prueba | Incluido en Microsoft 365 Apps |
| Git | No (recomendado) | Guardar el proyecto con historial de cambios | `winget install Git.Git` |

Para comprobarlo todo de una vez, haga doble clic en **`Diagnosticar-Equipo.cmd`**. No cambia nada; muestra qué está bien y qué falta.

### 2. Equipos de Contabilidad

**No se instala nada nuevo.** Necesitan:
- navegador, para trabajar los casos en SharePoint;
- **Excel de escritorio** (Microsoft 365 Apps), para actualizar el libro de análisis: *Datos › Actualizar todo*;
- su cuenta de Microsoft 365, con acceso al sitio (lo da el script 01 por el grupo de Contabilidad).

### 3. Equipos de los colaboradores

**No se instala nada.** Solo envían un correo desde el Outlook que ya usan (escritorio, web o celular) con la plantilla del comprobante.

## Qué se configura en Microsoft 365 (no se instala)

Esto hace que el sistema funcione de forma profesional y sin depender de una persona:

| Tema | Qué se hace | Quién |
|---|---|---|
| **Cuenta de servicio** | Una cuenta dedicada (no de una persona), con licencia de Microsoft 365 que incluya Power Automate, dueña de los flujos y de sus conexiones. Su contraseña queda bajo custodia de TI | Administrador de Microsoft 365 |
| **Acceso condicional y MFA** | Si la cuenta de servicio tiene MFA o directivas de acceso condicional, cree las conexiones de los flujos una vez de forma interactiva. Revise que ninguna directiva bloquee sus inicios de sesión o fuerce reautenticaciones frecuentes | Seguridad / TI |
| **Buzón compartido** | `legalizaciones@…` (no necesita licencia). A la cuenta de servicio: **Acceso total** y **Enviar como** | Administrador de Exchange |
| **Aplicación de Entra ID para PnP** | La registra el script 00 con permisos delegados. Solo la usan los scripts de TI | Administrador de Entra ID |
| **Sitio de SharePoint y permisos** | Los crea el script 01: solo TI, Contabilidad y la cuenta de servicio; sin uso compartido externo | Administrador de SharePoint |
| **Directivas de prevención de pérdida de datos (DLP) de Power Platform** | Compruebe que *Office 365 Outlook* y *SharePoint* estén en el **mismo grupo** (normalmente "Empresarial"); si no, Power Automate bloquea los flujos | Administrador de Power Platform |
| **Entorno de Power Platform** | Para la Fase 1 sirve el entorno predeterminado. Si la empresa exige separar pruebas y producción, cree un entorno de producción y construya los flujos ahí | Administrador de Power Platform |
| **Copropietarios de los flujos** | Agregue al menos a una persona de TI como copropietaria de LEG-01 y LEG-02 | TI |
| **Monitoreo** | Los flujos avisan a TI por correo cuando fallan (ámbito Capturar_error). Además, Power Automate envía a los dueños un resumen de fallas, y el Excel tiene la hoja Alertas | Automático |
| **Retención de documentos** | SharePoint guarda versiones y 93 días de papelera. Para la obligación legal de conservar soportes contables, configure una **directiva de retención** de Microsoft Purview sobre el sitio, con el plazo que indique Contabilidad o Legal (en México, por regla general, 5 años) | Cumplimiento / TI |
| **Auditoría** | El registro de auditoría de Microsoft 365 guarda quién accede y modifica. Además, la lista tiene historial de versiones | Automático (verificar que esté activo) |

## Cómo llevar el proyecto al equipo de la empresa

1. Descargue el proyecto al equipo de TI desde el [repositorio de GitHub](https://github.com/SantiagoOrrego72/Legalizaciones-TC-Emobility):
   - con Git: `git clone https://github.com/SantiagoOrrego72/Legalizaciones-TC-Emobility.git`;
   - sin Git: **Code › Download ZIP** y descomprímalo.

   También puede copiar la carpeta **Legalizaciones-TC-Emobility** completa desde otro equipo. Puede omitir `logs/` y `pruebas/datos/`, que se regeneran.
2. Si la copió desde un correo, un ZIP o una descarga, desbloquee los archivos. En PowerShell, desde la carpeta:
   ```powershell
   Get-ChildItem -Recurse | Unblock-File
   ```
3. Haga doble clic en **`Diagnosticar-Equipo.cmd`** y resuelva lo que marque como FALTA.
4. Si falta PowerShell 7 o PnP, ejecute **`Instalar-Prerrequisitos.cmd -Perfil Administracion`**. Con permisos de administrador instala PowerShell 7 con winget; sin ellos, usa la versión portátil.
5. Siga el [README](../README.md) desde la fase 2 (configuración).

## Requisitos de red

Si la empresa usa proxy o firewall, el equipo de TI necesita salida HTTPS (puerto 443) a:

| Destino | Para qué |
|---|---|
| `login.microsoftonline.com` | Iniciar sesión (scripts y Power Automate) |
| `<prefijo>.sharepoint.com` y `<prefijo>-admin.sharepoint.com` | Scripts de SharePoint y Power Query |
| `graph.microsoft.com` | PnP PowerShell |
| `make.powerautomate.com` | Construir los flujos |
| `www.powershellgallery.com` | Instalar PnP.PowerShell y Pester (solo la primera vez) |
| `github.com` | Descargar el proyecto y, si se usa, la versión portátil de PowerShell 7 |

`Diagnosticar-Equipo.cmd` comprueba estos accesos, usando el proxy configurado en Windows.

## Política de ejecución de scripts

Si una directiva de grupo de la empresa exige **AllSigned**, los scripts del proyecto no se ejecutarán sin firma. Hay dos opciones:
- TI los firma con el certificado de firma de código de la empresa (`Set-AuthenticodeSignature`);
- o autoriza **RemoteSigned** para el equipo de administración.

El diagnóstico muestra la política efectiva y si la impone una directiva.

## Resumen para TI

- **Instalar solo en el equipo de administración de TI:** PowerShell 7 y PnP.PowerShell. Excel de escritorio ya suele estar.
- **Configurar en Microsoft 365:** cuenta de servicio, buzón compartido, aplicación de Entra ID, sitio de SharePoint (script), DLP, copropietarios y retención.
- **Nada en los equipos de Contabilidad ni de los colaboradores.**
- **Nada que dependa de un computador encendido:** todo corre en Microsoft 365.
