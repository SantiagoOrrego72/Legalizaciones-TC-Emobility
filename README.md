# Legalización de gastos de tarjetas corporativas · Emobility · Fase 1

Automatización del proceso de legalización de gastos de tarjetas corporativas con **Microsoft 365**: Outlook, SharePoint, Excel Online y Power Automate. No usa IA, OCR, lectura automática de documentos ni licencias adicionales. La aprobación y la carga en Operam siguen siendo manuales.

**Para empezar, lea [`QUE-DEBO-HACER.md`](QUE-DEBO-HACER.md):** es la lista de lo que falta hacer, en orden y con responsables.

## Qué hace

1. El colaborador envía un correo con el asunto `Legalización Tarjeta Crédito - Nombre Apellido` y adjunta la factura (PDF), el comprobante quincenal (Excel) y, si lo tiene, el XML.
2. **LEG-01** (Power Automate):
   - detecta el correo, crea el caso `LEG-AAAAMM-NNNNN` en SharePoint (estado Pendiente) y guarda los adjuntos en `Pendientes/AAAA/MM/Colaborador`;
   - valida la documentación y posibles duplicados;
   - avisa al colaborador y, si hay un posible duplicado, a Contabilidad.
3. Contabilidad revisa en SharePoint y marca el caso **Aprobada** o **Rechazada**.
4. **LEG-02** mueve los archivos a Aprobadas o Rechazadas, registra la Fecha Cierre y avisa al colaborador.
5. En Excel, con Power Query, Contabilidad consolida los gastos aprobados, revisa las alertas y los carga en Operam.

## Estructura del proyecto (Clean Architecture)

```text
Legalizaciones-TC-Emobility/
├─ QUE-DEBO-HACER.md                   Lista de pasos pendientes (empiece aquí)
├─ README.md                           Esta guía de instalación
├─ Diagnosticar-Equipo.cmd             Revisa si el equipo tiene lo necesario (no cambia nada)
├─ Instalar-Prerrequisitos.cmd         Instala PowerShell 7 y PnP.PowerShell (y Pester y openpyxl en el perfil Completo)
├─ Abrir-PowerShell7.cmd               Abre PowerShell 7 en esta carpeta
├─ Ejecutar-Pruebas.cmd                Ejecuta las pruebas automáticas
├─ config/Config.Legalizaciones.psd1   ÚNICO archivo que se edita por entorno
├─ src/
│  ├─ 1-dominio/                       Reglas de negocio (fuente única) + implementación de referencia
│  ├─ 2-aplicacion/                    Casos de uso y puertos
│  ├─ 3-adaptadores/
│  │  ├─ sharepoint/                   Contrato de la lista y la biblioteca
│  │  ├─ power-automate/               Guías exactas de LEG-01 y LEG-02, expresiones, plantillas de correo
│  │  └─ excel/                        Consultas Power Query y contrato de la plantilla del comprobante
│  └─ 4-infraestructura/
│     ├─ sharepoint/                   Scripts PnP: registrar app, aprovisionar, verificar
│     ├─ power-automate/               Generador de la configuración de los flujos
│     ├─ excel/                        Generador del libro de análisis y de la plantilla
│     └─ herramientas/                 Instalador de prerrequisitos
├─ pruebas/                            Pruebas automáticas, plan de pruebas y escenarios de aceptación
├─ docs/                               Arquitectura, dominio, manuales y decisiones (ADR)
└─ salida/                             Archivos generados: libro de análisis, plantilla, configuración de los flujos
```

La arquitectura está explicada en [`docs/01-arquitectura.md`](docs/01-arquitectura.md).

---

## Instalación (orden exacto)

Tiempo estimado: 1 día de TI más la coordinación con los administradores. Los comandos se ejecutan **desde la raíz del proyecto, en PowerShell 7** (doble clic en `Abrir-PowerShell7.cmd`).

### Fase 0 · Reunir datos y permisos
Complete la sección 1 de [`QUE-DEBO-HACER.md`](QUE-DEBO-HACER.md): tenant, correos, grupo de Contabilidad, cuenta de servicio, buzón y país.

### Fase 1 · Preparar el equipo de administración de la empresa (15 minutos)
El sistema corre en Microsoft 365, no en un computador. El equipo de la empresa solo se usa para instalarlo y mantenerlo. Qué va en cada lugar: [`docs/05-instalacion-en-la-empresa.md`](docs/05-instalacion-en-la-empresa.md).

1. Descargue el proyecto al equipo de TI desde el [repositorio de GitHub](https://github.com/SantiagoOrrego72/Legalizaciones-TC-Emobility) y desbloquee los archivos: `Get-ChildItem -Recurse | Unblock-File`.
   - con Git: `git clone https://github.com/SantiagoOrrego72/Legalizaciones-TC-Emobility.git`;
   - sin Git: **Code › Download ZIP** y descomprímalo.
2. Doble clic en **`Diagnosticar-Equipo.cmd`**. Muestra qué está bien y qué falta.
3. Si falta algo, ejecute **`Instalar-Prerrequisitos.cmd -Perfil Administracion`**:
   - instala **PowerShell 7** con winget si es administrador, o la versión portátil si no;
   - instala el módulo **PnP.PowerShell**.
4. Opcional, para mantenimiento: `Instalar-Prerrequisitos.cmd` (perfil Completo, agrega Pester y openpyxl) y luego **`Ejecutar-Pruebas.cmd`**, que debe terminar con `RESULTADO: todas las pruebas pasan`.

### Fase 2 · Configuración
Edite `config/Config.Legalizaciones.psd1` y reemplace todos los valores entre `< >`. Cada sección está comentada. Revise `Pais` (México por defecto; Colombia disponible).

### Fase 3 · SharePoint (1 hora)
Guía detallada: [`src/4-infraestructura/sharepoint/LEEME.md`](src/4-infraestructura/sharepoint/LEEME.md).

```powershell
.\src\4-infraestructura\sharepoint\00-Registrar-AppEntraID.ps1
```
Solo una vez por tenant, y la ejecuta un administrador de Entra ID. Copie el Id. de aplicación en `Autenticacion.ClientId`.

```powershell
.\src\4-infraestructura\sharepoint\01-Aprovisionar-SharePoint.ps1
```
Requiere Administrador de SharePoint si crea el sitio.

```powershell
.\src\4-infraestructura\sharepoint\02-Verificar-SharePoint.ps1
```
Debe terminar con **FALLA: 0**.

### Fase 4 · Buzón y cuenta de servicio (administrador de Exchange / Microsoft 365)
1. Cuenta de servicio con licencia de Microsoft 365 (incluye Power Automate). Es la dueña de los flujos.
2. Buzón compartido, por ejemplo `legalizaciones@emobility.com`.
3. Sobre el buzón, a la cuenta de servicio: **Acceso total** (Full Access) y **Enviar como** (Send As).
4. Opcional: en el buzón, una carpeta **Procesadas** (paso opcional de LEG-01).

### Fase 5 · Configuración de los flujos
Complete la sección `Flujos` de la configuración y ejecute:
```powershell
.\src\4-infraestructura\power-automate\Generar-ConfiguracionFlujos.ps1
```
Genera `salida/Configuracion-Flujos.json`.

### Fase 6 · Flujos de Power Automate (medio día)
Con la cuenta de servicio en <https://make.powerautomate.com>, construya en este orden:
1. **LEG-01**, siguiendo [`src/3-adaptadores/power-automate/LEG-01-Recepcion.md`](src/3-adaptadores/power-automate/LEG-01-Recepcion.md). Pruébelo con el escenario E01.
2. **LEG-02**, siguiendo [`LEG-02-Cierre.md`](src/3-adaptadores/power-automate/LEG-02-Cierre.md). Pruébelo con E12.

Agregue un copropietario de TI a ambos flujos.

### Fase 7 · Libro de análisis (30 minutos)
```powershell
.\src\4-infraestructura\excel\Generar-LibroAnalisis.ps1
```
Requiere Excel de escritorio. Genera `salida/Analisis_Legalizaciones_TC.xlsx`, conectado al sitio del tenant.
1. Suba el libro a la biblioteca **Documentos** del sitio Legalizaciones TC.
2. Ábralo con Excel de escritorio, **Datos › Actualizar todo** e inicie sesión con una cuenta organizacional.
3. Si Excel pregunta por la privacidad, elija **Organizacional**.

Para conocer el libro sin SharePoint: `salida/Analisis_Legalizaciones_TC_DEMO.xlsx`, con datos de ejemplo.

### Fase 8 · Pruebas de aceptación (medio día)
Siga [`pruebas/PLAN-DE-PRUEBAS.md`](pruebas/PLAN-DE-PRUEBAS.md). Los escenarios cubren:
- legalización completa;
- incompleta;
- formato no soportado;
- posible duplicado;
- volumen alto;
- cierre de casos;
- análisis en Excel.

```powershell
python pruebas/aceptacion/generar_datos_prueba.py
```
```powershell
.\pruebas\aceptacion\Enviar-CorreosPrueba.ps1
```
```powershell
.\pruebas\aceptacion\Validar-Escenarios.ps1
```

### Fase 9 · Puesta en marcha
1. Complete la lista de verificación de la sección 8 del plan de pruebas.
2. Envíe a los colaboradores la plantilla `salida/Plantilla_Comprobante_Quincenal.xlsx` y las instrucciones de [`docs/04-instrucciones-colaboradores.md`](docs/04-instrucciones-colaboradores.md).
3. Contabilidad trabaja con [`docs/03-manual-contabilidad.md`](docs/03-manual-contabilidad.md).

---

## Pruebas automáticas

| Grupo | Qué comprueba | Cantidad |
|---|---|---|
| Dominio | Reglas de negocio con ejemplos verificados (los mismos que sirven para comprobar el flujo) | 74 |
| Arquitectura | Regla de dependencias; coherencia entre el dominio y SharePoint, los flujos, las plantillas y Power Query; expresiones balanceadas y sin errores de tipeo | 36 |
| Aprovisionamiento | Scripts de SharePoint contra un SharePoint simulado: creación, repetición sin cambios, verificación, cambios manuales, sitio compartido | 41 |
| Infraestructura | Comandos y parámetros PnP contra el módulo instalado | 4 |
| Excel | Power Query ejecutado de verdad en Excel: contrato de la plantilla y libro en modo Demostración y SharePoint | 18 |
| Python | Plantilla del comprobante, datos de prueba y contenido del libro (consultas M incrustadas) | 14 |

Ejecución: `Ejecutar-Pruebas.cmd`, o `.\pruebas\Invoke-Pruebas.ps1 [-SinExcel]`.

## Mantenimiento

| Necesito… | Cómo |
|---|---|
| Cambiar una regla (días de duplicados, mensajes, extensiones) | [`docs/02-dominio.md`](docs/02-dominio.md#cómo-cambiar-una-regla) |
| Agregar una persona de Contabilidad | Agregarla al grupo de Entra ID de Contabilidad, o a `Permisos.Contabilidad` y ejecutar `01` |
| Cambiar el país | `Pais` en la configuración, ejecutar `01`, regenerar la configuración de los flujos (pegarla en ambos) y regenerar el libro |
| Activar columnas extra del registro | Quitar el `#` en el esquema y ejecutar `01` y `02` |
| Ver por qué falló un correo | El caso tiene "ERROR TÉCNICO" en Observaciones y TI recibió el enlace a la ejecución |

## Documentación

- [`docs/recorrido/recorrido.html`](docs/recorrido/recorrido.html): recorrido interactivo con un simulador del proceso. Se abre con doble clic en el navegador.
- [`QUE-DEBO-HACER.md`](QUE-DEBO-HACER.md): pasos pendientes.
- [`docs/01-arquitectura.md`](docs/01-arquitectura.md): arquitectura y Clean Architecture aplicada.
- [`docs/02-dominio.md`](docs/02-dominio.md): reglas de negocio.
- [`src/2-aplicacion/CASOS-DE-USO.md`](src/2-aplicacion/CASOS-DE-USO.md): casos de uso y puertos.
- [`docs/05-instalacion-en-la-empresa.md`](docs/05-instalacion-en-la-empresa.md): qué se instala en cada equipo y qué se configura en Microsoft 365.
- [`docs/03-manual-contabilidad.md`](docs/03-manual-contabilidad.md): operación diaria.
- [`docs/04-instrucciones-colaboradores.md`](docs/04-instrucciones-colaboradores.md): comunicación a colaboradores.
- [`docs/decisiones/`](docs/decisiones/): decisiones de arquitectura (ADR).
- [`src/3-adaptadores/power-automate/`](src/3-adaptadores/power-automate/): flujos, catálogo de expresiones y plantillas de correo.
- [`pruebas/PLAN-DE-PRUEBAS.md`](pruebas/PLAN-DE-PRUEBAS.md): plan de pruebas.

## Fuera del alcance de la Fase 1

- Leer el contenido de los PDF: requiere OCR o IA (Fase 2: AI Builder o Azure Document Intelligence, como un adaptador nuevo).
- Validaciones fiscales con el XML (CFDI): en esta fase solo se guarda; Power Query podrá leerlo en una fase posterior.
- Integración automática con Operam: la carga sigue siendo manual.
