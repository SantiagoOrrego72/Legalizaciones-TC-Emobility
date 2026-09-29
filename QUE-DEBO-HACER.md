# Qué debo hacer

Todo lo que se podía construir y probar sin acceso a su tenant de Microsoft 365 **ya está hecho** (ver el final). Lo que queda necesita datos, permisos o acciones dentro de su Microsoft 365. Siga este orden.

## 1. Reunir estos datos

| Dato | Ejemplo | Dónde va | Quién lo tiene |
|---|---|---|---|
| Prefijo del tenant | `emobility` (de `https://emobility.sharepoint.com`) | `Tenant.Nombre` y `Tenant.DominioEntra` | TI |
| Correo de las personas de TI propietarias | `ti@emobility.com` | `Permisos.PropietariosTI` | TI |
| **Id. de objeto** del grupo de Contabilidad en Entra ID (o los correos de cada persona) | `3f2a…-…` | `Permisos.Contabilidad` | TI (Entra ID › Grupos) |
| Cuenta de servicio para los flujos | `svc.legalizaciones@emobility.com` | `Permisos.CuentaServicioFlujo` | TI |
| Buzón de legalizaciones | `legalizaciones@emobility.com` | `Flujos.BuzonLegalizaciones` | TI / Exchange |
| Correo de avisos a Contabilidad | `contabilidad@emobility.com` | `Flujos.CorreoContabilidad` | Contabilidad |
| Correo de avisos a TI | `soporte.ti@emobility.com` | `Flujos.CorreoSoporteTI` | TI |

Todo se escribe en **un solo archivo**: `config/Config.Legalizaciones.psd1`.

## 2. Confirmar estas decisiones (valores por defecto entre paréntesis)

- [ ] **País de operación** (`Pais = 'Mexico'`). Lo asumí por las facturas CFDI. Si es Colombia: `Pais = 'Colombia'`.
- [ ] **URL del sitio** (`/sites/Legalizaciones-TC`), en un sitio nuevo dedicado.
- [ ] **Plantilla del comprobante quincenal** (`salida/Plantilla_Comprobante_Quincenal.xlsx`). Si Contabilidad usa hoy otro formato, compártamelo y ajusto la lectura (ADR-009).
- [ ] **Aviso al colaborador al aprobar o rechazar** (`Flujos.NotificarCierre = $true`) (ADR-007).
- [ ] **Columnas extra del registro** (desactivadas): Id Mensaje Outlook, Tiene Factura, etc. Opcionales.

## 3. Pedir a los administradores

| Qué | A quién | Para qué |
|---|---|---|
| Ejecutar `00-Registrar-AppEntraID.ps1` y dar el consentimiento | Administrador de Entra ID (aplicaciones o global) | Que los scripts puedan conectarse a SharePoint |
| Ejecutar `01-Aprovisionar-SharePoint.ps1`, o darle a usted el rol | Administrador de SharePoint | Crear el sitio y bloquear el uso compartido externo |
| Crear el buzón compartido y darle a la cuenta de servicio **Acceso total** y **Enviar como** | Administrador de Exchange | Que LEG-01 lea los correos y responda desde ese buzón |
| Licencia de Microsoft 365 para la cuenta de servicio | Administrador de Microsoft 365 | Que sea dueña de los flujos |
| Revisar las directivas de prevención de pérdida de datos (DLP) de Power Platform: Outlook y SharePoint en el mismo grupo | Administrador de Power Platform | Que Power Automate no bloquee los flujos |
| Directiva de retención sobre el sitio, con el plazo legal de conservación de soportes | Cumplimiento / TI, con Contabilidad | Conservar las facturas el tiempo que exige la ley |

## 4. Ejecutar, en orden

0. [ ] **Preparar el equipo de TI de la empresa.**
   - Descargar el proyecto al equipo desde GitHub (o copiar esta carpeta).
   - Doble clic en `Diagnosticar-Equipo.cmd`.
   - Si falta algo: `Instalar-Prerrequisitos.cmd -Perfil Administracion`.

   Qué se instala y qué no: [docs/05-instalacion-en-la-empresa.md](docs/05-instalacion-en-la-empresa.md).
1. [ ] Completar `config/Config.Legalizaciones.psd1` con los datos del punto 1.
2. [ ] Registrar la aplicación de Entra ID (administrador):
   ```powershell
   .\src\4-infraestructura\sharepoint\00-Registrar-AppEntraID.ps1
   ```
   Luego pegar el Id. de aplicación en `Autenticacion.ClientId`.
3. [ ] Crear SharePoint:
   ```powershell
   .\src\4-infraestructura\sharepoint\01-Aprovisionar-SharePoint.ps1
   ```
4. [ ] Verificar SharePoint (debe dar FALLA: 0):
   ```powershell
   .\src\4-infraestructura\sharepoint\02-Verificar-SharePoint.ps1
   ```
5. [ ] Generar la configuración de los flujos:
   ```powershell
   .\src\4-infraestructura\power-automate\Generar-ConfiguracionFlujos.ps1
   ```
6. [ ] Construir **LEG-01** en Power Automate con la cuenta de servicio: [guía paso a paso](src/3-adaptadores/power-automate/LEG-01-Recepcion.md).
7. [ ] Construir **LEG-02**: [guía](src/3-adaptadores/power-automate/LEG-02-Cierre.md).
8. [ ] Generar el libro de análisis conectado a su tenant y subirlo a *Documentos* del sitio:
   ```powershell
   .\src\4-infraestructura\excel\Generar-LibroAnalisis.ps1
   ```
9. [ ] Ejecutar las pruebas de aceptación: [plan de pruebas](pruebas/PLAN-DE-PRUEBAS.md).
10. [ ] Enviar a los colaboradores la plantilla y las [instrucciones](docs/04-instrucciones-colaboradores.md), y compartir con Contabilidad su [manual](docs/03-manual-contabilidad.md).

Los comandos se ejecutan en **PowerShell 7**: doble clic en `Abrir-PowerShell7.cmd`, que abre la consola ya ubicada en esta carpeta. La guía completa, con tiempos y detalles, está en el [README](README.md).

## Lo que ya está hecho

- **Instalado en este equipo**, sin permisos de administrador:
  - PowerShell 7.6.6 portátil, verificado con su SHA256 oficial;
  - PnP.PowerShell 3.4.1;
  - Pester 6.2.0;
  - openpyxl 3.1.5.
- **Punto 1:** scripts de SharePoint (registrar app, aprovisionar, verificar). Probados contra un SharePoint simulado, y con cada comando y parámetro PnP comprobado contra el módulo real.
- **Punto 2:**
  - `salida/Analisis_Legalizaciones_TC_DEMO.xlsx`: libro de análisis con 18 consultas de Power Query, en modo demostración;
  - `salida/Plantilla_Comprobante_Quincenal.xlsx`: plantilla para los colaboradores;
  - el generador del libro conectado a su tenant (paso 8).
- **Punto 3:** guías exactas de LEG-01 y LEG-02 (cada acción, nombre y expresión), catálogo de expresiones y 6 plantillas de correo.
- **Punto 4:**
  - plan de pruebas con 14 escenarios;
  - adjuntos de prueba generados en `pruebas/datos/`;
  - script de envío de correos y script de validación en SharePoint;
  - 187 pruebas automáticas que pasan.
- **Punto 5:** este archivo, el README, la arquitectura, el dominio, los casos de uso, los manuales y 10 decisiones de arquitectura (ADR).
