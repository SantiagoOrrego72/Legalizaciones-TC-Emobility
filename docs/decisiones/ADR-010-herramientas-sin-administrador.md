# ADR-010 · PowerShell 7 portátil y aplicación propia de Entra ID para PnP

**Estado:** aceptada.

## Contexto
PnP.PowerShell 3.x exige PowerShell 7.4 o superior. Windows trae Windows PowerShell 5.1 y el usuario del equipo no es administrador, así que no puede instalar el paquete MSI de PowerShell 7. Además, desde septiembre de 2024 PnP exige que cada organización registre su propia aplicación en Entra ID.

## Decisión
- `Instalar-Prerrequisitos.ps1` instala PowerShell 7 en su **versión portátil oficial**: el ZIP que publica Microsoft en GitHub, verificado con su SHA256, en `%LOCALAPPDATA%\Programs\PowerShell-7`. Si el usuario es administrador, usa winget.
- Los módulos (PnP.PowerShell, Pester) se instalan solo para el usuario actual desde la PowerShell Gallery. openpyxl se instala desde PyPI.
- `00-Registrar-AppEntraID.ps1` registra la aplicación con permisos **delegados** mínimos: SharePoint `AllSites.FullControl` y Graph `User.Read`. La aplicación actúa con los permisos de quien inicia sesión.

## Consecuencias
- Se puede preparar el equipo sin pedir permisos de administrador local.
- Registrar la aplicación y dar su consentimiento sí requiere un administrador de Entra ID, una vez por tenant.
- `Abrir-PowerShell7.cmd` abre el PowerShell 7 correcto, instalado o portátil, en la carpeta del proyecto.
