<#
    Pruebas de los scripts de aprovisionamiento de SharePoint (Pester 5) contra un SharePoint
    simulado en memoria (SharePointSimulado.ps1). No necesitan tenant ni conexión.
    Cubren: validación de la configuración, primera ejecución, idempotencia, verificación,
    columnas opcionales, cambios manuales, sitio compartido y registro de la app.
#>

BeforeAll {
    $script:Origen = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
    . (Join-Path $PSScriptRoot 'SharePointSimulado.ps1')

    # Copia aislada del proyecto (misma estructura de carpetas) sin las líneas #Requires,
    # para que los scripts usen el SharePoint simulado y no el módulo PnP real.
    $script:Proyecto = Join-Path $TestDrive 'proyecto'
    foreach ($relativa in 'config', 'src\1-dominio', 'src\3-adaptadores\sharepoint', 'src\4-infraestructura\sharepoint') {
        $destino = Join-Path $Proyecto $relativa
        New-Item -ItemType Directory -Path $destino -Force | Out-Null
        Copy-Item (Join-Path $Origen "$relativa\*") $destino -Recurse -Force
    }
    $script:Scripts = Join-Path $Proyecto 'src\4-infraestructura\sharepoint'
    $utf8Bom = [System.Text.UTF8Encoding]::new($true)
    Get-ChildItem $Scripts -Filter *.ps1 | ForEach-Object {
        $texto = [System.IO.File]::ReadAllText($_.FullName) -replace '(?m)^#Requires.*$', ''
        [System.IO.File]::WriteAllText($_.FullName, $texto, $utf8Bom)
    }

    function script:New-ConfigPrueba([string]$Nombre, [hashtable]$Reemplazos = @{}) {
        $texto = [System.IO.File]::ReadAllText((Join-Path $Proyecto 'config\Config.Legalizaciones.psd1'))
        $base = [ordered]@{
            '<PREFIJO_TENANT>'               = 'contoso'
            '<CLIENT_ID>'                    = '11111111-2222-3333-4444-555555555555'
            '<correo.ti@dominio.com>'        = 'admin.ti@contoso.com'
            '<ID_OBJETO_GRUPO_CONTABILIDAD>' = 'AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE'
            '<cuenta.servicio@dominio.com>'  = 'svc.legalizaciones@contoso.com'
        }
        foreach ($clave in $base.Keys) { $texto = $texto.Replace($clave, $base[$clave]) }
        foreach ($clave in $Reemplazos.Keys) { $texto = $texto.Replace($clave, $Reemplazos[$clave]) }
        $ruta = Join-Path $Proyecto "config\$Nombre"
        [System.IO.File]::WriteAllText($ruta, $texto, [System.Text.UTF8Encoding]::new($true))
        return $ruta
    }
    function script:Invoke-Script([string]$Nombre, [hashtable]$Parametros = @{}) {
        $global:LASTEXITCODE = 0
        $salida = & (Join-Path $Scripts $Nombre) @Parametros 6>&1 | Out-String
        return [pscustomobject]@{ Salida = $salida; Codigo = $global:LASTEXITCODE }
    }
    function script:Get-Contador([string]$Salida, [string]$Estado) {
        $m = [regex]::Match($Salida, "(?m)^\s+$([regex]::Escape($Estado))\s+(\d+)\s*$")
        if ($m.Success) { return [int]$m.Groups[1].Value }
        return -1
    }
    function script:Get-Resumen([string]$Salida) {
        return [regex]::Match($Salida, 'OK: \d+\s+AVISO: \d+\s+FALLA: \d+').Value
    }

    # La sección Flujos se deja con marcadores: no debe impedir el aprovisionamiento.
    $script:CfgSitio = New-ConfigPrueba 'Config.Prueba.psd1'
    $script:CfgCompartido = New-ConfigPrueba 'Config.Compartido.psd1' @{ 'Crear        = $true' = 'Crear        = $false'; "Ambito = 'Sitio'" = "Ambito = 'ListaYBiblioteca'" }
    $script:CfgSinCuenta = New-ConfigPrueba 'Config.SinCuenta.psd1' @{ "CuentaServicioFlujo = 'svc.legalizaciones@contoso.com'" = "CuentaServicioFlujo = ''" }
    $script:CfgZonaAmbigua = New-ConfigPrueba 'Config.Zona.psd1' @{ "# ZonaHoraria = 'Guadalajara'" = "ZonaHoraria = '(UTC-06:00)'" }
    $script:CfgColombia = New-ConfigPrueba 'Config.Colombia.psd1' @{ "Pais = 'Mexico'" = "Pais = 'Colombia'" }
    $script:EsqCompleto = Join-Path $Proyecto 'src\3-adaptadores\sharepoint\Esquema.Completo.psd1'
    $esquema = [System.IO.File]::ReadAllText((Join-Path $Proyecto 'src\3-adaptadores\sharepoint\Esquema.Legalizaciones.psd1'))
    [System.IO.File]::WriteAllText($EsqCompleto, ($esquema -replace '(?m)^(\s*)# (@\{ NombreInterno)', '$1$2'), [System.Text.UTF8Encoding]::new($true))
    $script:EsqDistintoDominio = Join-Path $Proyecto 'src\3-adaptadores\sharepoint\Esquema.Distinto.psd1'
    [System.IO.File]::WriteAllText($EsqDistintoDominio, $esquema.Replace("'Pendiente', 'En revisión', 'Rechazada', 'Aprobada'", "'Pendiente', 'En revisión', 'Rechazada', 'Aprobada', 'Anulada'"), [System.Text.UTF8Encoding]::new($true))
}

Describe 'Validación de la configuración' {
    It 'rechaza la configuración con marcadores sin completar y dice cuáles' {
        { Invoke-Script '01-Aprovisionar-SharePoint.ps1' } | Should -Throw '*Config.Tenant.Nombre*'
    }
    It 'detecta un ámbito inválido y un GUID inválido' {
        $cfg = New-ConfigPrueba 'Config.Formato.psd1' @{ "Ambito = 'Sitio'" = "Ambito = 'Todo'"; 'AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE' = 'no-es-guid' }
        { Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $cfg } } | Should -Throw '*Permisos.Ambito*GUID*'
    }
    It 'rechaza un país no soportado' {
        $cfg = New-ConfigPrueba 'Config.Pais.psd1' @{ "Pais = 'Mexico'" = "Pais = 'Peru'" }
        { Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $cfg } } | Should -Throw '*Config.Pais*'
    }
    It 'se niega a aprovisionar si el esquema de SharePoint contradice los estados del dominio' {
        Reset-SP
        { Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgSitio; RutaEsquema = $EsqDistintoDominio } } | Should -Throw '*no coinciden con los estados del dominio*'
    }
}

Describe 'Sitio nuevo dedicado' {
    BeforeAll {
        Reset-SP
        $script:Primera = Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgSitio }
        $script:Lista = $global:SP.Listas['Registro_Legalizaciones_TC']
        $script:Biblioteca = $global:SP.Listas['Legalizaciones_TC']
    }
    It 'crea el sitio sin Grupo de Microsoft 365' {
        $global:SP.SitioExiste | Should -BeTrue
        ($global:SP.Log -match 'New-PnPSite') | Should -Not -BeNullOrEmpty
    }
    It 'aplica la zona horaria y el formato de México' {
        $global:SP.Web.RegionalSettings.TimeZone.Id | Should -Be 36
        $global:SP.Web.RegionalSettings.LocaleId | Should -Be 2058
    }
    It 'crea las carpetas raíz de la biblioteca' {
        ($Biblioteca.Carpetas -join ',') | Should -Be 'Pendientes,Aprobadas,Rechazadas'
    }
    It 'la biblioteca queda con versiones, sin desprotección obligatoria y con la columna ID Legalización indexada' {
        $Biblioteca.EnableVersioning | Should -BeTrue
        $Biblioteca.ForceCheckout | Should -BeFalse
        $Biblioteca.Campos['IdLegalizacion'].Title | Should -Be 'ID Legalización'
        $Biblioteca.Campos['IdLegalizacion'].Indexed | Should -BeTrue
    }
    It 'la vista de la biblioteca muestra el ID Legalización y ordena por nombre' {
        $vista = $Biblioteca.Vistas | Where-Object DefaultView
        $vista.ViewFields | Should -Contain 'IdLegalizacion'
        $vista.ViewQuery | Should -Match 'FileLeafRef'
    }
    It 'la lista queda con 100 versiones y sin adjuntos' {
        $Lista.EnableVersioning | Should -BeTrue
        $Lista.MajorVersionLimit | Should -Be 100
        $Lista.EnableAttachments | Should -BeFalse
    }
    It 'Title pasa a ser "ID Legalización", obligatoria, indexada y única' {
        $t = $Lista.Campos['Title']
        $t.Title | Should -Be 'ID Legalización'
        $t.Required | Should -BeTrue
        $t.Indexed | Should -BeTrue
        $t.EnforceUniqueValues | Should -BeTrue
    }
    It 'crea la columna <Nombre> (<Tipo>)' -ForEach @(
        @{ Nombre = 'FechaRecepcion'; Tipo = 'DateTime' }
        @{ Nombre = 'Colaborador'; Tipo = 'Text' }
        @{ Nombre = 'Correo'; Tipo = 'Text' }
        @{ Nombre = 'Asunto'; Tipo = 'Text' }
        @{ Nombre = 'NombreArchivo'; Tipo = 'Note' }
        @{ Nombre = 'Estado'; Tipo = 'Choice' }
        @{ Nombre = 'Observaciones'; Tipo = 'Note' }
        @{ Nombre = 'FechaCierre'; Tipo = 'DateTime' }
    ) {
        $Lista.Campos[$Nombre].TypeAsString | Should -Be $Tipo
    }
    It 'Estado tiene las opciones del dominio y Pendiente como valor predeterminado' {
        $xml = [xml]$Lista.Campos['Estado'].SchemaXml
        (@($xml.Field.CHOICES.CHOICE) -join '|') | Should -Be 'Pendiente|En revisión|Rechazada|Aprobada'
        $xml.Field.Default | Should -Be 'Pendiente'
    }
    It 'crea las vistas de trabajo y deja la predeterminada sin filtro' {
        $Lista.Vistas.Count | Should -Be 3
        ($Lista.Vistas | Where-Object DefaultView).ViewQuery | Should -Not -Match '<Where'
    }
    It 'Miembros queda con Colaborar (sin Edición) y contiene a Contabilidad y la cuenta del flujo' {
        $miembros = $global:SP.Grupos['Legalizaciones TC Miembros']
        ($miembros.Roles -join ',') | Should -Be 'Contributor'
        $miembros.Miembros.LoginName | Should -Contain 'c:0t.c|tenant|aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        $miembros.Miembros.Email | Should -Contain 'svc.legalizaciones@contoso.com'
    }
    It 'TI queda en Propietarios' {
        $global:SP.Grupos['Legalizaciones TC Propietarios'].Miembros.Email | Should -Contain 'admin.ti@contoso.com'
    }
    It 'bloquea el uso compartido externo y solo los propietarios comparten' {
        $global:SP.Sharing | Should -Be 'Disabled'
        $global:SP.Web.MembersCanShare | Should -BeFalse
    }
    It 'una segunda ejecución no cambia nada (idempotente)' {
        $segunda = Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgSitio }
        Get-Contador $segunda.Salida 'CREADO' | Should -Be 0
        Get-Contador $segunda.Salida 'ACTUALIZADO' | Should -Be 0
        Get-Contador $segunda.Salida 'AVISO' | Should -Be 0
    }
    It 'la verificación termina sin fallas ni avisos' {
        $r = Invoke-Script '02-Verificar-SharePoint.ps1' @{ RutaConfig = $CfgSitio }
        $r.Codigo | Should -Be 0
        Get-Resumen $r.Salida | Should -Match 'AVISO: 0\s+FALLA: 0'
    }
    It 'activa las 7 columnas recomendadas cuando se descomentan en el esquema' {
        $r = Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgSitio; RutaEsquema = $EsqCompleto }
        Get-Contador $r.Salida 'CREADO' | Should -Be 7
        $Lista.Campos['IdMensajeOutlook'].EnforceUniqueValues | Should -BeTrue
        $Lista.Campos['RevisadoPor'].TypeAsString | Should -Be 'User'
        $Lista.Campos['RutaCarpeta'].TypeAsString | Should -Be 'URL'
        $Lista.Campos['TieneFactura'].TypeAsString | Should -Be 'Boolean'
        (Invoke-Script '02-Verificar-SharePoint.ps1' @{ RutaConfig = $CfgSitio; RutaEsquema = $EsqCompleto }).Codigo | Should -Be 0
    }
}

Describe 'Cambios manuales en SharePoint' {
    BeforeAll {
        Reset-SP
        Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgSitio } | Out-Null
        $script:Lista = $global:SP.Listas['Registro_Legalizaciones_TC']
        $Lista.Campos['Correo'].Title = 'Email'
        $Lista.Campos['Estado'].SchemaXml = $Lista.Campos['Estado'].SchemaXml.Replace('En revisión', 'Revisando')
        $global:SP.Grupos['Legalizaciones TC Miembros'].Roles.Add('Editor')
        $global:SP.Grupos['Legalizaciones TC Visitantes'].Miembros.Add([pscustomobject]@{ Title = 'Alguien'; LoginName = 'i:0#.f|membership|alguien@contoso.com'; Email = 'alguien@contoso.com' })
    }
    It 'la verificación los detecta y termina con código 1' {
        $r = Invoke-Script '02-Verificar-SharePoint.ps1' @{ RutaConfig = $CfgSitio }
        $r.Codigo | Should -Be 1
        Get-Resumen $r.Salida | Should -Match 'FALLA: [1-9]'
    }
    It 'el aprovisionamiento corrige el nombre visible y quita Edición a Miembros' {
        $r = Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgSitio }
        $Lista.Campos['Correo'].Title | Should -Be 'Correo'
        ($global:SP.Grupos['Legalizaciones TC Miembros'].Roles -join ',') | Should -Be 'Contributor'
        $r.Salida | Should -Match 'Estado: en SharePoint las opciones son'
    }
    It 'no toca las opciones existentes de Estado (solo avisa)' {
        $Lista.Campos['Estado'].SchemaXml | Should -Match 'Revisando'
    }
    It 'se detiene si una columna existe con otro tipo, sin cambiarla' {
        $Lista.Campos['FechaCierre'].TypeAsString = 'Text'
        { Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgSitio } } | Should -Throw "*FechaCierre' de 'Registro_Legalizaciones_TC' ya existe con tipo 'Text'*"
        $Lista.Campos['FechaCierre'].TypeAsString | Should -Be 'Text'
    }
}

Describe 'Sitio compartido con otras áreas (Ambito = ListaYBiblioteca)' {
    BeforeAll {
        Reset-SP -SitioExiste
        $script:Primera = Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgCompartido }
        $script:Lista = $global:SP.Listas['Registro_Legalizaciones_TC']
        $script:Biblioteca = $global:SP.Listas['Legalizaciones_TC']
    }
    It 'no se conecta al centro de administración' {
        ($global:SP.Log -match '-admin\.sharepoint\.com') | Should -BeNullOrEmpty
    }
    It 'no cambia la zona horaria ni el uso compartido del sitio' {
        $global:SP.Web.RegionalSettings.TimeZone.Id | Should -Be 13
        $global:SP.Sharing | Should -Not -Be 'Disabled'
        $global:SP.Web.MembersCanShare | Should -BeTrue
    }
    It 'crea el grupo de Contabilidad y rompe la herencia de la lista y la biblioteca' {
        $global:SP.Grupos.Contains('Contabilidad - Legalizaciones TC') | Should -BeTrue
        $Lista.HasUniqueRoleAssignments | Should -BeTrue
        $Biblioteca.HasUniqueRoleAssignments | Should -BeTrue
        (@($Lista.Asignaciones | ForEach-Object { "$($_.Grupo)=$($_.Rol)" }) -join ';') | Should -Be 'Legalizaciones TC Propietarios=Control total;Contabilidad - Legalizaciones TC=Colaborar'
    }
    It 'una segunda ejecución no crea ni actualiza nada' {
        $r = Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgCompartido }
        Get-Contador $r.Salida 'CREADO' | Should -Be 0
        Get-Contador $r.Salida 'ACTUALIZADO' | Should -Be 0
    }
    It 'la verificación no tiene fallas' {
        (Invoke-Script '02-Verificar-SharePoint.ps1' @{ RutaConfig = $CfgCompartido }).Codigo | Should -Be 0
    }
}

Describe 'Otros casos' {
    It 'sin cuenta de servicio avisa y continúa' {
        Reset-SP
        $r = Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgSinCuenta }
        $r.Salida | Should -Match 'CuentaServicioFlujo está vacío'
        $global:SP.Grupos['Legalizaciones TC Miembros'].Miembros.Count | Should -Be 1
    }
    It 'rechaza una zona horaria que coincide con varias y lista las disponibles' {
        Reset-SP
        { Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgZonaAmbigua } } | Should -Throw '*coincide con 2 zonas horarias*'
    }
    It 'con Pais = Colombia aplica la zona de Bogotá y es-CO' {
        Reset-SP
        Invoke-Script '01-Aprovisionar-SharePoint.ps1' @{ RutaConfig = $CfgColombia } | Out-Null
        $global:SP.Web.RegionalSettings.TimeZone.Id | Should -Be 35
        $global:SP.Web.RegionalSettings.LocaleId | Should -Be 9226
    }
    It '00-Registrar-AppEntraID muestra el Id. de aplicación' {
        (Invoke-Script '00-Registrar-AppEntraID.ps1' @{ RutaConfig = $CfgSitio }).Salida | Should -Match '11111111-2222-3333-4444-555555555555'
    }
    It '00-Registrar-AppEntraID pide completar el dominio antes de registrar' {
        { Invoke-Script '00-Registrar-AppEntraID.ps1' } | Should -Throw '*DominioEntra*'
    }
}
