# SharePoint simulado en memoria para probar los scripts de aprovisionamiento sin tenant.
# Imita el comportamiento relevante de PnP.PowerShell 3.x y valida los parámetros que recibe.
# Lo usa pruebas/unitarias/Aprovisionamiento.Tests.ps1.

function global:Reset-SP {
    param([switch]$SitioExiste)
    $roles = @(
        [pscustomobject]@{ Name = 'Control total'; RoleTypeKind = 'Administrator' }
        [pscustomobject]@{ Name = 'Diseño'; RoleTypeKind = 'WebDesigner' }
        [pscustomobject]@{ Name = 'Edición'; RoleTypeKind = 'Editor' }
        [pscustomobject]@{ Name = 'Colaborar'; RoleTypeKind = 'Contributor' }
        [pscustomobject]@{ Name = 'Leer'; RoleTypeKind = 'Reader' }
    )
    $zonas = @(
        [pscustomobject]@{ Id = 13; Description = '(UTC-08:00) Hora del Pacífico (EE. UU. y Canadá)' }
        [pscustomobject]@{ Id = 36; Description = '(UTC-06:00) Guadalajara, Ciudad de México, Monterrey' }
        [pscustomobject]@{ Id = 35; Description = '(UTC-05:00) Bogotá, Lima, Quito' }
        [pscustomobject]@{ Id = 55; Description = '(UTC-06:00) América Central' }
    )
    $regional = [pscustomobject]@{ LocaleId = [uint32]1033; TimeZone = $zonas[0]; TimeZones = $zonas }
    $web = [pscustomobject]@{
        Title = 'Legalizaciones TC'; Url = 'https://contoso.sharepoint.com/sites/Legalizaciones-TC'
        Language = [uint32]3082; Description = ''; MembersCanShare = $true; RegionalSettings = $regional
    }
    $web | Add-Member -MemberType ScriptMethod -Name Update -Value { $global:SP.Log.Add('web.Update') }
    $global:SP = @{
        SitioExiste = [bool]$SitioExiste
        Sharing     = 'ExternalUserAndGuestSharing'
        Web         = $web
        Roles       = $roles
        Listas      = [ordered]@{}
        Grupos      = [ordered]@{}
        Log         = New-Object System.Collections.Generic.List[string]
        SiguienteId = 10
    }
    New-GrupoSP -Title 'Legalizaciones TC Propietarios' -Rol 'Administrator' -Asociado 'Owner' | Out-Null
    New-GrupoSP -Title 'Legalizaciones TC Miembros' -Rol 'Editor' -Asociado 'Member' | Out-Null
    New-GrupoSP -Title 'Legalizaciones TC Visitantes' -Rol 'Reader' -Asociado 'Visitor' | Out-Null
}

function global:New-GrupoSP {
    param($Title, $Rol, $Asociado)
    $id = $global:SP.SiguienteId
    $global:SP.SiguienteId++
    $g = [pscustomobject]@{
        Id = $id; Title = $Title; LoginName = $Title; Asociado = $Asociado
        Roles = New-Object System.Collections.Generic.List[string]
        Miembros = New-Object System.Collections.Generic.List[object]
    }
    if ($Rol) { $g.Roles.Add($Rol) }
    $global:SP.Grupos[$Title] = $g
    return $g
}

function global:Find-GrupoSP {
    param($Identity)
    if ($Identity -is [int]) { return $global:SP.Grupos.Values | Where-Object { $_.Id -eq $Identity } | Select-Object -First 1 }
    if ($Identity -isnot [string] -and $Identity.PSObject.Properties['Id']) { return Find-GrupoSP -Identity ([int]$Identity.Id) }
    return $global:SP.Grupos[[string]$Identity]
}

function global:New-ListaSP {
    param($Title, $Url, $BaseTemplate)
    $l = [pscustomobject]@{
        Title = $Title; Url = $Url; Id = [guid]::NewGuid(); BaseTemplate = $BaseTemplate
        Description = ''; EnableVersioning = ($BaseTemplate -eq 101); EnableMinorVersions = $false; MajorVersionLimit = [uint32]500
        ForceCheckout = $false; EnableAttachments = ($BaseTemplate -eq 100); EnableFolderCreation = ($BaseTemplate -eq 101)
        HasUniqueRoleAssignments = $false
        Campos = [ordered]@{}
        Vistas = New-Object System.Collections.Generic.List[object]
        Carpetas = New-Object System.Collections.Generic.List[string]
        Asignaciones = New-Object System.Collections.Generic.List[object]
    }
    if ($BaseTemplate -eq 100) {
        Add-CampoSP -Lista $l -InternalName 'Title' -Title 'Título' -Type 'Text' -Required $true -SchemaXml '<Field Type="Text" Name="Title" DisplayName="Título" Required="TRUE" />'
    }
    Add-CampoSP -Lista $l -InternalName 'Created' -Title 'Creado' -Type 'DateTime' -Required $false -SchemaXml '<Field Type="DateTime" Name="Created" Format="DateTime" />'
    $titulo = if ($BaseTemplate -eq 101) { 'Todos los documentos' } else { 'Todos los elementos' }
    $campos = if ($BaseTemplate -eq 101) { @('DocIcon', 'LinkFilename', 'Modified', 'Editor') } else { @('LinkTitle') }
    $l.Vistas.Add([pscustomobject]@{ Id = [guid]::NewGuid(); Title = $titulo; DefaultView = $true; ViewFields = $campos; ViewQuery = ''; RowLimit = [uint32]30; Paged = $true })
    $global:SP.Listas[$Title] = $l
    return $l
}

function global:Add-CampoSP {
    param($Lista, $InternalName, $Title, $Type, $Required, $SchemaXml)
    $Lista.Campos[$InternalName] = [pscustomobject]@{
        InternalName = $InternalName; Title = $Title; TypeAsString = $Type; Required = [bool]$Required
        Indexed = $false; EnforceUniqueValues = $false; Description = ''; SchemaXml = $SchemaXml
    }
}

function global:Assert-Bool { param($Nombre, $Valor) if ($Valor -isnot [bool]) { throw "Se esperaba Boolean en -$Nombre y llegó $($Valor.GetType().Name)" } }

# ---------------------------------------------------------------- conexión y sitio
function global:Connect-PnPOnline {
    param([string]$Url, [switch]$Interactive, [string]$ClientId, [switch]$ReturnConnection)
    $global:SP.Log.Add("Connect $Url")
    if (-not $Interactive) { throw 'Falta -Interactive' }
    if (-not $ClientId) { throw 'Falta -ClientId' }
    if ($Url -notlike '*-admin.sharepoint.com' -and -not $global:SP.SitioExiste) { throw "404: el sitio no existe ($Url)" }
    if ($ReturnConnection) { return [pscustomobject]@{ Url = $Url } }
}
function global:Get-PnPTenantSite {
    [CmdletBinding()] param([string]$Identity, $Connection)
    if (-not $Connection) { throw 'Get-PnPTenantSite sin conexión de administración' }
    if (-not $global:SP.SitioExiste) { throw "Cannot get site $Identity" }
    [pscustomobject]@{ Url = $Identity; SharingCapability = $global:SP.Sharing }
}
function global:New-PnPSite {
    param($Type, $Title, $Url, [uint32]$Lcid, [switch]$Wait, $Connection)
    if ($Type -ne 'TeamSiteWithoutMicrosoft365Group') { throw "Tipo inesperado $Type" }
    if (-not $Connection) { throw 'New-PnPSite sin conexión de administración' }
    if (-not $Wait) { throw 'Falta -Wait' }
    $global:SP.SitioExiste = $true
    $global:SP.Web.Language = $Lcid
    $global:SP.Log.Add("New-PnPSite $Url lcid=$Lcid")
}
function global:Set-PnPTenantSite {
    param($Identity, $SharingCapability, $Connection)
    if (-not $Connection) { throw 'Set-PnPTenantSite sin conexión de administración' }
    $global:SP.Sharing = [string]$SharingCapability
}
function global:Get-PnPWeb { param([string[]]$Includes) $global:SP.Web }
function global:Set-PnPWeb {
    param([string]$Description, [switch]$MembersCanShare)
    if ($PSBoundParameters.ContainsKey('Description')) { $global:SP.Web.Description = $Description }
    if ($PSBoundParameters.ContainsKey('MembersCanShare')) { $global:SP.Web.MembersCanShare = [bool]$MembersCanShare }
}
function global:Get-PnPProperty {
    param($ClientObject, [string[]]$Property)
    if ($Property.Count -ne 1) { return }
    if ($Property[0] -eq 'RoleAssignments') {
        if ([object]::ReferenceEquals($ClientObject, $global:SP.Web)) { return @(Get-AsignacionesWebSP) }
        return @(Get-AsignacionesListaSP -Lista $ClientObject)
    }
    return $ClientObject.($Property[0])
}
function global:Invoke-PnPQuery { $global:SP.Log.Add('Invoke-PnPQuery') }
function global:Get-PnPRoleDefinition { $global:SP.Roles }

# ---------------------------------------------------------------- listas y carpetas
function global:Get-PnPList {
    param($Identity, [string[]]$Includes)
    if ($null -eq $Identity) { return $global:SP.Listas.Values }
    return $global:SP.Listas[[string]$Identity]
}
function global:New-PnPList {
    param($Title, $Url, $Template, [switch]$OnQuickLaunch)
    $bt = switch ($Template) { 'DocumentLibrary' { 101 } 'GenericList' { 100 } default { throw "Plantilla $Template" } }
    New-ListaSP -Title $Title -Url $Url -BaseTemplate $bt
}
function global:Set-PnPList {
    param($Identity, $Description, $EnableVersioning, $EnableMinorVersions, $ForceCheckout, [uint32]$MajorVersions,
          $EnableAttachments, $EnableFolderCreation, [switch]$BreakRoleInheritance, [switch]$CopyRoleAssignments, [switch]$ClearSubScopes)
    $l = $global:SP.Listas[[string]$Identity]
    if (-not $l) { throw "No existe la lista $Identity" }
    foreach ($p in 'EnableVersioning', 'EnableMinorVersions', 'ForceCheckout', 'EnableAttachments', 'EnableFolderCreation') {
        if ($PSBoundParameters.ContainsKey($p)) { Assert-Bool $p $PSBoundParameters[$p]; $l.$p = $PSBoundParameters[$p] }
    }
    if ($PSBoundParameters.ContainsKey('Description')) { $l.Description = $Description }
    if ($PSBoundParameters.ContainsKey('MajorVersions')) { $l.MajorVersionLimit = $MajorVersions }
    if ($BreakRoleInheritance) { $l.HasUniqueRoleAssignments = $true; if (-not $CopyRoleAssignments) { $l.Asignaciones.Clear() } }
}
function global:Get-PnPFolderItem {
    param($FolderSiteRelativeUrl, $ItemType, $ItemName)
    $l = $global:SP.Listas.Values | Where-Object { $_.Url -eq $FolderSiteRelativeUrl } | Select-Object -First 1
    if (-not $l) { throw "No existe la carpeta $FolderSiteRelativeUrl" }
    $items = @($l.Carpetas | ForEach-Object { [pscustomobject]@{ Name = $_ } })
    if ($ItemName) { $items = @($items | Where-Object { $_.Name -eq $ItemName }) }
    $items
}
function global:Resolve-PnPFolder {
    param($SiteRelativePath)
    $partes = $SiteRelativePath -split '/'
    $l = $global:SP.Listas.Values | Where-Object { $_.Url -eq $partes[0] } | Select-Object -First 1
    if (-not $l) { throw "No existe la biblioteca $($partes[0])" }
    if ($l.Carpetas -notcontains $partes[1]) { $l.Carpetas.Add($partes[1]) }
    [pscustomobject]@{ Name = $partes[1] }
}

# ---------------------------------------------------------------- columnas
function global:Get-PnPField {
    param($List, $Identity)
    $l = $global:SP.Listas[[string]$List]
    if ($Identity) {
        $c = $l.Campos[[string]$Identity]
        if (-not $c) { throw "Campo $Identity no existe" }
        return $c
    }
    $l.Campos.Values
}
function global:Add-PnPFieldFromXml {
    param($List, $FieldXml)
    $l = $global:SP.Listas[[string]$List]
    $x = [xml]$FieldXml
    $n = $x.Field.Name
    if ($l.Campos.Contains($n)) { throw "Ya existe la columna $n" }
    if ($x.Field.DisplayName -ne $n -or $x.Field.StaticName -ne $n) { throw 'Se esperaba crear la columna con DisplayName = StaticName = Name' }
    Add-CampoSP -Lista $l -InternalName $n -Title $x.Field.DisplayName -Type $x.Field.Type -Required ($x.Field.Required -eq 'TRUE') -SchemaXml $FieldXml
    $l.Campos[$n].Description = $x.Field.Description
    $l.Campos[$n]
}
function global:Set-PnPField {
    param($List, $Identity, [hashtable]$Values)
    $c = Get-PnPField -List $List -Identity $Identity
    foreach ($k in $Values.Keys) {
        if (-not $c.PSObject.Properties[$k]) { throw "La columna no tiene la propiedad $k" }
        if ($k -eq 'EnforceUniqueValues' -and $Values[$k] -and -not $c.Indexed) { throw 'SharePoint: los valores únicos requieren que la columna esté indexada' }
        if ($k -eq 'Indexed' -and -not $Values[$k] -and $c.EnforceUniqueValues) { throw 'SharePoint: quite los valores únicos antes de quitar el índice' }
        $c.$k = $Values[$k]
    }
}

# ---------------------------------------------------------------- vistas
function global:Test-CamposVista {
    param($Lista, [string[]]$Campos)
    foreach ($f in $Campos) {
        if (@('LinkTitle', 'Editor', 'Modified', 'DocIcon', 'LinkFilename') -notcontains $f -and -not $Lista.Campos.Contains($f)) { throw "La vista usa una columna que no existe: $f" }
    }
}
function global:Get-PnPView { param($List, [string[]]$Includes) $global:SP.Listas[[string]$List].Vistas }
function global:Add-PnPView {
    param($List, $Title, [string[]]$Fields, $Query, [uint32]$RowLimit, [switch]$Paged)
    $l = $global:SP.Listas[[string]$List]
    Test-CamposVista -Lista $l -Campos $Fields
    [xml]"<Query>$Query</Query>" | Out-Null
    $l.Vistas.Add([pscustomobject]@{ Id = [guid]::NewGuid(); Title = $Title; DefaultView = $false; ViewFields = $Fields; ViewQuery = $Query; RowLimit = $RowLimit; Paged = [bool]$Paged })
}
function global:Set-PnPView {
    param($List, $Identity, [string[]]$Fields, [hashtable]$Values)
    $l = $global:SP.Listas[[string]$List]
    $v = $l.Vistas.ToArray() | Where-Object { $_.Id -eq $Identity -or $_.Title -eq $Identity } | Select-Object -First 1
    if (-not $v) { throw "No existe la vista $Identity" }
    if ($Fields) { Test-CamposVista -Lista $l -Campos $Fields; $v.ViewFields = $Fields }
    foreach ($k in $Values.Keys) {
        if ($k -eq 'ViewQuery') { [xml]"<Query>$($Values[$k])</Query>" | Out-Null }
        $v.$k = $Values[$k]
    }
}

# ---------------------------------------------------------------- grupos y permisos
function global:Get-PnPGroup {
    param($Identity, [switch]$AssociatedOwnerGroup, [switch]$AssociatedMemberGroup, [switch]$AssociatedVisitorGroup)
    if ($AssociatedOwnerGroup) { return $global:SP.Grupos.Values | Where-Object { $_.Asociado -eq 'Owner' } }
    if ($AssociatedMemberGroup) { return $global:SP.Grupos.Values | Where-Object { $_.Asociado -eq 'Member' } }
    if ($AssociatedVisitorGroup) { return $global:SP.Grupos.Values | Where-Object { $_.Asociado -eq 'Visitor' } }
    if ($Identity) { $g = Find-GrupoSP $Identity; if (-not $g) { throw 'Group cannot be found' }; return $g }
    $global:SP.Grupos.Values
}
function global:Get-PnPGroupPermissions {
    param($Identity)
    $g = Find-GrupoSP $Identity
    if ($g.Roles.Count -eq 0) { throw 'Can not find the principal with id' }
    foreach ($k in $g.Roles) { $global:SP.Roles | Where-Object { $_.RoleTypeKind -eq $k } }
}
function global:Set-PnPGroupPermissions {
    param($Identity, [string[]]$AddRole, [string[]]$RemoveRole)
    $g = Find-GrupoSP $Identity
    foreach ($r in $AddRole) {
        $k = ($global:SP.Roles | Where-Object { $_.Name -eq $r }).RoleTypeKind
        if (-not $k) { throw "El nivel '$r' no existe en el sitio" }
        if (-not $g.Roles.Contains($k)) { $g.Roles.Add($k) }
    }
    foreach ($r in $RemoveRole) {
        $k = ($global:SP.Roles | Where-Object { $_.Name -eq $r }).RoleTypeKind
        if (-not $k) { throw "El nivel '$r' no existe en el sitio" }
        [void]$g.Roles.Remove($k)
    }
}
function global:Get-PnPGroupMember { param($Group) (Find-GrupoSP $Group).Miembros }
function global:Add-PnPGroupMember {
    param([string]$LoginName, $Group)
    $g = Find-GrupoSP $Group
    if (-not $g) { throw "Grupo $Group no existe" }
    if ($LoginName -like '*|*') { $m = [pscustomobject]@{ Title = $LoginName; LoginName = $LoginName; Email = '' } }
    else { $m = [pscustomobject]@{ Title = $LoginName; LoginName = "i:0#.f|membership|$($LoginName.ToLowerInvariant())"; Email = $LoginName } }
    $g.Miembros.Add($m)
}
function global:New-PnPGroup { param($Title, $Description, $Owner) New-GrupoSP -Title $Title -Rol $null -Asociado $null }
function global:Set-PnPListPermission {
    param($Identity, $Group, $AddRole)
    $l = $global:SP.Listas[[string]$Identity]
    if (-not $l.HasUniqueRoleAssignments) { throw 'La lista hereda permisos: rompa la herencia primero' }
    if (-not ($global:SP.Roles | Where-Object { $_.Name -eq $AddRole })) { throw "El nivel '$AddRole' no existe en el sitio" }
    if (-not ($l.Asignaciones.ToArray() | Where-Object { $_.Grupo -eq $Group -and $_.Rol -eq $AddRole })) {
        $l.Asignaciones.Add([pscustomobject]@{ Grupo = $Group; Rol = $AddRole })
    }
}
function global:Get-AsignacionesWebSP {
    foreach ($g in $global:SP.Grupos.Values) {
        if ($g.Roles.Count -eq 0) { continue }
        [pscustomobject]@{
            Member = [pscustomobject]@{ Title = $g.Title; LoginName = $g.LoginName }
            RoleDefinitionBindings = @(foreach ($k in $g.Roles) { $global:SP.Roles | Where-Object { $_.RoleTypeKind -eq $k } })
        }
    }
}
function global:Get-AsignacionesListaSP {
    param($Lista)
    foreach ($a in $Lista.Asignaciones) {
        [pscustomobject]@{
            Member = [pscustomobject]@{ Title = $a.Grupo; LoginName = $a.Grupo }
            RoleDefinitionBindings = @($global:SP.Roles | Where-Object { $_.Name -eq $a.Rol })
        }
    }
}

# ---------------------------------------------------------------- Entra ID
function global:Register-PnPEntraIDAppForInteractiveLogin {
    param($ApplicationName, $Tenant, [string[]]$SharePointDelegatePermissions, [string[]]$GraphDelegatePermissions)
    if ($SharePointDelegatePermissions -notcontains 'AllSites.FullControl') { throw 'Falta AllSites.FullControl' }
    [pscustomobject]@{ 'AzureAppId/ClientId' = '11111111-2222-3333-4444-555555555555'; Tenant = $Tenant }
}
