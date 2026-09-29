<#
    Pruebas de la capa de INFRAESTRUCTURA contra el módulo PnP.PowerShell instalado (Pester 5).
    Sin conectarse a un tenant, comprueban que cada comando PnP que usan los scripts existe
    y que cada parámetro está bien escrito. Se omiten si PnP.PowerShell no está instalado.
#>

BeforeDiscovery {
    $script:SinPnP = -not (Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object { $_.Version -ge [version]'3.0.0' })
}

BeforeAll {
    $script:Raiz = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
    if (-not $SinPnP) { Import-Module PnP.PowerShell -MinimumVersion 3.0.0 -ErrorAction Stop }
    $script:Scripts = @(
        Get-ChildItem (Join-Path $Raiz 'src\4-infraestructura\sharepoint') -Filter *.ps1
        Get-Item (Join-Path $Raiz 'pruebas\aceptacion\Validar-Escenarios.ps1')
    )
    # Parámetros dinámicos: solo aparecen al elegir -Type y no los muestra Get-Command.
    $script:Dinamicos = @{ 'New-PnPSite' = @('Title', 'Url', 'Lcid', 'Owner', 'TimeZone', 'Description', 'Alias') }

    $script:Usos = foreach ($archivo in $Scripts) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($archivo.FullName, [ref]$null, [ref]$null)
        foreach ($comando in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
            $nombre = $comando.GetCommandName()
            if ($nombre -notmatch '-PnP') { continue }
            [pscustomobject]@{
                Archivo    = $archivo.Name
                Linea      = $comando.Extent.StartLineNumber
                Comando    = $nombre
                Parametros = @($comando.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] } | ForEach-Object { $_.ParameterName })
            }
        }
    }
}

Describe 'Comandos PnP.PowerShell usados por los scripts' -Skip:$SinPnP {
    It 'los scripts usan comandos PnP (la prueba revisa algo)' {
        @($Usos).Count | Should -BeGreaterThan 40
    }
    It 'cada comando existe en PnP.PowerShell' {
        foreach ($uso in $Usos) {
            Get-Command $uso.Comando -Module PnP.PowerShell -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty -Because "$($uso.Archivo):$($uso.Linea) usa $($uso.Comando)"
        }
    }
    It 'cada parámetro existe en su comando' {
        foreach ($uso in $Usos) {
            $comando = Get-Command $uso.Comando -Module PnP.PowerShell
            foreach ($parametro in $uso.Parametros) {
                $existe = $comando.Parameters.ContainsKey($parametro) -or
                          @($comando.Parameters.Values | Where-Object { $_.Aliases -contains $parametro }).Count -gt 0 -or
                          ($Dinamicos.ContainsKey($uso.Comando) -and $Dinamicos[$uso.Comando] -contains $parametro)
                $existe | Should -BeTrue -Because "$($uso.Archivo):$($uso.Linea) $($uso.Comando) -$parametro"
            }
        }
    }
    It 'New-PnPSite acepta -Title, -Url, -Lcid y -Wait con -Type TeamSiteWithoutMicrosoft365Group' {
        # Sin sesión iniciada, un parámetro inexistente falla al enlazar (NamedParameterNotFound)
        # y uno correcto falla después, por no haber iniciado sesión.
        $mensaje = try {
            New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title 'X' -Url 'https://contoso.sharepoint.com/sites/x' -Lcid 3082 -Wait -ErrorAction Stop
            ''
        }
        catch { $_.FullyQualifiedErrorId }
        $mensaje | Should -Not -Match 'NamedParameterNotFound'
    }
}
