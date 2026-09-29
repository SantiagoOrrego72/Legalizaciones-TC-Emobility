<#
    Pruebas de ARQUITECTURA (Pester 5): hacen cumplir la regla de dependencias de Clean Architecture
    y que los adaptadores (SharePoint, Power Automate, Excel) no contradigan al dominio.

      Dominio (dominio.json)  <-  Adaptadores (esquema, flujos, Power Query)  <-  Infraestructura (scripts)

    Si una prueba falla después de cambiar una regla, el mensaje dice qué pieza falta actualizar.
#>

BeforeAll {
    $script:Raiz = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
    $script:Dominio = Get-Content (Join-Path $Raiz 'src\1-dominio\dominio.json') -Raw -Encoding utf8 | ConvertFrom-Json
    $script:Esquema = Import-PowerShellDataFile (Join-Path $Raiz 'src\3-adaptadores\sharepoint\Esquema.Legalizaciones.psd1')
    $script:CarpetaPA = Join-Path $Raiz 'src\3-adaptadores\power-automate'
    $script:Leg01 = Get-Content (Join-Path $CarpetaPA 'LEG-01-Recepcion.md') -Raw -Encoding utf8
    $script:Leg02 = Get-Content (Join-Path $CarpetaPA 'LEG-02-Cierre.md') -Raw -Encoding utf8
    $script:Catalogo = Get-Content (Join-Path $CarpetaPA 'expresiones.md') -Raw -Encoding utf8

    function script:Get-BloquesExpresion([string]$Texto) {
        return @([regex]::Matches($Texto, '(?ms)^```expresion\r?\n(.*?)\r?\n```') | ForEach-Object { $_.Groups[1].Value.Trim() })
    }
    # Expresiones escritas en tablas (código en línea con paréntesis)
    function script:Get-ExpresionesEnTablas([string]$Texto) {
        $resultado = @()
        foreach ($linea in ($Texto -split "`r?`n" | Where-Object { $_.StartsWith('|') })) {
            # Los segmentos impares al separar por ` son el contenido de los fragmentos de código.
            $segmentos = $linea.Split('`')
            for ($i = 1; $i -lt $segmentos.Count; $i += 2) {
                if ($segmentos[$i].Contains('(')) { $resultado += $segmentos[$i] }
            }
        }
        return $resultado
    }
    # Paréntesis, corchetes y comillas simples balanceados ('' es una comilla escapada dentro de una cadena).
    function script:Test-Balanceada([string]$Expresion) {
        $e = $Expresion.Trim()
        if ($e.StartsWith('@')) { $e = $e.Substring(1) }
        $parentesis = 0; $corchetes = 0; $enCadena = $false
        for ($i = 0; $i -lt $e.Length; $i++) {
            $c = [string]$e[$i]
            if ($enCadena) {
                if ($c -eq "'") {
                    if ($i + 1 -lt $e.Length -and [string]$e[$i + 1] -eq "'") { $i++ } else { $enCadena = $false }
                }
                continue
            }
            if ($c -eq "'") { $enCadena = $true }
            elseif ($c -eq '(') { $parentesis++ }
            elseif ($c -eq ')') { $parentesis--; if ($parentesis -lt 0) { return $false } }
            elseif ($c -eq '[') { $corchetes++ }
            elseif ($c -eq ']') { $corchetes--; if ($corchetes -lt 0) { return $false } }
        }
        return (-not $enCadena) -and ($parentesis -eq 0) -and ($corchetes -eq 0)
    }
    # Quita el contenido de las cadenas para analizar solo el código de la expresión.
    function script:Remove-Cadenas([string]$Expresion) {
        return [regex]::Replace($Expresion, "'(?:[^']|'')*'", "''")
    }

    # Configuración de los flujos generada con un entorno de prueba
    $config = [System.IO.File]::ReadAllText((Join-Path $Raiz 'config\Config.Legalizaciones.psd1'))
    foreach ($par in @(
            @('<PREFIJO_TENANT>', 'contoso'), @('<legalizaciones@dominio.com>', 'legalizaciones@contoso.com'),
            @('<contabilidad@dominio.com>', 'contabilidad@contoso.com'), @('<soporte.ti@dominio.com>', 'ti@contoso.com'))) {
        $config = $config.Replace($par[0], $par[1])
    }
    $rutaConfig = Join-Path $TestDrive 'Config.Prueba.psd1'
    [System.IO.File]::WriteAllText($rutaConfig, $config, [System.Text.UTF8Encoding]::new($true))
    $rutaJson = Join-Path $TestDrive 'Configuracion-Flujos.json'
    & (Join-Path $Raiz 'src\4-infraestructura\power-automate\Generar-ConfiguracionFlujos.ps1') -RutaConfig $rutaConfig -RutaSalida $rutaJson 6>$null
    $script:ConfigFlujos = Get-Content $rutaJson -Raw -Encoding utf8 | ConvertFrom-Json

    $script:FuncionesPowerAutomate = @(
        'add', 'addDays', 'and', 'base64ToBinary', 'body', 'coalesce', 'concat', 'contains', 'convertFromUtc', 'createArray',
        'decodeUriComponent', 'empty', 'equals', 'first', 'formatNumber', 'greater', 'guid', 'if', 'int', 'intersection',
        'item', 'items', 'join', 'json', 'last', 'length', 'mul', 'not', 'or', 'outputs', 'replace', 'result', 'split',
        'startsWith', 'string', 'substring', 'take', 'toLower', 'triggerBody', 'trim', 'union', 'utcNow', 'variables', 'workflow'
    )
}

Describe 'Regla de dependencias (Clean Architecture)' {
    It 'el dominio no depende de SharePoint, Outlook, Excel ni de las capas externas' {
        $tokens = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $Raiz 'src\1-dominio\Legalizaciones.Dominio.psm1'), [ref]$tokens, [ref]$null)
        $comandos = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() })
        $comandos | Where-Object { $_ -match '-PnP|^Connect-|^Invoke-WebRequest$|^Invoke-RestMethod$|^New-Object$|^Send-' } | Should -BeNullOrEmpty
        $codigo = ($tokens | Where-Object { $_.Kind -ne 'Comment' } | ForEach-Object Text) -join ' '
        $codigo | Should -Not -Match 'ComObject|3-adaptadores|4-infraestructura'
    }
    It 'dominio.json no contiene datos de un entorno (URL, tenant, correos)' {
        $texto = Get-Content (Join-Path $Raiz 'src\1-dominio\dominio.json') -Raw -Encoding utf8
        $texto | Should -Not -Match 'https?://|sharepoint\.com|@[a-z0-9-]+\.[a-z]'
    }
    It 'la capa de adaptadores no contiene scripts ejecutables (solo contratos, consultas y documentación)' {
        @(Get-ChildItem (Join-Path $Raiz 'src\3-adaptadores') -Recurse -File -Include *.ps1, *.psm1, *.py) | Should -BeNullOrEmpty
    }
}

Describe 'Dominio y adaptador de SharePoint' {
    It 'las opciones de Estado son los estados del dominio, en el mismo orden' {
        $estado = $Esquema.Lista.Columnas | Where-Object { $_.NombreInterno -eq 'Estado' }
        @($estado.Opciones) | Should -Be @($Dominio.caso.estados)
        $estado.ValorPredeterminado | Should -Be $Dominio.caso.estadoInicial
    }
    It 'las vistas solo filtran por estados que existen' {
        foreach ($vista in $Esquema.Lista.Vistas) {
            foreach ($m in [regex]::Matches($vista.Consulta, '<Value Type="Choice">([^<]+)</Value>')) {
                @($Dominio.caso.estados) | Should -Contain $m.Groups[1].Value -Because "vista '$($vista.Titulo)'"
            }
        }
    }
    It 'la biblioteca tiene todas las carpetas raíz que usa el dominio' {
        foreach ($raizCarpeta in ($Dominio.carpetas.raizPorEstado.PSObject.Properties.Value | Select-Object -Unique)) {
            @($Esquema.Biblioteca.CarpetasRaiz) | Should -Contain $raizCarpeta
        }
    }
}

Describe 'Dominio y configuración de los flujos (Configuracion-Flujos.json)' {
    It 'lleva el prefijo, estados, extensiones, ventana de duplicados y mensajes del dominio' {
        $ConfigFlujos.prefijoId | Should -Be $Dominio.caso.prefijoId
        $ConfigFlujos.formatoConsecutivo | Should -Be ('0' * $Dominio.caso.digitosConsecutivo)
        $ConfigFlujos.estadoInicial | Should -Be $Dominio.caso.estadoInicial
        $ConfigFlujos.estadoAprobado | Should -Be $Dominio.caso.estadoAprobado
        $ConfigFlujos.estadoRechazado | Should -Be $Dominio.caso.estadoRechazado
        $ConfigFlujos.diasVentanaDuplicado | Should -Be $Dominio.duplicados.diasVentana
        @($ConfigFlujos.estadosExcluidosDuplicado) | Should -Be @($Dominio.duplicados.estadosExcluidos)
        foreach ($tipo in $Dominio.documentos.PSObject.Properties) {
            @($ConfigFlujos.extensiones.($tipo.Name)) | Should -Be @($tipo.Value.extensiones)
        }
        foreach ($mensaje in $Dominio.mensajes.PSObject.Properties) {
            $ConfigFlujos.mensajes.($mensaje.Name) | Should -BeExactly $mensaje.Value
        }
    }
    It 'toma la zona horaria de Windows del país configurado' {
        $ConfigFlujos.zonaHorariaWindows | Should -Be 'Central Standard Time (Mexico)'
    }
    It 'solo factura y comprobante son obligatorios (la expresión Documentacion_completa depende de eso)' {
        $obligatorios = @($Dominio.documentos.PSObject.Properties | Where-Object { $_.Value.obligatorio } | ForEach-Object Name)
        $obligatorios | Should -Be @('factura', 'comprobante')
    }
}

Describe 'Documentación de los flujos (<Nombre>)' -ForEach @(
    @{ Nombre = 'LEG-01'; Archivo = 'LEG-01-Recepcion.md' }
    @{ Nombre = 'LEG-02'; Archivo = 'LEG-02-Cierre.md' }
) {
    BeforeAll {
        $script:Texto = Get-Content (Join-Path $CarpetaPA $Archivo) -Raw -Encoding utf8
        $script:Expresiones = @(Get-BloquesExpresion $Texto) + @(Get-ExpresionesEnTablas $Texto)
        $script:Definidas = @([regex]::Matches($Texto, '(?m)^#{2,4} .*?`([A-Za-z][A-Za-z0-9_]*)`') | ForEach-Object { $_.Groups[1].Value })
    }
    It 'todas las expresiones tienen paréntesis, corchetes y comillas balanceados' {
        $Expresiones.Count | Should -BeGreaterThan 10
        foreach ($expresion in $Expresiones) { Test-Balanceada $expresion | Should -BeTrue -Because $expresion }
    }
    It 'solo usan funciones que existen en Power Automate (sin errores de tipeo)' {
        foreach ($expresion in $Expresiones) {
            foreach ($m in [regex]::Matches((Remove-Cadenas $expresion), '([A-Za-z_][A-Za-z0-9_]*)\(')) {
                $FuncionesPowerAutomate | Should -Contain $m.Groups[1].Value -Because $expresion
            }
        }
    }
    It 'cada acción referenciada con outputs/body/items/result está definida en la guía' {
        foreach ($m in [regex]::Matches($Texto, "(?:outputs|body|items|result)\('([A-Za-z0-9_]+)'\)")) {
            $Definidas | Should -Contain $m.Groups[1].Value
        }
    }
    It 'cada clave de Configuracion que usan las expresiones existe en Configuracion-Flujos.json' {
        foreach ($m in [regex]::Matches($Texto, "outputs\('Configuracion'\)\?\['([A-Za-z]+)'\](?:\?\['([A-Za-z]+)'\])?")) {
            $clave = $m.Groups[1].Value
            $ConfigFlujos.PSObject.Properties.Name | Should -Contain $clave
            if ($m.Groups[2].Success -and $ConfigFlujos.$clave -is [pscustomobject]) {
                $ConfigFlujos.$clave.PSObject.Properties.Name | Should -Contain $m.Groups[2].Value
            }
        }
    }
    It 'las variables usadas están declaradas' {
        foreach ($m in [regex]::Matches($Texto, "variables\('([A-Za-z]+)'\)")) {
            $Texto | Should -Match ('\| `Inicializar_{0}` \| `{0}` \|' -f $m.Groups[1].Value)
        }
    }
}

Describe 'Reglas del dominio dentro de las expresiones' {
    It 'la condición del desencadenador de LEG-01 usa el texto clave del dominio' {
        $Leg01 | Should -Match ([regex]::Escape("'" + $Dominio.correo.textoClaveAsunto + "')"))
    }
    It 'la condición del desencadenador de LEG-02 usa los estados finales del dominio' {
        foreach ($estado in $Dominio.caso.estadoAprobado, $Dominio.caso.estadoRechazado) {
            $Leg02 | Should -Match ([regex]::Escape("['Value'], '$estado')"))
        }
    }
    It 'la limpieza de carpetas quita exactamente los caracteres no permitidos del dominio' {
        foreach ($texto in $Leg01, $Leg02) {
            $bloque = (Get-BloquesExpresion $texto | Where-Object { $_ -match "^trim\(take\(trim\(replace" } | Select-Object -First 1)
            $quitados = @([regex]::Matches($bloque, "'((?:[^']|'')*)', ''\)") | ForEach-Object { $_.Groups[1].Value.Replace("''", "'") })
            ($quitados | Sort-Object) | Should -Be (@($Dominio.carpetas.caracteresNoPermitidos) | Sort-Object)
        }
    }
    It 'cada expresión del catálogo aparece idéntica en LEG-01 o LEG-02' {
        $flujos = $Leg01 + "`n" + $Leg02
        $bloques = Get-BloquesExpresion $Catalogo
        $bloques.Count | Should -BeGreaterThan 10
        foreach ($bloque in $bloques) { $flujos.Contains($bloque) | Should -BeTrue -Because $bloque }
    }
    It 'el catálogo tiene paréntesis y comillas balanceados' {
        foreach ($bloque in Get-BloquesExpresion $Catalogo) { Test-Balanceada $bloque | Should -BeTrue -Because $bloque }
    }
}

Describe 'Plantillas de correo' {
    It '<Plantilla>: cada marcador {{...}} se reemplaza en <Flujo>' -ForEach @(
        @{ Plantilla = 'recepcion-completa.html'; Flujo = 'LEG-01'; Compose = 'Plantilla_completa' }
        @{ Plantilla = 'recepcion-incompleta.html'; Flujo = 'LEG-01'; Compose = 'Plantilla_incompleta' }
        @{ Plantilla = 'aviso-duplicado.html'; Flujo = 'LEG-01'; Compose = 'Plantilla_duplicado' }
        @{ Plantilla = 'error-tecnico.html'; Flujo = 'LEG-01'; Compose = 'Plantilla_error' }
        @{ Plantilla = 'cierre-aprobada.html'; Flujo = 'LEG-02'; Compose = 'Plantilla_aprobada' }
        @{ Plantilla = 'cierre-rechazada.html'; Flujo = 'LEG-02'; Compose = 'Plantilla_rechazada' }
        @{ Plantilla = 'error-tecnico.html'; Flujo = 'LEG-02'; Compose = 'Plantilla_error' }
    ) {
        $html = Get-Content (Join-Path $CarpetaPA "plantillas-correo\$Plantilla") -Raw -Encoding utf8
        $texto = if ($Flujo -eq 'LEG-01') { $Leg01 } else { $Leg02 }
        $texto | Should -Match ([regex]::Escape($Plantilla)) -Because 'la guía debe indicar qué archivo pegar'
        foreach ($marcador in ([regex]::Matches($html, '\{\{[A-Z_]+\}\}') | ForEach-Object Value | Select-Object -Unique)) {
            $texto | Should -Match ([regex]::Escape("'$marcador'")) -Because "$Plantilla usa $marcador"
        }
    }
    It 'ninguna plantilla empieza con @ ni contiene @{ (Power Automate lo tomaría como expresión)' {
        foreach ($archivo in Get-ChildItem (Join-Path $CarpetaPA 'plantillas-correo') -Filter *.html) {
            $html = Get-Content $archivo.FullName -Raw -Encoding utf8
            $html.TrimStart().StartsWith('@') | Should -BeFalse
            $html | Should -Not -Match '@\{'
        }
    }
}

Describe 'Dominio y adaptador de Excel (Power Query)' {
    BeforeAll {
        $script:Generador = Get-Content (Join-Path $Raiz 'src\4-infraestructura\excel\Generar-LibroAnalisis.ps1') -Raw -Encoding utf8
        $script:Plantilla = Get-Content (Join-Path $Raiz 'src\3-adaptadores\excel\plantilla-comprobante.json') -Raw -Encoding utf8 | ConvertFrom-Json
        $script:Consultas = Get-ChildItem (Join-Path $Raiz 'src\3-adaptadores\excel\powerquery') -Filter *.pq
    }
    It 'cada marcador {{...}} de las consultas lo reemplaza el generador' {
        foreach ($consulta in $Consultas) {
            $texto = Get-Content $consulta.FullName -Raw -Encoding utf8
            foreach ($m in [regex]::Matches($texto, '\{\{([A-Z_]+)\}\}')) {
                $Generador | Should -Match ('(?m)^\s+{0}\s+=' -f $m.Groups[1].Value) -Because "$($consulta.Name) usa $($m.Value)"
            }
        }
    }
    It 'cada consulta .pq está registrada en el generador' {
        foreach ($consulta in $Consultas) {
            $Generador | Should -Match ('(?m)^\s+{0}\s+=' -f [regex]::Escape($consulta.BaseName)) -Because $consulta.Name
        }
    }
    It 'la plantilla del comprobante tiene las columnas que usan las consultas' {
        $columnas = @($Plantilla.columnas | ForEach-Object nombre)
        foreach ($obligatoria in 'Fecha', 'Proveedor', 'Total', 'Subtotal', 'IVA', 'Número Factura') { $columnas | Should -Contain $obligatoria }
        $consolidado = Get-Content (Join-Path $Raiz 'src\3-adaptadores\excel\powerquery\Consolidado_Operam.pq') -Raw -Encoding utf8
        $expandidas = [regex]::Match($consolidado, '"Lineas",\s*\{([^}]*)\}').Groups[1].Value
        foreach ($m in [regex]::Matches($expandidas, '"([^"]+)"')) {
            ($columnas + 'Archivo') | Should -Contain $m.Groups[1].Value
        }
    }
    It 'Registro lee de SharePoint los nombres internos del contrato' {
        $registro = Get-Content (Join-Path $Raiz 'src\3-adaptadores\excel\powerquery\Registro.pq') -Raw -Encoding utf8
        foreach ($columna in $Esquema.Lista.Columnas) { $registro | Should -Match ('"{0}"' -f $columna.NombreInterno) }
    }
}
