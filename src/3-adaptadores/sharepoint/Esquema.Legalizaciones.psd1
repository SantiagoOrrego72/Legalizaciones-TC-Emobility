# =============================================================================
#  Esquema.Legalizaciones.psd1 · Capa de ADAPTADORES (SharePoint)
#  CONTRATO DE DATOS entre SharePoint, Power Automate y Power Query.
#
#  Power Automate (src/3-adaptadores/power-automate) y Power Query
#  (src/3-adaptadores/excel/powerquery) usan los NOMBRES INTERNOS definidos aquí.
#  Una vez en producción:
#    - NO cambie NombreInterno, Tipo ni las Url: rompe el flujo y las consultas.
#    - Sí puede cambiar NombreVisible, Descripcion, Requerida e Indexada:
#      01-Aprovisionar-SharePoint.ps1 los sincroniza al volver a ejecutarlo.
#    - Los ESTADOS y las carpetas raíz vienen del dominio (src/1-dominio/dominio.json).
#      El script se detiene si este archivo no coincide con el dominio.
#
#  Tipos admitidos: Texto, TextoMultilinea, FechaHora, Fecha, Opcion, SiNo,
#                   Hipervinculo, Persona
# =============================================================================
@{
    Biblioteca = @{
        Titulo       = 'Legalizaciones_TC'
        Url          = 'Legalizaciones_TC'
        Descripcion  = 'Soportes de legalizaciones de tarjeta corporativa: factura (PDF), comprobante (Excel) y XML (CFDI).'
        # Power Automate crea dentro de cada una la estructura AAAA/MM/NombreColaborador.
        CarpetasRaiz = @('Pendientes', 'Aprobadas', 'Rechazadas')

        # Columna de la biblioteca (no del registro): une cada archivo con su caso.
        # La escribe LEG-01 al guardar el adjunto y la usa LEG-02 para mover los archivos al cerrar el caso.
        Columnas = @(
            @{
                NombreInterno = 'IdLegalizacion'
                NombreVisible = 'ID Legalización'
                Tipo          = 'Texto'
                Requerida     = $false
                Indexada      = $true
                Descripcion   = 'Caso del registro al que pertenece el archivo. Lo escribe Power Automate.'
            }
        )

        # Vista predeterminada de la biblioteca: los archivos de cada caso quedan juntos
        # porque su nombre empieza por el ID Legalización.
        VistaPredeterminada = @{
            Campos   = @('DocIcon', 'LinkFilename', 'IdLegalizacion', 'Modified', 'Editor')
            Consulta = '<OrderBy><FieldRef Name="FileLeafRef" /></OrderBy>'
        }
    }

    Lista = @{
        Titulo               = 'Registro_Legalizaciones_TC'
        Url                  = 'Lists/Registro_Legalizaciones_TC'
        Descripcion          = 'Registro operativo: un elemento por cada correo de legalización recibido.'
        VersionesPrincipales = 100

        # Campo acordado "ID". SharePoint reserva el nombre "ID" para su número interno,
        # así que se reutiliza la columna obligatoria Title con el nombre visible "ID Legalización".
        # Valores únicos: SharePoint rechaza un ID repetido aunque dos flujos corran a la vez.
        ColumnaTitulo = @{
            NombreVisible = 'ID Legalización'
            Descripcion   = 'Identificador del caso. Lo asigna Power Automate; no lo modifique.'
            Requerida     = $true
            Indexada      = $true
            ValoresUnicos = $true
        }

        # Campos acordados, en el orden acordado.
        Columnas = @(
            @{
                NombreInterno = 'FechaRecepcion'
                NombreVisible = 'Fecha Recepción'
                Tipo          = 'FechaHora'
                Requerida     = $true
                Indexada      = $true
                Descripcion   = 'Fecha y hora de llegada del correo. La registra Power Automate.'
            }
            @{
                NombreInterno = 'Colaborador'
                NombreVisible = 'Colaborador'
                Tipo          = 'Texto'
                Requerida     = $true
                Indexada      = $false
                Descripcion   = 'Nombre del colaborador, tomado del asunto del correo.'
            }
            @{
                NombreInterno = 'Correo'
                NombreVisible = 'Correo'
                Tipo          = 'Texto'
                Requerida     = $true
                Indexada      = $true
                Descripcion   = 'Correo del remitente.'
            }
            @{
                NombreInterno = 'Asunto'
                NombreVisible = 'Asunto'
                Tipo          = 'Texto'
                Requerida     = $true
                Indexada      = $false
                Descripcion   = 'Asunto del correo recibido.'
            }
            @{
                # Varias líneas: un correo trae 2 o 3 adjuntos y los nombres pueden pasar de 255 caracteres.
                NombreInterno = 'NombreArchivo'
                NombreVisible = 'Nombre Archivo'
                Tipo          = 'TextoMultilinea'
                Requerida     = $false
                Indexada      = $false
                Descripcion   = 'Nombres de los archivos guardados en la biblioteca, uno por línea.'
            }
            @{
                NombreInterno       = 'Estado'
                NombreVisible       = 'Estado'
                Tipo                = 'Opcion'
                Opciones            = @('Pendiente', 'En revisión', 'Rechazada', 'Aprobada')
                ValorPredeterminado = 'Pendiente'
                Requerida           = $true
                Indexada            = $true
                Descripcion         = 'Contabilidad lo cambia a Aprobada o Rechazada al terminar la revisión.'
            }
            @{
                NombreInterno = 'Observaciones'
                NombreVisible = 'Observaciones'
                Tipo          = 'TextoMultilinea'
                Requerida     = $false
                Indexada      = $false
                Descripcion   = 'Documentos faltantes, aviso de posible duplicado o motivo de rechazo.'
            }
            @{
                NombreInterno = 'FechaCierre'
                NombreVisible = 'Fecha Cierre'
                Tipo          = 'FechaHora'
                Requerida     = $false
                Indexada      = $false
                Descripcion   = 'Fecha en que el caso pasó a Aprobada o Rechazada.'
            }

            # --- Columnas RECOMENDADAS, todavía no acordadas ------------------------------
            # Para activarlas: quite el '#' de las que quiera y vuelva a ejecutar 01 y 02.
            # @{ NombreInterno = 'IdMensajeOutlook'; NombreVisible = 'Id Mensaje Outlook'; Tipo = 'Texto';        Requerida = $false; Indexada = $true;  ValoresUnicos = $true; Descripcion = 'Id. del correo en Outlook. Evita registrar dos veces el mismo correo.' }
            # @{ NombreInterno = 'TieneFactura';     NombreVisible = 'Tiene Factura';      Tipo = 'SiNo';         Requerida = $false; Indexada = $false; Descripcion = 'Llegó al menos un PDF.' }
            # @{ NombreInterno = 'TieneComprobante'; NombreVisible = 'Tiene Comprobante';  Tipo = 'SiNo';         Requerida = $false; Indexada = $false; Descripcion = 'Llegó el Excel quincenal.' }
            # @{ NombreInterno = 'TieneXML';         NombreVisible = 'Tiene XML';          Tipo = 'SiNo';         Requerida = $false; Indexada = $false; Descripcion = 'Llegó el XML (CFDI).' }
            # @{ NombreInterno = 'PosibleDuplicado'; NombreVisible = 'Posible Duplicado';  Tipo = 'SiNo';         Requerida = $false; Indexada = $true;  Descripcion = 'Marcado por la validación heurística del flujo.' }
            # @{ NombreInterno = 'RutaCarpeta';      NombreVisible = 'Ruta Carpeta';       Tipo = 'Hipervinculo'; Requerida = $false; Indexada = $false; Descripcion = 'Enlace a la carpeta con los soportes del caso.' }
            # @{ NombreInterno = 'RevisadoPor';      NombreVisible = 'Revisado Por';       Tipo = 'Persona';      Requerida = $false; Indexada = $false; Descripcion = 'Persona de Contabilidad que cerró el caso.' }
        )

        # Vista predeterminada ("Todos los elementos"). Se deja SIN filtro a propósito:
        # Power Query puede leer la vista predeterminada y un filtro dejaría registros fuera del análisis.
        VistaPredeterminada = @{
            Campos   = @('LinkTitle', 'FechaRecepcion', 'Colaborador', 'Correo', 'Asunto', 'NombreArchivo', 'Estado', 'Observaciones', 'FechaCierre')
            Consulta = '<OrderBy><FieldRef Name="FechaRecepcion" Ascending="FALSE" /></OrderBy>'
        }

        # Vistas de trabajo de Contabilidad. "Editor" = Modificado por (quién hizo el último cambio).
        Vistas = @(
            @{
                # Bandeja de trabajo: lo más antiguo primero.
                Titulo   = 'Pendientes de revisión'
                Campos   = @('LinkTitle', 'FechaRecepcion', 'Colaborador', 'Correo', 'NombreArchivo', 'Estado', 'Observaciones')
                Consulta = '<Where><Or><Eq><FieldRef Name="Estado" /><Value Type="Choice">Pendiente</Value></Eq><Eq><FieldRef Name="Estado" /><Value Type="Choice">En revisión</Value></Eq></Or></Where><OrderBy><FieldRef Name="FechaRecepcion" Ascending="TRUE" /></OrderBy>'
            }
            @{
                Titulo   = 'Cerradas'
                Campos   = @('LinkTitle', 'FechaRecepcion', 'Colaborador', 'Estado', 'FechaCierre', 'Observaciones', 'Editor')
                Consulta = '<Where><Or><Eq><FieldRef Name="Estado" /><Value Type="Choice">Aprobada</Value></Eq><Eq><FieldRef Name="Estado" /><Value Type="Choice">Rechazada</Value></Eq></Or></Where><OrderBy><FieldRef Name="FechaCierre" Ascending="FALSE" /></OrderBy>'
            }
        )
    }
}
