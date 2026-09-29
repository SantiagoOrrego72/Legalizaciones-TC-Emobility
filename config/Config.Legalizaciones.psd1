# =============================================================================
#  Config.Legalizaciones.psd1 · Configuración del ENTORNO (Legalizaciones TC)
#
#  Único archivo que se edita para instalar la solución en un tenant.
#  - Reemplace los valores entre < >. Cada script revisa solo las secciones que usa
#    y se detiene si en ellas queda algún marcador < > sin completar.
#  - Para un entorno de pruebas, copie este archivo (ej. Config.Pruebas.psd1),
#    cambie Sitio.RutaRelativa y pase la copia con -RutaConfig.
#  - Las reglas de negocio NO están aquí (src/1-dominio/dominio.json) ni la
#    estructura de SharePoint (src/3-adaptadores/sharepoint/Esquema.Legalizaciones.psd1).
# =============================================================================
@{
    # País de operación. Define en un solo lugar la zona horaria y el formato regional de
    # SharePoint, la zona horaria de Power Automate y la conversión de fechas de Excel.
    #   'Mexico'   -> (UTC-06:00) Guadalajara, Ciudad de México, Monterrey · es-MX
    #   'Colombia' -> (UTC-05:00) Bogotá, Lima, Quito · es-CO
    # ASUMIDO: México, porque las facturas son CFDI.
    Pais = 'Mexico'

    Tenant = @{
        # Prefijo del tenant. Si SharePoint es https://emobility.sharepoint.com -> 'emobility'
        Nombre       = '<PREFIJO_TENANT>'
        # Dominio inicial de Entra ID (casi siempre <prefijo>.onmicrosoft.com).
        # Solo lo usa 00-Registrar-AppEntraID.ps1.
        DominioEntra = '<PREFIJO_TENANT>.onmicrosoft.com'
    }

    Autenticacion = @{
        # Id. de aplicación (cliente) que devuelve 00-Registrar-AppEntraID.ps1,
        # o el de una aplicación de PnP PowerShell que la organización ya tenga registrada.
        ClientId  = '<CLIENT_ID>'
        NombreApp = 'PnP PowerShell - Legalizaciones TC'
    }

    Sitio = @{
        # $true  = el script crea un sitio dedicado (requiere rol Administrador de SharePoint).
        # $false = usa un sitio que ya existe en RutaRelativa.
        Crear        = $true
        RutaRelativa = '/sites/Legalizaciones-TC'
        Titulo       = 'Legalizaciones TC'
        Descripcion  = 'Legalización de gastos de tarjetas corporativas - Contabilidad'

        # Idioma del sitio. Solo se aplica al crearlo y no se puede cambiar después. 3082 = español.
        Lcid         = 3082

        # Solo con Permisos.Ambito = 'Sitio':
        # Bloquea compartir con personas de fuera de la organización (requiere Administrador de SharePoint).
        DeshabilitarComparticionExterna = $true
        # Solo los propietarios (TI) pueden compartir archivos, carpetas o el sitio.
        SoloPropietariosComparten       = $true

        # Opcional, solo si el país no está en la lista: texto de la zona horaria de SharePoint
        # (debe aparecer en UNA sola zona) y configuración regional (LCID). Normalmente se dejan comentados.
        # ZonaHoraria = 'Guadalajara'
        # LocaleId    = 2058
    }

    Permisos = @{
        # 'Sitio'            = sitio dedicado (recomendado): los permisos se dan en el sitio.
        # 'ListaYBiblioteca' = sitio compartido con otras áreas: la lista y la biblioteca dejan
        #                      de heredar permisos y solo acceden los propietarios y Contabilidad.
        Ambito = 'Sitio'

        # Grupo de SharePoint para Contabilidad. Solo se usa con Ambito = 'ListaYBiblioteca'
        # (con 'Sitio' se usa el grupo de Miembros del sitio).
        NombreGrupoSharePoint = 'Contabilidad - Legalizaciones TC'

        # Formato de cada persona o grupo:
        #   @{ Tipo = 'Usuario';        Valor = 'correo@dominio.com' }
        #   @{ Tipo = 'GrupoSeguridad'; Valor = 'Id. de objeto del grupo en Entra ID (GUID)' }
        #   @{ Tipo = 'GrupoM365';      Valor = 'Id. de objeto del grupo en Entra ID (GUID)' }
        # El Id. de objeto está en Entra ID > Grupos > (el grupo) > Información general.

        # Control total: personas de TI que administran la solución.
        PropietariosTI = @(
            @{ Tipo = 'Usuario'; Valor = '<correo.ti@dominio.com>' }
        )

        # Nivel Colaborar: ver, crear, editar y eliminar elementos y archivos.
        # No pueden cambiar columnas, vistas ni permisos, ni borrar la lista.
        # Se recomienda un grupo en lugar de personas sueltas.
        Contabilidad = @(
            @{ Tipo = 'GrupoSeguridad'; Valor = '<ID_OBJETO_GRUPO_CONTABILIDAD>' }
            # @{ Tipo = 'Usuario'; Valor = 'persona.contabilidad@dominio.com' }
        )

        # Cuenta con la que se crean las conexiones de Power Automate (recomendado: cuenta de
        # servicio con licencia de Microsoft 365). Recibe el mismo nivel que Contabilidad.
        # Déjela en '' si aún no existe; agréguela después y vuelva a ejecutar 01.
        CuentaServicioFlujo = '<cuenta.servicio@dominio.com>'
    }

    Flujos = @{
        # Buzón al que los colaboradores envían las legalizaciones. Recomendado: un buzón
        # compartido (ej. legalizaciones@emobility.com) al que la cuenta de servicio tenga acceso.
        BuzonLegalizaciones = '<legalizaciones@dominio.com>'
        # $true si BuzonLegalizaciones es un buzón compartido; $false si es el buzón propio
        # de la cuenta de servicio. Cambia el desencadenador y la acción de envío (ver LEG-01).
        BuzonCompartido     = $true
        # Avisos internos (posibles duplicados). Varios correos se separan con punto y coma.
        CorreoContabilidad  = '<contabilidad@dominio.com>'
        # Avisos de errores técnicos de los flujos.
        CorreoSoporteTI     = '<soporte.ti@dominio.com>'
        # Avisar al colaborador cuando su caso se aprueba o se rechaza (flujo LEG-02).
        NotificarCierre     = $true
    }

    Registro = @{
        # Carpeta (relativa a la raíz del proyecto) donde queda el registro de cada ejecución.
        CarpetaLogs = 'logs'
    }
}
