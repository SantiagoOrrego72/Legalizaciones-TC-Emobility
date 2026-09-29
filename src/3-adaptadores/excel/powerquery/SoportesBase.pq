// Archivos de la biblioteca de soportes, con su caso, tipo y carpeta de estado.
// Incluye el contenido binario (lo usa Gastos); no se carga a una hoja.
let
    Columnas = {"ID Legalización", "Nombre Archivo", "Nombre Guardado", "Tipo", "Carpeta Estado", "Extensión", "Ruta", "Fecha Creación", "Content"},

    TipoDocumento = (extension as nullable text) as text =>
        let
            Ext = Text.Lower(if extension = null then "" else extension),
            Tipos = Record.FieldNames(Dominio[TiposDocumento]),
            Coinciden = List.Select(Tipos, (t) => List.Contains(Record.Field(Dominio[TiposDocumento], t), Ext))
        in
            if List.IsEmpty(Coinciden) then "no soportado" else List.First(Coinciden),

    DesdeSharePoint = () =>
        let
            Archivos = SharePoint.Files(pSitioUrl, [ApiVersion = 15]),
            Biblioteca = Table.SelectRows(Archivos, each Text.Contains([Folder Path], pRutaBiblioteca)),
            ConId = Table.AddColumn(Biblioteca, "ID Legalización", each
                if Text.StartsWith([Name], Dominio[PrefijoId] & "-") and Text.Contains([Name], "_") then Text.BeforeDelimiter([Name], "_") else null),
            ConNombre = Table.AddColumn(ConId, "Nombre Archivo", each if [#"ID Legalización"] = null then [Name] else Text.AfterDelimiter([Name], "_")),
            ConTipo = Table.AddColumn(ConNombre, "Tipo", each TipoDocumento([Extension])),
            ConCarpeta = Table.AddColumn(ConTipo, "Carpeta Estado", each Text.BeforeDelimiter(Text.AfterDelimiter([Folder Path], pRutaBiblioteca), "/")),
            ConRuta = Table.AddColumn(ConCarpeta, "Ruta", each [Folder Path] & [Name]),
            Renombrado = Table.RenameColumns(ConRuta, {{"Name", "Nombre Guardado"}, {"Extension", "Extensión"}, {"Date created", "Fecha Creación"}}),
            Seleccion = Table.SelectColumns(Renombrado, Columnas)
        in
            Seleccion,

    Demostracion = () =>
        let
            Ahora = DateTime.LocalNow(),
            Ruta = "https://ejemplo.sharepoint.com/sites/Legalizaciones-TC/Legalizaciones_TC/",
            Fila = (id as text, nombre as text, carpeta as text, dias as number) as list =>
                let Ext = "." & List.Last(Text.Split(nombre, "."))
                in {id, nombre, id & "_" & nombre, TipoDocumento(Ext), carpeta, Ext, Ruta & carpeta & "/" & id & "_" & nombre, Ahora - #duration(dias, 0, 0, 0), null},
            Filas = {
                Fila("LEG-202608-00095", "B017_factura.pdf", "Aprobadas", 35), Fila("LEG-202608-00095", "Gastos_Q2.xlsx", "Aprobadas", 35),
                Fila("LEG-202609-00101", "A001_factura.pdf", "Aprobadas", 9), Fila("LEG-202609-00101", "Gastos_Q1.xlsx", "Aprobadas", 9), Fila("LEG-202609-00101", "A001_cfdi.xml", "Aprobadas", 9),
                Fila("LEG-202609-00102", "C220_factura.pdf", "Pendientes", 7), Fila("LEG-202609-00102", "Gastos_Q1.xlsx", "Pendientes", 7),
                Fila("LEG-202609-00103", "D031_factura.pdf", "Pendientes", 6),
                Fila("LEG-202609-00104", "A001_factura.pdf", "Pendientes", 1), Fila("LEG-202609-00104", "Gastos_Q1.xlsx", "Pendientes", 1),
                Fila("LEG-202609-00105", "E550_factura.pdf", "Rechazadas", 4), Fila("LEG-202609-00105", "Gastos_Q1.xlsx", "Rechazadas", 4),
                Fila("LEG-202609-00106", "F100_factura.pdf", "Pendientes", 3), Fila("LEG-202609-00106", "Gastos_formato_libre.xlsx", "Pendientes", 3)
            }
        in
            #table(Columnas, Filas),

    Vacio = () => #table(Columnas, {}),

    Base = if pModo = "SharePoint" then DesdeSharePoint() else if pModo = "Demostracion" then Demostracion() else Vacio(),
    Tipado = Table.TransformColumnTypes(Base, {
        {"ID Legalización", type text}, {"Nombre Archivo", type text}, {"Nombre Guardado", type text}, {"Tipo", type text},
        {"Carpeta Estado", type text}, {"Extensión", type text}, {"Ruta", type text}, {"Fecha Creación", type datetime}, {"Content", type binary}
    })
in
    Tipado
