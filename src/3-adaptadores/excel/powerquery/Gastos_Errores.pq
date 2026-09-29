// Comprobantes (.xlsx) que no se pudieron leer: no usan la plantilla o están dañados.
let
    Columnas = {"ID Legalización", "Archivo", "Ruta", "Error"},

    DesdeSharePoint = () =>
        let
            Comprobantes = Table.SelectRows(SoportesBase, each [Tipo] = "comprobante"),
            Intentos = Table.AddColumn(Comprobantes, "Resultado", each try fnLeerComprobante([Content])),
            Fallidos = Table.SelectRows(Intentos, each [Resultado][HasError]),
            ConError = Table.AddColumn(Fallidos, "Error", each [Resultado][Error][Message]),
            Renombrado = Table.RenameColumns(ConError, {{"Nombre Guardado", "Archivo"}}),
            Seleccion = Table.SelectColumns(Renombrado, Columnas)
        in
            Seleccion,

    Demostracion = () => #table(Columnas, {
        {"LEG-202609-00106", "LEG-202609-00106_Gastos_formato_libre.xlsx",
         "https://ejemplo.sharepoint.com/sites/Legalizaciones-TC/Legalizaciones_TC/Pendientes/LEG-202609-00106_Gastos_formato_libre.xlsx",
         "No se encontró la tabla {{TABLA_PLANTILLA}}: el archivo no usa la plantilla del comprobante."}
    }),

    Vacio = () => #table(Columnas, {}),

    Base = if pModo = "SharePoint" then DesdeSharePoint() else if pModo = "Demostracion" then Demostracion() else Vacio(),
    Tipado = Table.TransformColumnTypes(Base, {{"ID Legalización", type text}, {"Archivo", type text}, {"Ruta", type text}, {"Error", type text}})
in
    Tipado
