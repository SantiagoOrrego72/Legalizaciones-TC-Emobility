// Líneas de gasto de todos los comprobantes quincenales (.xlsx) de la biblioteca que usan la plantilla.
// Diferencia = Subtotal + IVA - Total (debe ser 0; si no, aparece en Alertas).
let
    ColumnasPlantilla = {{COLUMNAS_PLANTILLA}},
    Columnas = {"ID Legalización", "Archivo"} & ColumnasPlantilla,

    DesdeSharePoint = () =>
        let
            Comprobantes = Table.SelectRows(SoportesBase, each [Tipo] = "comprobante"),
            Leidos = Table.AddColumn(Comprobantes, "Lineas", each try fnLeerComprobante([Content]) otherwise null),
            Validos = Table.SelectRows(Leidos, each [Lineas] <> null),
            SoloCaso = Table.SelectColumns(Validos, {"ID Legalización", "Nombre Guardado", "Lineas"}),
            Renombrado = Table.RenameColumns(SoloCaso, {{"Nombre Guardado", "Archivo"}}),
            Expandido = Table.ExpandTableColumn(Renombrado, "Lineas", ColumnasPlantilla)
        in
            Expandido,

    Demostracion = () =>
        let
            Hoy = Date.From(DateTime.LocalNow()),
            Dia = (dias as number) as date => Date.AddDays(Hoy, -dias),
            Filas = {
                {"LEG-202608-00095", "LEG-202608-00095_Gastos_Q2.xlsx", Dia(40), "B017", "Hotel Centro SA de CV", "HCE010101AB1", "Hospedaje visita a cliente", "CC-VENTAS", 2500, 400, 2900, "MXN", "7C1B9E2A-0000-4000-8000-000000000095"},
                {"LEG-202608-00095", "LEG-202608-00095_Gastos_Q2.xlsx", Dia(39), "B018", "Taxis Ejecutivos SA", "TEJ020202CD2", "Traslado aeropuerto", "CC-VENTAS", 350, 56, 406, "MXN", null},
                {"LEG-202609-00101", "LEG-202609-00101_Gastos_Q1.xlsx", Dia(12), "A001", "Papelería Moderna SA", "PMO030303EF3", "Material de oficina", "CC-ADMIN", 1000, 160, 1160, "MXN", "5D2A1B3C-0000-4000-8000-000000000101"},
                {"LEG-202609-00101", "LEG-202609-00101_Gastos_Q1.xlsx", Dia(11), "A002", "Restaurante El Puerto", "REP040404GH4", "Comida con proveedor", "CC-COMPRAS", 800, 128, 928, "MXN", null},
                {"LEG-202609-00101", "LEG-202609-00101_Gastos_Q1.xlsx", Dia(10), "A003", "Gasolinera Norte", "GNO050505IJ5", "Combustible vehículo de reparto", "CC-LOGISTICA", 1200, 192, 1392, "MXN", null},
                {"LEG-202609-00102", "LEG-202609-00102_Gastos_Q1.xlsx", Dia(9), "C220", "Aerolínea Nacional", "ANA060606KL6", "Vuelo a Monterrey", "CC-VENTAS", 3500, 560, 4060, "MXN", null},
                {"LEG-202609-00102", "LEG-202609-00102_Gastos_Q1.xlsx", Dia(8), "C221", "Hotel Norte", "HNO070707MN7", "Hospedaje Monterrey", "CC-VENTAS", 1800, 288, 2100, "MXN", null},
                {"LEG-202609-00104", "LEG-202609-00104_Gastos_Q1.xlsx", Dia(12), "A001", "Papelería Moderna SA", "PMO030303EF3", "Material de oficina", "CC-ADMIN", 1000, 160, 1160, "MXN", "5D2A1B3C-0000-4000-8000-000000000101"},
                {"LEG-202609-00105", "LEG-202609-00105_Gastos_Q1.xlsx", Dia(6), "E550", "Tienda Departamental", "TDE080808OP8", "Compra personal", "CC-ADMIN", 900, 144, 1044, "MXN", null}
            }
        in
            #table(Columnas, Filas),

    Vacio = () => #table(Columnas, {}),

    Base = if pModo = "SharePoint" then DesdeSharePoint() else if pModo = "Demostracion" then Demostracion() else Vacio(),
    Tipado = Table.TransformColumnTypes(Base, {
        {"ID Legalización", type text}, {"Archivo", type text}, {"Fecha", type date}, {"Número Factura", type text},
        {"Proveedor", type text}, {"Id Fiscal Proveedor", type text}, {"Concepto", type text}, {"Centro de Costo", type text},
        {"Subtotal", type number}, {"IVA", type number}, {"Total", type number}, {"Moneda", type text}, {"Folio Fiscal", type text}
    }),
    Cero = (x as nullable number) as number => if x = null then 0 else x,
    ConDiferencia = Table.AddColumn(Tipado, "Diferencia", each Number.Round(Cero([Subtotal]) + Cero([IVA]) - Cero([Total]), 2), type number)
in
    ConDiferencia
