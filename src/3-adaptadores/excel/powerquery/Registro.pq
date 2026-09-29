// Registro operativo (lista de SharePoint) con fechas locales y columnas de análisis.
// pModo: "SharePoint" = datos reales · "Demostracion" = datos de ejemplo · "Vacio" = sin filas.
let
    Columnas = {"IdElemento", "ID Legalización", "Fecha Recepción", "Colaborador", "Correo", "Asunto", "Nombre Archivo", "Estado", "Observaciones", "Fecha Cierre"},

    DesdeSharePoint = () =>
        let
            Origen = SharePoint.Tables(pSitioUrl, [ApiVersion = 15]),
            Items = Origen{[Title = pListaRegistro]}[Items],
            ColumnaId = if Table.HasColumns(Items, "Id") then "Id" else "ID",
            Seleccion = Table.SelectColumns(Items, {ColumnaId, "Title", "FechaRecepcion", "Colaborador", "Correo", "Asunto", "NombreArchivo", "Estado", "Observaciones", "FechaCierre"}, MissingField.UseNull),
            Renombrado = Table.RenameColumns(Seleccion, {{ColumnaId, "IdElemento"}, {"Title", "ID Legalización"}, {"FechaRecepcion", "Fecha Recepción"}, {"NombreArchivo", "Nombre Archivo"}, {"FechaCierre", "Fecha Cierre"}}),
            EstadoTexto = Table.TransformColumns(Renombrado, {{"Estado", each if _ is record then Record.FieldOrDefault(_, "Value") else _}}),
            FechasLocales = Table.TransformColumns(EstadoTexto, {{"Fecha Recepción", fnUtcALocal}, {"Fecha Cierre", fnUtcALocal}})
        in
            FechasLocales,

    Demostracion = () =>
        let
            Ahora = DateTime.LocalNow(),
            Hace = (dias as number, horas as number) as datetime => Ahora - #duration(dias, horas, 0, 0),
            Filas = {
                {95, "LEG-202608-00095", Hace(35, 2), "Pedro Soto (DEMO)", "pedro.soto@ejemplo.com", "Legalización Tarjeta Crédito - Pedro Soto", "B017_factura.pdf#(lf)Gastos_Q2.xlsx", "Aprobada", null, Hace(30, 0)},
                {101, "LEG-202609-00101", Hace(9, 3), "Ana Pérez (DEMO)", "ana.perez@ejemplo.com", "Legalización Tarjeta Crédito - Ana Pérez", "A001_factura.pdf#(lf)Gastos_Q1.xlsx#(lf)A001_cfdi.xml", "Aprobada", null, Hace(2, 0)},
                {102, "LEG-202609-00102", Hace(7, 5), "Luis Gómez (DEMO)", "luis.gomez@ejemplo.com", "Legalización Tarjeta Crédito - Luis Gómez", "C220_factura.pdf#(lf)Gastos_Q1.xlsx", "En revisión", null, null},
                {103, "LEG-202609-00103", Hace(6, 1), "María Díaz (DEMO)", "maria.diaz@ejemplo.com", "Legalización Tarjeta Crédito - María Díaz", "D031_factura.pdf", "Pendiente", "Falta el comprobante quincenal en Excel (.xlsx).", null},
                {104, "LEG-202609-00104", Hace(1, 4), "Ana Pérez (DEMO)", "ana.perez@ejemplo.com", "RV: Legalización Tarjeta Crédito - Ana Pérez", "A001_factura.pdf#(lf)Gastos_Q1.xlsx", "Pendiente", "[Interno] Posible duplicado de LEG-202609-00101: mismo remitente y mismo nombre de archivo en los últimos 45 días.", null},
                {105, "LEG-202609-00105", Hace(4, 2), "Jorge Ruiz (DEMO)", "jorge.ruiz@ejemplo.com", "Legalización Tarjeta Crédito - Jorge Ruiz", "E550_factura.pdf#(lf)Gastos_Q1.xlsx", "Rechazada", "La factura no corresponde a un gasto de la tarjeta corporativa.", Hace(1, 0)},
                {106, "LEG-202609-00106", Hace(3, 6), "Carla Méndez (DEMO)", "carla.mendez@ejemplo.com", "Legalización Tarjeta Crédito - Carla Méndez", "F100_factura.pdf#(lf)Gastos_formato_libre.xlsx", "Aprobada", null, null}
            }
        in
            #table(Columnas, Filas),

    Vacio = () => #table(Columnas, {}),

    Base = if pModo = "SharePoint" then DesdeSharePoint() else if pModo = "Demostracion" then Demostracion() else Vacio(),

    ConMes = Table.AddColumn(Base, "Mes Recepción", each if [#"Fecha Recepción"] = null then null else Date.ToText(DateTime.Date([#"Fecha Recepción"]), "yyyy-MM")),
    ConDias = Table.AddColumn(ConMes, "Días Abierto", each
        if List.Contains(Dominio[EstadosAbiertos], [Estado]) and [#"Fecha Recepción"] <> null
        then Number.RoundDown(Duration.TotalDays(DateTime.LocalNow() - [#"Fecha Recepción"]))
        else null),
    ConDuplicado = Table.AddColumn(ConDias, "Posible Duplicado", each [Observaciones] <> null and Text.Contains([Observaciones], Dominio[TextoAvisoDuplicado])),
    ConEnlace = Table.AddColumn(ConDuplicado, "Enlace", each if [IdElemento] = null then null else pSitioUrl & "/" & Dominio[UrlListaRelativa] & "/DispForm.aspx?ID=" & Text.From([IdElemento])),
    SinId = Table.RemoveColumns(ConEnlace, {"IdElemento"}),
    Tipado = Table.TransformColumnTypes(SinId, {
        {"ID Legalización", type text}, {"Fecha Recepción", type datetime}, {"Colaborador", type text}, {"Correo", type text},
        {"Asunto", type text}, {"Nombre Archivo", type text}, {"Estado", type text}, {"Observaciones", type text},
        {"Fecha Cierre", type datetime}, {"Mes Recepción", type text}, {"Días Abierto", Int64.Type},
        {"Posible Duplicado", type logical}, {"Enlace", type text}
    }),
    Ordenado = Table.Sort(Tipado, {{"Fecha Recepción", Order.Descending}})
in
    Ordenado
