// Controles de calidad sobre el registro, los soportes y los comprobantes.
let
    Formatear = (tabla as table, alerta as text, detalle as function) as table =>
        let
            ConAlerta = Table.AddColumn(tabla, "Alerta", each alerta),
            ConDetalle = Table.AddColumn(ConAlerta, "Detalle", detalle),
            Columnas = Table.SelectColumns(ConDetalle, {"Alerta", "ID Legalización", "Colaborador", "Detalle"}, MissingField.UseNull)
        in
            Columnas,
    ColaboradorDe = (id as nullable text) as nullable text => try Registro{[#"ID Legalización" = id]}[Colaborador] otherwise null,

    Abiertos = Table.SelectRows(Registro, each List.Contains(Dominio[EstadosAbiertos], [Estado])),
    Cerrados = Table.SelectRows(Registro, each List.Contains(Dominio[EstadosFinales], [Estado])),
    IdsConComprobante = List.Buffer(List.Distinct(Gastos[#"ID Legalización"])),
    IdsConArchivos = List.Buffer(List.Distinct(Soportes[#"ID Legalización"])),

    A1 = Formatear(Table.SelectRows(Abiertos, each [#"Días Abierto"] <> null and [#"Días Abierto"] > pDiasAlertaAbierto),
        "Caso abierto hace más de " & Text.From(pDiasAlertaAbierto) & " días", each "Estado " & [Estado] & " desde hace " & Text.From([#"Días Abierto"]) & " días."),
    A2 = Formatear(Table.SelectRows(Abiertos, each [#"Posible Duplicado"]),
        "Posible duplicado sin resolver", each [Observaciones]),
    A3 = Formatear(Table.SelectRows(Cerrados, each [#"Fecha Cierre"] = null),
        "Caso cerrado sin Fecha Cierre", each "Estado " & [Estado] & ". Revise el flujo LEG-02 (debió registrar la fecha y mover los archivos)."),
    A4 = Formatear(Table.SelectRows(Abiertos, each [#"Fecha Cierre"] <> null),
        "Caso reabierto", each "Tiene Fecha Cierre pero está en " & [Estado] & ". Aprobada y Rechazada son estados finales."),
    A5 = Formatear(Table.SelectRows(Registro, each [Estado] = Dominio[EstadoAprobado] and not List.Contains(IdsConComprobante, [#"ID Legalización"])),
        "Aprobada sin comprobante legible", each "No hay líneas de gasto para cargar a Operam."),
    A6 = Formatear(Table.AddColumn(Gastos_Errores, "Colaborador", each ColaboradorDe([#"ID Legalización"])),
        "Comprobante ilegible", each [Archivo] & ": " & [Error]),
    A7 = Formatear(Table.AddColumn(Table.SelectRows(Gastos, each [Diferencia] <> 0), "Colaborador", each ColaboradorDe([#"ID Legalización"])),
        "Totales no cuadran", each "Factura " & (if [#"Número Factura"] = null then "(sin número)" else [#"Número Factura"]) & ": Subtotal + IVA - Total = " & Number.ToText([Diferencia], "N2")),
    A8 = Formatear(Table.SelectRows(Registro, each not List.Contains(IdsConArchivos, [#"ID Legalización"])),
        "Caso sin archivos en la biblioteca", each "Revise la carpeta del caso y la columna ID Legalización de los archivos."),

    Combinado = Table.Combine({A1, A2, A3, A4, A5, A6, A7, A8}),
    Tipado = Table.TransformColumnTypes(Combinado, {{"Alerta", type text}, {"ID Legalización", type text}, {"Colaborador", type text}, {"Detalle", type text}})
in
    Tipado
