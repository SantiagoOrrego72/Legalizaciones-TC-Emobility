// Gastos de los casos APROBADOS que todavía no se registraron en Control_Operam.
// Es la información que Contabilidad carga manualmente en Operam.
let
    Aprobadas = Table.SelectRows(Registro, each [Estado] = Dominio[EstadoAprobado]),
    Cargados = List.Buffer(Control_Operam[#"ID Legalización"]),
    PorCargar = Table.SelectRows(Aprobadas, each not List.Contains(Cargados, [#"ID Legalización"])),
    Casos = Table.SelectColumns(PorCargar, {"ID Legalización", "Fecha Cierre", "Colaborador", "Correo"}),
    Unido = Table.NestedJoin(Casos, {"ID Legalización"}, Gastos, {"ID Legalización"}, "Lineas", JoinKind.Inner),
    Expandido = Table.ExpandTableColumn(Unido, "Lineas",
        {"Fecha", "Número Factura", "Proveedor", "Id Fiscal Proveedor", "Concepto", "Centro de Costo", "Subtotal", "IVA", "Total", "Moneda", "Folio Fiscal", "Archivo"}),
    Ordenado = Table.Sort(Expandido, {{"ID Legalización", Order.Ascending}, {"Fecha", Order.Ascending}})
in
    Ordenado
