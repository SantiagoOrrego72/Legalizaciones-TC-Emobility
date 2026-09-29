// Casos por mes de recepción y estado.
let
    Base = Table.SelectColumns(Registro, {"Mes Recepción", "Estado"}),
    Agrupado = Table.Group(Base, {"Mes Recepción", "Estado"}, {{"Casos", each Table.RowCount(_), Int64.Type}}),
    Pivote = Table.Pivot(Agrupado, Dominio[Estados], "Estado", "Casos", List.Sum),
    SinNulos = Table.ReplaceValue(Pivote, null, 0, Replacer.ReplaceValue, Dominio[Estados]),
    ConTotal = Table.AddColumn(SinNulos, "Total", each List.Sum(Record.FieldValues(Record.SelectFields(_, Dominio[Estados])))),
    Tipado = Table.TransformColumnTypes(ConTotal, List.Transform(Dominio[Estados] & {"Total"}, each {_, Int64.Type}) & {{"Mes Recepción", type text}}),
    Ordenado = Table.Sort(Tipado, {{"Mes Recepción", Order.Descending}})
in
    Ordenado
