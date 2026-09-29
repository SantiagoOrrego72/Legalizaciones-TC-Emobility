// Archivos de la biblioteca (sin el contenido binario), para revisar qué soportes tiene cada caso.
let
    SinContenido = Table.RemoveColumns(SoportesBase, {"Content"}),
    Ordenado = Table.Sort(SinContenido, {{"ID Legalización", Order.Descending}, {"Nombre Archivo", Order.Ascending}})
in
    Ordenado
