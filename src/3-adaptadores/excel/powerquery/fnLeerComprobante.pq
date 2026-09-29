// Lee las filas de la tabla tblGastos de un comprobante quincenal (.xlsx).
// Falla (y el archivo aparece en Errores_Lectura) si el archivo no usa la plantilla.
(contenido as binary) as table =>
let
    Columnas = {{COLUMNAS_PLANTILLA}},
    Libro = Excel.Workbook(contenido, true, true),
    Tabla = Libro{[Item = "{{TABLA_PLANTILLA}}", Kind = "Table"]}[Data],
    Seleccion = Table.SelectColumns(Tabla, Columnas, MissingField.Error),
    ConDatos = Table.SelectRows(Seleccion, each List.NonNullCount({[Fecha], [Proveedor], [Total]}) > 0)
in
    ConDatos
