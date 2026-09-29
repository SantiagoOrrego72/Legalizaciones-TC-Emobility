// Casos ya cargados en Operam. Lee la tabla manual tblControlOperam de la hoja Control_Operam.
let
    Origen = Excel.CurrentWorkbook(){[Name = "tblControlOperam"]}[Content],
    ConId = Table.SelectRows(Origen, each [#"ID Legalización"] <> null and Text.Trim(Text.From([#"ID Legalización"])) <> ""),
    Limpio = Table.TransformColumns(ConId, {{"ID Legalización", each Text.Trim(Text.From(_)), type text}})
in
    Limpio
