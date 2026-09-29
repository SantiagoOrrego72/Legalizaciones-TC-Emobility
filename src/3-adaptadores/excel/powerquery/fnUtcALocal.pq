// Convierte una fecha UTC (como la guarda SharePoint) a la hora local del país.
// pDesfaseHorasUTC es fijo porque México y Colombia no tienen horario de verano.
(valor as any) as nullable datetime =>
    if valor = null then
        null
    else if Value.Is(valor, type datetimezone) then
        DateTimeZone.RemoveZone(DateTimeZone.SwitchZone(valor, pDesfaseHorasUTC))
    else if Value.Is(valor, type datetime) then
        valor + #duration(0, pDesfaseHorasUTC, 0, 0)
    else
        DateTimeZone.RemoveZone(DateTimeZone.SwitchZone(DateTimeZone.From(valor), pDesfaseHorasUTC))
