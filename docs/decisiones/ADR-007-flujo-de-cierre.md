# ADR-007 · Flujo de cierre LEG-02

**Estado:** aceptada; pendiente de confirmar con Contabilidad si quiere el aviso de cierre al colaborador.

## Contexto
La estructura acordada tiene las carpetas Aprobadas y Rechazadas y el campo Fecha Cierre, pero el proceso acordado no dice quién mueve los archivos ni quién llena la fecha. Hacerlo a mano es lento y propenso a errores: archivos en la carpeta equivocada o fechas vacías.

## Decisión
Un segundo flujo, **LEG-02**, se dispara cuando un caso pasa a Aprobada o Rechazada y todavía no tiene Fecha Cierre. Hace lo siguiente:
1. mueve sus archivos de `Pendientes` a `Aprobadas` o `Rechazadas`, con la misma estructura AAAA/MM/Colaborador;
2. registra la Fecha Cierre;
3. avisa al colaborador con las Observaciones como motivo. Se puede desactivar con `Flujos.NotificarCierre`.

La decisión de aprobar o rechazar sigue siendo manual.

## Consecuencias
- La Fecha Cierre y la ubicación de los archivos son siempre coherentes con el estado.
- Si algo falla, no se registra la fecha y el Excel lo muestra como alerta ("Caso cerrado sin Fecha Cierre"). Al guardar de nuevo el caso, se vuelve a intentar.
- Contabilidad no debe llenar la Fecha Cierre a mano ni mover los archivos a mano.
- Si no se quiere este flujo, las carpetas y la fecha deberán manejarse manualmente.
