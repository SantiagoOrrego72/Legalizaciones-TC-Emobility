# ADR-005 · Detección heurística de duplicados

**Estado:** aceptada.

## Contexto
Hay que advertir posibles duplicados sin leer el contenido de los documentos (no se permite OCR ni IA). El Excel quincenal suele llamarse igual en cada envío ("Gastos.xlsx").

## Decisión
Un caso es **posible duplicado** si existe otro caso:
- del **mismo remitente**;
- recibido en los **últimos 45 días** (`duplicados.diasVentana`);
- que **no esté Rechazado**, porque un reenvío corregido después de un rechazo es normal;
- y que comparta al menos un **nombre de archivo PDF o XML**, sin distinguir mayúsculas. El Excel no se compara.

Es un aviso **interno**: queda en Observaciones y se envía a Contabilidad. El colaborador no lo recibe y el caso no se bloquea.

## Consecuencias
- Detecta el caso típico: reenviar el mismo correo o la misma factura.
- Puede haber falsos positivos (facturas distintas con el mismo nombre genérico) y falsos negativos (misma factura con otro nombre). Contabilidad decide.
- Una fase futura puede comparar el UUID del CFDI desde el XML con Power Query, sin cambiar este flujo: se agrega un adaptador.
- Para repetir las pruebas de aceptación, los casos de prueba anteriores se marcan como Rechazados; así no cuentan como duplicados.
