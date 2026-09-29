# Registro de decisiones de arquitectura (ADR)

Cada decisión importante queda escrita con su contexto, la decisión tomada y sus consecuencias. Para cambiar una decisión, se agrega un ADR nuevo que la reemplace; no se borra la anterior.

| ADR | Decisión | Estado |
|---|---|---|
| [ADR-001](ADR-001-enfoque-hibrido.md) | Enfoque híbrido: SharePoint como registro operativo, Excel + Power Query como análisis y Power Automate como orquestador | Aceptada |
| [ADR-002](ADR-002-clean-architecture-low-code.md) | Clean Architecture en una solución low-code: dominio como fuente única y pruebas de arquitectura | Aceptada |
| [ADR-003](ADR-003-id-en-title.md) | El campo "ID" es la columna Title, con valores únicos y formato LEG-AAAAMM-NNNNN | Aceptada |
| [ADR-004](ADR-004-un-registro-por-correo.md) | Un registro por correo; Nombre Archivo de varias líneas | Aceptada |
| [ADR-005](ADR-005-duplicados-heuristicos.md) | Detección heurística de duplicados (remitente + nombre de PDF/XML + 45 días) | Aceptada |
| [ADR-006](ADR-006-procesamiento-secuencial.md) | Procesamiento secuencial (simultaneidad 1) | Aceptada |
| [ADR-007](ADR-007-flujo-de-cierre.md) | Flujo de cierre LEG-02 para mover archivos y registrar la Fecha Cierre | Aceptada, pendiente de confirmar con Contabilidad |
| [ADR-008](ADR-008-libro-excel-generado-con-com.md) | El libro de análisis se genera con Excel (COM); la plantilla y los datos de prueba, con Python | Aceptada |
| [ADR-009](ADR-009-plantilla-comprobante.md) | Plantilla estándar para el comprobante quincenal | Propuesta, pendiente de confirmar con Contabilidad |
| [ADR-010](ADR-010-herramientas-sin-administrador.md) | PowerShell 7 portátil y aplicación propia de Entra ID para PnP | Aceptada |
