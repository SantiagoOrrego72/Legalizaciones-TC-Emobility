# ADR-009 · Plantilla estándar del comprobante quincenal

**Estado:** propuesta; pendiente de confirmar con Contabilidad.

## Contexto
La arquitectura acordada dice que Excel "procesa el contenido del Excel de gastos quincenales". Para que Power Query lo lea de forma confiable, todos los comprobantes deben tener la misma estructura. No conocemos el formato que usan hoy los colaboradores.

## Decisión
- Proponer una plantilla (`salida/Plantilla_Comprobante_Quincenal.xlsx`) con una tabla de Excel llamada `tblGastos` y estas columnas: Fecha, Número Factura, Proveedor, Id Fiscal Proveedor, Concepto, Centro de Costo, Subtotal, IVA, Total, Moneda y Folio Fiscal. Tiene validaciones de fecha, número y moneda.
- Su contrato está en `src/3-adaptadores/excel/plantilla-comprobante.json` y lo usan tanto el generador de la plantilla como la consulta `fnLeerComprobante`.
- Un comprobante que no usa la plantilla se registra igual (es un .xlsx válido), pero aparece en la hoja *Errores_Lectura* y en *Alertas* del análisis.

## Consecuencias
- Hay que comunicar la plantilla a los colaboradores (`docs/04-instrucciones-colaboradores.md`).
- Si Contabilidad prefiere mantener el formato actual, se ajusta el contrato y la consulta `fnLeerComprobante` a ese formato; el resto no cambia.
