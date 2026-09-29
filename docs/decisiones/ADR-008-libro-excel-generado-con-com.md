# ADR-008 · Libro de análisis generado con Excel (COM); plantilla y datos de prueba con Python

**Estado:** aceptada.

## Contexto
Se pidió un generador del Excel de análisis "en Python con openpyxl, o equivalente", con la estructura base de la conexión Power Query al registro. openpyxl crea tablas, validaciones y formatos, pero **no puede crear consultas de Power Query**: se guardan en un paquete binario (DataMashup) que solo Excel genera.

## Decisión
- El **libro de análisis** se genera con `Generar-LibroAnalisis.ps1`, que automatiza Excel de escritorio por COM. Crea las consultas M reales (desde `src/3-adaptadores/excel/powerquery/*.pq`), las tablas conectadas, la lista desplegable de Estado, los formatos y las fórmulas.
- Para dar formato sin conexión, las consultas tienen un parámetro `pModo`:
  - `Demostracion`: datos de ejemplo, para conocer el libro;
  - `Vacio`: tablas vacías;
  - `SharePoint`: datos reales.
  El generador arma el libro con datos de ejemplo, lo vacía y lo deja en `SharePoint`. Se comprobó que las validaciones y los formatos se conservan al vaciar y volver a llenar las tablas.
- La **plantilla del comprobante** y los **datos de prueba** se generan con Python y openpyxl, porque no necesitan Power Query.

## Consecuencias
- El libro queda listo: basta con subirlo a SharePoint y actualizar los datos.
- Generar el libro requiere Excel de escritorio; usarlo, no.
- Las pruebas ejecutan Power Query de verdad con Excel (`Excel.Tests.ps1`) y leen el paquete DataMashup desde Python (`test_plantilla_y_datos.py`).
- Actualizar desde Excel para la web depende de lo que su tenant permita con orígenes de SharePoint. La forma segura es actualizar desde Excel de escritorio.
