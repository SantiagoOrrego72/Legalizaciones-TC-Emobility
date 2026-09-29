# ADR-003 · El "ID" del caso es la columna Title con valores únicos

**Estado:** aceptada.

## Contexto
El registro acordado tiene un campo "ID". SharePoint ya reserva el nombre "ID" para su número interno y no permite otra columna con ese nombre. Además, dos correos pueden procesarse casi al mismo tiempo, y el ID tiene que ser legible en correos, archivos y Operam.

## Decisión
- Reutilizar la columna obligatoria `Title` con el nombre visible **"ID Legalización"**, indexada y con **valores únicos**.
- Formato `LEG-AAAAMM-NNNNN`: mes local de recepción + número interno de SharePoint con 5 dígitos. El flujo crea el elemento con un ID temporal (`TMP-<guid>`) y lo reemplaza por el definitivo en cuanto SharePoint asigna el número.
- El mismo ID es prefijo del nombre de cada archivo (`LEG-…_factura.pdf`) y valor de la columna "ID Legalización" de la biblioteca.

## Consecuencias
- El ID nunca se repite, porque el número interno es único, y SharePoint rechaza cualquier duplicado.
- El ID lleva el año y el mes: LEG-02 obtiene de ahí la carpeta de destino.
- Si la recepción falla entre la creación y la actualización, el caso queda con `TMP-…` o `ERR-…` y con "ERROR TÉCNICO" en Observaciones, y TI recibe un aviso.
