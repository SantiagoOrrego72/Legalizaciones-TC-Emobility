# ADR-004 · Un registro por correo

**Estado:** aceptada.

## Contexto
Cada correo trae 2 o 3 adjuntos (factura, comprobante y, a veces, XML), pero el campo acordado se llama "Nombre Archivo", en singular.

## Decisión
- Un elemento del registro por **correo**, es decir, por caso de legalización.
- "Nombre Archivo" es una columna de varias líneas con los nombres **originales** de los adjuntos guardados, uno por línea. Puede pasar de 255 caracteres.
- Cada archivo guardado lleva además el ID del caso, en su nombre y en la columna de la biblioteca.

## Consecuencias
- Contabilidad aprueba o rechaza el caso completo, como hoy.
- SharePoint no permite filtrar por columnas de varias líneas. La búsqueda de duplicados filtra por Correo y Fecha Recepción, que están indexadas, y compara los nombres dentro del flujo.
- Si en el futuro se necesita aprobar factura por factura, habrá que agregar una lista hija de líneas. El dominio ya distingue el caso de sus documentos.
