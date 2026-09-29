# ADR-006 · Procesamiento secuencial (simultaneidad 1)

**Estado:** aceptada.

## Contexto
Pueden llegar varias legalizaciones el mismo día, incluso el mismo minuto. Si dos ejecuciones del flujo buscan duplicados al mismo tiempo, ninguna ve el caso de la otra.

## Decisión
- Activar el **control de simultaneidad con grado 1** en el desencadenador de LEG-01 y de LEG-02. Los correos se procesan uno detrás de otro, en orden de llegada.
- Dividir el desencadenador (una ejecución por correo) y tener un ID único garantizado por SharePoint (ADR-003) como segunda protección.

## Consecuencias
- Con unos 127 casos al mes (6 al día), la espera es de segundos: cada ejecución tarda menos de un minuto.
- La detección de duplicados ve siempre los casos anteriores ya registrados.
- Una vez activado, el control de simultaneidad no se puede desactivar sin volver a crear el desencadenador. La guía lo advierte.
