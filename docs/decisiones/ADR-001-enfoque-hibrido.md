# ADR-001 · Enfoque híbrido SharePoint + Excel, con Power Automate como orquestador

**Estado:** aceptada.

## Contexto
Hay que registrar unas 127 legalizaciones al mes con varias personas trabajando el mismo día y dejar trazabilidad. Además, hay que analizar el contenido de los comprobantes Excel. Solo se permiten herramientas de Microsoft 365, sin IA ni desarrollo a la medida.

## Decisión
- **SharePoint** guarda los archivos y el registro operativo en tiempo real: un caso por correo y su estado. Resuelve la concurrencia: varias personas editan elementos distintos sin bloquear un archivo.
- **Excel con Power Query** es la capa de análisis, conectada al registro y a los comprobantes. No se usa Excel como registro operativo porque un libro compartido no maneja bien las escrituras simultáneas de un flujo y de varias personas.
- **Power Automate** solo mueve la información entre Outlook, SharePoint y Excel. No guarda datos propios.
- La aprobación y la carga en Operam siguen siendo manuales.

## Consecuencias
- Cada herramienta hace lo que hace bien, y no hace falta ninguna licencia adicional.
- El PDF de la factura no se lee en esta fase: solo se valida que exista y su extensión.
- Un cambio de estado se hace en SharePoint; Excel solo lo refleja al actualizar.
