# Adaptador de Power Automate

Dos flujos de nube implementan los casos de uso del proceso. Power Automate solo **orquesta**: no guarda datos propios. Lee y escribe en Outlook y SharePoint, y toma las reglas del dominio desde el JSON de configuración.

| Flujo | Qué hace | Guía paso a paso |
|---|---|---|
| **LEG-01 Recepción de legalizaciones** | Pasos 2 a 6: detecta el correo, registra el caso, guarda los adjuntos, valida y notifica | [LEG-01-Recepcion.md](LEG-01-Recepcion.md) |
| **LEG-02 Cierre de casos** | Paso 7 (complemento): al aprobar o rechazar, mueve los archivos, registra la Fecha Cierre y avisa al colaborador | [LEG-02-Cierre.md](LEG-02-Cierre.md) |

Archivos de apoyo:
- [`expresiones.md`](expresiones.md): catálogo de expresiones de reglas de negocio, con ejemplos verificados.
- [`plantillas-correo/`](plantillas-correo/): cuerpos HTML de los correos con marcadores `{{...}}`.
- `salida/Configuracion-Flujos.json`: se genera con `src/4-infraestructura/power-automate/Generar-ConfiguracionFlujos.ps1` y se pega en la acción *Configuracion* de ambos flujos.

## Orden de construcción

1. Terminar el punto de SharePoint (aprovisionar y verificar sin fallas).
2. Tener la cuenta de servicio y el buzón compartido con **Acceso total** y **Enviar como** para esa cuenta.
3. Completar `config/Config.Legalizaciones.psd1` (sección `Flujos`) y generar `salida/Configuracion-Flujos.json`.
4. Construir **LEG-01** y probarlo con el escenario E01.
5. Construir **LEG-02** y probarlo con los escenarios E12 y E13.
6. Ejecutar el plan de pruebas completo (`pruebas/PLAN-DE-PRUEBAS.md`).

## Licencias

Solo se usan conectores **estándar**: Office 365 Outlook y SharePoint, además de las acciones integradas (Redactar, Filtrar matriz, Seleccionar, Condición, Ámbito, Variables, Terminar). Todo está incluido en las licencias de Microsoft 365 con Power Automate. No se usa HTTP (premium), AI Builder ni otros conectores de IA.

## Cambiar una regla

1. Cambie `src/1-dominio/dominio.json`. Por ejemplo, los días de la ventana de duplicados o los textos de los mensajes.
2. Ejecute las pruebas (`pruebas/Invoke-Pruebas.ps1`): le dirán qué otras piezas hay que ajustar.
3. Genere de nuevo `salida/Configuracion-Flujos.json` y reemplace el contenido de *Configuracion* en LEG-01 y LEG-02.
4. Solo si cambian los nombres de los estados, actualice además las condiciones de los desencadenadores (E-01 y E-12 del catálogo), las opciones de la columna Estado en SharePoint y el libro de Excel (vuelva a generarlo).
