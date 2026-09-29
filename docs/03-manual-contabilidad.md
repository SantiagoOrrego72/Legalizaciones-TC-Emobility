# Manual de operación · Contabilidad

Qué hace Contabilidad cada día con la solución. Lo automático lo hacen los flujos; las decisiones siguen siendo de Contabilidad.

## Dónde está todo

| Qué | Dónde |
|---|---|
| Casos (registro) | Sitio **Legalizaciones TC** › lista **Registro_Legalizaciones_TC** |
| Soportes (archivos) | Sitio **Legalizaciones TC** › biblioteca **Legalizaciones_TC** › `Pendientes` / `Aprobadas` / `Rechazadas` › año › mes › colaborador |
| Análisis y consolidación | Libro **Analisis_Legalizaciones_TC.xlsx** en la biblioteca *Documentos* del mismo sitio |
| Avisos de posible duplicado | Correo de Contabilidad (asunto `[Legalizaciones] Posible duplicado - …`) |

## Rutina diaria: revisar casos

1. Abra la lista del registro y elija la vista **Pendientes de revisión**. Muestra los casos Pendiente y En revisión, los más antiguos primero.
2. Abra un caso por su **ID Legalización** (`LEG-AAAAMM-NNNNN`) y lea **Observaciones**:
   - *Falta la factura* o *Falta el comprobante*: el colaborador ya recibió un correo con lo que falta. Espere su reenvío, que llegará como un **caso nuevo**.
   - *Posible duplicado de LEG-…*: compare con el caso citado antes de aprobar.
   - *ERROR TÉCNICO*: el correo no se terminó de procesar. TI ya recibió un aviso; no apruebe ese caso hasta que TI lo revise.
3. Busque los soportes en la biblioteca con el ID. Todos los archivos del caso empiezan con el ID y tienen la columna *ID Legalización*.
4. Si va a tomarse tiempo, cambie el Estado a **En revisión** para que el equipo sepa que alguien lo tiene.
5. Decida:
   - **Aprobada**: si quiere, escriba una nota en Observaciones y guarde.
   - **Rechazada**: escriba **primero el motivo en Observaciones**, porque el correo al colaborador lo incluye, y luego cambie el Estado y guarde.
6. En uno o dos minutos el flujo LEG-02 mueve los archivos a Aprobadas o Rechazadas, llena la **Fecha Cierre** y avisa al colaborador.

El correo de cierre incluye las Observaciones, **menos las líneas que empiezan con `[Interno]`**. Los avisos automáticos de duplicado y de error técnico ya llevan ese prefijo. Si quiere dejar una nota solo para el equipo, empiece la línea con `[Interno]`.

### Casos especiales

| Situación | Qué hacer |
|---|---|
| El colaborador reenvió completo un caso que estaba incompleto | El reenvío es un caso nuevo (y puede marcarse como posible duplicado del anterior; es normal). Rechace el caso incompleto con la observación *"Reemplazado por LEG-…"* y trabaje el nuevo |
| Es un duplicado real | Rechace el caso nuevo con *"Duplicado de LEG-…"* |
| Aprobé o rechacé por error | No lo cambie a otro estado: los estados finales no se reabren. Pida a TI que lo corrija. El historial de versiones del elemento muestra quién cambió qué y cuándo |
| Un caso cerrado sigue sin Fecha Cierre | LEG-02 falló al mover los archivos. TI recibió un aviso; cuando lo corrija, guarde de nuevo el caso |

### No haga esto
- No cambie el **ID Legalización** ni el nombre de los archivos.
- No llene la **Fecha Cierre** a mano: impide que LEG-02 mueva los archivos.
- No mueva los archivos entre carpetas a mano.
- No suba soportes a mano sin llenar su columna *ID Legalización*, porque LEG-02 no los encontraría.

## Rutina de consolidación y carga en Operam

1. Abra **Analisis_Legalizaciones_TC.xlsx** con **Excel de escritorio**: en SharePoint, *Abrir* › *Abrir en la aplicación de escritorio*.
2. **Datos › Actualizar todo**. La primera vez:
   - si pide iniciar sesión, elija **Cuenta organizacional** e inicie sesión con su cuenta de Microsoft 365;
   - si pregunta por los niveles de privacidad, elija **Organizacional** para el sitio de SharePoint.
3. Revise la hoja **Alertas** y resuelva lo que corresponda: casos viejos, duplicados sin resolver, totales que no cuadran, comprobantes ilegibles.
4. La hoja **Consolidado_Operam** tiene las líneas de gasto de los casos **Aprobados** que aún no se cargaron. Cárguelas en Operam como hasta ahora.
5. En la hoja **Control_Operam**, agregue una fila por caso cargado: ID Legalización, Fecha Carga, Cargado Por y Referencia Operam.
6. **Actualizar todo** otra vez: esos casos desaparecen del consolidado y no se cargan dos veces.
7. Guarde el libro.

La hoja **Resumen** muestra los casos por mes y estado, y la hoja **Registro** tiene todos los casos. El estado **no se cambia en Excel**: se cambia en SharePoint, y Excel lo refleja al actualizar.

## Trazabilidad y auditoría
- **Quién cambió el estado y cuándo:** en la lista, menú del elemento › *Historial de versiones*.
- **Qué llegó en cada correo:** columna *Nombre Archivo* y archivos con el mismo ID.
- **Qué se cargó a Operam y cuándo:** hoja Control_Operam.
