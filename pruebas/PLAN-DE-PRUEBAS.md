# Plan de pruebas

## 1. Niveles de prueba

| Nivel | Qué prueba | Cómo se ejecuta | Necesita tenant |
|---|---|---|---|
| **Automáticas** | Reglas del dominio, coherencia entre capas, scripts de SharePoint (contra un SharePoint simulado), comandos PnP, Power Query real en Excel, plantilla y datos de prueba | `./pruebas/Invoke-Pruebas.ps1` (PowerShell 7) | No |
| **Aceptación (punta a punta)** | Correo real → LEG-01 → SharePoint → avisos → LEG-02 → Excel | Escenarios E01 a E14 de este documento | Sí |
| **Puesta en marcha** | Lista de verificación antes de usarlo con colaboradores | Sección 8 | Sí |

Las pruebas automáticas se ejecutan **antes** de cada cambio que se lleve al tenant y deben terminar con `RESULTADO: todas las pruebas pasan`.

## 2. Preparación de las pruebas de aceptación

1. SharePoint aprovisionado y verificado (`02-Verificar-SharePoint.ps1` con FALLA: 0).
2. LEG-01 y LEG-02 construidos y **activados**.
3. Adjuntos de prueba generados:
   ```powershell
   python pruebas/aceptacion/generar_datos_prueba.py
   ```
   Crea en `pruebas/datos/` facturas PDF, un XML con estructura CFDI 4.0, comprobantes Excel (con y sin plantilla) y archivos no válidos (.png, .xls). Todos son ficticios y dicen "DOCUMENTO DE PRUEBA".
4. Una cuenta de correo para enviar las pruebas (la suya). Ella recibirá los avisos al colaborador.
5. Si ya se ejecutaron pruebas antes, cambie a **Rechazada** los casos de prueba anteriores (asunto con "Prueba E…" o "Prueba V…"). Los casos rechazados no cuentan como duplicados; si no se hace esto, los nuevos saldrán como posibles duplicados de los anteriores.

## 3. Escenarios

Cada asunto lleva la etiqueta **"Prueba <id>"**, con la que se encuentra su caso. El detalle está en [`aceptacion/Escenarios.json`](aceptacion/Escenarios.json).

| Id | Escenario pedido | Adjuntos | Resultado esperado en SharePoint | Correo esperado |
|---|---|---|---|---|
| **E01** | Legalización completa | factura_A001.pdf, Gastos_Quincena.xlsx | Caso Pendiente, sin observaciones, 2 archivos en `Pendientes/AAAA/MM/Prueba E01` | "Legalización recibida" al remitente |
| **E02** | Completa con XML | factura_A002.pdf, Gastos_Quincena.xlsx, cfdi_A001.xml | Pendiente, sin observaciones, 3 archivos | "Legalización recibida" |
| **E03** | Incompleta: sin comprobante | factura_B100.pdf | Pendiente; Observaciones: "Falta el comprobante quincenal…"; 1 archivo | "Legalización incompleta" que dice qué falta |
| **E04** | Incompleta: sin factura | Gastos_Quincena.xlsx | Pendiente; "Falta la factura en PDF."; 1 archivo | "Legalización incompleta" |
| **E05** | Incompleta: sin adjuntos | (ninguno) | Pendiente; faltan factura y comprobante; 0 archivos | "Legalización incompleta" |
| **E06** | **Formato no soportado** (factura como imagen) | factura_foto.png, Gastos_Quincena.xlsx | Pendiente; "Falta la factura…" y "…no se guardaron… factura_foto.png"; 1 archivo (la imagen **no** se guarda) | "Legalización incompleta" |
| **E07** | Completa con un archivo extra no válido | factura_C200.pdf, Gastos_Quincena.xlsx, Gastos_Antiguo.xls | Pendiente; la observación menciona Gastos_Antiguo.xls; 2 archivos | "Legalización recibida" con notas |
| **E08** | Asunto sin el nombre | factura_C201.pdf, Gastos_Quincena.xlsx | Pendiente; "El asunto no trae el nombre…"; Colaborador = parte de su correo antes de la @ | "Legalización recibida" con notas |
| **E09** | Correo que no es una legalización | factura_C202.pdf | **No** se crea caso | Ninguno |
| **E10** | **Posible duplicado** (reenvío de E01) | factura_A001.pdf, Gastos_Quincena.xlsx | Pendiente; "Posible duplicado de <ID de E01>…" | "Legalización recibida" al remitente **y** "[Legalizaciones] Posible duplicado" a Contabilidad |
| **E14** | Comprobante sin la plantilla | factura_C202.pdf, Gastos_Formato_Libre.xlsx | Pendiente, sin observaciones, 2 archivos | "Legalización recibida". En Excel aparece en *Errores_Lectura* y *Alertas* |
| **E11** | **Volumen alto**: 15 correos seguidos el mismo día | factura_V01…V15.pdf + Gastos_Quincena.xlsx | 15 casos, todos con ID distinto, sin observaciones ni errores | 15 "Legalización recibida" |
| **E12** | Cierre como Aprobada (manual) | — | Caso de E01: Aprobada, con Fecha Cierre, 2 archivos en `Aprobadas/…` | "Legalización aprobada" |
| **E13** | Cierre como Rechazada (manual) | — | Caso de E03: Rechazada, con Fecha Cierre, 1 archivo en `Rechazadas/…` | "Legalización rechazada" con el motivo |

## 4. Ejecución

### 4.1 Enviar los correos

**Opción A: con el script** (requiere Outlook clásico):
```powershell
.\pruebas\aceptacion\Enviar-CorreosPrueba.ps1
```
Muestra la lista de correos y pide confirmación antes de enviar. Envía E01 a E10, E14 y luego los 15 de E11, con 5 segundos entre uno y otro.

**Opción B: a mano.** Envíe desde su correo, **en el orden de la tabla** (E10 siempre después de E01), un correo por escenario, con el asunto exacto de `Escenarios.json` y los adjuntos de `pruebas/datos/`.

### 4.2 Validar en SharePoint
Espere a que terminen las ejecuciones de LEG-01 (en Power Automate: *Mis flujos* › LEG-01 › *Historial de ejecuciones*) y ejecute:
```powershell
.\pruebas\aceptacion\Validar-Escenarios.ps1
```
El resultado queda en pantalla y en `logs/validacion-escenarios_*.csv`, y debe terminar con **FALLA: 0**.

### 4.3 Revisar los correos (a mano)
Compruebe en su bandeja la columna "Correo esperado" (y en la de Contabilidad, el aviso de E10). Revise que:
- el número de caso del correo coincida con el de SharePoint;
- en los incompletos se diga exactamente lo que falta.

### 4.4 Cierre (E12 y E13)
1. En el caso de **Prueba E01**, cambie el Estado a **Aprobada** y guarde.
2. En el caso de **Prueba E03**, escriba un motivo en Observaciones, cambie a **Rechazada** y guarde.
3. Espere de 1 a 2 minutos y ejecute:
   ```powershell
   .\pruebas\aceptacion\Validar-Escenarios.ps1 -Escenarios E12, E13
   ```
4. Revise los correos "Legalización aprobada" y "Legalización rechazada".

### 4.5 Análisis en Excel
1. Genere el libro con el tenant real (`Generar-LibroAnalisis.ps1`), súbalo a *Documentos* del sitio y ábralo con Excel de escritorio.
2. **Datos › Actualizar todo** e inicie sesión.
3. Compruebe:
   - **Registro** muestra todos los casos de prueba, con fechas en la hora local;
   - **Gastos** tiene 3 líneas por cada comprobante de plantilla (Gastos_Quincena.xlsx);
   - **Errores_Lectura** muestra el comprobante de E14;
   - **Consolidado_Operam** muestra las 3 líneas del caso de E01 (Aprobada en E12);
   - al anotar ese ID en **Control_Operam** y actualizar, desaparece del consolidado;
   - **Alertas** muestra "Comprobante ilegible" (E14) y "Posible duplicado sin resolver" (E10).

### 4.6 Error técnico (opcional, con TI)
Para probar el manejo de errores sin afectar datos reales:
1. En una **copia** de LEG-01, cambie a propósito *Ruta de acceso de la carpeta* de `Crear_archivo` por una biblioteca que no existe.
2. Envíe un correo de prueba y compruebe:
   - la ejecución queda como **Con errores**;
   - el caso queda con "ERROR TÉCNICO…" en Observaciones;
   - TI recibe "[Legalizaciones] Error técnico en LEG-01" con el enlace a la ejecución.
3. Elimine la copia del flujo.

## 5. Criterios de aceptación

- [ ] `Invoke-Pruebas.ps1`: todas las pruebas automáticas pasan.
- [ ] `Validar-Escenarios.ps1` (E01 a E11 y E14): **FALLA: 0**.
- [ ] `Validar-Escenarios.ps1 -Escenarios E12, E13`: **FALLA: 0**.
- [ ] Todos los correos esperados llegaron con el número de caso correcto, y ningún aviso de error inesperado.
- [ ] Volumen (E11): los 15 casos se procesaron en menos de 15 minutos, sin ejecuciones fallidas.
- [ ] Excel: los puntos de la sección 4.5 se cumplen.

## 6. Registro de resultados

| Fecha | Quién | Nivel / escenarios | Resultado | CSV / observaciones |
|---|---|---|---|---|
| | | Automáticas | | |
| | | Aceptación E01–E11, E14 | | |
| | | Cierre E12–E13 | | |
| | | Excel (4.5) | | |

## 7. Después de las pruebas

- Cambie los casos de prueba a **Rechazada** con la observación "PRUEBA". Si TI prefiere, puede eliminarlos a mano.
- Borre de la hoja Control_Operam las filas de prueba.

## 8. Lista de verificación de puesta en marcha

- [ ] Configuración completa (`config/Config.Legalizaciones.psd1`), con país confirmado.
- [ ] SharePoint verificado sin fallas; permisos solo para TI, Contabilidad y la cuenta del flujo.
- [ ] Buzón compartido con *Acceso total* y *Enviar como* para la cuenta de servicio.
- [ ] LEG-01 y LEG-02 activados, con dueño la cuenta de servicio y un copropietario de TI.
- [ ] Pruebas de aceptación aprobadas (sección 5).
- [ ] Libro de análisis en *Documentos* del sitio, probado por Contabilidad.
- [ ] Plantilla del comprobante e instrucciones enviadas a los colaboradores (`docs/04-instrucciones-colaboradores.md`).
- [ ] Contabilidad leyó `docs/03-manual-contabilidad.md`.
