# Arquitectura

## 1. Contexto

Emobility legaliza unas **127 facturas de tarjetas corporativas al mes**. Hoy todo se hace por correo: cerca de 5 minutos de revisión por factura, sin trazabilidad y con riesgo de duplicados y errores. Operam es el ERP de destino.

**Restricciones de la Fase 1** (no negociables):
- Solo Microsoft 365: Outlook, SharePoint, Excel Online y Power Automate.
- Sin IA, OCR, AI Builder, Azure Document Intelligence ni lectura automática de documentos. Del PDF solo se valida que exista y su extensión.
- Sin licencias adicionales ni desarrollo de software a la medida en producción.
- La aprobación y la carga en Operam siguen siendo manuales, a propósito.

## 2. Vista general

```text
 Colaborador ──correo (PDF + Excel [+ XML])──► Buzón de legalizaciones (Outlook)
                                                         │
                                                         ▼
                                   ┌────────── LEG-01 Recepción (Power Automate) ──────────┐
                                   │ registra el caso · guarda adjuntos · valida · notifica │
                                   └───────────────┬───────────────────────┬───────────────┘
                                                   ▼                       ▼
                       SharePoint: Registro_Legalizaciones_TC   SharePoint: Legalizaciones_TC
                       (un caso por correo, con estado)          (Pendientes/Aprobadas/Rechazadas/AAAA/MM/Colaborador)
                                                   ▲                       ▲
              Contabilidad revisa y cambia el Estado ──► LEG-02 Cierre: mueve archivos, fecha y avisa
                                                   │
                                                   ▼
                                  Excel + Power Query (Registro, Gastos, Alertas, Consolidado)
                                                   │
                                                   ▼
                                   Contabilidad carga manualmente en Operam
```

## 3. Clean Architecture en una solución low-code

Clean Architecture ordena un sistema en capas concéntricas. Las **reglas de negocio** quedan en el centro, independientes de las herramientas, y **las dependencias apuntan siempre hacia adentro**. Aquí lo que corre en producción no es código compilado, sino configuración de Power Automate, estructura de SharePoint y consultas de Power Query. Se aplican los mismos principios:

```text
┌──────────────────────────────────────────────────────────────────────┐
│ 4. Infraestructura   scripts PnP, generador Excel (COM), Python,     │
│                      instalador de prerrequisitos                     │
│  ┌────────────────────────────────────────────────────────────────┐  │
│  │ 3. Adaptadores   contrato SharePoint, flujos LEG-01/LEG-02,     │  │
│  │                  consultas Power Query, plantillas de correo     │  │
│  │  ┌──────────────────────────────────────────────────────────┐  │  │
│  │  │ 2. Aplicación   casos de uso CU-01..CU-07 y puertos       │  │  │
│  │  │  ┌────────────────────────────────────────────────────┐  │  │  │
│  │  │  │ 1. Dominio   dominio.json: estados, documentos,     │  │  │  │
│  │  │  │              duplicados, ID, carpetas, mensajes     │  │  │  │
│  │  │  └────────────────────────────────────────────────────┘  │  │  │
│  │  └──────────────────────────────────────────────────────────┘  │  │
│  └────────────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────────────┘
                      las dependencias apuntan hacia adentro ▲
```

| Capa | Contenido | Carpeta | Puede depender de |
|---|---|---|---|
| **1 · Dominio** | Reglas de negocio puras (`dominio.json`) e implementación de referencia probada (`Legalizaciones.Dominio.psm1`) | `src/1-dominio` | Nada |
| **2 · Aplicación** | Casos de uso del proceso y puertos: las capacidades que necesitan, sin decir con qué herramienta | `src/2-aplicacion` | Dominio |
| **3 · Adaptadores** | Traducción a cada herramienta: esquema de SharePoint, guías exactas de los flujos, consultas M, plantillas | `src/3-adaptadores` | Aplicación, Dominio |
| **4 · Infraestructura** | Lo que crea o configura las herramientas: aprovisionamiento PnP, generadores, instalador | `src/4-infraestructura` | Todas |

### Cómo se hace cumplir la regla de dependencias

1. **Fuente única.** Las reglas de negocio se escriben **solo** en `src/1-dominio/dominio.json`. Las capas externas se derivan de ahí:
   - `Generar-ConfiguracionFlujos.ps1` convierte el dominio en el JSON que usan los flujos (acción *Configuracion*): extensiones, ventana de duplicados, mensajes y formato del ID;
   - `Generar-LibroAnalisis.ps1` inserta el dominio en la consulta `Dominio` del libro de Excel (estados, tipos de documento);
   - `01-Aprovisionar-SharePoint.ps1` **se niega a ejecutarse** si las opciones de Estado del esquema no coinciden con los estados del dominio.
2. **Pruebas de arquitectura** (`pruebas/unitarias/Arquitectura.Tests.ps1`). Comprueban que:
   - el dominio no llama a SharePoint, Outlook ni Excel, ni tiene datos del entorno;
   - los adaptadores no contienen scripts;
   - cada clave, acción, variable y marcador que usan los flujos existe;
   - las expresiones del catálogo están copiadas idénticas y balanceadas;
   - las consultas de Power Query usan los nombres del contrato.
3. **Especificación ejecutable.** Cada regla tiene ejemplos verificados (`pruebas/unitarias/Dominio.Tests.ps1`), que son los mismos que se usan para comprobar a mano el flujo (`expresiones.md`).

## 4. Puertos y adaptadores

Los casos de uso piden **capacidades** (puertos). Cada herramienta es un adaptador intercambiable. Esto permite que las fases futuras agreguen, por ejemplo, lectura automática de facturas **sin tocar el dominio**.

| Puerto | Adaptador en Fase 1 | Posible adaptador futuro |
|---|---|---|
| Buzón de entrada | Outlook: desencadenador de LEG-01 | El mismo |
| Registro de casos | Lista de SharePoint `Registro_Legalizaciones_TC` | La misma o Dataverse |
| Almacén de soportes | Biblioteca de SharePoint `Legalizaciones_TC` | La misma |
| Notificador | Outlook: envío desde el buzón compartido | Teams |
| Lector de comprobantes | Power Query sobre la tabla `tblGastos` de la plantilla | El mismo |
| Lector de facturas (PDF) | **No existe en Fase 1** (requeriría OCR o IA) | AI Builder / Azure Document Intelligence |
| Lector de XML (CFDI) | Solo se verifica que exista | Power Query sobre el XML (validaciones fiscales) |
| Análisis y consolidación | Excel + Power Query | Power BI |
| ERP | Carga manual en Operam | Integración con Operam |

## 5. Vista de ejecución (8 pasos acordados)

| Paso | Qué pasa | Componente | Caso de uso |
|---|---|---|---|
| 1 | El colaborador envía el correo con el asunto acordado y los adjuntos | Outlook | — |
| 2 | Se detecta el correo (condición del asunto, sin importar tildes ni mayúsculas) | LEG-01, desencadenador | CU-01 |
| 3 | Se guarda cada adjunto válido en su carpeta, con el ID como prefijo | LEG-01, bloque F | CU-01 |
| 4 | Se crea el caso con estado Pendiente y un ID único | LEG-01, bloques D y G | CU-01 |
| 5 | Validaciones: factura, comprobante, formatos, posible duplicado | LEG-01, bloques B, C y E | CU-02, CU-03 |
| 6 | Avisos: completa, incompleta (qué falta) o posible duplicado (a Contabilidad) | LEG-01, bloque H | CU-04 |
| 7 | Contabilidad revisa en SharePoint y aprueba o rechaza; LEG-02 mueve los archivos, fecha el cierre y avisa | SharePoint + LEG-02 | CU-05 |
| 8 | Contabilidad consolida en Excel y carga en Operam | Excel/Power Query + persona | CU-06, CU-07 |

## 6. Vista de despliegue

| Dónde | Qué |
|---|---|
| SharePoint Online, sitio `/sites/Legalizaciones-TC` | Lista del registro, biblioteca de soportes y el libro de análisis (subido a *Documentos*) |
| Power Automate, entorno predeterminado | LEG-01 y LEG-02, con dueño la cuenta de servicio |
| Exchange Online | Buzón compartido de legalizaciones: la cuenta de servicio tiene *Acceso total* y *Enviar como* |
| Entra ID | Aplicación "PnP PowerShell - Legalizaciones TC", con permisos delegados, solo para los scripts de TI |
| Equipo de TI (esta carpeta) | Scripts, generadores, pruebas y documentación. No se ejecuta nada en producción desde aquí |

## 7. Atributos de calidad

| Atributo | Cómo se logra |
|---|---|
| **Trazabilidad** | ID único por caso en el registro y en cada archivo (prefijo y columna). Historial de versiones de la lista: quién cambió cada estado y cuándo. Observaciones con los hallazgos |
| **Concurrencia** | Simultaneidad 1 en ambos flujos (los correos se procesan en orden). ID basado en el número interno de SharePoint, con valores únicos |
| **Seguridad** | Sitio dedicado sin Grupo de Microsoft 365; solo TI (Control total), Contabilidad y la cuenta del flujo (Colaborar). Sin uso compartido externo; solo los propietarios comparten. Aplicación de Entra ID con permisos delegados |
| **Manejo de errores** | Patrón Intentar/Capturar en ambos flujos: caso marcado con "ERROR TÉCNICO", aviso a TI con enlace a la ejecución y ejecución terminada como fallida. Alertas de calidad en Excel |
| **Mantenibilidad** | Fuente única de reglas, contratos explícitos, scripts que se pueden repetir (idempotentes) y 187 pruebas automáticas |
| **Costo** | Solo conectores estándar incluidos en Microsoft 365; nada premium ni de Azure |

## 8. Decisiones

Las decisiones de arquitectura y su justificación están en [`docs/decisiones/`](decisiones/).
