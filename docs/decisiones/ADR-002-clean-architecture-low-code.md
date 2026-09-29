# ADR-002 · Clean Architecture en una solución low-code

**Estado:** aceptada.

## Contexto
Se pidió usar Clean Architecture. En producción no hay código compilado: hay configuración de Power Automate, estructura de SharePoint y consultas de Power Query. El riesgo típico de estas soluciones es que la misma regla termine escrita en varios lugares (los estados en SharePoint, en el flujo y en Excel) y que se desalineen con el tiempo.

## Decisión
- Organizar el repositorio en cuatro capas (dominio, aplicación, adaptadores, infraestructura) con la regla de dependencias hacia adentro.
- Escribir las reglas **una sola vez** en `src/1-dominio/dominio.json` y **derivar** de ahí la configuración de los flujos, la consulta `Dominio` del Excel y la validación del esquema de SharePoint.
- Tener una **implementación de referencia** de las reglas en PowerShell (`Legalizaciones.Dominio.psm1`) con ejemplos verificados. No corre en producción: es la especificación ejecutable con la que se comprueba el flujo.
- Hacer cumplir la arquitectura con **pruebas** (`Arquitectura.Tests.ps1`) en vez de con revisión manual.

## Consecuencias
- Cambiar una regla es cambiar el dominio, regenerar y ejecutar las pruebas; las pruebas dicen qué más falta.
- Algunas cosas no pueden leer la configuración y se duplican a propósito, cubiertas por pruebas: las condiciones de los desencadenadores de Power Automate y la limpieza de nombres de carpeta, que Power Automate no puede hacer con un bucle.
- No viola la restricción de "sin desarrollo a la medida": los scripts crean y verifican las herramientas, no son una aplicación en producción.
