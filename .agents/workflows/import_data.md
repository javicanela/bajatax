---
description: Guía de Importación Inteligente de Datos (REGISTROS)
---
# Workflow: Import Data

Este workflow explica el proceso de ingesta de datos en `AUTOMATIZACION_v7.xlsm`.

## Pasos Analíticos para el Agente

Cuando se requiera importar un CSV o Excel sucio mediante Python o explicar la importación en VBA, sigue consultando la `docs/SPEC_TECNICA_DETALLADA.md`.
Recuerda que el motor del código de importación prioriza la **Detección por Contenido** (Regex de 12-13 caracteres para RFC, etc.) por encima del mapeo de headers directos (ya que los archivos origen cambian mucho).

Una vez importados a la hoja `REGISTROS`, ocurre la distribución:
1. Se verifica si el RFC existe en `DIRECTORIO`. Si no, se da de alta.
2. Se copian los datos transaccionales a `OPERACIONES`.
3. Se previenen duplicados usando la clave compuesta: `RFC + Concepto + Fecha Cobranza + Monto`
