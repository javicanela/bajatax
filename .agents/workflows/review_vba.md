---
description: Revisar y validar un módulo VBA existente
---
# Workflow: Review VBA Module

Este workflow instruye a Antigravity a auditar código VBA de acuerdo a los estándares de BajaTax.

## Pasos

1. **Identificar archivo**: El usuario indica qué archivo `src/vba-modules/*.bas` requiere revisión.
2. **Validar el Código (Subagente Reviewer)**: Antigravity lee el contenido del archivo `.bas` e instancia un subagente utilizando el prompt `.agents/prompts/reviewer.md` para que lo evalúe.
3. **Presentar Reporte**: Antigravity muestra el reporte del subagente. Si hubo fallos críticos, se solicita permiso al usuario para que Antigravity aplique las correcciones directamente.
