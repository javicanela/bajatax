---
description: Generar un nuevo módulo VBA y validar su código
---
# Workflow: Generate VBA Module

Este workflow instruye a Antigravity a crear y validar un módulo VBA de forma autónoma.

## Pasos

1. **Recibir instrucciones**: Lee el requerimiento del usuario sobre qué debe hacer el nuevo módulo.
2. **Generar el Código (Subagente Coder)**: Lanza un subagente utilizando el prompt `.agents/prompts/coder.md` para que genere el código fuente del módulo `.bas` en la carpeta temporal de trabajo o lo proponga en tu memoria.
3. **Validar el Código (Subagente Reviewer)**: Lanza un segundo subagente utilizando el prompt `.agents/prompts/reviewer.md` proporcionándole el código generado en el paso 2 para que lo revise contra el checklist completo (Mac, Errores, Seguridad, etc.).
4. **Iterar**: Si el *Reviewer* encuentra fallos, haz que el *Coder* los corrija. Repite hasta obtener un veredicto de `✅ APROBADO`.
5. **Guardar Archivo**: Guarda el código final en `src/vba-modules/`.
6. **Sugerir Inyección**: Una vez guardado, sugiere al usuario ejecutar el script local `src/python/instalar_vba.py` para inyectarlo en el archivo `.xlsm`.
