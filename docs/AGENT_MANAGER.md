# Agent Manager: Workspaces Paralelos

Para orquestar múltiples agentes en paralelo sin que se interfieran (ej: un agente trabajando en la importación y otro en WhatsApp), Antigravity utiliza entornos aislados a través de **VS Code Workspaces**.

## ¿Por qué usar Workspaces separados?
Cada ventana de VS Code ejecuta **su propia instancia aislada del agente**. Al abrir un "Sub-Workspace", le limitamos la visibilidad (a través de `files.exclude`) solo a los archivos relevantes de su módulo. Esto logra tres cosas:
1. **Foco**: El agente de importación no "ve" ni se confunde con los archivos de WhatsApp.
2. **Contexto optimizado**: Se gastan menos tokens y el agente es más rápido y asertivo.
3. **Paralelismo Seguro**: Puedes tener un agente corriendo scripts o tests en una ventana, mientras en la otra discutes arquitectura con otro agente.

## ¿Cómo Activar Múltiples Agentes?

1. En tu ventana principal (donde estás ahora), ve a **File > New Window** (Archivo > Nueva Ventana) o usa el atajo `Ctrl+Shift+N` (Windows) / `Cmd+Shift+N` (Mac).
2. En la ventana nueva, ve a **File > Open Workspace from File...**
3. Selecciona uno de los archivos `.code-workspace` generados en la raíz del proyecto:
   - `bajatax-import.code-workspace`: Excluye carpetas de WhatsApp y PDF, ideal para enfocar al agente en la limpieza de datos y Excel.
   - `bajatax-whatsapp.code-workspace`: Excluye las carpetas de Importación, enfocado en Evolution API y envíos.
4. Una vez abierto el workspace, inicia el Agente en esa nueva ventana y dile: *"Actúa como el Subagente de [WhatsApp/Importación]. Revisa tu entorno y dime con qué empezamos."*

En tu ventana actual (BajaTax Original), puedes seguir operando a **Antigravity como el Orquestador Central**, revisando el código que los otros subagentes modifican.
