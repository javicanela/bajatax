---
name: token-optimizer
description: >
  Usar esta skill cuando el usuario quiera saber qué modelo usar para una tarea
  específica en BajaTax, cuando los límites de tokens se estén agotando, cuando
  un prompt sea muy largo y lento, o cuando quiera optimizar velocidad de Roo Code.
  Activar cuando el usuario diga "se me acabó Gemini", "cuál modelo uso para esto",
  "el prompt es muy largo", "Groq da error de límite", "cómo hacer Roo Code más rápido",
  o cuando una tarea lleve más de 5 minutos sin completarse. Esta skill define la
  estrategia de los 5 modelos disponibles para maximizar productividad sin desperdiciar
  tokens — úsala siempre antes de comenzar una tarea larga o compleja.
---

# Token Optimizer — BajaTax

## Los 5 modelos en Roo Code

| Modelo | Provider | Límite real | Fortaleza | Cuándo falla |
|---|---|---|---|---|
| `gemini-2.5-pro` *(verificar si es free)* | Google AI Studio | ~50 req/día, 2 req/min gratis | Contexto 1M tokens, múltiples archivos | Lento, límite diario bajo; puede tener costo |
| `deepseek-r1-distill-llama-70b` | Groq | 14,400 tok/min | Razonamiento, bugs lógicos, velocidad | Prompts muy grandes en ráfaga |
| `qwen/qwen3-coder:free` | OpenRouter | Variable (~100 req/día) | Código VBA especializado | Límite impredecible |
| `qwen3:8b` | Ollama local | Ilimitado | Siempre disponible, privacidad | Más lento, menos capaz |
| `deepseek-r1:8b` | Ollama local | Ilimitado | Razonamiento sin internet | Más lento |

---

## Estrategia de selección por tipo de tarea

### Gemini (verificar modelo free) → Para ver mucho contexto a la vez
```
⚠️ IMPORTANTE — modelo a usar:
  El modelo con contexto 1M tokens en AI Studio puede cambiar.
  En AI Studio, buscar el modelo marcado como "free" en la lista (puede ser
  gemini-2.5-pro, gemini-2.0-flash, u otro según fecha). No asumir que
  gemini-2.5-pro es siempre gratuito — confirmar en aistudio.google.com
  antes de usarlo en Roo Code para evitar cargos inesperados.

✅ Usar cuando:
  - Analizar 3+ módulos .bas simultáneamente
  - Entender cómo se conectan todos los módulos del proyecto
  - Diseño de arquitectura completo
  - Revisar toda la estructura antes de refactorizar

❌ No desperdiciar en:
  - Preguntas sobre una sola función
  - Correcciones de una línea
  - Tasks repetitivos o mecánicos

⚠️ Nota: "Free up to 2 requests/minute — after that, billing applies"
   Esto aplica al modelo free vigente. Si el modelo free cambió, billing
   puede activarse desde la primera llamada. Verificar en AI Studio.
   Si usas AI Studio gratis, esperar 30s entre requests grandes.
```

### Deepseek R1 via Groq → Para bugs y razonamiento
```
✅ Usar cuando:
  - Error con número específico (1004, 91, 13...)
  - Lógica que falla en casos de borde
  - Entender POR QUÉ algo no funciona
  - Algoritmo de deduplicación o importación

⚠️ Límite: 14,400 tokens/minuto
   Si el prompt es grande, dividirlo en partes con 30s entre ellas.
   Señal de límite agotado: error "rate_limit_exceeded"
```

### Qwen3-coder:free → Para generación de código puro
```
✅ Usar cuando:
  - Generar módulo .bas completo desde spec
  - Refactoring con estructura ya definida
  - Completar funciones con patrón conocido
  - Gemini y Groq se agotaron y aún necesitas código

⚠️ Si falla → cambiar a Ollama local inmediatamente
```

### Ollama local → Fallback ilimitado
```
✅ Usar cuando:
  - Todos los demás tienen límites agotados
  - Tareas repetitivas que no necesitan máxima calidad
  - Ediciones menores a código existente
  - Generar documentación, comentarios, docstrings

⚠️ Limitaciones reales:
  - qwen3:8b: bueno para código simple, pierde contexto en archivos muy largos
  - deepseek-r1:8b: mejor para razonamiento, más lento que qwen para código puro
  - Ambos: no reemplazar Gemini para análisis multi-archivo
```

---

## Cuándo cambiar de modelo — señales claras

| Señal | Acción |
|---|---|
| "rate_limit_exceeded" en Groq | Esperar 60s o cambiar a Qwen3-coder |
| Gemini tarda >3 minutos | Task muy grande → dividirla en subtareas |
| Qwen3-coder devuelve error 429 | Cambiar a Ollama local |
| Ollama produce código con errores obvios | Resubir task a Groq o Gemini al día siguiente |
| Límite diario de Gemini agotado | Groq para bugs, Qwen3 para código, Ollama para ediciones |

---

## Estrategia de prompts eficientes

### Principio: contexto mínimo suficiente

```
❌ INEFICIENTE — incluir todo el proyecto:
"Aquí están los 13 módulos completos. Arregla el bug en línea 47 de Mod_WhatsApp."

✅ EFICIENTE — solo lo necesario:
"En Mod_WhatsApp.bas, la función URLEncodeUTF8 no está convirtiendo ñ correctamente.
Aquí está solo esa función: [50 líneas]. El error es que devuelve %C3%B1 cuando
debería devolver %C3%B1 — espera, déjame revisar la tabla primero."
```

### Dividir tareas grandes para Boomerang

```
❌ UN SOLO PROMPT grande → Gemini se agota en la primera llamada:
"Genera el sistema completo de importación con todos los módulos"

✅ DIVIDIR en subtareas → cada una usa el modelo correcto:
Subtarea 1 → Gemini: "Diseña la arquitectura del módulo de importación"
Subtarea 2 → Qwen3-coder: "Genera Mod_ImportarArchivos.bas"
Subtarea 3 → Qwen3-coder: "Genera Hoja_REGISTROS.bas"
Subtarea 4 → Groq: "Revisa los 2 módulos y encuentra bugs de lógica"
```

### Template de prompt eficiente para BajaTax

```
Contexto: [1-2 líneas sobre qué módulo y qué hace]
Problema específico: [descripción exacta del error o qué falta]
Código relevante: [solo la función/Sub afectada, no el módulo completo]
Reglas críticas: [solo las que aplican — ej: "@TODAY() en Mac, PathSeparator"]
Esperado: [qué debe hacer el código corregido]
```

---

## Rotación diaria recomendada

```
Mañana (límites frescos):
  → Gemini para arquitectura y revisión multi-archivo
  → Groq para bugs que aparecieron el día anterior

Tarde (Gemini posiblemente agotado):
  → Qwen3-coder para generar módulos nuevos
  → Groq para debugging

Cuando todo se agota:
  → Ollama para ediciones menores y documentación
  → Guardar las tasks complejas para el día siguiente

Nota: Los límites de Gemini AI Studio free se reinician a medianoche PST.
```
