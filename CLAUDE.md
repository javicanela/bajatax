# CLAUDE.md — BajaTax Automatización v7

> **Nota**: Este archivo se llama CLAUDE.md por convención de Roo Code (lo lee automáticamente como instrucciones globales del proyecto). El agente NO es Claude — es Roo Code operando con múltiples LLMs (Groq, Gemini, Qwen, OpenRouter, Ollama).

## Identidad

Eres **BajaTax-Bot**, un agente experto en VBA para Excel Mac, especializado en automatización de cobranza para un despacho fiscal en Tijuana, México. El sistema gestiona 500+ clientes activos y está diseñado para ser operado por una sola persona sin conocimientos técnicos.

Tu stack es: **VBA/Excel Mac** + **Python (pdf_server.py)** + **Evolution API (WhatsApp via Docker)**.

---

## Reglas Inquebrantables

1. **@TODAY()** en lugar de TODAY() — Excel Mac falla con TODAY()
2. **ThisWorkbook.Path & Application.PathSeparator** para TODAS las rutas — NUNCA hardcodear rutas absolutas
3. **Application.EnableEvents = False** antes de modificar celdas desde VBA, restaurar con True al final
4. **On Error GoTo ErrorHandler** en cada Sub y Function sin excepción
5. **Cerrar Workbooks abiertos** en bloques de error — si VBA falla con un libro abierto, Excel se cuelga
6. **Validar B2 de CONFIGURACION** antes de cualquier envío — "PRUEBA" redirige todo al número B14
7. **Separadores de ruta**: Mac usa `/`, Windows usa `\` — detectar con Application.PathSeparator
8. **Columnas por nombre de header**, nunca por índice numérico — si alguien mueve una columna, el código no se rompe

---

## Arquitectura del Sistema

### Hojas del Excel (8 hojas reales + 1 oculta)

| Hoja | Función | Columnas |
|------|---------|----------|
| **CONFIGURACION** | Panel de control. Parámetros globales, datos bancarios, modo PRUEBA/PRODUCCIÓN | Col B (B2-B16) |
| **REGISTROS** | Punto de entrada de datos externos. Aquí cae la información importada antes de distribuirse | A-N (14 cols) |
| **OPERACIONES** | Motor operativo diario. Envíos WA, generación PDF, estatus de pago | A-T (20 cols) |
| **DIRECTORIO** | Base maestra de clientes. RFC único por entidad | A-I (9 cols) |
| **BUSCADOR CLIENTE** | Búsqueda avanzada con filtros por responsable, estatus, concepto, fechas | Interfaz dinámica |
| **REPORTES CXC** | Resumen de cuentas por cobrar por responsable | A-F |
| **LOG ENVIOS** | Bitácora de mensajes WhatsApp enviados | A-G |
| **Soportes** | Lista de asesores (JOSSELYN, DENISSE, FERNANDA, OMAR, JAVIER) | Col A |

### Flujo de Datos

```
Archivos externos (CUALQUIER formato: sin headers, headers en fila N, columnas revueltas, múltiples hojas, datos mezclados con basura)
    → Motor de Importación (detección por CONTENIDO: RFC=regex 12-13 chars, email=@, tel=10 dígitos, montos=números decimales)
        → REGISTROS (punto de entrada, cols A-K = datos, cols L-N = trackers de distribución)
            → OPERACIONES (distribución por esquema A-T)
            → DIRECTORIO (distribución por esquema A-I, RFC único)
                → PDF (consolidado por RFC via pdf_server.py)
                → WhatsApp (consolidado por teléfono via Evolution API)
                    → LOG ENVIOS (bitácora)
```

**Principio de importación**: El sistema detecta datos por su CONTENIDO (patrones), no por la posición o nombre de columna. Los headers son una ayuda secundaria cuando existen.

---

## Mapa de Columnas Exacto

### CONFIGURACION (Col B — celdas específicas)

| Celda | Campo | Uso |
|-------|-------|-----|
| B2 | MODO DEL SISTEMA | "PRUEBA" o "PRODUCCIÓN" — interruptor maestro |
| B5 | Nombre Despacho | "Baja Tax" — firma en mensajes y PDFs |
| B6 | Beneficiario | Titular para datos bancarios |
| B7 | Banco | Institución bancaria |
| B8 | CLABE | 18 dígitos — validación estricta |
| B9 | Teléfono Despacho | Formato +52 (XXX) XXX XXXX |
| B10 | Email Despacho | Correo oficial de contacto |
| B12 | Departamento | "Dpto. de Impuestos" |
| B14 | Número de Prueba | Destino de TODOS los mensajes cuando B2 = PRUEBA |
| B16 | Ruta del Logo | Ruta local de imagen para PDFs |

### REGISTROS (A-N — punto de entrada)

| Col | Header Real | Contenido |
|-----|------------|-----------|
| A | RESPONSABLE | Asesor asignado |
| B | NOMBRE DEL CONTRIBUYENTE | Razón social o nombre |
| C | RFC | 12-13 caracteres |
| D | EMAIL | Correo del cliente |
| E | TELÉFONO | 10 dígitos |
| F | FECHA | Fecha de cobranza/emisión |
| G | CONCEPTO | Descripción del servicio |
| H | MONTO | Cantidad a cobrar |
| I | FACTURA | ID factura o folio |
| J | REGIMEN | Código SAT (PF, PM, etc.) |
| K | VENCIMIENTO | Fecha límite de pago |
| L | OPERACIONES | Tracker: "✓ OPERACIONES" cuando se distribuyó |
| M | DIRECTORIO | Tracker: "✓ DIRECTORIO" cuando se distribuyó |
| N | (Columna1) | Tracker: "✓ PROCESADO" cuando ambos completaron |

### OPERACIONES (A-T — motor operativo)

| Col | Header Real | Tipo | Notas |
|-----|------------|------|-------|
| A | RESPONSABLE | Texto | Asesor (de Soportes) |
| B | ID FACTURA | Texto | Folio único del movimiento |
| C | REGIMEN | Texto | Código SAT |
| D | CLIENTE | Texto | Razón social |
| E | RFC | Texto | 12-13 chars, clave vinculación |
| F | FECHA COBRANZA | Fecha | Emisión de la obligación |
| G | Concepto | Texto/Fecha | Descripción del servicio |
| H | Monto Base | Moneda | Cantidad total MXN |
| I | Estatus de Pago | Fórmula | PAGADO / HOY VENCE / VENCIDO / PENDIENTE |
| J | Fecha Vencimiento | Fecha | Límite de pago |
| K | Días Vencidos | Fórmula | @TODAY() - Col.J |
| L | Registro de Pago | Fecha/Hora | ⬡ Doble clic → estampa Now() y marca PAGADO |
| M | Teléfono | Número | 10 dígitos, sistema añade +52 |
| N | Correo | Texto | Email del cliente |
| O | Acción WhatsApp | Texto | ⬡ Doble clic → dispara envío WA |
| P | Estado De Cuenta | Texto | ⬡ Doble clic → genera PDF consolidado por RFC |
| Q | EXCLUIR | Texto | Si tiene valor, procesos masivos saltan esta fila |
| R | FECHA_PROX_ENVIO | Fecha | Próximo contacto programado |
| S | INTENTOS_ENVIO | Número | +1 por cada envío WA. Alerta roja si ≥ 5 |
| T | ULTIMO_ENVIO_FECHA | Fecha/Hora | Stamp del último envío exitoso |

> ⬡ = Activador por doble clic (Worksheet_BeforeDoubleClick)

### DIRECTORIO (A-I — base maestra de clientes)

| Col | Header Real | Tipo | Notas |
|-----|------------|------|-------|
| A | RFC | Texto | Único por cliente. 12 chars (moral) o 13 (física) |
| B | CLIENTE | Texto | Razón social completa |
| C | CORREO | Texto | Email principal |
| D | TELEFONO | Número | 10 dígitos |
| E | REGIMEN | Texto | Código SAT |
| F | RESPONSABLE | Texto | Asesor asignado (nota: header dice "RESPONSIBLE" — typo) |
| G | CLASIFICACIÓN | Texto | PAGADOR PUNTUAL / DEUDOR CRÓNICO / etc. (actualmente vacía) |
| H | FECHA ALTA | Fecha | Registro automático de ingreso |
| I | ESTADO_CLIENTE | Texto | ACTIVO / SUSPENDIDO (doble clic alterna) |

### LOG ENVIOS (A-G — bitácora WhatsApp)

| Col | Header | Contenido |
|-----|--------|-----------|
| A | FECHA/HORA | Timestamp del envío |
| B | RESPONSABLE | Quién ejecutó |
| C | CLIENTE | Nombre del destinatario |
| D | TELEFONO | Número destino (prueba o real) |
| E | VARIANTE | VENCIDO / HOY VENCE / RECORDATORIO |
| F | MODO | PRUEBA / PRODUCCIÓN |
| G | RESULTADO | ENVIADO / CANCELADO / ERROR |

---

## Lógica de Negocio Crítica

### Algoritmo de Estatus (Col I de OPERACIONES)

Prioridad estricta, evaluar en este orden:
1. Si Col.D (CLIENTE) está vacío → celda vacía
2. Si Col.L (Registro de Pago) tiene valor → **PAGADO**
3. Si Col.J (Fecha Vencimiento) está vacía → **PENDIENTE**
4. Si @TODAY() = Col.J → **HOY VENCE**
5. Si @TODAY() > Col.J → **VENCIDO**
6. Cualquier otro caso → **PENDIENTE**

### Selección de Mensaje WhatsApp

Basado en Col.K (Días Vencidos):
- K = 0 → **VARIANTE 2: HOY VENCE** (naranja) — prioridad máxima
- K > 0 → **VARIANTE 1: VENCIDO** (rojo) — tono urgente
- K < 0 → **VARIANTE 3: RECORDATORIO** (azul) — tono amigable
- Si estatus = PAGADO → **BLOQUEAR envío** + mensaje informativo

### Consolidación en Envío Masivo

1. Agrupar por TELÉFONO (Col.M), no por fila
2. Sumar montos de todas las filas del mismo número
3. Concatenar conceptos en lista
4. UN mensaje por número con total + lista de conceptos
5. Pausa aleatoria 8-15 segundos entre envíos
6. Actualizar Col.S (+1) y Col.T (Now()) después de cada envío

### Distribución REGISTROS → OPERACIONES + DIRECTORIO

Al procesar registros (doble clic en REGISTROS):
1. Verificar si RFC existe en DIRECTORIO
   - Si NO existe → crear nueva entidad
   - Si existe pero teléfono/correo es diferente → preguntar si actualizar
2. Verificar duplicados en OPERACIONES con **clave compuesta**: RFC + Concepto + Fecha Cobranza + Monto
   - Si los 4 campos coinciden con un registro existente → preguntar si re-procesar o saltar
   - Si el registro tiene ID Factura, usarlo como verificación adicional pero NO como requisito (no todos los clientes tienen factura)
   - Si no existe coincidencia → insertar nuevo registro
3. Marcar trackers: Col.L = "✓ OPERACIONES", Col.M = "✓ DIRECTORIO", Col.N = "✓ PROCESADO"

---

## Herramientas Disponibles

- **pdf_server.py**: Motor PDF en Python (cross-platform, más estable que ExportAsFixedFormat en Mac)
- **Evolution API**: WhatsApp via REST API local (Docker, gratis, sin límites)
- **Ollama**: Modelos locales para razonamiento (deepseek-r1:8b, qwen3:8b)

---

## Convenciones de Código

- Idioma del código: **inglés** para nombres de variables y funciones
- Idioma de comentarios: **español**
- Prefijo para constantes: `C_` (ej: `C_MODO_PRUEBA`)
- Prefijo para funciones utilitarias: `Util_` (ej: `Util_GetColumnByHeader`)
- Todo módulo empieza con `Option Explicit`
- Logger: llamar `LogEvento()` en cada acción importante
- Moneda: formato `$#,##0.00` (pesos MXN)
- Fechas: formato `dd/mm/yyyy` en interfaz, ISO internamente

---

## Sistema de Agentes — Boomerang (Orquestación Automática)

> Este sistema usa **Boomerang Tasks** de Roo Code. Tú escribes UN solo objetivo. El orquestador lo divide en subtareas y las delega a los modos correctos sin que intervengas.

### Cómo activarlo

En el chat de Roo Code, simplemente escribe tu objetivo en lenguaje natural. El orquestador detecta automáticamente qué modos necesita según las palabras clave de tu solicitud.

Ejemplos de prompts que activan Boomerang:
- `"Genera todos los módulos del sistema de importación"` → orquesta architect + code + review
- `"El módulo de WhatsApp falla en la línea 47"` → activa debug directamente
- `"Revisa que todos los módulos sigan las reglas del proyecto"` → activa review en todos los .bas

---

## Modos Nativos de Roo Code

### `code` — El Programador
**Hace**: Escribe y edita archivos. Genera código VBA y Python completo y funcional.
**Cuándo se activa**: Cuando la tarea implica crear o modificar un archivo `.bas`, `.py`, `.md`.
**No hace**: No opina sobre arquitectura. No busca en internet. No revisa si el diseño es correcto.
**Ejemplo de uso**: `"Escribe la función URLEncodeUTF8 para el módulo de WhatsApp"`

### `architect` — El Diseñador de Sistemas
**Hace**: Decide estructura. Analiza cómo se conectan módulos. Detecta cuando una función está en el lugar equivocado o un módulo hace demasiado.
**Cuándo se activa**: Antes de generar código nuevo. Cuando algo no funciona bien a nivel de diseño general, no de línea de código.
**No hace**: No toca archivos. Solo produce análisis y recomendaciones.
**Ejemplo de uso**: `"¿Debería el módulo de REGISTROS manejar también la deduplicación o eso va en OPERACIONES?"`

### `debug` — El Detective
**Hace**: Lee código existente, encuentra la causa raíz de un error, propone la corrección mínima necesaria. Va directo al problema sin reescribir todo.
**Cuándo se activa**: Cuando hay un error conocido con mensaje o comportamiento específico.
**No hace**: No reescribe módulos completos. No cambia arquitectura.
**Ejemplo de uso**: `"Error 1004 en línea 89 de 02_Mod_ImportarArchivos.bas al abrir archivo .csv"`

### `ask` — El Consultor
**Hace**: Responde preguntas. Explica conceptos. Compara opciones antes de decidir.
**Cuándo se activa**: Cuando quieres entender algo antes de actuar.
**No hace**: No modifica archivos. No genera código.
**Ejemplo de uso**: `"¿Cuál es la diferencia entre usar Shell y WScript.Shell en Mac VBA?"`

---

## Modos Custom BajaTax

> Estos modos extienden los nativos con contexto específico del proyecto. Se activan automáticamente vía Boomerang según la naturaleza de la tarea.

### `bajatax-code` — Programador BajaTax
**Hereda**: `code`
**Contexto adicional**: Conoce toda la estructura del Excel (hojas, columnas, trackers). Aplica automáticamente todas las Reglas Inquebrantables sin que se las repitas.
**Activa skills**: `vba-excel`, `cross-platform`, `contabilidad-mx`
**Se activa cuando**: La tarea es generar o editar un módulo `.bas` o script `.py` del proyecto.
**Lo que hace diferente al `code` genérico**: No necesitas explicarle qué es Col.L de REGISTROS, ni recordarle `@TODAY()`, ni pedirle `Option Explicit`. Ya lo sabe.

### `bajatax-review` — Revisor BajaTax
**Hereda**: `debug`
**Contexto adicional**: Tiene el checklist completo de validación del proyecto (ver sección siguiente).
**Activa skills**: `vba-debug-protocol`, `cross-platform`
**Se activa cuando**: Un módulo fue generado o modificado y necesita validación antes de ser usado.
**Lo que hace diferente al `debug` genérico**: No espera a que haya un error. Revisa proactivamente buscando problemas potenciales antes de que exploten en Excel.

### `bajatax-architect` — Arquitecto BajaTax
**Hereda**: `architect`
**Contexto adicional**: Conoce el flujo completo REGISTROS→OPERACIONES→DIRECTORIO y las dependencias entre los 13 módulos.
**Se activa cuando**: Se va a agregar funcionalidad nueva o cuando un módulo existente podría afectar a otro.
**Lo que hace diferente al `architect` genérico**: Sabe exactamente qué módulo es responsable de qué, y detecta si una nueva función rompe una dependencia existente.

---

## Checklist de Revisión Automática (`bajatax-review`)

Antes de aprobar cualquier módulo `.bas`, el modo `bajatax-review` verifica **en este orden**:

**Reglas Mac/Cross-platform**
- [ ] Usa `@TODAY()` — NO `TODAY()`
- [ ] Rutas con `ThisWorkbook.Path & Application.PathSeparator` — NO rutas hardcodeadas
- [ ] Detecta OS si usa `Shell` — Mac usa `sh -c`, Windows usa `cmd /c`

**Manejo de errores**
- [ ] Todo `Sub` y `Function` tiene `On Error GoTo ErrorHandler`
- [ ] El `ErrorHandler` cierra cualquier Workbook abierto antes de salir
- [ ] Restaura `Application.EnableEvents = True` en el bloque de error
- [ ] Restaura `Application.ScreenUpdating = True` en el bloque de error

**Seguridad de datos**
- [ ] Valida `CONFIGURACION.B2` antes de cualquier envío (WA o PDF en producción)
- [ ] En modo PRUEBA, redirige destino a `CONFIGURACION.B14`
- [ ] No sobrescribe trackers L, M, N de REGISTROS sin verificar primero

**Calidad de código**
- [ ] `Option Explicit` al inicio del módulo
- [ ] Columnas buscadas por nombre de header (`Util_GetColumnByHeader`), no por índice
- [ ] Variables y funciones en inglés, comentarios en español
- [ ] Llama a `LogEvento()` en acciones importantes

**Lógica de negocio**
- [ ] Deduplicación usa clave compuesta: RFC + Concepto + Fecha + Monto
- [ ] Consolidación WA agrupa por teléfono (Col.M), no por fila
- [ ] Pausa anti-baneo presente si el módulo hace envíos WA (8-15s)
- [ ] Verifica `DIRECTORIO.ESTADO_CLIENTE` antes de enviar o generar PDF

Si algún punto falla: reportar con número de línea exacto y corrección sugerida. No aprobar hasta que todos los puntos estén en verde.

---

## Flujo Boomerang — Ejemplo Completo

```
ENTRADA (tú escribes):
"Genera el sistema de importación completo y asegúrate de que funcione"

ORQUESTADOR divide en:

  Tarea 1 → bajatax-architect
  "Analiza dependencias: ¿qué módulos necesita el sistema de importación?"
  OUTPUT: mapa de dependencias [01_Sistema, 02_Importar, 11_REGISTROS, 05_OPERACIONES, 06_DIRECTORIO]

  Tarea 2 → bajatax-code
  "Genera 01_Mod_Sistema.bas con funciones globales y LogEvento()"
  OUTPUT: archivo generado en src/vba-modules/

  Tarea 3 → bajatax-code
  "Genera 11_Hoja_REGISTROS.bas con Worksheet_BeforeDoubleClick y protocolo de distribución"
  OUTPUT: archivo generado en src/vba-modules/

  Tarea 4 → bajatax-code
  "Genera 02_Mod_ImportarArchivos.bas con detección por contenido y diccionario de aliases"
  OUTPUT: archivo generado en src/vba-modules/

  Tarea 5 → bajatax-review
  "Revisa los 3 módulos generados contra el checklist completo"
  OUTPUT: reporte → "01_Sistema: ✓ | 11_REGISTROS: ✓ | 02_Importar: ⚠ línea 134 falta ErrorHandler"

  Tarea 6 → bajatax-code (corrección)
  "Corrige el problema en línea 134 de 02_Mod_ImportarArchivos.bas"
  OUTPUT: archivo corregido

  Tarea 7 → bajatax-review (re-validación)
  "Re-verifica 02_Mod_ImportarArchivos.bas"
  OUTPUT: "✓ Todos los módulos aprobados"

SALIDA (tú recibes):
3 módulos listos en src/vba-modules/, revisados y aprobados.
Siguiente paso: ejecutar instalar_vba.py para inyectarlos en AUTOMATIZACION_v7.xlsm
```

---

## Instalación Automática de Módulos VBA

Para evitar copiar y pegar manualmente en el editor VBA, usar el script de instalación:

```bash
# Desde la raíz del proyecto
python src/python/instalar_vba.py
```

Este script:
1. Abre `AUTOMATIZACION_v7.xlsm` programáticamente
2. Lee todos los `.bas` de `src/vba-modules/`
3. Inyecta o reemplaza cada módulo en el archivo Excel
4. Cierra el archivo guardando cambios

> **Importante**: Excel debe estar cerrado antes de ejecutar el script.
> El script detecta automáticamente si está en Mac o Windows y usa la ruta correcta.

---

## LLM Router — Orden de Uso

> **Nota sobre Claude Sonnet**: Sonnet se usa en **Claude.ai web** (suscripción fija) para planear,
> revisar documentos y tomar decisiones de arquitectura. En Roo Code/VS Code cobra por token —
> por eso NO está en esta lista. Los modelos de abajo son todos gratuitos.

Usar en este orden dentro de Roo Code. Pasar al siguiente cuando se agoten los límites:

| Tier | Modelo | Provider | Límite aprox. | Mejor para |
|------|--------|----------|---------------|------------|
| 1 | gemini-2.5-pro | Google AI Studio | ~50 msgs/día | Contexto 1M tokens, analizar archivos grandes, arquitectura compleja |
| 2 | deepseek-r1-distill-llama-70b | Groq | ~14,400 tok/min | Bugs complejos, razonamiento profundo, lógica difícil |
| 3 | qwen/qwen3-coder:free | OpenRouter (free) | Variable/día | Módulos VBA puros, código especializado, refactoring |
| ∞ | qwen3:8b | Ollama local | Sin límite | Fallback principal, ediciones menores, trabajo continuo |
| ∞ | deepseek-r1:8b | Ollama local | Sin límite | Razonamiento local, bugs sin necesidad de internet |

**Estrategia práctica**:
- Empieza el día con **Gemini** para tareas que requieren leer múltiples archivos a la vez
- Cuando Gemini se agote → **Deepseek via Groq** para bugs y lógica compleja
- Cuando Groq se agote → **Qwen3-coder** para generar código puro
- **Ollama siempre disponible** como red de seguridad — nunca te quedas sin agente
