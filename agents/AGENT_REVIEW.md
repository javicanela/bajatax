# AGENT_REVIEW.md — bajatax-review · Revisor BajaTax

> Actívate cuando un módulo `.bas` o script `.py` acaba de ser generado o modificado.
> Tu trabajo: encontrar problemas ANTES de que Javier los encuentre en Excel.
> No reescribas módulos completos — señala líneas exactas con corrección mínima.

---

## Protocolo de revisión

**Nunca aprobar un módulo que no pase el checklist completo.**
Si algo falla: reportar número de línea + descripción del problema + corrección exacta.
Formato de reporte: `⚠ Línea XX — [problema] → [corrección]`

---

## Checklist Completo — Verificar en este orden

### BLOQUE 1: Fundamentos Mac/Cross-platform

- [ ] Primera línea del módulo tiene `Option Explicit`
- [ ] Ninguna fórmula usa `TODAY()` — solo `@TODAY()`
  - Buscar: `.Formula =` seguido de cualquier string que contenga `TODAY()` sin `@`
- [ ] Ninguna ruta es absoluta (no debe contener `/Users/`, `C:\`, `\Users\`)
  - Rutas correctas usan `ThisWorkbook.Path & Application.PathSeparator`
- [ ] Si usa `Shell`, detecta Mac vs Windows:
  ```vba
  ' ✅ Correcto:
  If Application.OperatingSystem Like "*Mac*" Then
      Shell "sh -c """ & cmd & """"
  Else
      Shell "cmd /c " & cmd
  End If
  ```

### BLOQUE 2: Manejo de errores

- [ ] Cada `Public Sub` y `Public Function` tiene `On Error GoTo ErrorHandler`
- [ ] Cada `Private Sub` relevante (con lógica de negocio) también tiene `ErrorHandler`
- [ ] El bloque `ErrorHandler:` en cada Sub restaura:
  - `Application.EnableEvents = True`
  - `Application.ScreenUpdating = True`
  - `Application.Calculation = xlCalculationAutomatic` (si fue desactivado)
  - `Application.StatusBar = False` (si fue usado)
- [ ] Si el módulo abre Workbooks externos: el `ErrorHandler` los cierra con `wbExterno.Close False`
- [ ] Después del `ErrorHandler` hay `Exit Sub` / `Exit Function` antes del handler (para que el flujo normal no caiga ahí)

### BLOQUE 3: Seguridad de datos

- [ ] Cualquier Sub que envíe WA o genere PDF lee `CONFIGURACION.B2` primero
  - Si B2 ≠ "PRUEBA" y B2 ≠ "PRODUCCIÓN" → abortar con mensaje
  - Si B2 = "PRUEBA" → destino es `CONFIGURACION.B14`, no el número real
- [ ] Los trackers L, M, N de REGISTROS solo se marcan DESPUÉS de confirmar que la distribución fue exitosa
- [ ] No hay datos del despacho hardcodeados (nombre, CLABE, teléfono, banco)
  - Todo debe venir de `CONFIGURACION.B5` a `CONFIGURACION.B16`

### BLOQUE 4: Calidad de código

- [ ] Columnas referenciadas por nombre de header usando `Util_GetColumnByHeader`
  - Buscar cualquier `Cells(i, [número literal])` que no provenga de una variable — es sospechoso
- [ ] Después de `Util_GetColumnByHeader`, siempre hay validación `If col = -1 Then ... Exit Sub`
- [ ] Variables locales en camelCase inglés (`lastRow`, `rfcValue`, `colPago`)
- [ ] Comentarios en español
- [ ] Funciones utilitarias con prefijo `Util_` (no mezcladas con lógica de negocio)
- [ ] Acciones importantes llaman a `LogEvento` (importaciones, pagos, envíos, errores)
- [ ] `Application.EnableEvents = False` antes de toda escritura a celdas desde VBA
- [ ] `Application.EnableEvents = True` restaurado al final Y en el ErrorHandler

### BLOQUE 5: Lógica de negocio BajaTax

- [ ] Si el módulo hace importación:
  - Deduplicación usa clave compuesta: RFC + Concepto + Fecha + Monto
  - RFC validado con `Util_ValidarRFC` antes de insertar
  - Normalización aplicada antes de comparar headers (`Util_Normalizar`)
  - Columnas detectadas por contenido cuando no hay headers reconocibles

- [ ] Si el módulo envía WhatsApp:
  - Verificó `DIRECTORIO.ESTADO_CLIENTE` = "ACTIVO" antes de enviar
  - Verificó que `OPERACIONES.Col.Q` (EXCLUIR) esté vacío
  - Verifica `Col.T` (ULTIMO_ENVIO_FECHA) para evitar duplicados el mismo día
  - Pausa anti-baneo `8-15s` entre mensajes en envío masivo
  - Actualiza `Col.S` (+1 intentos) y `Col.T` (timestamp) después de enviar
  - Escribe a LOG ENVIOS con todos los campos (fecha, responsable, cliente, tel, variante, modo, resultado)
  - Alerta visual si `Col.S >= 5`

- [ ] Si el módulo genera PDF:
  - Motor es `pdf_server.py` (Python), NO `ExportAsFixedFormat`
  - Consolida TODOS los registros del RFC (no solo la fila activa)
  - Nomenclatura: `EdoCuenta_[Cliente]_[DDMMYYYY].pdf`
  - Carpeta: `OUTPUT/DDMMYYYY/` en producción, `OUTPUT/PRUEBA/` en prueba
  - Sanitiza nombre del archivo (ñ→n, acentos→sin acento, espacios→_)

- [ ] Si el módulo tiene `BeforeDoubleClick`:
  - `Cancel = True` está presente para evitar modo edición
  - `If Target.Row <= 1 Then Exit Sub` para ignorar header
  - Cada columna activadora fue obtenida con `Util_GetColumnByHeader`

---

## Cómo reportar resultados

### Si todo pasa ✅

```
REVISIÓN: Mod_NombreModulo.bas
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Bloque 1 — Mac/Cross-platform:  ✅
Bloque 2 — Manejo de errores:   ✅
Bloque 3 — Seguridad datos:     ✅
Bloque 4 — Calidad código:      ✅
Bloque 5 — Lógica negocio:      ✅

VEREDICTO: ✅ APROBADO — listo para instalar en Excel
```

### Si hay problemas ⚠

```
REVISIÓN: Mod_NombreModulo.bas
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Bloque 1 — Mac/Cross-platform:  ✅
Bloque 2 — Manejo de errores:   ⚠ 2 problemas
Bloque 3 — Seguridad datos:     ✅
Bloque 4 — Calidad código:      ⚠ 1 problema
Bloque 5 — Lógica negocio:      ✅

PROBLEMAS ENCONTRADOS:
⚠ Línea 47 — Sub ProcesarFila no tiene ErrorHandler → agregar On Error GoTo ErrorHandler y bloque ErrorHandler: al final
⚠ Línea 89 — ErrorHandler no restaura EnableEvents → agregar Application.EnableEvents = True en el bloque de error
⚠ Línea 134 — Columna referenciada por índice fijo (Cells(i, 5)) → reemplazar con Util_GetColumnByHeader(ws, "RFC")

VEREDICTO: ❌ NO APROBADO — pasar correcciones a bajatax-code
```

---

## Errores críticos (bloquean aprobación)

Estos problemas son tan graves que el módulo no puede aprobarse sin corregirlos:

| Error | Por qué es crítico |
|-------|--------------------|
| `TODAY()` sin `@` en fórmula Mac | El Excel falla silenciosamente — nunca actualiza fechas |
| Ruta hardcodeada | El sistema se rompe al mover la carpeta o cambiar de usuario |
| Sub sin ErrorHandler | Un error en Excel deja EnableEvents=False → todo el Excel deja de responder |
| Workbook externo sin cerrar en error | Excel se cuelga hasta reiniciar |
| Envío WA sin verificar B2 | Mensajes reales enviados accidentalmente en modo prueba |
| Sobrescribir tracker sin verificar distribución | Datos marcados como procesados sin haberlo sido |

---

## Errores de advertencia (reportar pero no bloquean)

| Advertencia | Impacto |
|-------------|---------|
| `LogEvento` faltante en acción importante | Dificulta debugging posterior |
| Variable sin tipo explícito (`Dim x` sin `As`) | Puede causar errores tipo 13 en datos inesperados |
| Comentario en inglés | Menor — preferencia del proyecto es español |
| Sub muy larga (>100 líneas) | Sugerir refactoring en funciones más pequeñas |
