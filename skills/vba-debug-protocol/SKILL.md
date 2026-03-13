---
name: vba-debug-protocol
description: >
  Usar esta skill cuando haya un error, bug, o comportamiento inesperado en los
  módulos VBA de BajaTax. Activar cuando el usuario diga "da error", "no funciona",
  "se cuelga Excel", "dice error [número]", "el macro no termina", "se cierra Excel",
  "no importa los datos", "no manda el WhatsApp", "no genera el PDF", o cuando pegue
  un mensaje de error de VBA. También activar cuando Excel queda en estado raro
  después de una macro, cuando el código falla solo a veces, o cuando el usuario
  quiera revisar código antes de usarlo. Esta skill diagnostica PRIMERO y propone
  solución DESPUÉS. Nunca proponer un fix sin identificar causa raíz — los fixes
  sin diagnóstico introducen nuevos bugs.
---

# VBA Debug Protocol — BajaTax

## Ley de hierro

```
SIN INVESTIGACIÓN DE CAUSA RAÍZ → SIN FIX PROPUESTO
```

El orden siempre es: **Clasificar → Localizar → Entender → Corregir → Verificar**.
Si ya intentaste 3+ fixes sin éxito, el problema es arquitectural — no intentes un 4to fix.

---

## Fase 1: Clasificar el tipo de error

### Compile Error
Aparece ANTES de ejecutar. VBA resalta línea en azul al abrir módulo.
Causas: variable no declarada, paréntesis sin cerrar, referencia a librería faltante.
**Acción**: resolver línea resaltada, verificar Tools → References.

### Runtime Error — diagnosticar por número

| Error | Significado | Causa típica en BajaTax |
|---|---|---|
| **1004** | Operation failed | Nombre hoja incorrecto, rango no existe, hoja protegida |
| **91** | Object not set | `Set` faltante, variable = Nothing |
| **13** | Type mismatch | Texto donde espera número/fecha |
| **438** | Property not supported | Tools → References faltante |
| **9** | Subscript out of range | Nombre hoja incorrecto, array fuera de límites |
| **76** | Path not found | Carpeta no existe — usar `Util_EnsureFolderExists` |
| **1001** | File not found | Archivo no existe — verificar con `Dir()` antes |
| **5** | Invalid call | Argumento fuera de rango |

### Logic Error (el más difícil)
No genera error visible — código corre pero produce resultado incorrecto.
Señales: datos duplicados, celdas vacías, importación parcial, números incorrectos.
Requiere `MsgBox` o `Debug.Print` para rastrear valores en runtime.

### Silent Error (el más peligroso)
`On Error Resume Next` oculta todo. **Primera acción siempre**: desactivarlo temporalmente.

---

## Fase 2: Instrumentar antes de proponer fix

Para bugs en sistemas multi-componente (VBA → Python → archivo → Excel), agregar
diagnóstico en CADA frontera antes de proponer cualquier solución:

```vba
' Verificar en cada punto de la cadena:

' 1. ¿La variable tiene el valor correcto al entrar?
MsgBox "ENTRADA — RFC: [" & rfc & "] | Fila: " & fila & " | Col: " & colRFC

' 2. ¿El archivo existe antes de abrirlo?
If Dir(rutaArchivo) = "" Then
    MsgBox "DIAGNÓSTICO: Archivo no existe en: " & rutaArchivo
    Exit Sub
End If

' 3. ¿La hoja tiene el nombre correcto?
Dim sh As Worksheet
For Each sh In ThisWorkbook.Sheets
    Debug.Print "Hoja: [" & sh.Name & "]"  ' Ver en Ctrl+G
Next sh

' 4. ¿El JSON se escribió antes de llamar a Python?
If Dir(jsonPath) = "" Then
    MsgBox "DIAGNÓSTICO: JSON temporal no se creó en: " & jsonPath
    Exit Sub
End If

' 5. ¿Python generó el PDF?
Application.Wait Now() + TimeValue("00:00:05")  ' Dar tiempo a Python
If Dir(pdfPath) = "" Then
    MsgBox "DIAGNÓSTICO: PDF no apareció después de 5s en: " & pdfPath
End If
```

**Ejecutar UNA VEZ con diagnóstico → analizar dónde rompe → corregir ese punto.**

---

## Fase 3: Localizar línea exacta

```vba
' Técnica 1 — MsgBox temporal:
MsgBox "Punto A: RFC=[" & rfcValue & "] Monto=[" & montoValue & "]"

' Técnica 2 — Debug.Print (Ventana Inmediata: Ctrl+G en editor VBA):
Debug.Print Now() & " | Fila " & i & " | RFC: " & ws.Cells(i, colRFC).Value

' Técnica 3 — Verificar objeto antes de usar:
If ws Is Nothing Then
    MsgBox "ERROR: hoja no asignada — el nombre puede estar mal escrito"
    Exit Sub
End If

' Técnica 4 — F9 breakpoint + F8 paso a paso en el editor VBA
```

---

## Checklist por módulo

### 02_Mod_ImportarArchivos
- [ ] Archivo externo CERRADO antes de `Workbooks.Open`?
- [ ] Ruta usa `Application.PathSeparator`?
- [ ] `Dir(filePath)` verificado antes de abrir?
- [ ] `ErrorHandler` cierra `wbExterno.Close False`?
- [ ] Headers buscados en primeras 20 filas (no solo fila 1)?
- [ ] `EnableEvents = False` al escribir a REGISTROS?

### 03_Mod_WhatsApp
- [ ] `CONFIGURACION.B2` validado antes de enviar?
- [ ] Número tiene formato `52XXXXXXXXXX` (12 dígitos, sin +)?
- [ ] `URLEncodeUTF8` cubre á,é,í,ó,ú,ñ,espacios,`Chr(10)`?
- [ ] Consolidación agrupa por teléfono Col.M, no por fila?
- [ ] Pausa 8-15s entre envíos presente?

### 04_Mod_PDF
- [ ] `pdf_server.py` existe en `src/python/`?
- [ ] Python 3 accesible desde Shell en esa máquina?
- [ ] JSON temporal escrito ANTES de llamar Python?
- [ ] Se espera y verifica que PDF apareció en OUTPUT/?
- [ ] JSON temporal limpiado después?

### 11_Hoja_REGISTROS
- [ ] `BeforeDoubleClick` tiene `Cancel = True`?
- [ ] Trackers marcados en orden L → M → N?
- [ ] Deduplicación usa clave compuesta RFC+Concepto+Fecha+Monto?
- [ ] Se pregunta antes de actualizar DIRECTORIO?

---

## Defense-in-depth — validación en capas

Cuando encuentres un bug causado por datos inválidos, no basta con un solo `If`.
Agregar validación en CADA capa que ese dato atraviesa:

```vba
' Capa 1 — Entrada: rechazar datos obviamente inválidos
If Len(Trim(rfc)) < 12 Or Len(Trim(rfc)) > 13 Then
    MsgBox "RFC inválido: " & rfc, vbCritical: Exit Sub
End If

' Capa 2 — Lógica de negocio: verificar que tiene sentido para esta operación
If Not Util_ValidateRFC(rfc) Then
    LogEvento "RFC_INVALIDO", rfc, "SALTADO"
    GoTo SiguienteFila
End If

' Capa 3 — Guard de ambiente: prevenir operaciones peligrosas
' Ej: si modo = PRODUCCIÓN pero teléfono tiene solo 5 dígitos → bloquear
If modo = "PRODUCCIÓN" And Len(Util_NormalizePhone(tel)) <> 10 Then
    LogEvento "TEL_INVALIDO", cliente & " | " & tel, "BLOQUEADO"
    GoTo SiguienteFila
End If

' Capa 4 — Debug: registrar contexto para investigación futura
LogEvento "PROCESANDO", cliente & " | RFC:" & rfc & " | Tel:" & tel, "OK"
```

---

## Fixes más comunes

### Error 91 — Object not set
```vba
' ❌ ws no fue asignada:
Dim ws As Worksheet
ws.Cells(1,1).Value = "dato"  ' Error 91

' ✅ Set es obligatorio para todos los objetos:
Set ws = ThisWorkbook.Sheets("OPERACIONES")
If ws Is Nothing Then MsgBox "Hoja OPERACIONES no encontrada": Exit Sub
```

### Error 1004 — Nombre de hoja
```vba
' Nombres exactos (mayúsculas y espacios importan):
' "OPERACIONES", "REGISTROS", "DIRECTORIO", "CONFIGURACION",
' "LOG ENVIOS", "BUSCADOR CLIENTE", "REPORTES CXC", "Soportes"

' Debug — listar hojas reales:
For Each sh In ThisWorkbook.Sheets
    Debug.Print "[" & sh.Name & "]"  ' Ctrl+G para ver
Next sh
```

### Excel congelado después de error
```vba
' Ejecutar en Ventana Inmediata (Ctrl+G):
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.StatusBar = False
```

### 3+ fixes fallidos — señal de problema arquitectural
Si ya intentaste 3 correcciones diferentes y el bug persiste:
- **PARAR** — no intentar un 4to fix
- El patrón indica acoplamiento o diseño incorrecto
- Discutir con el usuario si hay que refactorizar el módulo

---

## Racionalizaciones peligrosas

| Excusa | Realidad |
|---|---|
| "Quick fix por ahora, investigar después" | El fix incorrecto introduce nuevos bugs |
| "Solo probar este cambio a ver si funciona" | Si no entiendes el problema, el cambio es una apuesta |
| "Ya probé manualmente, funciona" | Las pruebas manuales no son reproducibles |
| "Es probablemente X" | Probablemente ≠ certeza. Verificar antes de cambiar |
| "Un fix más" (después de 2+ fallidos) | 3+ fixes fallidos = problema arquitectural, no de líneas |

---

## Verificación post-fix — no declarar victoria sin esto

1. El caso que fallaba ahora funciona
2. Los casos que funcionaban siguen funcionando
3. `ErrorHandler` presente en el Sub modificado
4. No se introdujo `On Error Resume Next`
5. `Application.EnableEvents` y `ScreenUpdating` se restauran en todos los paths
6. El fix corrige la causa raíz, no solo el síntoma

> Para reglas Mac/Windows: skill **cross-platform**
> Para columnas y hojas: sección "Mapa de Columnas" en **GEMINI.md**
> Para loops lentos: skill **excel-performance**
