---
name: vba-excel
description: >
  Usar esta skill para CUALQUIER tarea que involucre escribir, corregir, o mejorar
  código VBA para Excel en BajaTax. Activar cuando el usuario mencione macros,
  módulos .bas, Sub, Function, Worksheet, Workbook, Range, Cells, errores de VBA
  (1004, 91, 13, 438), doble clic, BeforeDoubleClick, botones en Excel, o cualquier
  automatización dentro de AUTOMATIZACION_v7.xlsm. También activar cuando el
  usuario diga "no funciona el Excel", "el botón no hace nada", "se cuelga Excel",
  o pegue código VBA en el chat. Esta skill conoce la estructura exacta del proyecto.
  Para diagnóstico de errores activos, ver vba-debug-protocol.
  Para compatibilidad Mac/Windows, ver cross-platform.
---

# VBA Excel — BajaTax

## Contexto del proyecto

Archivo: `AUTOMATIZACION_v7.xlsm` — Mac (principal) y Windows (secundario).
Módulos `.bas` en `src/vba-modules/`, inyectados con `src/python/instalar_vba.py`.
8 hojas: CONFIGURACION, REGISTROS, OPERACIONES, DIRECTORIO,
BUSCADOR CLIENTE, REPORTES CXC, LOG ENVIOS, Soportes.

---

## Reglas absolutas — sin excepción

```vba
Option Explicit  ' Inicio de CADA módulo

' Fecha en Mac — siempre @TODAY(), nunca TODAY():
ws.Range("K2").Formula = "=@TODAY()-J2"

' Rutas — nunca absolutas:
ThisWorkbook.Path & Application.PathSeparator & "OUTPUT"  ' ✅
"C:\Users\Javi\OUTPUT"                                     ' ❌

' Eventos — apagar antes de modificar celdas:
Application.EnableEvents = False
ws.Cells(fila, col).Value = valor
Application.EnableEvents = True

' Error handler — en cada Sub y Function:
On Error GoTo ErrorHandler
' ... código ...
Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If Not wbExterno Is Nothing Then wbExterno.Close False
    LogEvento "ERROR", "NombreSub", Err.Number & ": " & Err.Description
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
```

---

## Buscar columnas por nombre — nunca por índice

```vba
Function Util_GetColumnByHeader(ws As Worksheet, headerName As String) As Long
    Dim cell As Range
    For Each cell In ws.Rows(1).Cells
        If Trim(LCase(cell.Value)) = Trim(LCase(headerName)) Then
            Util_GetColumnByHeader = cell.Column
            Exit Function
        End If
        If cell.Column > 50 Then Exit For
    Next cell
    Util_GetColumnByHeader = -1
End Function

' Uso:
Dim colRFC As Long
colRFC = Util_GetColumnByHeader(wsOps, "RFC")
If colRFC = -1 Then MsgBox "Columna RFC no encontrada", vbCritical: Exit Sub
```

---

## Logger — LogEvento()

```vba
Sub LogEvento(tipo As String, origen As String, detalle As String)
    On Error Resume Next  ' Logger nunca rompe el flujo principal
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets("LOG ENVIOS")
    If wsLog Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Dim r As Long: r = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(r, 1).Value = Now()
    wsLog.Cells(r, 2).Value = tipo
    wsLog.Cells(r, 3).Value = origen
    wsLog.Cells(r, 4).Value = detalle
    Application.EnableEvents = True
    On Error GoTo 0
End Sub
```

---

## Patrones fundamentales

### Leer configuración completa
```vba
' Siempre leer de CONFIGURACION, nunca hardcodear valores del despacho
Dim wsConfig As Worksheet
Set wsConfig = ThisWorkbook.Sheets("CONFIGURACION")
Dim modo As String:      modo      = wsConfig.Range("B2").Value
Dim beneficiario As String: beneficiario = wsConfig.Range("B6").Value
Dim banco As String:     banco     = wsConfig.Range("B7").Value
Dim clabe As String:     clabe     = wsConfig.Range("B8").Value
Dim telDespacho As String: telDespacho = wsConfig.Range("B9").Value
Dim emailDespacho As String: emailDespacho = wsConfig.Range("B10").Value
Dim depto As String:     depto     = wsConfig.Range("B12").Value
Dim telPrueba As String: telPrueba = wsConfig.Range("B14").Value

' Validar modo ANTES de cualquier envío:
If modo <> "PRUEBA" And modo <> "PRODUCCIÓN" Then
    MsgBox "Configure CONFIGURACION.B2 con PRUEBA o PRODUCCIÓN", vbCritical
    Exit Sub
End If
```

### Última fila con datos
```vba
Dim wsOps As Worksheet
Set wsOps = ThisWorkbook.Sheets("OPERACIONES")
Dim lastRow As Long
lastRow = wsOps.Cells(wsOps.Rows.Count, 1).End(xlUp).Row
If lastRow < 2 Then MsgBox "No hay datos en OPERACIONES", vbInformation: Exit Sub
```

### Doble clic — patrón BajaTax
```vba
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrorHandler
    If Target.Row <= 1 Then Exit Sub  ' Ignorar header
    Application.EnableEvents = False
    
    Dim colPago As Long: colPago = Util_GetColumnByHeader(Me, "Registro de Pago")
    Dim colWA As Long:   colWA   = Util_GetColumnByHeader(Me, "Acción WhatsApp")
    Dim colPDF As Long:  colPDF  = Util_GetColumnByHeader(Me, "Estado De Cuenta")
    
    Cancel = True  ' Evitar modo edición
    
    Select Case Target.Column
        Case colPago
            Target.Value = Now()
            Target.NumberFormat = "dd/mm/yyyy hh:mm"
            LogEvento "PAGO", "Hoja_OPERACIONES", "Fila " & Target.Row & " registrada"
        Case colWA:  Call Mod_WhatsApp.EnviarMensaje(Target.Row)
        Case colPDF: Call Mod_PDF.GenerarEstadoCuenta(Target.Row)
    End Select
    
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    MsgBox "Error doble clic: " & Err.Description, vbCritical
End Sub
```

### Barra de progreso
```vba
For i = 2 To lastRow
    Application.StatusBar = "Procesando " & (i-1) & " de " & (lastRow-1) & "..."
    DoEvents
    ' ... lógica ...
Next i
Application.StatusBar = False
```

---

## Convenciones de nombres

| Elemento | Patrón | Ejemplo |
|---|---|---|
| Módulos | `Mod_` + tema | `Mod_WhatsApp`, `Mod_PDF` |
| Hojas | `Hoja_` + nombre | `Hoja_REGISTROS` |
| Variables | camelCase inglés | `lastRow`, `rfcValue` |
| Funciones utilitarias | `Util_` + verbo | `Util_GetColumnByHeader` |
| Constantes | `C_` + MAYÚSCULAS | `C_MODO_PRUEBA` |
| Comentarios | Español | `' Buscar última fila` |

---

## Errores frecuentes

| Error | Causa en BajaTax | Fix inmediato |
|---|---|---|
| 1004 | Nombre hoja incorrecto / hoja protegida | Verificar mayúsculas exactas del nombre |
| 91 | `Set` faltante en variable objeto | `Set ws = ThisWorkbook.Sheets(...)` |
| 13 | Texto donde se espera número/fecha | `CDate()`, `CDbl()`, verificar con `IsNumeric()` |
| 438 | Propiedad no existe | Tools → References en editor VBA |
| 9 | Nombre de hoja no encontrado | `On Error Resume Next` + verificar `Is Nothing` |

> Para diagnóstico profundo: skill **vba-debug-protocol**
> Para rutas y Shell: skill **cross-platform**
> Para loops lentos en archivos grandes: skill **excel-performance**
