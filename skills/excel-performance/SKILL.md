---
name: excel-performance
description: >
  Usar esta skill cuando el código VBA de BajaTax es lento, Excel se congela, o
  un loop procesa muchas filas y tarda demasiado. Activar cuando el usuario diga
  "el macro tarda mucho", "Excel se cuelga mientras corre", "el envío masivo es
  muy lento", "procesar 500 clientes tarda horas", o cuando un Sub necesite iterar
  sobre más de 100 filas de OPERACIONES o REGISTROS. También activar cuando se
  escriban loops que modifiquen celdas, lean rangos grandes, o abran archivos
  externos repetidamente. Esta skill convierte código lento en código rápido
  con cambios mínimos y predecibles.
---

# Excel Performance — BajaTax

## El trío de velocidad — siempre juntos

```vba
' Al inicio de cualquier Sub que procese muchas filas:
Application.ScreenUpdating = False    ' Excel no redibuja la pantalla
Application.Calculation = xlCalculationManual  ' No recalcula fórmulas en cada cambio
Application.EnableEvents = False      ' No dispara eventos al modificar celdas

' Al final — también en ErrorHandler:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

' Nota: ScreenUpdating y Calculation son independientes de EnableEvents.
' Los 3 deben restaurarse siempre. Si olvidas uno, Excel queda en estado raro.
```

---

## Leer todo el rango de una vez — el cambio más impactante

```vba
' ❌ LENTO — una llamada a Excel por cada celda (500 filas = 500 llamadas):
Dim i As Long
For i = 2 To lastRow
    Dim rfc As String: rfc = ws.Cells(i, colRFC).Value
    Dim monto As Double: monto = ws.Cells(i, colMonto).Value
    ' ... procesar ...
Next i

' ✅ RÁPIDO — una sola llamada, todo en memoria (500 filas = 1 llamada):
Dim datos As Variant
datos = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, colMax)).Value
' datos es un array 2D: datos(fila, columna)
Dim i As Long
For i = 1 To UBound(datos, 1)  ' Índice empieza en 1
    Dim rfc As String: rfc = datos(i, colRFC)
    Dim monto As Double
    If IsNumeric(datos(i, colMonto)) Then monto = CDbl(datos(i, colMonto))
    ' ... procesar solo con datos en memoria, sin tocar Excel ...
Next i

' Escribir resultados de vuelta en una sola operación:
ws.Range(ws.Cells(2, colResultado), ws.Cells(lastRow, colResultado)).Value = resultados
```

---

## Evitar autofit y formato en cada iteración

```vba
' ❌ LENTO — format dentro del loop:
For i = 2 To lastRow
    ws.Cells(i, colMonto).NumberFormat = "$#,##0.00"
    ws.Cells(i, colMonto).Font.Bold = True
Next i

' ✅ RÁPIDO — format al rango completo después del loop:
ws.Range(ws.Cells(2, colMonto), ws.Cells(lastRow, colMonto)).NumberFormat = "$#,##0.00"
ws.Range(ws.Cells(2, colMonto), ws.Cells(lastRow, colMonto)).Font.Bold = True
```

---

## DoEvents — cuándo usarlo

```vba
' DoEvents permite que Excel responda al usuario durante loops largos.
' Usarlo cada N iteraciones, no en cada iteración (tiene costo):

Const DOE_INTERVAL As Integer = 50  ' Cada 50 filas

For i = 2 To lastRow
    Application.StatusBar = "Procesando " & (i-1) & " de " & (lastRow-1)
    
    If i Mod DOE_INTERVAL = 0 Then
        DoEvents  ' Permite que Excel respire y el usuario vea el progreso
    End If
    
    ' ... lógica ...
Next i
Application.StatusBar = False
```

---

## Abrir archivos externos una sola vez

```vba
' ❌ LENTO — abrir y cerrar el mismo archivo en cada iteración del loop:
For i = 2 To lastRow
    Dim wbExt As Workbook
    Set wbExt = Workbooks.Open(rutaArchivo)
    ' leer algo
    wbExt.Close False
Next i

' ✅ RÁPIDO — abrir una vez, usar múltiples veces, cerrar al final:
Dim wbExt As Workbook
Set wbExt = Workbooks.Open(rutaArchivo, ReadOnly:=True)
For i = 2 To lastRow
    ' leer de wbExt sin abrir/cerrar
Next i
wbExt.Close False
```

---

## Benchmark orientativo para BajaTax

| Operación | Sin optimizar | Optimizado | Ganancia |
|---|---|---|---|
| Leer 500 filas de OPERACIONES | ~15s | ~0.3s | 50x |
| Escribir trackers en 500 filas de REGISTROS | ~20s | ~1s | 20x |
| Aplicar formato a 500 filas | ~8s | ~0.5s | 16x |
| Envío masivo WA (500 clientes, pausa 10s) | ~90min | ~90min | Sin ganancia — limitado por pausa anti-baneo |

El envío masivo WA es lento por diseño (anti-baneo). Todo lo demás puede ser casi instantáneo.

---

## Template completo para Subs de proceso masivo

```vba
Sub ProcesarMasivo()
    On Error GoTo ErrorHandler
    
    ' 1. Optimizar Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim wsOps As Worksheet
    Set wsOps = ThisWorkbook.Sheets("OPERACIONES")
    Dim lastRow As Long
    lastRow = wsOps.Cells(wsOps.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then GoTo Salir
    
    ' 2. Leer todo en memoria
    Dim colMax As Long: colMax = 20  ' Columna T
    Dim datos As Variant
    datos = wsOps.Range(wsOps.Cells(2,1), wsOps.Cells(lastRow, colMax)).Value
    
    ' 3. Procesar en memoria
    Dim resultados() As Variant
    ReDim resultados(1 To UBound(datos,1), 1 To 1)
    
    Dim i As Long
    Dim procesados As Long, errores As Long
    For i = 1 To UBound(datos, 1)
        If i Mod 50 = 0 Then
            Application.StatusBar = "Procesando " & i & " de " & UBound(datos,1)
            DoEvents
        End If
        ' ... lógica en memoria ...
        resultados(i,1) = "RESULTADO"
        procesados = procesados + 1
    Next i
    
    ' 4. Escribir resultados de vuelta en una operación
    wsOps.Range(wsOps.Cells(2, colResultado), _
                wsOps.Cells(lastRow, colResultado)).Value = resultados
    
Salir:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Completado: " & procesados & " procesados, " & errores & " errores", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Error en fila " & i & ": " & Err.Description, vbCritical
End Sub
```
