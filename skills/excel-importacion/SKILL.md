---
name: excel-importacion
description: >
  Usar esta skill para cualquier tarea relacionada con importar datos externos al
  Excel de BajaTax. Activar cuando el usuario mencione importación, archivos de
  clientes, Contalink, mapeo de columnas, detección automática, headers, columnas
  revueltas, archivos sin headers, datos mezclados, deduplicación, distribución
  a OPERACIONES o DIRECTORIO, o cuando el código necesite leer un .xlsx o .csv
  externo y colocar datos en REGISTROS. También activar cuando el usuario diga
  "el archivo tiene formato raro", "las columnas están en otro orden", o "algunos
  archivos no tienen headers". Esta skill contiene el algoritmo completo de
  importación inteligente de BajaTax — úsala siempre que haya datos externos.
---

# Excel Importación Inteligente — BajaTax

## El problema central

Los archivos de clientes llegan en formatos completamente distintos:
headers en distintas filas, columnas en cualquier orden, sin headers,
datos horizontales, basura mezclada, múltiples hojas. El sistema debe
manejar todo sin pedirle al usuario que reformatee el archivo.

---

## Algoritmo de detección en dos fases

### Fase 1: Escaneo automático

```vba
Sub Fase1_EscanearArchivo(wbExt As Workbook)
    ' 1. Analizar CADA hoja del workbook externo
    Dim ws As Worksheet
    For Each ws In wbExt.Sheets
        ' 2. Buscar la fila de headers (donde hay más etiquetas que valores)
        Dim headerRow As Long
        headerRow = DetectarFilaHeaders(ws)
        
        ' 3. Si no hay headers → clasificar columnas por CONTENIDO
        If headerRow = -1 Then
            ClasificarPorContenido ws
        Else
            ClasificarPorHeaders ws, headerRow
        End If
    Next ws
End Sub

Function DetectarFilaHeaders(ws As Worksheet) As Long
    ' Buscar en primeras 20 filas la que tenga más texto tipo-etiqueta
    Dim fila As Long, mejorFila As Long, mejorScore As Integer
    For fila = 1 To 20
        Dim score As Integer: score = 0
        Dim col As Long
        For col = 1 To 15
            Dim val As String: val = Trim(CStr(ws.Cells(fila, col).Value))
            If val = "" Then GoTo SiguienteCol
            ' Un header tiene texto, NO números ni fechas
            If Not IsNumeric(val) And Not IsDate(val) And Len(val) < 50 Then
                score = score + 1
            End If
SiguienteCol:
        Next col
        If score > mejorScore Then
            mejorScore = score: mejorFila = fila
        End If
    Next fila
    If mejorScore >= 3 Then DetectarFilaHeaders = mejorFila Else DetectarFilaHeaders = -1
End Function
```

### Fase 2: Confirmación del usuario

Antes de importar, mostrar ventana con el mapeo propuesto:
```
"Columna C → RFC (92% confianza)"
"Columna E → Email (88% confianza)"
"Columna H → Monto (85% confianza)"
"Columna F → ? (desconocida) — ¿qué campo es?"

[Confirmar] [Corregir mapeo]
```

El usuario confirma o corrige. Solo entonces se importan los datos.

---

## Detección por contenido (método primario)

Analizar muestra de 10 celdas por columna y clasificar por patrones:

```vba
Function ClasificarColumna(ws As Worksheet, col As Long, dataStartRow As Long) As String
    Dim rfcCount As Integer, emailCount As Integer
    Dim phoneCount As Integer, montoCount As Integer
    Dim dateCount As Integer, total As Integer

    Dim i As Long
    For i = dataStartRow To Application.Min(dataStartRow + 9, _
            ws.Cells(ws.Rows.Count, col).End(xlUp).Row)
        Dim v As String: v = Trim(CStr(ws.Cells(i, col).Value))
        If v = "" Then GoTo Next_i
        total = total + 1

        ' RFC: 12-13 alfanumérico con patrón fiscal
        If Len(v) = 12 Or Len(v) = 13 Then
            If v Like "[A-Za-z][A-Za-z][A-Za-z]*[0-9][0-9][0-9][0-9][0-9][0-9]*" Then
                rfcCount = rfcCount + 1
            End If
        End If
        ' Email: contiene @ y punto después
        If InStr(v,"@") > 0 And InStr(v,".") > InStr(v,"@") Then
            emailCount = emailCount + 1
        End If
        ' Teléfono: 10 dígitos numéricos
        Dim digitsOnly As String: digitsOnly = ""
        Dim j As Integer
        For j = 1 To Len(v)
            If IsNumeric(Mid(v,j,1)) Then digitsOnly = digitsOnly & Mid(v,j,1)
        Next j
        If Len(digitsOnly) = 10 Then phoneCount = phoneCount + 1
        ' Monto: número con decimales o símbolo $
        If IsNumeric(Replace(Replace(Replace(v,",",""),"$","")," ","")) And _
           (InStr(v,".") > 0 Or InStr(v,"$") > 0) Then montoCount = montoCount + 1
        ' Fecha
        If IsDate(v) Then dateCount = dateCount + 1
Next_i:
    Next i

    If total = 0 Then ClasificarColumna = "UNKNOWN": Exit Function
    Dim threshold As Double: threshold = 0.6

    If rfcCount / total >= threshold Then    ClasificarColumna = "RFC"
    ElseIf emailCount / total >= threshold Then ClasificarColumna = "EMAIL"
    ElseIf phoneCount / total >= threshold Then ClasificarColumna = "TELEFONO"
    ElseIf montoCount / total >= threshold Then ClasificarColumna = "MONTO"
    ElseIf dateCount / total >= threshold Then  ClasificarColumna = "FECHA"
    Else:                                        ClasificarColumna = "UNKNOWN"
    End If
End Function
```

---

## Orden de prioridad de detección

1. **Contenido** (patrones regex) — SIEMPRE primero
2. **Aliases de headers** (si la hoja tiene headers reconocibles) — segundo
3. **Posición heurística** (último recurso para columnas muy vacías) — tercero
4. **Preguntar al usuario** (si confianza < 60% en alguna columna importante)

---

## Distribución REGISTROS → OPERACIONES + DIRECTORIO

Al procesar una fila de REGISTROS (doble clic):

```vba
Sub ProcesarRegistro(wsReg As Worksheet, fila As Long)
    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    Dim rfc As String:    rfc    = wsReg.Cells(fila, "C").Value
    Dim nombre As String: nombre = wsReg.Cells(fila, "B").Value
    Dim tel As String:    tel    = wsReg.Cells(fila, "E").Value
    Dim email As String:  email  = wsReg.Cells(fila, "D").Value

    ' PASO 1: DIRECTORIO
    Dim wsDir As Worksheet
    Set wsDir = ThisWorkbook.Sheets("DIRECTORIO")
    Dim colRFCDir As Long: colRFCDir = Util_GetColumnByHeader(wsDir, "RFC")
    Dim rfcRow As Long:    rfcRow    = BuscarRFC(wsDir, rfc, colRFCDir)

    If rfcRow = -1 Then
        ' Crear nueva entidad
        Dim newRow As Long
        newRow = wsDir.Cells(wsDir.Rows.Count, colRFCDir).End(xlUp).Row + 1
        wsDir.Cells(newRow, colRFCDir).Value                              = rfc
        wsDir.Cells(newRow, Util_GetColumnByHeader(wsDir,"CLIENTE")).Value = nombre
        wsDir.Cells(newRow, Util_GetColumnByHeader(wsDir,"CORREO")).Value  = email
        wsDir.Cells(newRow, Util_GetColumnByHeader(wsDir,"TELEFONO")).Value = tel
        wsDir.Cells(newRow, Util_GetColumnByHeader(wsDir,"FECHA ALTA")).Value = Now()
        wsDir.Cells(newRow, Util_GetColumnByHeader(wsDir,"ESTADO_CLIENTE")).Value = "ACTIVO"
    Else
        ' Verificar si hay cambios en contacto
        Dim telExistente As String
        telExistente = wsDir.Cells(rfcRow, Util_GetColumnByHeader(wsDir,"TELEFONO")).Value
        If telExistente <> tel And tel <> "" Then
            If MsgBox("Teléfono diferente para " & nombre & vbLf & _
                      "Actual: " & telExistente & vbLf & "Nuevo: " & tel & vbLf & _
                      "¿Actualizar?", vbYesNo) = vbYes Then
                wsDir.Cells(rfcRow, Util_GetColumnByHeader(wsDir,"TELEFONO")).Value = tel
            End If
        End If
    End If
    wsReg.Cells(fila, "M").Value = "✓ DIRECTORIO"

    ' PASO 2: OPERACIONES — verificar duplicado con clave compuesta
    Dim wsOps As Worksheet
    Set wsOps = ThisWorkbook.Sheets("OPERACIONES")
    Dim concepto As String:    concepto   = wsReg.Cells(fila, "G").Value
    Dim fechaCobro As String:  fechaCobro = CStr(wsReg.Cells(fila, "F").Value)
    Dim monto As Double:       monto      = CDbl(wsReg.Cells(fila, "H").Value)

    Dim esDuplicado As Boolean
    esDuplicado = VerificarDuplicado(wsOps, rfc, concepto, fechaCobro, monto)

    If esDuplicado Then
        If MsgBox("Registro duplicado: " & nombre & " / " & concepto & vbLf & _
                  "¿Re-procesar de todas formas?", vbYesNo) = vbNo Then
            GoTo Paso3
        End If
    End If

    ' Insertar en OPERACIONES
    Dim opRow As Long
    opRow = wsOps.Cells(wsOps.Rows.Count, 1).End(xlUp).Row + 1
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"RESPONSABLE")).Value    = wsReg.Cells(fila,"A").Value
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"ID FACTURA")).Value      = wsReg.Cells(fila,"I").Value
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"REGIMEN")).Value         = wsReg.Cells(fila,"J").Value
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"CLIENTE")).Value         = nombre
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"RFC")).Value             = rfc
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"FECHA COBRANZA")).Value  = wsReg.Cells(fila,"F").Value
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"Concepto")).Value        = concepto
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"Monto Base")).Value      = monto
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"Fecha Vencimiento")).Value = wsReg.Cells(fila,"K").Value
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"Teléfono")).Value        = tel
    wsOps.Cells(opRow, Util_GetColumnByHeader(wsOps,"Correo")).Value          = email
    wsReg.Cells(fila, "L").Value = "✓ OPERACIONES"

Paso3:
    ' PASO 3: Marcar como procesado
    wsReg.Cells(fila, "N").Value = "✓ PROCESADO"
    wsReg.Cells(fila, "N").Interior.Color = RGB(198, 239, 206)  ' Verde
    LogEvento "DISTRIBUCIÓN", nombre & " | " & rfc, "OK"

    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    LogEvento "DISTRIBUCIÓN_ERROR", "Fila " & fila & ": " & Err.Description, "FALLO"
    MsgBox "Error procesando fila " & fila & ": " & Err.Description, vbCritical
End Sub

Function VerificarDuplicado(wsOps As Worksheet, rfc As String, _
    concepto As String, fechaCobro As String, monto As Double) As Boolean
    Dim colRFC As Long:     colRFC     = Util_GetColumnByHeader(wsOps,"RFC")
    Dim colConcepto As Long: colConcepto = Util_GetColumnByHeader(wsOps,"Concepto")
    Dim colFecha As Long:   colFecha   = Util_GetColumnByHeader(wsOps,"FECHA COBRANZA")
    Dim colMonto As Long:   colMonto   = Util_GetColumnByHeader(wsOps,"Monto Base")
    Dim lastRow As Long:    lastRow    = wsOps.Cells(wsOps.Rows.Count,1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If wsOps.Cells(i,colRFC).Value = rfc And _
           wsOps.Cells(i,colConcepto).Value = concepto And _
           CStr(wsOps.Cells(i,colFecha).Value) = fechaCobro And _
           wsOps.Cells(i,colMonto).Value = monto Then
            VerificarDuplicado = True: Exit Function
        End If
    Next i
    VerificarDuplicado = False
End Function
```

---

## Situaciones que el sistema maneja

| Situación | Manejo |
|---|---|
| Sin headers | Detectar por contenido de columnas |
| Headers en fila 2,3,4... | Buscar en primeras 20 filas |
| Columnas en cualquier orden | Detectar por contenido, no por posición |
| Múltiples hojas | Analizar cada una, reportar qué encontró en cada hoja |
| Filas vacías / basura | Saltar filas donde RFC y Nombre están vacíos |
| Teléfonos con/sin formato | Normalizar con `Util_NormalizePhone()` |
| RFC mayúsculas/minúsculas | Normalizar con `UCase(Trim(rfc))` |
| Archivos .csv | Abrir con `Workbooks.Open` igual que .xlsx |
| Columna desconocida | Preguntar al usuario, no descartar |
