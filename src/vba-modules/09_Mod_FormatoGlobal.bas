Attribute VB_Name = "Mod_FormatoGlobal"
Option Explicit

'================================================================
' MODULO: Mod_FormatoGlobal
' PROPOSITO:
'   - Aplicar un formato visual profesional y consistente
'     a las hojas principales (OPERACIONES, REGISTROS, DIRECTORIO)
'   - Ordenar los registros importados para que sean fáciles de leer
'
' COMPATIBILIDAD:
'   - Solo usa objetos nativos de Excel (ListObject, Range, etc.)
'   - Funciona igual en Windows y macOS
'
' COMO INSTALAR:
'   1) Editor VBA > Insertar > Modulo
'   2) Renombrar a "Mod_FormatoGlobal"
'   3) Copiar y pegar todo este código
'   4) Guardar el archivo .xlsm
'
' COMO USAR:
'   - Ejecutar manualmente desde el Editor VBA:
'       AplicarFormatoProfesional
'   - O asignar un botón en la hoja CONFIGURACION:
'       Texto sugerido: "REFRESCAR FORMATO"
'       Macro: AplicarFormatoProfesional
'================================================================

'---------------------------------------------------------------
'  Punto de entrada principal
'---------------------------------------------------------------
Public Sub AplicarFormatoProfesional()
    Dim wsOp As Worksheet
    Dim wsReg As Worksheet
    Dim wsDir As Worksheet

    Set wsOp = ObtenerHoja("OPERACIONES")
    Set wsReg = ObtenerHoja("REGISTROS")
    Set wsDir = ObtenerHoja("DIRECTORIO")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Salir

    If Not wsOp Is Nothing Then
        FormatearHojaComoTabla wsOp, "tblOPERACIONES", 1, COL_OP_ULT_ENVIO
    End If

    If Not wsReg Is Nothing Then
        FormatearHojaComoTabla wsReg, "tblREGISTROS", 1, COL_REG_PROCESADO
        OrdenarRegistrosImportados wsReg
    End If

    If Not wsDir Is Nothing Then
        FormatearHojaComoTabla wsDir, "tblDIRECTORIO", 1, COL_DIR_ESTADO
    End If

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'---------------------------------------------------------------
'  FormatearHojaComoTabla
'  - Crea o recrea una tabla con filtros en la hoja indicada
'  - Aplica un estilo visual limpio y profesional
'---------------------------------------------------------------
Private Sub FormatearHojaComoTabla( _
    ByVal ws As Worksheet, _
    ByVal nombreTabla As String, _
    ByVal filaEncabezados As Long, _
    ByVal ultimaColumna As Long)

    On Error GoTo Fin

    Dim uFila As Long
    uFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If uFila < filaEncabezados Then GoTo Fin

    Dim rngTabla As Range
    Set rngTabla = ws.Range(ws.Cells(filaEncabezados, 1), ws.Cells(uFila, ultimaColumna))

    ' Borrar tabla anterior con el mismo nombre (si existe)
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If StrComp(lo.Name, nombreTabla, vbTextCompare) = 0 Then
            lo.Unlist
            Exit For
        End If
    Next lo

    ' Crear nueva tabla
    Set lo = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=rngTabla, _
        XlListObjectHasHeaders:=xlYes)

    On Error Resume Next
    lo.Name = nombreTabla
    lo.TableStyle = "TableStyleMedium2"
    On Error GoTo 0

    ' Formato general de la hoja
    With rngTabla
        .EntireColumn.AutoFit
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With

    ' Encabezados centrados y con fondo suave
    With ws.Range(ws.Cells(filaEncabezados, 1), ws.Cells(filaEncabezados, ultimaColumna))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With

    ' Congelar fila de encabezados
    CongelarEncabezados ws, filaEncabezados

Fin:
End Sub

'---------------------------------------------------------------
'  OrdenarRegistrosImportados
'  - Ordena REGISTROS por Responsable, Nombre, Fecha y Monto
'  - Hace más legible la importación masiva
'---------------------------------------------------------------
Private Sub OrdenarRegistrosImportados(ByVal ws As Worksheet)
    Dim uFila As Long
    uFila = ws.Cells(ws.Rows.Count, COL_REG_NOMBRE).End(xlUp).Row
    If uFila <= 2 Then Exit Sub

    Dim tieneTabla As Boolean
    tieneTabla = (ws.ListObjects.Count > 0)

    If tieneTabla Then
        With ws.ListObjects(1).Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("A2:A" & uFila), Order:=xlAscending   ' Responsable
            .SortFields.Add Key:=ws.Range("B2:B" & uFila), Order:=xlAscending   ' Nombre
            .SortFields.Add Key:=ws.Range("F2:F" & uFila), Order:=xlAscending   ' Fecha
            .SortFields.Add Key:=ws.Range("H2:H" & uFila), Order:=xlDescending  ' Monto
            .Header = xlYes
            .Apply
        End With
    Else
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("A2:A" & uFila), Order:=xlAscending
            .SortFields.Add Key:=ws.Range("B2:B" & uFila), Order:=xlAscending
            .SortFields.Add Key:=ws.Range("F2:F" & uFila), Order:=xlAscending
            .SortFields.Add Key:=ws.Range("H2:H" & uFila), Order:=xlDescending
            .Header = xlYes
            .SetRange ws.Range("A1:N" & uFila)
            .Apply
        End With
    End If
End Sub

'---------------------------------------------------------------
'  CongelarEncabezados
'  - Congela la fila de encabezados para que siempre sea visible
'---------------------------------------------------------------
Private Sub CongelarEncabezados(ByVal ws As Worksheet, ByVal filaEnc As Long)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(filaEnc + 1, 1).Select
    ActiveWindow.FreezePanes = True
End Sub

