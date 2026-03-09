Attribute VB_Name = "Hoja_DIRECTORIO"
'================================================================
' CODIGO DE HOJA: DIRECTORIO -- BajaTax v4 FINAL
' Columnas A-I:
'   A=RFC  B=CLIENTE  C=CORREO  D=NUMERO  E=REGIMEN
'   F=RESPONSABLE  G=CLASIFICACION  H=FECHA_ALTA  I=ESTADO_CLIENTE
'   Z=Botones accion (Z1-Z2)
'
' BeforeDoubleClick:
'   - Col Z, fila 1 -> InicializarEncabezadosDirectorio
'   - Col Z, fila 2 -> ColorizarEstados
'   - Col I (ESTADO), fila>=2 -> toggle ACTIVO / SUSPENDIDO
'   - Cols A-H, fila>=2 -> edicion controlada con sync a OP/REG
' Worksheet_Change:
'   - Post-edicion: sincronizar a OPERACIONES y REGISTROS
'   - Col A vaciada, fila>=2 -> limpiar fila
'   - Col Z vaciada, fila 1-2 -> restaurar boton Z
'================================================================
Option Explicit

Private Const COL_ESTADO_LOCAL As Integer = 9  ' I

' --- Variables de control de edicion ---
Private mEditando   As Boolean
Private mEditRow    As Long
Private mEditCol    As Integer
Private mEditOldVal As String

'================================================================
'  InicializarBotonesZ_DIRECTORIO -- escribe/restaura Z1-Z2
'================================================================
Public Sub InicializarBotonesZ_DIRECTORIO()
    Dim bEv As Boolean: bEv = Application.EnableEvents
    Application.EnableEvents = False

    With Me.Cells(1, 26)
        .Value               = ChrW(&H21BA) & " INICIALIZAR"
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .Font.Bold           = True
        .Font.Size           = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With

    With Me.Cells(2, 26)
        .Value               = ChrW(&H2605) & " COLORIZAR"
        .Interior.Color      = RGB(112, 173, 71)
        .Font.Color          = RGB(255, 255, 255)
        .Font.Bold           = True
        .Font.Size           = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With

    Application.EnableEvents = bEv
End Sub

'================================================================
'  BeforeDoubleClick -- Z1/Z2, col I (toggle), cols A-H (edicion)
'================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Row < 1 Then Exit Sub
    If Target.CountLarge > 1 Then Exit Sub

    '--- Botones Z (col 26, filas 1-2) -------------------------
    If Target.Column = 26 Then
        Cancel = True
        On Error GoTo FinDC
        Select Case Target.Row
            Case 1: Call InicializarEncabezadosDirectorio
            Case 2: Call ColorizarEstados
        End Select
        Exit Sub
    End If

    '--- Solo filas de datos (fila 2+) -------------------------
    If Target.Row < 2 Then Exit Sub

    '--- Toggle estado col I (fila 2+) -------------------------
    If Target.Column = COL_ESTADO_LOCAL Then
        Cancel = True
        Application.EnableEvents = False
        On Error GoTo FinDC

        Dim sEstado  As String: sEstado  = UCase(Trim(CStr(Target.Value)))
        Dim sCliente As String: sCliente = Trim(CStr(Me.Cells(Target.Row, COL_DIR_CLIENTE).Value))
        If sCliente = "" Then GoTo FinDC

        If sEstado = "SUSPENDIDO" Then
            ' --- REACTIVAR ---
            Dim rAct As Integer
            rAct = MsgBox("Reactivar a " & sCliente & "?" & vbCrLf & _
                          "El cliente podr" & ChrW(&HE1) & " recibir mensajes y PDFs nuevamente.", _
                          vbYesNo + vbQuestion, "BajaTax - Reactivar Cliente")
            If rAct = vbYes Then
                Target.Value               = "ACTIVO"
                Target.Interior.Color      = RGB(198, 239, 206)
                Target.Font.Color          = RGB(0, 97, 0)
                Target.Font.Bold           = True
                Target.HorizontalAlignment = xlCenter
                ' Quitar formato de fila suspendida
                Dim c As Integer
                For c = 1 To 9
                    If c <> COL_ESTADO_LOCAL Then
                        Me.Cells(Target.Row, c).Interior.ColorIndex = xlAutomatic
                        Me.Cells(Target.Row, c).Font.ColorIndex = xlAutomatic
                        Me.Cells(Target.Row, c).Font.Bold = False
                    End If
                Next c
            End If
        Else
            ' --- SUSPENDER ---
            Dim rSus As Integer
            rSus = MsgBox("SUSPENDER a " & sCliente & "?" & vbCrLf & vbCrLf & _
                          "El cliente quedar" & ChrW(&HE1) & " bloqueado para env" & ChrW(&HED) & "os de WA y PDFs." & vbCrLf & _
                          "Para reactivar, vuelve a hacer doble clic en la celda I.", _
                          vbYesNo + vbExclamation, "BajaTax - Suspender Cliente")
            If rSus = vbYes Then
                ' Celda I: SUSPENDIDO en negritas, fondo rojo tenue
                Target.Value               = "SUSPENDIDO"
                Target.Interior.Color      = RGB(255, 199, 206)
                Target.Font.Color          = RGB(156, 0, 6)
                Target.Font.Bold           = True
                Target.HorizontalAlignment = xlCenter
                ' Fila: texto negro, fondo rojo muy tenue
                For c = 1 To 9
                    If c <> COL_ESTADO_LOCAL Then
                        Me.Cells(Target.Row, c).Interior.Color = RGB(255, 230, 230)
                        Me.Cells(Target.Row, c).Font.Color = RGB(0, 0, 0)
                        Me.Cells(Target.Row, c).Font.Bold = True
                    End If
                Next c
            End If
        End If
        GoTo FinDC
    End If

    '--- Edicion controlada cols A-H (fila 2+) -----------------
    If Target.Column >= 1 And Target.Column <= 8 Then
        If Trim(CStr(Target.Value)) = "" Then Exit Sub
        Cancel = True

        Dim resp As Integer
        resp = MsgBox("Deseas editar este dato?" & vbCrLf & vbCrLf & _
                      "  Campo: " & NombreColumnaDir(Target.Column) & vbCrLf & _
                      "  Valor: " & CStr(Target.Value), _
                      vbYesNo + vbQuestion, "BajaTax - Editar Directorio")

        If resp = vbYes Then
            mEditRow    = Target.Row
            mEditCol    = Target.Column
            mEditOldVal = CStr(Target.Value)
            mEditando   = True
            Cancel      = False   ' permitir edicion
        End If
    End If

    Exit Sub

FinDC:
    Application.EnableEvents = True
End Sub

'================================================================
'  Worksheet_Change -- post-edicion sync, col A limpiar, col Z
'================================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    '--- Botones Z vaciados (col 26, filas 1-2) ----------------
    If Target.Column = 26 And Target.Row >= 1 And Target.Row <= 2 Then
        If Trim(CStr(Target.Value)) = "" Then
            Application.EnableEvents = False
            Call InicializarBotonesZ_DIRECTORIO
            Application.EnableEvents = True
        End If
        Exit Sub
    End If

    '--- Post-edicion: sincronizar a OPERACIONES y REGISTROS ---
    If mEditando Then
        If Target.Row = mEditRow And Target.Column = mEditCol Then
            mEditando = False
            Dim nuevoVal As String: nuevoVal = Trim(CStr(Target.Value))

            Application.EnableEvents = False
            On Error GoTo FinSync

            If nuevoVal = "" Then
                ' Celda vaciada: ofrecer restaurar
                Dim rRestore As Integer
                Application.EnableEvents = True
                rRestore = MsgBox("La celda quedo vacia." & vbCrLf & vbCrLf & _
                                  "Restaurar el valor original?" & vbCrLf & _
                                  "  (" & NombreColumnaDir(mEditCol) & " = """ & mEditOldVal & """)", _
                                  vbYesNo + vbQuestion, "BajaTax - Celda Vaciada")
                Application.EnableEvents = False

                If rRestore = vbYes Then
                    Target.Value = mEditOldVal
                End If
            Else
                ' Valor nuevo: preguntar si sincronizar
                Dim rSync As Integer
                Application.EnableEvents = True
                rSync = MsgBox("Quieres cambiar el dato en la hoja de OPERACIONES y REGISTROS?" & vbCrLf & vbCrLf & _
                               "  Campo: " & NombreColumnaDir(mEditCol) & vbCrLf & _
                               "  Antes: " & mEditOldVal & vbCrLf & _
                               "  Nuevo: " & nuevoVal, _
                               vbYesNo + vbQuestion, "BajaTax - Sincronizar Cambio")
                Application.EnableEvents = False

                If rSync = vbYes Then
                    Call SincronizarEdicionDirectorio(mEditRow, mEditCol, nuevoVal, mEditOldVal)
                End If
            End If

FinSync:
            Application.EnableEvents = True
            Exit Sub
        Else
            mEditando = False
        End If
        Exit Sub
    End If

    '--- Col A (RFC) vaciada -> limpiar fila --------------------
    If Target.Column <> 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo FinCh

    Dim celda As Range
    For Each celda In Target
        If celda.Row >= 2 Then
            If Trim(CStr(celda.Value)) = "" Then
                Me.Rows(celda.Row).ClearContents
                Me.Rows(celda.Row).Interior.ColorIndex = xlNone
                Me.Rows(celda.Row).Font.ColorIndex     = xlAutomatic
                Me.Rows(celda.Row).Font.Bold           = False
            End If
        End If
    Next celda

FinCh:
    Application.EnableEvents = True
End Sub

'================================================================
'  NombreColumnaDir -- nombre legible de columna DIRECTORIO
'================================================================
Private Function NombreColumnaDir(NC As Integer) As String
    Select Case NC
        Case 1:  NombreColumnaDir = "RFC"
        Case 2:  NombreColumnaDir = "Cliente"
        Case 3:  NombreColumnaDir = "Correo"
        Case 4:  NombreColumnaDir = "Numero"
        Case 5:  NombreColumnaDir = "Regimen"
        Case 6:  NombreColumnaDir = "Responsable"
        Case 7:  NombreColumnaDir = "Clasificacion"
        Case 8:  NombreColumnaDir = "Fecha Alta"
        Case 9:  NombreColumnaDir = "Estado"
        Case Else: NombreColumnaDir = "Col " & NC
    End Select
End Function

'================================================================
'  InicializarEncabezadosDirectorio
'================================================================
Public Sub InicializarEncabezadosDirectorio()
    Dim bEv As Boolean: bEv = Application.EnableEvents
    Application.EnableEvents = False

    Dim enc As Variant
    enc = Array("RFC", "CLIENTE", "CORREO", _
                "NUMERO", _
                "REGIMEN", _
                "RESPONSABLE", _
                "CLASIFICACION", _
                "FECHA ALTA", _
                "ESTADO CLIENTE")

    Dim col As Integer
    For col = 0 To 8
        With Me.Cells(1, col + 1)
            .Value          = enc(col)
            .Font.Bold      = True
            .Font.Color     = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
        End With
    Next col

    Application.EnableEvents = bEv

    ' Congelar fila 1
    On Error Resume Next
    Me.Activate
    If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False
    Me.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    MsgBox "Encabezados DIRECTORIO actualizados (A-I).", vbInformation, "BajaTax"
End Sub

'================================================================
'  ColorizarEstados -- aplica colores a toda la columna I (ESTADO)
'================================================================
Public Sub ColorizarEstados()
    Dim bEv As Boolean: bEv = Application.EnableEvents
    Application.EnableEvents = False

    Dim uFila As Long
    uFila = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To uFila
        Dim sRFC As String: sRFC = Trim(CStr(Me.Cells(i, 1).Value))
        If sRFC = "" Then GoTo SigE

        Dim sEst As String: sEst = UCase(Trim(CStr(Me.Cells(i, COL_ESTADO_LOCAL).Value)))
        Dim c As Integer

        If sEst = "" Then
            Me.Cells(i, COL_ESTADO_LOCAL).Value               = "ACTIVO"
            Me.Cells(i, COL_ESTADO_LOCAL).Interior.Color      = RGB(198, 239, 206)
            Me.Cells(i, COL_ESTADO_LOCAL).Font.Color          = RGB(0, 97, 0)
            Me.Cells(i, COL_ESTADO_LOCAL).Font.Bold           = True
            Me.Cells(i, COL_ESTADO_LOCAL).HorizontalAlignment = xlCenter
        ElseIf sEst = "SUSPENDIDO" Then
            Me.Cells(i, COL_ESTADO_LOCAL).Interior.Color      = RGB(255, 199, 206)
            Me.Cells(i, COL_ESTADO_LOCAL).Font.Color          = RGB(156, 0, 6)
            Me.Cells(i, COL_ESTADO_LOCAL).Font.Bold           = True
            Me.Cells(i, COL_ESTADO_LOCAL).HorizontalAlignment = xlCenter
            ' Fila: texto negro, fondo rojo tenue
            For c = 1 To 8
                Me.Cells(i, c).Interior.Color = RGB(255, 230, 230)
                Me.Cells(i, c).Font.Color = RGB(0, 0, 0)
                Me.Cells(i, c).Font.Bold = True
            Next c
        ElseIf sEst = "ACTIVO" Then
            Me.Cells(i, COL_ESTADO_LOCAL).Interior.Color      = RGB(198, 239, 206)
            Me.Cells(i, COL_ESTADO_LOCAL).Font.Color          = RGB(0, 97, 0)
            Me.Cells(i, COL_ESTADO_LOCAL).Font.Bold           = True
            Me.Cells(i, COL_ESTADO_LOCAL).HorizontalAlignment = xlCenter
        End If
SigE:
    Next i

    Application.EnableEvents = bEv
End Sub
