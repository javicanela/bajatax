Attribute VB_Name = "Hoja_OPERACIONES"
'================================================================
' CODIGO DE HOJA: OPERACIONES -- BajaTax v4 FINAL
'
' Botones accion : doble-clic en col Z (filas 1-5)
' Botones WA/PDF : doble-clic en cols O (15) / P (16)
' Registro pago  : doble-clic en col L (12)
' Auto-regenera  : cols I (formula), K (formula), O, P, Z
'
' REGLA COL I: NO se escribe "PAGADO" directo.
'   La formula IF(L<>"","PAGADO",...) lo maneja automaticamente.
'   Solo se gestiona el COLOR de la celda.
'
' TABLA CON FILTROS: AutoFilter habilitado en fila 1 (A1:T1)
'================================================================
Option Explicit

Private bOcupado As Boolean

'----------------------------------------------------------------
'  Texto de cada boton Z (filas 1-5)
'----------------------------------------------------------------
Private Function TextoBotonZ(fila As Long) As String
    Select Case fila
        Case 1: TextoBotonZ = ChrW(&H25B6) & " IMPORTAR"
        Case 2: TextoBotonZ = ChrW(&H25B6) & " PROCESAR TODO"
        Case 3: TextoBotonZ = ChrW(&H25B6) & " ENVIO MASIVO WA"
        Case 4: TextoBotonZ = ChrW(&H25A0) & " PDF MASIVO"
        Case 5: TextoBotonZ = ChrW(&H21BA) & " REGENERAR"
        Case Else: TextoBotonZ = ""
    End Select
End Function

'----------------------------------------------------------------
'  InicializarBotonesZ -- escribe/restaura Z1-Z5 + activa AutoFilter
'----------------------------------------------------------------
Public Sub InicializarBotonesZ()
    Dim bEv As Boolean: bEv = Application.EnableEvents
    Application.EnableEvents = False
    Dim i As Long
    For i = 1 To 5
        With Me.Cells(i, 26)
            .Value               = TextoBotonZ(i)
            .Interior.Color      = RGB(68, 114, 196)
            .Font.Color          = RGB(255, 255, 255)
            .Font.Bold           = True
            .Font.Size           = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment   = xlCenter
            .WrapText            = False
        End With
    Next i

    ' Activar AutoFilter (tabla con filtros) en fila 1
    On Error Resume Next
    If Not Me.AutoFilterMode Then
        Me.Range("A1:T1").AutoFilter
    End If
    On Error GoTo 0

    Application.EnableEvents = bEv
End Sub

'================================================================
'  BeforeDoubleClick
'  Maneja: Z1-Z5 (botones accion), O (WA), P (PDF), L (pago)
'================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.CountLarge > 1 Then Exit Sub
    If bOcupado Then Exit Sub

    Dim NL As Long: NL = Target.Row
    Dim NC As Long: NC = Target.Column

    '--- Botones Z (col 26, filas 1-5) -------------------------
    If NC = 26 And NL >= 1 And NL <= 5 Then
        Cancel = True
        bOcupado = True
        On Error GoTo FinDC
        Select Case NL
            Case 1: Call ImportarArchivosExternos
            Case 2: Call ProcesarTodoBajaTax
            Case 3: Call EnvioMasivoAutomatico
            Case 4: Call GenerarPDFMasivo
            Case 5: Call RegenerarFaltantes
        End Select
        GoTo FinDC
    End If

    '--- Solo filas de datos (fila 2+) -------------------------
    If NL < 2 Then Exit Sub

    Dim sCliente As String
    sCliente = Trim(CStr(Me.Cells(NL, COL_OP_CLIENTE).Value))

    '--- Boton WA (col O = 15) ---------------------------------
    If NC = COL_OP_WA Then
        If sCliente = "" Then Exit Sub
        Cancel = True
        bOcupado = True
        Application.EnableEvents = False
        On Error GoTo FinDC
        Call EnviarMensajeInteligente(NL)
        GoTo FinDC
    End If

    '--- Boton PDF (col P = 16) --------------------------------
    If NC = COL_OP_PDF Then
        If sCliente = "" Then Exit Sub
        Cancel = True
        bOcupado = True
        Application.EnableEvents = False
        On Error GoTo FinDC
        Call GenerarEstadoCuentaPDF(NL)
        GoTo FinDC
    End If

    '--- Registro / Cancelacion de pago (col L = 12) ----------
    If NC = COL_OP_REG_PAGO Then
        If sCliente = "" Then Exit Sub
        Cancel = True
        bOcupado = True
        Application.EnableEvents = False
        On Error GoTo FinDC

        Dim sPago As String: sPago = Trim(CStr(Target.Value))

        If sPago = "" Then
            '=========== REGISTRAR PAGO NUEVO ===================
            Dim dMonto    As Double: dMonto    = 0
            Dim sConcepto As String: sConcepto = Trim(CStr(Me.Cells(NL, COL_OP_CONCEPTO).Value))
            On Error Resume Next
            dMonto = CDbl(Me.Cells(NL, COL_OP_MONTO).Value)
            On Error GoTo FinDC

            Dim conf As Integer
            Application.EnableEvents = True
            conf = MsgBox("Confirmar pago recibido de:" & vbCrLf & vbCrLf & _
                          "  Cliente:  " & sCliente & vbCrLf & _
                          "  Monto:    " & Format(dMonto, "$#,##0.00") & vbCrLf & _
                          "  Concepto: " & sConcepto & vbCrLf & vbCrLf & _
                          "Fecha a registrar: " & Format(Date, "dd/mm/yyyy"), _
                          vbYesNo + vbQuestion, "BajaTax - Registrar Pago")
            Application.EnableEvents = False

            If conf = vbYes Then
                ' 1. Escribir en col L (fecha de pago)
                Target.Value               = Format(Now, "dd/mm/yyyy hh:mm")
                Target.Interior.Color      = RGB(198, 239, 206)
                Target.Font.Color          = RGB(0, 97, 0)
                Target.HorizontalAlignment = xlCenter

                ' 2. Forzar recalculo -> col I (formula) muestra "PAGADO"
                Application.Calculate

                ' 3. Colorear celda I (el VALOR lo maneja la formula)
                With Me.Cells(NL, COL_OP_ESTATUS)
                    .Interior.Color = RGB(198, 239, 206)
                    .Font.Color     = RGB(0, 97, 0)
                    .Font.Bold      = True
                End With

                ' 4. Escribir "PAGADO" en O y P (reemplazar botones WA/PDF)
                With Me.Cells(NL, COL_OP_WA)
                    .Value               = "PAGADO"
                    .Interior.Color      = RGB(198, 239, 206)
                    .Font.Color          = RGB(0, 97, 0)
                    .Font.Bold           = True
                    .HorizontalAlignment = xlCenter
                End With
                With Me.Cells(NL, COL_OP_PDF)
                    .Value               = "PAGADO"
                    .Interior.Color      = RGB(198, 239, 206)
                    .Font.Color          = RGB(0, 97, 0)
                    .Font.Bold           = True
                    .HorizontalAlignment = xlCenter
                End With

                ' 5. Colorear resto de la fila
                Me.Range(Me.Cells(NL, 1), Me.Cells(NL, COL_OP_ULT_ENVIO)) _
                    .Interior.Color = RGB(198, 239, 206)

                ' 6. Registrar en log
                Call RegistrarPagoEnLog(sCliente, dMonto, sConcepto)
            End If

        Else
            '=========== CANCELAR PAGO EXISTENTE ================
            Dim confElim As Integer
            Application.EnableEvents = True
            confElim = MsgBox("Cancelar el registro de pago de:" & vbCrLf & vbCrLf & _
                              "  Cliente: " & sCliente & vbCrLf & _
                              "  Fecha:   " & sPago & vbCrLf & vbCrLf & _
                              "El estatus volvera al calculado automaticamente.", _
                              vbYesNo + vbExclamation, "BajaTax - Cancelar Pago")
            Application.EnableEvents = False

            If confElim = vbYes Then
                ' 1. Limpiar col L -> formula de col I recalcula sola
                Target.ClearContents
                Target.Interior.ColorIndex = xlAutomatic
                Target.Font.ColorIndex     = xlAutomatic

                ' 2. Forzar recalculo
                Application.Calculate

                ' 3. Quitar color de celda I
                With Me.Cells(NL, COL_OP_ESTATUS)
                    .Interior.ColorIndex = xlAutomatic
                    .Font.ColorIndex     = xlAutomatic
                    .Font.Bold           = False
                End With

                ' 4. Quitar color de fila
                Me.Range(Me.Cells(NL, 1), Me.Cells(NL, COL_OP_ULT_ENVIO)) _
                    .Interior.ColorIndex = xlAutomatic

                ' 5. Restaurar botones WA y PDF
                Call InicializarBotonFila(Me, NL)
            End If
        End If
    End If

FinDC:
    bOcupado = False
    Application.EnableEvents = True
End Sub

'================================================================
'  Worksheet_Change -- auto-regenerar celdas borradas
'================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    If bOcupado Then Exit Sub
    If Target.CountLarge > 100 Then Exit Sub

    bOcupado = True
    Application.EnableEvents = False
    On Error GoTo FinCh

    Dim NL As Long: NL = Target.Row
    Dim NC As Long: NC = Target.Column

    '--- Botones Z vaciados (col 26, filas 1-5) ----------------
    If NC = 26 And NL >= 1 And NL <= 5 Then
        If Trim(CStr(Target.Value)) = "" Then
            Call InicializarBotonesZ
        End If
        GoTo FinCh
    End If

    If NL < 2 Then GoTo FinCh

    '--- Col I (Estatus) vaciada -> restaurar formula ----------
    If NC = COL_OP_ESTATUS Then
        If Trim(CStr(Target.Value)) = "" Then
            If Trim(CStr(Me.Cells(NL, COL_OP_CLIENTE).Value)) <> "" Then
                Me.Cells(NL, COL_OP_ESTATUS).Formula = _
                    "=IF(D" & NL & "="""","""",IF(L" & NL & "<>"""",""PAGADO""," & _
                    "IF(J" & NL & "="""",""PENDIENTE"",IF(TODAY()>J" & NL & _
                    ",""VENCIDO"",IF(TODAY()=J" & NL & ",""HOY VENCE"",""PENDIENTE"")))))"
            End If
        End If
        GoTo FinCh
    End If

    '--- Col K (Dias_Venc) vaciada -> restaurar formula --------
    If NC = COL_OP_DIAS_VENC Then
        If Trim(CStr(Target.Value)) = "" Then
            If Trim(CStr(Me.Cells(NL, COL_OP_CLIENTE).Value)) <> "" Then
                Me.Cells(NL, COL_OP_DIAS_VENC).Formula = _
                    "=IFERROR(IF(J" & NL & "="""","""",TODAY()-J" & NL & "),"""")"
            End If
        End If
        GoTo FinCh
    End If

    '--- Col O (WA) vaciada -> restaurar boton ----------------
    If NC = COL_OP_WA Then
        If Trim(CStr(Target.Value)) = "" Then
            If Trim(CStr(Me.Cells(NL, COL_OP_CLIENTE).Value)) <> "" And _
               Trim(CStr(Me.Cells(NL, COL_OP_MONTO).Value)) <> "" Then
                ' Solo restaurar boton si NO esta PAGADO
                If UCase(Trim(CStr(Me.Cells(NL, COL_OP_ESTATUS).Value))) <> "PAGADO" Then
                    Call InicializarBotonFila(Me, NL)
                End If
            End If
        End If
        GoTo FinCh
    End If

    '--- Col P (PDF) vaciada -> restaurar boton ---------------
    If NC = COL_OP_PDF Then
        If Trim(CStr(Target.Value)) = "" Then
            If Trim(CStr(Me.Cells(NL, COL_OP_CLIENTE).Value)) <> "" And _
               Trim(CStr(Me.Cells(NL, COL_OP_MONTO).Value)) <> "" Then
                If UCase(Trim(CStr(Me.Cells(NL, COL_OP_ESTATUS).Value))) <> "PAGADO" Then
                    Call InicializarBotonFila(Me, NL)
                End If
            End If
        End If
        GoTo FinCh
    End If

    '--- Col D (Cliente) o H (Monto) con nuevo valor -> init botones
    If NC = COL_OP_CLIENTE Or NC = COL_OP_MONTO Then
        Dim kFila As Long
        For kFila = NL To NL + Target.Rows.Count - 1
            If Trim(CStr(Me.Cells(kFila, COL_OP_CLIENTE).Value)) <> "" Then
                If Trim(CStr(Me.Cells(kFila, COL_OP_WA).Value)) = "" Then
                    Call InicializarBotonFila(Me, kFila)
                End If
            End If
        Next kFila
    End If

FinCh:
    bOcupado = False
    Application.EnableEvents = True
End Sub

'================================================================
'  Worksheet_Activate -- asegurar AutoFilter activo
'================================================================
Private Sub Worksheet_Activate()
    On Error Resume Next
    If Not Me.AutoFilterMode Then
        If Me.Cells(1, 1).Value <> "" Then
            Me.Range("A1:T1").AutoFilter
        End If
    End If
    On Error GoTo 0
End Sub

'================================================================
'  RegistrarPagoEnLog -- escribe en hoja LOG ENVIOS
'================================================================
Private Sub RegistrarPagoEnLog(sCliente As String, dMonto As Double, sConcepto As String)
    Dim wsLog As Worksheet
    Set wsLog = ObtenerHoja("LOG ENVIOS")
    If wsLog Is Nothing Then Set wsLog = ObtenerHoja("LOG ENV" & ChrW(&HCD) & "OS")
    If wsLog Is Nothing Then Exit Sub

    Dim uFila As Long
    uFila = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    With wsLog
        .Cells(uFila, 1).Value          = Format(Now, "dd/mm/yyyy hh:mm:ss")
        .Cells(uFila, 2).Value          = ""
        .Cells(uFila, 3).Value          = sCliente
        .Cells(uFila, 4).Value          = Format(dMonto, "$#,##0.00")
        .Cells(uFila, 5).Value          = "PAGO REGISTRADO"
        .Cells(uFila, 6).Value          = ModoSistema()
        .Cells(uFila, 7).Value          = "PAGO OK"
        .Cells(uFila, 1).Interior.Color = RGB(198, 239, 206)
    End With
End Sub
