Attribute VB_Name = "WhatsApp"
'================================================================
' MODULO: WhatsApp -- BajaTax v4 FINAL
' - Mensajes con formato *bold* para WhatsApp
' - Envio masivo CONSOLIDADO por telefono
' - Anti-ban: intervalo 8-15 s aleatorio
' - Verificacion SUSPENDIDO antes de enviar
'================================================================
Option Explicit

' ---------------------------------------------------------------
'  EnviarMensajeInteligente -- envio individual desde OPERACIONES
' ---------------------------------------------------------------
Public Sub EnviarMensajeInteligente(ByVal NL As Long)
    If Not HojasOK() Then Exit Sub

    Dim wsOp   As Worksheet: Set wsOp   = ObtenerHoja("OPERACIONES")
    Dim wsConf As Worksheet: Set wsConf = ObtenerHoja("CONFIGURACION")

    Dim sCliente  As String: sCliente  = Trim(CStr(wsOp.Cells(NL, COL_OP_CLIENTE).Value))
    Dim sRFC      As String: sRFC      = Trim(CStr(wsOp.Cells(NL, COL_OP_RFC).Value))
    Dim sConcepto As String: sConcepto = Trim(CStr(wsOp.Cells(NL, COL_OP_CONCEPTO).Value))
    Dim sResp     As String: sResp     = Trim(CStr(wsOp.Cells(NL, COL_OP_RESPONSABLE).Value))
    Dim sEstatus  As String: sEstatus  = UCase(Trim(CStr(wsOp.Cells(NL, COL_OP_ESTATUS).Value)))
    Dim sTelRaw   As String: sTelRaw   = Trim(CStr(wsOp.Cells(NL, COL_OP_TELEFONO).Value))
    Dim fVencim   As Variant: fVencim  = wsOp.Cells(NL, COL_OP_VENCIMIENTO).Value
    Dim dMonto    As Double: dMonto    = 0
    On Error Resume Next
    dMonto = CDbl(wsOp.Cells(NL, COL_OP_MONTO).Value)
    On Error GoTo 0

    ' Validaciones basicas
    If sCliente = "" Then
        MsgBox "Fila " & NL & " sin cliente.", vbExclamation, "BajaTax": Exit Sub
    End If
    If sEstatus = "PAGADO" Then
        MsgBox sCliente & " ya esta PAGADO.", vbInformation, "BajaTax": Exit Sub
    End If
    If dMonto <= 0 Then
        MsgBox "Monto $0. No se envia a " & sCliente & ".", vbExclamation, "BajaTax": Exit Sub
    End If

    ' Verificar SUSPENDIDO
    If RFCSuspendido(sRFC) Then
        MsgBox "El cliente " & sCliente & " esta SUSPENDIDO en el DIRECTORIO." & Chr(13) & _
               "No se enviara mensaje.", vbExclamation, "BajaTax -- Cliente Suspendido"
        Exit Sub
    End If

    ' Calcular dias y fecha
    Dim diasVenc As Long: diasVenc = 0
    Dim sFecha   As String: sFecha  = "sin fecha"
    If IsDate(fVencim) Then
        diasVenc = DateDiff("d", CDate(fVencim), Date)
        sFecha   = Format(CDate(fVencim), "dd/mm/yyyy")
    End If

    Dim sMonto As String: sMonto = Format(dMonto, "$#,##0.00")

    ' Determinar variante
    Dim sVariante As String
    If diasVenc = 0 Then
        sVariante = "HOY_VENCE"
    ElseIf diasVenc > 0 Then
        sVariante = "VENCIDO"
    Else
        sVariante = "RECORDATORIO"
    End If

    ' Construir mensaje
    Dim sMensaje As String
    sMensaje = ConstruirMensaje(sVariante, sCliente, sMonto, sConcepto, sFecha, diasVenc)

    ' Determinar numero destino
    Dim sTelDestino As String
    Dim sModo As String: sModo = ModoSistema()

    If sModo = "PRUEBA" Then
        Dim sTelPrueba As String: sTelPrueba = Trim(CStr(wsConf.Range("B14").Value))
        If sTelPrueba = "" Then
            MsgBox "MODO PRUEBA: no hay numero en B14.", vbExclamation, "BajaTax": Exit Sub
        End If
        sTelDestino = LimpiarTelefono(sTelPrueba)
    Else
        If sTelRaw = "" Then
            MsgBox sCliente & " no tiene telefono en columna M.", vbExclamation, "BajaTax": Exit Sub
        End If
        sTelDestino = LimpiarTelefono(sTelRaw)
    End If

    If Len(sTelDestino) < 10 Then
        MsgBox "Telefono invalido: '" & sTelRaw & "'", vbExclamation, "BajaTax": Exit Sub
    End If

    ' Confirmar reenvio si ya fue enviado
    Dim sBotActual As String
    sBotActual = Trim(CStr(wsOp.Cells(NL, COL_OP_WA).Value))
    If InStr(sBotActual, "REENVIAR") > 0 Then
        Dim rta As Integer
        rta = MsgBox("Ya fue enviado anteriormente." & Chr(13) & _
                     "Si=Restaurar boton  No=Reenviar  Cancelar=Nada", _
                     vbYesNoCancel + vbQuestion, "BajaTax -- Reenvio")
        Select Case rta
            Case vbYes:    Call RestaurarBotonWA(NL, diasVenc): Exit Sub
            Case vbCancel: Exit Sub
        End Select
    End If

    ' Vista previa en modo PRUEBA
    If sModo = "PRUEBA" Then
        Dim prev As Integer
        prev = MsgBox("*** MODO PRUEBA ***" & Chr(13) & _
                      "Envio al: " & sTelDestino & Chr(13) & Chr(13) & _
                      Left(sMensaje, 500) & IIf(Len(sMensaje) > 500, "...", ""), _
                      vbYesNo + vbInformation, "BajaTax -- Vista previa")
        If prev = vbNo Then Exit Sub
    End If

    ' Abrir WhatsApp
    Dim sURL As String
    sURL = "https://wa.me/" & sTelDestino & "?text=" & CodificarWhatsApp(sMensaje)
    On Error GoTo ErrorAbrirURL
    ActiveWorkbook.FollowHyperlink Address:=sURL
    GoTo URLAbierta

ErrorAbrirURL:
    If Not EsMac() Then Shell "cmd /c start " & sURL, vbHide
    Resume URLAbierta

URLAbierta:
    On Error GoTo 0

    ' Actualizar boton con timestamp
    wsOp.Cells(NL, COL_OP_WA).Value = "REENVIAR " & Format(Now, "dd/mm HH:mm")
    Select Case sVariante
        Case "VENCIDO":    wsOp.Cells(NL, COL_OP_WA).Interior.Color = RGB(255, 199, 206)
        Case "HOY_VENCE":  wsOp.Cells(NL, COL_OP_WA).Interior.Color = RGB(255, 235, 156)
        Case "RECORDATORIO": wsOp.Cells(NL, COL_OP_WA).Interior.Color = RGB(198, 224, 180)
    End Select

    ' Registrar en LOG y actualizar contadores
    Call RegistrarLogEnvio(sResp, sCliente, sVariante, sMonto, sConcepto, sTelDestino, sModo)

    Dim intentos As Long: intentos = 0
    On Error Resume Next
    intentos = CLng(wsOp.Cells(NL, COL_OP_INTENTOS).Value)
    On Error GoTo 0
    wsOp.Cells(NL, COL_OP_INTENTOS).Value  = intentos + 1
    wsOp.Cells(NL, COL_OP_ULT_ENVIO).Value = Format(Now, "dd/mm/yyyy hh:mm")
End Sub

' ---------------------------------------------------------------
'  EnvioMasivoConsolidado
'  Agrupa por telefono -> 1 mensaje por telefono (suma adeudos)
'  Anti-ban: 8-15 s aleatorios entre envios
' ---------------------------------------------------------------
Public Sub EnvioMasivoConsolidado()
    If Not HojasOK() Then Exit Sub

    Dim wsOp  As Worksheet: Set wsOp  = ObtenerHoja("OPERACIONES")
    Dim sModo As String:    sModo     = ModoSistema()

    Dim resp As Integer
    resp = MsgBox("ENVIO MASIVO CONSOLIDADO" & Chr(13) & Chr(13) & _
                  "Modo: " & sModo & Chr(13) & Chr(13) & _
                  "Se enviara UN mensaje por telefono, consolidando " & Chr(13) & _
                  "todos los adeudos pendientes de ese numero." & Chr(13) & Chr(13) & _
                  "Intervalo anti-ban: " & ANTI_BAN_MIN & "-" & ANTI_BAN_MAX & " segundos." & Chr(13) & _
                  ChrW(&H26A0) & " Continuar?", _
                  vbYesNo + vbExclamation, "BajaTax -- Envio Masivo")
    If resp = vbNo Then Exit Sub

    Dim uFila As Long
    uFila = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row

    ' -- Recopilar grupos por telefono --------------------------
    ' Estructura paralela: arrays con datos agrupados (dinamicos, sin limite fijo)
    Dim arrTel()       As String   ' telefonos unicos (limpiado)
    Dim arrCliente()   As String   ' nombre del cliente
    Dim arrMonto()     As Double   ' suma de montos
    Dim arrConceptos() As String   ' conceptos concatenados con *
    Dim arrRows()      As String   ' filas involucradas "2,5,7"
    Dim arrVariante()  As String   ' variante de mayor urgencia
    Dim arrRFC()       As String   ' RFC del cliente
    Dim nGrupos As Long: nGrupos = 0

    Dim wsConf As Worksheet: Set wsConf = ObtenerHoja("CONFIGURACION")

    Dim i As Long
    For i = 2 To uFila
        Dim sCliente  As String: sCliente  = Trim(CStr(wsOp.Cells(i, COL_OP_CLIENTE).Value))
        If sCliente = "" Then GoTo SigEnv

        Dim sEstatus  As String: sEstatus  = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_ESTATUS).Value)))
        Dim sPago     As String: sPago     = Trim(CStr(wsOp.Cells(i, COL_OP_REG_PAGO).Value))
        Dim sExcluir  As String: sExcluir  = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_EXCLUIR).Value)))
        Dim sRFC      As String: sRFC      = Trim(CStr(wsOp.Cells(i, COL_OP_RFC).Value))
        Dim sTelRaw   As String: sTelRaw   = Trim(CStr(wsOp.Cells(i, COL_OP_TELEFONO).Value))
        Dim sBotWA    As String: sBotWA    = Trim(CStr(wsOp.Cells(i, COL_OP_WA).Value))

        ' Filtros de seguridad
        If sPago <> "" Then GoTo SigEnv
        If sExcluir = "SI" Or sExcluir = "S" & ChrW(237) Or sExcluir = "X" Then GoTo SigEnv
        If sEstatus = "PAGADO" Then GoTo SigEnv
        If sEstatus <> "PENDIENTE" And sEstatus <> "VENCIDO" And sEstatus <> "HOY VENCE" Then GoTo SigEnv
        If InStr(sBotWA, "REENVIAR") > 0 Then GoTo SigEnv
        If sTelRaw = "" Then GoTo SigEnv

        ' Verificar SUSPENDIDO
        If RFCSuspendido(sRFC) Then GoTo SigEnv

        ' Limpiar telefono
        Dim sTelLimpio As String
        If sModo = "PRUEBA" Then
            sTelLimpio = LimpiarTelefono(Trim(CStr(wsConf.Range("B14").Value)))
        Else
            sTelLimpio = LimpiarTelefono(sTelRaw)
        End If
        If Len(sTelLimpio) < 10 Then GoTo SigEnv

        ' Determinar variante de esta fila
        Dim fVencim As Variant: fVencim = wsOp.Cells(i, COL_OP_VENCIMIENTO).Value
        Dim diasVenc As Long:   diasVenc = 0
        If IsDate(fVencim) Then diasVenc = DateDiff("d", CDate(fVencim), Date)
        Dim varFila As String
        If diasVenc > 0 Then
            varFila = "VENCIDO"
        ElseIf diasVenc = 0 Then
            varFila = "HOY_VENCE"
        Else
            varFila = "RECORDATORIO"
        End If

        ' Leer monto y concepto
        Dim dMonto    As Double: dMonto    = 0
        Dim sConcepto As String: sConcepto = Trim(CStr(wsOp.Cells(i, COL_OP_CONCEPTO).Value))
        On Error Resume Next
        dMonto = CDbl(wsOp.Cells(i, COL_OP_MONTO).Value)
        On Error GoTo 0

        ' Filtro: no enviar si monto es 0 o negativo
        If dMonto <= 0 Then GoTo SigEnv

        ' Buscar grupo existente con ese telefono
        Dim g As Long: g = -1
        Dim k As Long
        For k = 0 To nGrupos - 1
            If arrTel(k) = sTelLimpio Then g = k: Exit For
        Next k

        ' Prioridad: VENCIDO > HOY_VENCE > RECORDATORIO
        Dim prioridad As Integer
        If varFila = "VENCIDO" Then prioridad = 3
        ElseIf varFila = "HOY_VENCE" Then prioridad = 2
        Else prioridad = 1
        End If

        If g = -1 Then
            ' Nuevo grupo
            g = nGrupos

            ' Asegurar capacidad de arrays dinamicos (sin limite fijo)
            If nGrupos = 0 Then
                ReDim arrTel(0 To 0)
                ReDim arrCliente(0 To 0)
                ReDim arrMonto(0 To 0)
                ReDim arrConceptos(0 To 0)
                ReDim arrRows(0 To 0)
                ReDim arrVariante(0 To 0)
                ReDim arrRFC(0 To 0)
            Else
                ReDim Preserve arrTel(0 To nGrupos)
                ReDim Preserve arrCliente(0 To nGrupos)
                ReDim Preserve arrMonto(0 To nGrupos)
                ReDim Preserve arrConceptos(0 To nGrupos)
                ReDim Preserve arrRows(0 To nGrupos)
                ReDim Preserve arrVariante(0 To nGrupos)
                ReDim Preserve arrRFC(0 To nGrupos)
            End If

            arrTel(g)       = sTelLimpio
            arrCliente(g)   = sCliente
            arrMonto(g)     = dMonto
            arrConceptos(g) = ChrW(&H2022) & " *" & sConcepto & "*"
            arrRows(g)      = CStr(i)
            arrVariante(g)  = varFila
            arrRFC(g)       = sRFC
            nGrupos = nGrupos + 1
        Else
            ' Agregar al grupo existente
            arrMonto(g) = arrMonto(g) + dMonto
            arrConceptos(g) = arrConceptos(g) & Chr(10) & ChrW(&H2022) & " *" & sConcepto & "*"
            arrRows(g) = arrRows(g) & "," & CStr(i)
            ' Escalar variante si tiene mas urgencia
            Dim priActual As Integer
            If arrVariante(g) = "VENCIDO" Then priActual = 3
            ElseIf arrVariante(g) = "HOY_VENCE" Then priActual = 2
            Else priActual = 1
            End If
            If prioridad > priActual Then arrVariante(g) = varFila
        End If

SigEnv:
    Next i

    If nGrupos = 0 Then
        MsgBox "No hay mensajes pendientes de envio.", vbInformation, "BajaTax"
        Exit Sub
    End If

    Dim contEnviados As Long: contEnviados = 0
    Dim contErr      As Long: contErr      = 0

    ' -- Enviar un mensaje por grupo ----------------------------
    Dim g2 As Long
    For g2 = 0 To nGrupos - 1
        Dim sMsg As String
        Dim sMonto As String: sMonto = Format(arrMonto(g2), "$#,##0.00")

        ' Decidir entre mensaje simple o consolidado
        If InStr(arrRows(g2), ",") > 0 Then
            ' Multiples adeudos -> mensaje consolidado
            sMsg = ConstruirMensajeConsolidado(arrCliente(g2), sMonto, arrConceptos(g2))
        Else
            ' Un solo adeudo -> mensaje individual con variante
            Dim fV2 As Variant: fV2 = wsOp.Cells(CLng(arrRows(g2)), COL_OP_VENCIMIENTO).Value
            Dim dV2 As Long: dV2 = 0
            If IsDate(fV2) Then dV2 = DateDiff("d", CDate(fV2), Date)
            Dim sFech2 As String: sFech2 = IIf(IsDate(fV2), Format(CDate(fV2), "dd/mm/yyyy"), "s/f")
            Dim sCon2  As String: sCon2  = Trim(CStr(wsOp.Cells(CLng(arrRows(g2)), COL_OP_CONCEPTO).Value))
            sMsg = ConstruirMensaje(arrVariante(g2), arrCliente(g2), sMonto, sCon2, sFech2, dV2)
        End If

        ' Agregar variacion anti-deteccion (espacio invisible al final)
        sMsg = sMsg & ChrW(&H200B)

        ' Abrir WhatsApp
        Dim sURL As String
        sURL = "https://wa.me/" & arrTel(g2) & "?text=" & CodificarWhatsApp(sMsg)

        On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:=sURL
        If Err.Number <> 0 Then
            If Not EsMac() Then Shell "cmd /c start " & sURL, vbHide
            Err.Clear
        End If
        On Error GoTo 0

        ' Marcar todas las filas del grupo como enviadas
        Dim rowArr() As String: rowArr = Split(arrRows(g2), ",")
        Dim r As Integer
        For r = 0 To UBound(rowArr)
            Dim nRow As Long: nRow = CLng(rowArr(r))
            wsOp.Cells(nRow, COL_OP_WA).Value = "REENVIAR " & Format(Now, "dd/mm HH:mm")
            Select Case arrVariante(g2)
                Case "VENCIDO":      wsOp.Cells(nRow, COL_OP_WA).Interior.Color = RGB(255, 199, 206)
                Case "HOY_VENCE":    wsOp.Cells(nRow, COL_OP_WA).Interior.Color = RGB(255, 235, 156)
                Case "RECORDATORIO": wsOp.Cells(nRow, COL_OP_WA).Interior.Color = RGB(198, 224, 180)
            End Select
            Dim int2 As Long: int2 = 0
            On Error Resume Next: int2 = CLng(wsOp.Cells(nRow, COL_OP_INTENTOS).Value): On Error GoTo 0
            wsOp.Cells(nRow, COL_OP_INTENTOS).Value  = int2 + 1
            wsOp.Cells(nRow, COL_OP_ULT_ENVIO).Value = Format(Now, "dd/mm/yyyy hh:mm")
        Next r

        ' Registrar en LOG
        Call RegistrarLogEnvio("MASIVO", arrCliente(g2), arrVariante(g2), sMonto, _
                               "Consolidado", arrTel(g2), ModoSistema())

        contEnviados = contEnviados + 1

        ' Feedback en barra de estado
        Application.StatusBar = "BajaTax: " & contEnviados & " de " & nGrupos & " enviados..."

        ' Anti-ban: espera aleatoria 8-15 s
        If g2 < nGrupos - 1 Then
            Dim waitSecs As Long
            waitSecs = ANTI_BAN_MIN + Int((ANTI_BAN_MAX - ANTI_BAN_MIN + 1) * Rnd)
            Application.Wait Now + TimeSerial(0, 0, waitSecs)
        End If
    Next g2

    Application.StatusBar = False

    MsgBox "Envio masivo completado." & Chr(13) & Chr(13) & _
           "  Telefonos contactados: " & contEnviados & Chr(13) & _
           "  (grupos consolidados de " & uFila - 1 & " registros)", _
           vbInformation, "BajaTax -- Envio Masivo"
End Sub

' ---------------------------------------------------------------
'  EnvioMasivoAutomatico -- alias para compatibilidad
' ---------------------------------------------------------------
Public Sub EnvioMasivoAutomatico()
    Call EnvioMasivoConsolidado
End Sub

' ---------------------------------------------------------------
'  RestaurarBotonWA
' ---------------------------------------------------------------
Public Sub RestaurarBotonWA(ByVal NL As Long, ByVal diasVenc As Long)
    Dim wsOp As Worksheet: Set wsOp = ObtenerHoja("OPERACIONES")
    If wsOp Is Nothing Then Exit Sub

    Dim sTexto As String
    Dim colorFondo As Long

    If diasVenc = 0 Then
        sTexto = SimboloWA() & " HOY VENCE" & Chr(10) & "ENVIAR WA"
        colorFondo = RGB(255, 235, 156)
    ElseIf diasVenc > 0 Then
        sTexto = SimboloWA() & " VENCIDO" & Chr(10) & "ENVIAR WA"
        colorFondo = RGB(255, 199, 206)
    Else
        sTexto = SimboloWA() & " RECORDATORIO" & Chr(10) & "ENVIAR WA"
        colorFondo = RGB(198, 224, 180)
    End If

    With wsOp.Cells(NL, COL_OP_WA)
        .Value               = sTexto
        .Interior.Color      = colorFondo
        .HorizontalAlignment = xlCenter
        .Font.Bold           = True
    End With
End Sub

' ---------------------------------------------------------------
'  LimpiarTelefono -- normaliza a formato 52XXXXXXXXXX
' ---------------------------------------------------------------
Public Function LimpiarTelefono(sTel As String) As String
    Dim res As String: res = ""
    Dim i   As Integer
    For i = 1 To Len(sTel)
        Dim c As String: c = Mid(sTel, i, 1)
        If c >= "0" And c <= "9" Then res = res & c
    Next i
    If Len(res) = 10 Then res = "52" & res
    LimpiarTelefono = res
End Function

' ---------------------------------------------------------------
'  CodificarWhatsApp -- UTF-8 URL encoding completo con %0A para \n
' ---------------------------------------------------------------
Public Function CodificarWhatsApp(sTexto As String) As String
    Dim i As Integer
    Dim resultado As String: resultado = ""

    For i = 1 To Len(sTexto)
        Dim c  As String: c  = Mid(sTexto, i, 1)
        Dim cw As Long:   cw = AscW(c)

        Select Case cw
            ' Alfanumerico y seguros
            Case 48 To 57:   resultado = resultado & c
            Case 65 To 90:   resultado = resultado & c
            Case 97 To 122:  resultado = resultado & c
            Case 45, 95, 46, 126: resultado = resultado & c
            Case 42: resultado = resultado & c             ' * para bold
            ' Espacios y saltos
            Case 32: resultado = resultado & "%20"
            Case 10: resultado = resultado & "%0A"         ' salto de linea
            Case 13: ' ignorar CR
            ' Vocales acentuadas minusculas
            Case 225: resultado = resultado & "%C3%A1"
            Case 233: resultado = resultado & "%C3%A9"
            Case 237: resultado = resultado & "%C3%AD"
            Case 243: resultado = resultado & "%C3%B3"
            Case 250: resultado = resultado & "%C3%BA"
            ' Vocales acentuadas mayusculas
            Case 193: resultado = resultado & "%C3%81"
            Case 201: resultado = resultado & "%C3%89"
            Case 205: resultado = resultado & "%C3%8D"
            Case 211: resultado = resultado & "%C3%93"
            Case 218: resultado = resultado & "%C3%9A"
            ' N / n
            Case 241: resultado = resultado & "%C3%B1"
            Case 209: resultado = resultado & "%C3%91"
            ' Signos especiales
            Case 161: resultado = resultado & "%C2%A1"
            Case 191: resultado = resultado & "%C2%BF"
            Case 252: resultado = resultado & "%C3%BC"
            Case 220: resultado = resultado & "%C3%9C"
            ' Puntuacion
            Case 38:  resultado = resultado & "%26"
            Case 43:  resultado = resultado & "%2B"
            Case 61:  resultado = resultado & "%3D"
            Case 63:  resultado = resultado & "%3F"
            Case 35:  resultado = resultado & "%23"
            Case 37:  resultado = resultado & "%25"
            Case 47:  resultado = resultado & "%2F"
            Case 58:  resultado = resultado & "%3A"
            Case 64:  resultado = resultado & "%40"
            ' Bullet point y similares
            Case 8226: resultado = resultado & "%E2%80%A2"  ' bullet
            Case 8203: resultado = resultado & "%E2%80%8B"  ' zero-width space (variacion anti-ban)
            Case 33 To 127:
                resultado = resultado & "%" & UCase(Hex(cw))
            Case Else
                ' UTF-8 de 2 bytes (128-2047)
                If cw >= 128 And cw <= 2047 Then
                    Dim b1 As Long: b1 = &HC0 Or (cw \ 64)
                    Dim b2 As Long: b2 = &H80 Or (cw Mod 64)
                    resultado = resultado & "%" & UCase(Hex(b1)) & "%" & UCase(Hex(b2))
                ' UTF-8 de 3 bytes (2048-65535)
                ElseIf cw >= 2048 And cw <= 65535 Then
                    Dim b1t As Long: b1t = &HE0 Or (cw \ 4096)
                    Dim b2t As Long: b2t = &H80 Or ((cw \ 64) Mod 64)
                    Dim b3t As Long: b3t = &H80 Or (cw Mod 64)
                    resultado = resultado & "%" & UCase(Hex(b1t)) & "%" & _
                                UCase(Hex(b2t)) & "%" & UCase(Hex(b3t))
                Else
                    resultado = resultado & c
                End If
        End Select
    Next i

    CodificarWhatsApp = resultado
End Function

' ---------------------------------------------------------------
'  RegistrarLogEnvio
' ---------------------------------------------------------------
Private Sub RegistrarLogEnvio(sResp As String, sCliente As String, _
                               sVariante As String, sMonto As String, _
                               sConcepto As String, sTelDestino As String, _
                               sModo As String)
    Dim wsLog As Worksheet
    Set wsLog = ObtenerHoja("LOG ENVIOS")
    If wsLog Is Nothing Then Set wsLog = ObtenerHoja("LOG ENV" & ChrW(205) & "OS")
    If wsLog Is Nothing Then Exit Sub

    Dim uFila As Long
    uFila = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    With wsLog
        .Cells(uFila, 1).Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
        .Cells(uFila, 2).Value = sResp
        .Cells(uFila, 3).Value = sCliente
        .Cells(uFila, 4).Value = sTelDestino
        .Cells(uFila, 5).Value = sVariante
        .Cells(uFila, 6).Value = sModo
        .Cells(uFila, 7).Value = "ENVIADO"
    End With
End Sub
