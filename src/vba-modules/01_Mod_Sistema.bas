Attribute VB_Name = "Mod_Sistema"
'================================================================
' MODULO: Mod_Sistema  -- BajaTax v4 FINAL
' Instalar PRIMERO. Contiene todas las constantes y helpers.
'================================================================
Option Explicit

' ---------------------------------------------------------------
'  COLUMNAS -- OPERACIONES (A=1 ... T=20)
' ---------------------------------------------------------------
Public Const COL_OP_RESPONSABLE  As Integer = 1   ' A
Public Const COL_OP_ID_FACTURA   As Integer = 2   ' B
Public Const COL_OP_REGIMEN      As Integer = 3   ' C
Public Const COL_OP_CLIENTE      As Integer = 4   ' D
Public Const COL_OP_RFC          As Integer = 5   ' E
Public Const COL_OP_FECHA_COB    As Integer = 6   ' F
Public Const COL_OP_CONCEPTO     As Integer = 7   ' G
Public Const COL_OP_MONTO        As Integer = 8   ' H
Public Const COL_OP_ESTATUS      As Integer = 9   ' I
Public Const COL_OP_VENCIMIENTO  As Integer = 10  ' J
Public Const COL_OP_DIAS_VENC    As Integer = 11  ' K
Public Const COL_OP_REG_PAGO     As Integer = 12  ' L
Public Const COL_OP_TELEFONO     As Integer = 13  ' M
Public Const COL_OP_CORREO       As Integer = 14  ' N
Public Const COL_OP_WA           As Integer = 15  ' O
Public Const COL_OP_PDF          As Integer = 16  ' P
Public Const COL_OP_EXCLUIR      As Integer = 17  ' Q
Public Const COL_OP_PROX_ENVIO   As Integer = 18  ' R
Public Const COL_OP_INTENTOS     As Integer = 19  ' S
Public Const COL_OP_ULT_ENVIO    As Integer = 20  ' T

' ---------------------------------------------------------------
'  COLUMNAS -- REGISTROS
' ---------------------------------------------------------------
Public Const COL_REG_RESPONSABLE  As Integer = 1   ' A
Public Const COL_REG_NOMBRE       As Integer = 2   ' B
Public Const COL_REG_RFC          As Integer = 3   ' C
Public Const COL_REG_EMAIL        As Integer = 4   ' D
Public Const COL_REG_TELEFONO     As Integer = 5   ' E
Public Const COL_REG_FECHA        As Integer = 6   ' F
Public Const COL_REG_CONCEPTO     As Integer = 7   ' G
Public Const COL_REG_MONTO        As Integer = 8   ' H
Public Const COL_REG_FACTURA      As Integer = 9   ' I
Public Const COL_REG_REGIMEN      As Integer = 10  ' J
Public Const COL_REG_VENCIMIENTO  As Integer = 11  ' K
Public Const COL_REG_IND_OPER     As Integer = 12  ' L
Public Const COL_REG_IND_DIR      As Integer = 13  ' M
Public Const COL_REG_PROCESADO    As Integer = 14  ' N

' ---------------------------------------------------------------
'  COLUMNAS -- DIRECTORIO v4 (A-I)
' ---------------------------------------------------------------
Public Const COL_DIR_RFC           As Integer = 1  ' A
Public Const COL_DIR_CLIENTE       As Integer = 2  ' B  (antes NOMBRE)
Public Const COL_DIR_CORREO        As Integer = 3  ' C  (antes EMAIL)
Public Const COL_DIR_NUMERO        As Integer = 4  ' D  (antes TELEFONO)
Public Const COL_DIR_REGIMEN       As Integer = 5  ' E
Public Const COL_DIR_RESPONSABLE   As Integer = 6  ' F  NUEVO
Public Const COL_DIR_CLASIFICACION As Integer = 7  ' G  NUEVO
Public Const COL_DIR_FECHA_ALTA    As Integer = 8  ' H  NUEVO
Public Const COL_DIR_ESTADO        As Integer = 9  ' I  NUEVO (ACTIVO / SUSPENDIDO)

' ---------------------------------------------------------------
'  ANTI-BANEO: intervalo en segundos entre envios WA
' ---------------------------------------------------------------
Public Const ANTI_BAN_MIN As Integer = 8
Public Const ANTI_BAN_MAX As Integer = 15

' ---------------------------------------------------------------
'  FILA INICIO RESULTADOS -- BUSCADOR CLIENTE
' ---------------------------------------------------------------
Public Const BUSC_FILA_HEADERS  As Long = 6  ' fila de encabezados
Public Const BUSC_FILA_DATOS    As Long = 7  ' primera fila de datos

' ---------------------------------------------------------------
'  Simbolos Unicode BMP (compatibles Mac/Win)
' ---------------------------------------------------------------
Public Function SimboloWA() As String
    SimboloWA = ChrW(&H25B6)        ' triangulo
End Function
Public Function SimboloPDF() As String
    SimboloPDF = ChrW(&H25A0)       ' cuadrado
End Function
Public Function SimboloCheck() As String
    SimboloCheck = ChrW(&H2713)     ' check
End Function
Public Function SimboloCheckBold() As String
    SimboloCheckBold = ChrW(&H2714) ' check bold
End Function

' ---------------------------------------------------------------
'  Utilidades generales
' ---------------------------------------------------------------
Public Function EsMac() As Boolean
    EsMac = (Application.OperatingSystem Like "Mac*")
End Function

Public Function ObtenerHoja(sNombre As String) As Worksheet
    On Error Resume Next
    Set ObtenerHoja = ThisWorkbook.Sheets(sNombre)
    On Error GoTo 0
End Function

Public Function HojasOK() As Boolean
    If ObtenerHoja("OPERACIONES") Is Nothing Or _
       ObtenerHoja("CONFIGURACION") Is Nothing Then
        MsgBox "Faltan hojas requeridas (OPERACIONES, CONFIGURACION).", _
               vbCritical, "BajaTax"
        HojasOK = False
    Else
        HojasOK = True
    End If
End Function

Public Function LeerConfig(sCelda As String) As String
    Dim ws As Worksheet
    Set ws = ObtenerHoja("CONFIGURACION")
    If ws Is Nothing Then LeerConfig = "": Exit Function
    LeerConfig = Trim(CStr(ws.Range(sCelda).Value))
End Function

Public Function ModoSistema() As String
    ModoSistema = UCase(LeerConfig("B2"))
End Function

' ---------------------------------------------------------------
'  RFCSuspendido -- devuelve True si el RFC esta SUSPENDIDO
' ---------------------------------------------------------------
Public Function RFCSuspendido(sRFC As String) As Boolean
    RFCSuspendido = False
    If Trim(sRFC) = "" Then Exit Function

    Dim wsDir As Worksheet
    Set wsDir = ObtenerHoja("DIRECTORIO")
    If wsDir Is Nothing Then Exit Function

    Dim uFila As Long
    uFila = wsDir.Cells(wsDir.Rows.Count, COL_DIR_RFC).End(xlUp).Row

    Dim i As Long
    For i = 2 To uFila
        If UCase(Trim(CStr(wsDir.Cells(i, COL_DIR_RFC).Value))) = UCase(Trim(sRFC)) Then
            If UCase(Trim(CStr(wsDir.Cells(i, COL_DIR_ESTADO).Value))) = "SUSPENDIDO" Then
                RFCSuspendido = True
            End If
            Exit For
        End If
    Next i
End Function

' ---------------------------------------------------------------
'  ConstruirMensaje -- mensaje individual con plantillas v4
'  Usa variables desde CONFIGURACION, formato WhatsApp (*bold*)
' ---------------------------------------------------------------
Public Function ConstruirMensaje(sVariante As String, _
                                  sCliente As String, _
                                  sMonto As String, _
                                  sConcepto As String, _
                                  sFecha As String, _
                                  diasVenc As Long) As String
    Dim sDesp  As String: sDesp  = LeerConfig("B5")
    Dim sBene  As String: sBene  = LeerConfig("B6")
    Dim sBanco As String: sBanco = LeerConfig("B7")
    Dim sCLABE As String: sCLABE = LeerConfig("B8")
    Dim sTel   As String: sTel   = LeerConfig("B9")
    Dim sEmail As String: sEmail = LeerConfig("B10")
    Dim sDepto As String: sDepto = LeerConfig("B12")

    Dim nl As String: nl = Chr(10)
    Dim msg As String

    Select Case sVariante

        Case "VENCIDO"
            Dim sDiasV As String: sDiasV = CStr(diasVenc)
            msg = sDesp & " - Recordatorio de Pago Vencido" & nl & _
                  "Estimado *" & sCliente & "*," & nl & _
                  "Su cuenta presenta un saldo vencido de *" & sMonto & _
                  "* correspondiente a: *" & sConcepto & "*" & nl & _
                  "Fecha de vencimiento: *" & sFecha & "* (*" & sDiasV & "* d" & ChrW(237) & "as de retraso)" & nl & _
                  "Le pedimos regularizar su situaci" & ChrW(243) & "n a la brevedad " & _
                  "para evitar la suspensi" & ChrW(243) & "n de servicios." & nl & _
                  "Apreciamos su pronto pago:" & nl & _
                  "*Datos para Transferencia:*" & nl & _
                  "*Beneficiario:* " & sBene & nl & _
                  "*Banco:* " & sBanco & nl & _
                  "*CLABE:* " & sCLABE & nl & _
                  "Cualquier duda estamos a sus ordenes." & nl & _
                  "*" & sDepto & "* | " & sTel & " | " & sEmail

        Case "HOY_VENCE"
            msg = sDesp & " - Vencimiento Hoy" & nl & _
                  "Estimado *" & sCliente & "*," & nl & _
                  "Le recordamos que hoy *" & sFecha & "* es la fecha l" & ChrW(237) & "mite para realizar su pago." & nl & _
                  "Saldo pendiente: *" & sMonto & "*" & nl & _
                  "Concepto: *" & sConcepto & "*" & nl & _
                  "Evite recargos realizando su pago el d" & ChrW(237) & "a de hoy. Apreciamos su puntualidad:" & nl & _
                  "*Datos para Transferencia:*" & nl & _
                  "*Beneficiario:* " & sBene & nl & _
                  "*Banco:* " & sBanco & nl & _
                  "*CLABE:* " & sCLABE & nl & _
                  "Cualquier duda estamos a sus ordenes." & nl & _
                  "*" & sDepto & "* | " & sTel

        Case "RECORDATORIO"
            Dim diasFaltan As Long: diasFaltan = Abs(diasVenc)
            msg = sDesp & " - Pr" & ChrW(243) & "ximo Vencimiento" & nl & _
                  "Estimado *" & sCliente & "*," & nl & _
                  "Le recordamos que el pr" & ChrW(243) & "ximo *" & sFecha & _
                  "* es la fecha l" & ChrW(237) & "mite para realizar su pago." & nl & _
                  "Saldo pendiente: *" & sMonto & "*" & nl & _
                  "Concepto: *" & sConcepto & "* (*" & CStr(diasFaltan) & "* d" & ChrW(237) & "as restantes)" & nl & _
                  "Agradecemos de antemano su gesti" & ChrW(243) & "n." & nl & _
                  "*Beneficiario:* " & sBene & nl & _
                  "*Banco:* " & sBanco & nl & _
                  "*CLABE:* " & sCLABE & nl & _
                  "*" & sDepto & "* | " & sTel

        Case Else
            msg = "Estimado " & sCliente & ", tiene un saldo de " & sMonto & ". Concepto: " & sConcepto

    End Select

    ' Normalizar saltos de linea
    msg = Replace(msg, Chr(13) & Chr(10), Chr(10))
    msg = Replace(msg, Chr(13), Chr(10))

    ConstruirMensaje = msg
End Function

' ---------------------------------------------------------------
'  ConstruirMensajeConsolidado
'  Para envio masivo cuando hay multiples adeudos en un mismo telefono
' ---------------------------------------------------------------
Public Function ConstruirMensajeConsolidado(sCliente As String, _
                                              sMontoTotal As String, _
                                              sConceptos As String) As String
    Dim sDesp  As String: sDesp  = LeerConfig("B5")
    Dim sBano  As String: sBano  = LeerConfig("B7")
    Dim sCLABE As String: sCLABE = LeerConfig("B8")
    Dim sTel   As String: sTel   = LeerConfig("B9")
    Dim sDepto As String: sDepto = LeerConfig("B12")

    Dim nl As String: nl = Chr(10)
    Dim msg As String

    msg = sDesp & " - Recordatorio de Saldo Pendiente" & nl & _
          "Estimado *" & sCliente & "*," & nl & _
          "Su cuenta presenta un saldo pendiente por la suma de *" & sMontoTotal & _
          "* correspondiente a los siguientes conceptos:" & nl & _
          sConceptos & nl & _
          "Le pedimos regularizar su situaci" & ChrW(243) & "n a la brevedad." & nl & _
          "*Datos para Transferencia:*" & nl & _
          "*Banco:* " & sBano & " | *CLABE:* " & sCLABE & nl & _
          sDepto & " | " & sTel

    msg = Replace(msg, Chr(13) & Chr(10), Chr(10))
    msg = Replace(msg, Chr(13), Chr(10))

    ConstruirMensajeConsolidado = msg
End Function

' ---------------------------------------------------------------
'  InicializarBotonFila -- botones WA y PDF en OPERACIONES
' ---------------------------------------------------------------
Public Sub InicializarBotonFila(wsOp As Worksheet, ByVal NL As Long)
    Dim fVencim  As Variant
    Dim diasVenc As Long
    fVencim = wsOp.Cells(NL, COL_OP_VENCIMIENTO).Value
    diasVenc = 0
    If IsDate(fVencim) Then
        diasVenc = DateDiff("d", CDate(fVencim), Date)
    End If

    Dim sTextoWA  As String
    Dim colorBoton As Long

    If diasVenc = 0 Then
        sTextoWA   = SimboloWA() & " HOY VENCE" & Chr(10) & "ENVIAR WA"
        colorBoton = RGB(255, 235, 156)
    ElseIf diasVenc > 0 Then
        sTextoWA   = SimboloWA() & " VENCIDO" & Chr(10) & "ENVIAR WA"
        colorBoton = RGB(255, 199, 206)
    Else
        sTextoWA   = SimboloWA() & " RECORDATORIO" & Chr(10) & "ENVIAR WA"
        colorBoton = RGB(198, 224, 180)
    End If

    With wsOp.Cells(NL, COL_OP_WA)
        .Value              = sTextoWA
        .Interior.Color     = colorBoton
        .HorizontalAlignment = xlCenter
        .VerticalAlignment  = xlCenter
        .WrapText           = True
        .Font.Bold          = True
        .Font.Size          = 9
    End With

    With wsOp.Cells(NL, COL_OP_PDF)
        .Value              = SimboloPDF() & " GENERAR PDF"
        .Interior.Color     = RGB(189, 215, 238)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment  = xlCenter
        .WrapText           = True
        .Font.Bold          = True
        .Font.Size          = 9
    End With
End Sub

' ---------------------------------------------------------------
'  InicializarBotones -- batch sobre toda la hoja OPERACIONES
' ---------------------------------------------------------------
Public Function InicializarBotones(wsOp As Worksheet) As Integer
    Dim uFila As Long
    uFila = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row

    Dim cont As Integer: cont = 0
    Dim i As Long
    For i = 2 To uFila
        Dim sCliente As String
        sCliente = Trim(CStr(wsOp.Cells(i, COL_OP_CLIENTE).Value))
        If sCliente = "" Then GoTo SigFila

        Dim sPago As String
        sPago = Trim(CStr(wsOp.Cells(i, COL_OP_REG_PAGO).Value))
        If sPago <> "" Then GoTo SigFila

        Dim sBotWA As String
        sBotWA = Trim(CStr(wsOp.Cells(i, COL_OP_WA).Value))
        If InStr(sBotWA, "REENVIAR") = 0 Then
            Call InicializarBotonFila(wsOp, i)
        End If
        cont = cont + 1
SigFila:
    Next i

    InicializarBotones = cont
End Function

' ---------------------------------------------------------------
'  ProcesarTodoBajaTax -- REGISTROS -> OPERACIONES + DIRECTORIO
' ---------------------------------------------------------------
Public Sub ProcesarTodoBajaTax()
    If Not HojasOK() Then Exit Sub

    Dim wsReg As Worksheet: Set wsReg = ObtenerHoja("REGISTROS")
    Dim wsDir As Worksheet: Set wsDir = ObtenerHoja("DIRECTORIO")
    Dim wsOp  As Worksheet: Set wsOp  = ObtenerHoja("OPERACIONES")

    If wsReg Is Nothing Or wsDir Is Nothing Or wsOp Is Nothing Then
        MsgBox "No se encuentran hojas REGISTROS / DIRECTORIO / OPERACIONES.", _
               vbCritical, "BajaTax"
        Exit Sub
    End If

    Dim resp As Integer
    resp = MsgBox("Se procesaran los registros pendientes." & Chr(13) & Chr(13) & _
                  "Los duplicados (mismo concepto+monto) pediran confirmacion." & Chr(13) & _
                  ChrW(&H26A0) & " Deseas continuar?", _
                  vbYesNo + vbQuestion, "BajaTax - Procesar Registros")
    If resp = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    On Error GoTo ErrorHandler

    Application.Calculate

    Dim contBotones   As Integer: contBotones   = InicializarBotones(wsOp)
    Dim contNuevosDir As Integer: contNuevosDir = 0
    Dim contNuevosOp  As Integer: contNuevosOp  = 0
    Dim contDupl      As Integer: contDupl      = 0
    Dim contOmitidos  As Integer: contOmitidos  = 0

    Dim uFilaReg As Long
    uFilaReg = wsReg.Cells(wsReg.Rows.Count, COL_REG_NOMBRE).End(xlUp).Row

    Dim i As Long
    For i = 2 To uFilaReg
        Dim sProcesado As String
        sProcesado = Trim(CStr(wsReg.Cells(i, COL_REG_PROCESADO).Value))
        If InStr(UCase(sProcesado), "PROCESADO") > 0 Then GoTo SigReg

        Dim sNombre   As String: sNombre   = Trim(CStr(wsReg.Cells(i, COL_REG_NOMBRE).Value))
        Dim sRFC      As String: sRFC      = Trim(CStr(wsReg.Cells(i, COL_REG_RFC).Value))
        Dim sEmail    As String: sEmail    = Trim(CStr(wsReg.Cells(i, COL_REG_EMAIL).Value))
        Dim sTel      As String: sTel      = Trim(CStr(wsReg.Cells(i, COL_REG_TELEFONO).Value))
        Dim sRegimen  As String: sRegimen  = Trim(CStr(wsReg.Cells(i, COL_REG_REGIMEN).Value))
        Dim sConcepto As String: sConcepto = Trim(CStr(wsReg.Cells(i, COL_REG_CONCEPTO).Value))
        Dim sFactura  As String: sFactura  = Trim(CStr(wsReg.Cells(i, COL_REG_FACTURA).Value))
        Dim sResp     As String: sResp     = Trim(CStr(wsReg.Cells(i, COL_REG_RESPONSABLE).Value))
        Dim fFecha    As Variant: fFecha   = wsReg.Cells(i, COL_REG_FECHA).Value
        Dim fVencim   As Variant: fVencim  = wsReg.Cells(i, COL_REG_VENCIMIENTO).Value
        Dim dMonto    As Double: dMonto    = 0
        On Error Resume Next
        dMonto = CDbl(wsReg.Cells(i, COL_REG_MONTO).Value)
        On Error GoTo ErrorHandler

        If sNombre = "" Then GoTo SigReg

        ' Verificar duplicado en OPERACIONES
        Dim hayDup As Boolean: hayDup = False
        Dim uFilaOp As Long
        uFilaOp = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row

        Dim kOp As Long
        For kOp = 2 To uFilaOp
            Dim sNOp  As String: sNOp  = Trim(CStr(wsOp.Cells(kOp, COL_OP_CLIENTE).Value))
            Dim sCOp  As String: sCOp  = Trim(CStr(wsOp.Cells(kOp, COL_OP_CONCEPTO).Value))
            Dim dMOp  As Double: dMOp  = 0
            On Error Resume Next
            dMOp = CDbl(wsOp.Cells(kOp, COL_OP_MONTO).Value)
            On Error GoTo ErrorHandler
            If UCase(sNOp) = UCase(sNombre) And _
               UCase(sCOp) = UCase(sConcepto) And _
               Abs(dMOp - dMonto) < 0.01 Then
                hayDup = True: Exit For
            End If
        Next kOp

        If hayDup Then
            Application.ScreenUpdating = True
            Dim rDup As Integer
            rDup = MsgBox("POSIBLE DUPLICADO (Fila " & i & "):" & Chr(13) & _
                          "  " & sNombre & " / " & sConcepto & " / " & Format(dMonto, "$#,##0.00") & Chr(13) & _
                          "SI=Agregar  NO=Omitir", _
                          vbYesNo + vbExclamation, "BajaTax - Duplicado")
            Application.ScreenUpdating = False
            If rDup = vbNo Then
                contOmitidos = contOmitidos + 1
                wsReg.Cells(i, COL_REG_PROCESADO).Value = "OMITIDO"
                wsReg.Cells(i, COL_REG_PROCESADO).Font.Color = RGB(156, 101, 0)
                wsReg.Cells(i, COL_REG_PROCESADO).Interior.Color = RGB(255, 235, 156)
                GoTo ActualizarDir
            Else
                contDupl = contDupl + 1
            End If
        End If

        ' Agregar a OPERACIONES
        Dim uNOp As Long
        uNOp = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row + 1

        wsOp.Cells(uNOp, COL_OP_RESPONSABLE).Value = sResp
        wsOp.Cells(uNOp, COL_OP_ID_FACTURA).Value  = sFactura
        wsOp.Cells(uNOp, COL_OP_REGIMEN).Value     = sRegimen
        wsOp.Cells(uNOp, COL_OP_CLIENTE).Value     = sNombre
        wsOp.Cells(uNOp, COL_OP_RFC).Value         = sRFC
        If IsDate(fFecha) Then wsOp.Cells(uNOp, COL_OP_FECHA_COB).Value = CDate(fFecha)
        wsOp.Cells(uNOp, COL_OP_CONCEPTO).Value    = sConcepto
        wsOp.Cells(uNOp, COL_OP_MONTO).Value       = dMonto
        If IsDate(fVencim) Then wsOp.Cells(uNOp, COL_OP_VENCIMIENTO).Value = CDate(fVencim)
        wsOp.Cells(uNOp, COL_OP_TELEFONO).Value    = sTel
        wsOp.Cells(uNOp, COL_OP_CORREO).Value      = sEmail

        wsOp.Cells(uNOp, COL_OP_ESTATUS).Formula = _
            "=IF(D" & uNOp & "="""","""",IF(L" & uNOp & "<>"""",""PAGADO""," & _
            "IF(J" & uNOp & "="""",""PENDIENTE"",IF(TODAY()>J" & uNOp & _
            ",""VENCIDO"",IF(TODAY()=J" & uNOp & ",""HOY VENCE"",""PENDIENTE"")))))"
        wsOp.Cells(uNOp, COL_OP_DIAS_VENC).Formula = _
            "=IFERROR(IF(J" & uNOp & "="""","""",TODAY()-J" & uNOp & "),"""")"

        Call InicializarBotonFila(wsOp, uNOp)
        contNuevosOp = contNuevosOp + 1

        wsReg.Cells(i, COL_REG_IND_OPER).Value      = SimboloCheck() & " OPERACIONES"
        wsReg.Cells(i, COL_REG_IND_OPER).Font.Color = RGB(0, 70, 127)
        wsReg.Cells(i, COL_REG_PROCESADO).Value             = SimboloCheck() & " PROCESADO"
        wsReg.Cells(i, COL_REG_PROCESADO).Font.Color        = RGB(0, 97, 0)
        wsReg.Cells(i, COL_REG_PROCESADO).Interior.Color    = RGB(198, 239, 206)

ActualizarDir:
        ' Verificar / Actualizar DIRECTORIO
        Dim existeDir As Boolean: existeDir = False
        Dim uFilaDir  As Long
        uFilaDir = wsDir.Cells(wsDir.Rows.Count, COL_DIR_RFC).End(xlUp).Row

        Dim jDir As Long
        For jDir = 2 To uFilaDir
            Dim sDirRFC As String: sDirRFC = Trim(CStr(wsDir.Cells(jDir, COL_DIR_RFC).Value))
            If sRFC <> "" And UCase(sDirRFC) = UCase(sRFC) Then
                existeDir = True
                If Trim(CStr(wsDir.Cells(jDir, COL_DIR_CORREO).Value)) = "" And sEmail <> "" Then
                    wsDir.Cells(jDir, COL_DIR_CORREO).Value = sEmail
                End If
                If Trim(CStr(wsDir.Cells(jDir, COL_DIR_NUMERO).Value)) = "" And sTel <> "" Then
                    wsDir.Cells(jDir, COL_DIR_NUMERO).Value = sTel
                End If
                Exit For
            ElseIf sRFC = "" And UCase(Trim(CStr(wsDir.Cells(jDir, COL_DIR_CLIENTE).Value))) = UCase(sNombre) Then
                existeDir = True: Exit For
            End If
        Next jDir

        If Not existeDir Then
            Dim uNDir As Long
            uNDir = wsDir.Cells(wsDir.Rows.Count, COL_DIR_RFC).End(xlUp).Row + 1
            wsDir.Cells(uNDir, COL_DIR_RFC).Value          = sRFC
            wsDir.Cells(uNDir, COL_DIR_CLIENTE).Value      = sNombre
            wsDir.Cells(uNDir, COL_DIR_CORREO).Value       = sEmail
            wsDir.Cells(uNDir, COL_DIR_NUMERO).Value       = sTel
            wsDir.Cells(uNDir, COL_DIR_REGIMEN).Value      = sRegimen
            wsDir.Cells(uNDir, COL_DIR_RESPONSABLE).Value  = sResp
            wsDir.Cells(uNDir, COL_DIR_ESTADO).Value       = "ACTIVO"
            wsDir.Cells(uNDir, COL_DIR_FECHA_ALTA).Value   = Format(Date, "dd/mm/yyyy")
            wsDir.Range(wsDir.Cells(uNDir, 1), wsDir.Cells(uNDir, 9)).Interior.Color = RGB(198, 239, 206)
            contNuevosDir = contNuevosDir + 1
        End If

        wsReg.Cells(i, COL_REG_IND_DIR).Value      = SimboloCheck() & " DIRECTORIO"
        wsReg.Cells(i, COL_REG_IND_DIR).Font.Color = RGB(0, 97, 0)

SigReg:
    Next i

    contBotones = InicializarBotones(wsOp)

    Application.ScreenUpdating = True
    Application.EnableEvents   = True

    MsgBox "Sincronizacion completada." & Chr(13) & Chr(13) & _
           "  Nuevos en OPERACIONES: " & contNuevosOp & Chr(13) & _
           "  Duplicados confirmados: " & contDupl & Chr(13) & _
           "  Omitidos:               " & contOmitidos & Chr(13) & _
           "  Nuevos en DIRECTORIO:   " & contNuevosDir & Chr(13) & _
           "  Botones inicializados:  " & contBotones, _
           vbInformation, "BajaTax"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    MsgBox "Error: " & Err.Description & " (#" & Err.Number & ")", _
           vbCritical, "BajaTax"
End Sub

' ---------------------------------------------------------------
'  RegenerarFaltantes -- mantenimiento
' ---------------------------------------------------------------
Public Sub RegenerarFaltantes()
    If Not HojasOK() Then Exit Sub
    Dim wsReg As Worksheet: Set wsReg = ObtenerHoja("REGISTROS")
    If wsReg Is Nothing Then Exit Sub

    Dim uFila  As Long: uFila = wsReg.Cells(wsReg.Rows.Count, COL_REG_NOMBRE).End(xlUp).Row
    Dim lista  As String: lista  = ""
    Dim total  As Integer: total = 0

    Dim i As Long
    For i = 2 To uFila
        Dim sN As String: sN = Trim(CStr(wsReg.Cells(i, COL_REG_NOMBRE).Value))
        If sN = "" Then GoTo SigF
        If InStr(CStr(wsReg.Cells(i, COL_REG_IND_DIR).Value), "DIRECTORIO") = 0 Then
            total = total + 1
            lista = lista & "  - " & sN & Chr(13)
        End If
SigF:
    Next i

    If total = 0 Then
        MsgBox "Todo sincronizado. No hay clientes faltantes.", vbInformation, "BajaTax"
        Exit Sub
    End If

    Dim r As Integer
    r = MsgBox(total & " cliente(s) sin entrada en DIRECTORIO:" & Chr(13) & Chr(13) & _
               lista & Chr(13) & "Agregar ahora?", vbYesNo + vbQuestion, "BajaTax")
    If r = vbYes Then Call ProcesarTodoBajaTax
End Sub
