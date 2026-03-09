Attribute VB_Name = "Mod_BuscadorCliente"
'================================================================
' MODULO: Mod_BuscadorCliente -- BajaTax v4 FINAL
'
' Layout hoja BUSCADOR CLIENTE (v4):
'   Fila 3 -- Filtros:
'     B3 = Responsable (dropdown)       C3 = Regimen (dropdown)
'     D3 = Estatus (dropdown)           E3 = Cliente (texto)
'     F3 = RFC (texto)                  G3 = Concepto (texto)
'     H3 = BUSCAR (cell-btn, doble-clic)
'     I3 = Ordenar (dropdown)           J3 = Direccion (dropdown)
'     K3 = LIMPIAR (cell-btn, doble-clic)
'
'   Fila 6 -- Encabezados resultados (fijos):
'     A=No | B=Cliente | C=Responsable | D=RFC | E=Regimen |
'     F=Concepto | G=Monto | H=Vencimiento | I=Estatus | J=WA | K=PDF
'
'   Fila 7+ -- Resultados dinamicos
'================================================================
Option Explicit

' Constantes de celdas de filtros
Private Const CEL_RESP    As String = "B3"
Private Const CEL_REG     As String = "C3"
Private Const CEL_EST     As String = "D3"
Private Const CEL_CLI     As String = "E3"
Private Const CEL_RFC     As String = "F3"
Private Const CEL_CONC    As String = "G3"
Private Const CEL_BUSCAR  As String = "H3"
Private Const CEL_ORDEN   As String = "I3"
Private Const CEL_DIR     As String = "J3"
Private Const CEL_LIMPIAR As String = "K3"

' Columnas de resultados
Private Const RES_NO      As Integer = 1   ' A
Private Const RES_CLI     As Integer = 2   ' B
Private Const RES_RESP    As Integer = 3   ' C
Private Const RES_RFC     As Integer = 4   ' D
Private Const RES_REG     As Integer = 5   ' E
Private Const RES_CONC    As Integer = 6   ' F
Private Const RES_MONTO   As Integer = 7   ' G
Private Const RES_VENC    As Integer = 8   ' H
Private Const RES_EST     As Integer = 9   ' I
Private Const RES_WA      As Integer = 10  ' J
Private Const RES_PDF     As Integer = 11  ' K

'================================================================
'  EjecutarBusqueda -- lee filtros, filtra OPERACIONES, escribe resultados
'================================================================
Public Sub EjecutarBusqueda()
    Dim wsBusc As Worksheet: Set wsBusc = ObtenerHoja("BUSCADOR CLIENTE")
    Dim wsOp   As Worksheet: Set wsOp   = ObtenerHoja("OPERACIONES")
    If wsBusc Is Nothing Or wsOp Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    On Error GoTo FinBusq

    ' --- Leer filtros ------------------------------------------
    Dim sFResp  As String: sFResp  = UCase(Trim(CStr(wsBusc.Range(CEL_RESP).Value)))
    Dim sFReg   As String: sFReg   = UCase(Trim(CStr(wsBusc.Range(CEL_REG).Value)))
    Dim sFEst   As String: sFEst   = UCase(Trim(CStr(wsBusc.Range(CEL_EST).Value)))
    Dim sFCli   As String: sFCli   = UCase(Trim(CStr(wsBusc.Range(CEL_CLI).Value)))
    Dim sFRFC   As String: sFRFC   = UCase(Trim(CStr(wsBusc.Range(CEL_RFC).Value)))
    Dim sFConc  As String: sFConc  = UCase(Trim(CStr(wsBusc.Range(CEL_CONC).Value)))
    Dim sFOrden As String: sFOrden = UCase(Trim(CStr(wsBusc.Range(CEL_ORDEN).Value)))
    Dim sFDir   As String: sFDir   = UCase(Trim(CStr(wsBusc.Range(CEL_DIR).Value)))

    ' Normalizar "TODOS"
    If sFResp = "TODOS" Or sFResp = "" Then sFResp = ""
    If sFReg  = "TODOS" Or sFReg  = "" Then sFReg  = ""
    If sFEst  = "TODOS" Or sFEst  = "" Then sFEst  = ""

    ' --- Limpiar area de resultados ----------------------------
    Call LimpiarResultados(wsBusc)

    ' --- Arrays temporales (max 2000 filas) --------------------
    Dim nMax As Integer: nMax = 2000
    Dim arrCli(2000)   As String
    Dim arrResp(2000)  As String
    Dim arrRFC(2000)   As String
    Dim arrReg(2000)   As String
    Dim arrConc(2000)  As String
    Dim arrMonto(2000) As Double
    Dim arrVenc(2000)  As Variant
    Dim arrEst(2000)   As String
    Dim arrDias(2000)  As Long
    Dim arrFOp(2000)   As Long
    Dim nRes As Integer: nRes = 0

    Dim uFila As Long
    uFila = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row

    Dim i As Long
    For i = 2 To uFila
        If nRes >= nMax Then GoTo EscribirRes

        Dim sCli  As String: sCli  = Trim(CStr(wsOp.Cells(i, COL_OP_CLIENTE).Value))
        If sCli = "" Then GoTo SigOp

        Dim sEst  As String: sEst  = Trim(CStr(wsOp.Cells(i, COL_OP_ESTATUS).Value))
        Dim sRFC  As String: sRFC  = Trim(CStr(wsOp.Cells(i, COL_OP_RFC).Value))
        Dim sRsp  As String: sRsp  = Trim(CStr(wsOp.Cells(i, COL_OP_RESPONSABLE).Value))
        Dim sReg  As String: sReg  = Trim(CStr(wsOp.Cells(i, COL_OP_REGIMEN).Value))
        Dim sConc As String: sConc = Trim(CStr(wsOp.Cells(i, COL_OP_CONCEPTO).Value))
        Dim fVenc As Variant: fVenc = wsOp.Cells(i, COL_OP_VENCIMIENTO).Value
        Dim dMon  As Double: dMon = 0
        On Error Resume Next: dMon = CDbl(wsOp.Cells(i, COL_OP_MONTO).Value): On Error GoTo FinBusq

        Dim nDias As Long: nDias = 0
        If IsDate(fVenc) Then nDias = DateDiff("d", CDate(fVenc), Date)

        ' --- Aplicar filtros ------------------------------------
        If sFResp <> "" Then
            If UCase(sRsp) <> sFResp Then GoTo SigOp
        End If
        If sFReg <> "" Then
            If UCase(sReg) <> sFReg Then GoTo SigOp
        End If
        If sFEst <> "" Then
            If UCase(sEst) <> sFEst Then GoTo SigOp
        End If
        If sFCli <> "" Then
            If InStr(UCase(sCli), sFCli) = 0 Then GoTo SigOp
        End If
        If sFRFC <> "" Then
            If InStr(UCase(sRFC), sFRFC) = 0 Then GoTo SigOp
        End If
        If sFConc <> "" Then
            If InStr(UCase(sConc), sFConc) = 0 Then GoTo SigOp
        End If

        ' --- Acumular ------------------------------------------
        arrCli(nRes)  = sCli:   arrResp(nRes) = sRsp
        arrRFC(nRes)  = sRFC:   arrReg(nRes)  = sReg
        arrConc(nRes) = sConc:  arrMonto(nRes) = dMon
        arrVenc(nRes) = fVenc:  arrEst(nRes)  = sEst
        arrDias(nRes) = nDias:  arrFOp(nRes)  = i
        nRes = nRes + 1
SigOp:
    Next i

EscribirRes:
    ' --- Ordenar -----------------------------------------------
    If nRes > 1 Then
        Call OrdenarResultados(arrCli, arrResp, arrRFC, arrReg, arrConc, _
                               arrMonto, arrVenc, arrEst, arrDias, arrFOp, _
                               nRes, sFOrden, sFDir)
    End If

    ' --- Escribir en hoja --------------------------------------
    Dim colorZebra As Long: colorZebra = RGB(242, 242, 242)

    Dim r As Integer
    For r = 0 To nRes - 1
        Dim filaR As Long: filaR = BUSC_FILA_DATOS + r
        Dim bgColor As Long
        bgColor = IIf(r Mod 2 = 0, RGB(255, 255, 255), colorZebra)

        ' Datos A-H
        wsBusc.Cells(filaR, RES_NO).Value   = r + 1
        wsBusc.Cells(filaR, RES_CLI).Value  = arrCli(r)
        wsBusc.Cells(filaR, RES_RESP).Value = arrResp(r)
        wsBusc.Cells(filaR, RES_RFC).Value  = arrRFC(r)
        wsBusc.Cells(filaR, RES_REG).Value  = arrReg(r)
        wsBusc.Cells(filaR, RES_CONC).Value = arrConc(r)
        wsBusc.Cells(filaR, RES_MONTO).Value = arrMonto(r)
        wsBusc.Cells(filaR, RES_MONTO).NumberFormat = "$#,##0.00"

        Dim sVF As String
        sVF = IIf(IsDate(arrVenc(r)), Format(CDate(arrVenc(r)), "dd/mm/yyyy"), "")
        wsBusc.Cells(filaR, RES_VENC).Value = sVF

        ' Col I -- Estatus con color
        Dim cF As Long: Dim cT As Long
        Select Case UCase(arrEst(r))
            Case "VENCIDO":    cF = RGB(255, 199, 206): cT = RGB(156, 0, 6)
            Case "HOY VENCE":  cF = RGB(255, 235, 156): cT = RGB(156, 101, 0)
            Case "PENDIENTE":  cF = RGB(221, 235, 247): cT = RGB(31, 78, 121)
            Case "PAGADO":     cF = RGB(198, 239, 206): cT = RGB(0, 97, 0)
            Case Else:         cF = bgColor:            cT = RGB(0, 0, 0)
        End Select

        With wsBusc.Cells(filaR, RES_EST)
            .Value               = arrEst(r)
            .Interior.Color      = cF
            .Font.Color          = cT
            .Font.Bold           = True
            .HorizontalAlignment = xlCenter
        End With

        ' Col J -- Boton WA
        Dim cWA As Long
        cWA = IIf(arrDias(r) > 0, RGB(255, 199, 206), _
              IIf(arrDias(r) = 0, RGB(255, 235, 156), RGB(198, 224, 180)))
        With wsBusc.Cells(filaR, RES_WA)
            .Value               = ChrW(&H25B6) & " WA"
            .Interior.Color      = cWA
            .Font.Color          = RGB(0, 0, 0)
            .Font.Bold           = True
            .HorizontalAlignment = xlCenter
        End With

        ' Col K -- Boton PDF
        With wsBusc.Cells(filaR, RES_PDF)
            .Value               = ChrW(&H25A0) & " PDF"
            .Interior.Color      = RGB(189, 215, 238)
            .Font.Color          = RGB(0, 0, 0)
            .Font.Bold           = True
            .HorizontalAlignment = xlCenter
        End With

        ' Color de fondo en A-H
        Dim c As Integer
        For c = RES_NO To RES_VENC
            wsBusc.Cells(filaR, c).Interior.Color = bgColor
        Next c
    Next r

    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Exit Sub

FinBusq:
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
End Sub

'================================================================
'  LimpiarBuscador -- resetea filtros y borra resultados
'================================================================
Public Sub LimpiarBuscador()
    Dim wsBusc As Worksheet: Set wsBusc = ObtenerHoja("BUSCADOR CLIENTE")
    If wsBusc Is Nothing Then Exit Sub

    Application.EnableEvents = False

    ' Resetear dropdowns a "TODOS"
    wsBusc.Range(CEL_RESP).Value = "TODOS"
    wsBusc.Range(CEL_REG).Value  = "TODOS"
    wsBusc.Range(CEL_EST).Value  = "TODOS"
    wsBusc.Range(CEL_ORDEN).Value = "Vencimiento"
    wsBusc.Range(CEL_DIR).Value  = "Mayor a menor"

    ' Limpiar campos de texto libre
    wsBusc.Range(CEL_CLI).ClearContents
    wsBusc.Range(CEL_RFC).ClearContents
    wsBusc.Range(CEL_CONC).ClearContents

    ' Restaurar activador H3
    With wsBusc.Range(CEL_BUSCAR)
        .Value               = "BUSCAR " & ChrW(&H25B6)
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' Restaurar boton K3 LIMPIAR
    Call InicializarBotonK3_Buscador(wsBusc)

    ' Limpiar resultados
    Call LimpiarResultados(wsBusc)

    Application.EnableEvents = True
End Sub

'================================================================
'  PopularDropdownsRegimen -- pone valores unicos de col C en C3
'================================================================
Public Sub PopularDropdownsRegimen()
    Dim wsBusc As Worksheet: Set wsBusc = ObtenerHoja("BUSCADOR CLIENTE")
    Dim wsOp   As Worksheet: Set wsOp   = ObtenerHoja("OPERACIONES")
    If wsBusc Is Nothing Or wsOp Is Nothing Then Exit Sub

    ' Recolectar unicos de col C (Regimen)
    Dim colReg As New Collection
    On Error Resume Next
    Dim uF As Long: uF = wsOp.Cells(wsOp.Rows.Count, COL_OP_REGIMEN).End(xlUp).Row
    Dim i As Long
    For i = 2 To uF
        Dim sR As String: sR = Trim(CStr(wsOp.Cells(i, COL_OP_REGIMEN).Value))
        If sR <> "" Then colReg.Add sR, UCase(sR)
    Next i
    On Error GoTo 0

    ' Construir lista "TODOS,R1,R2,..."
    Dim lista As String: lista = "TODOS"
    Dim k As Integer
    For k = 1 To colReg.Count
        lista = lista & "," & colReg(k)
    Next k

    ' Aplicar Data Validation a C3
    With wsBusc.Range(CEL_REG).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:=lista
        .ShowError = False
    End With
End Sub

'================================================================
'  InicializarBotonK3_Buscador -- K3 = LIMPIAR (rojo)
'================================================================
Public Sub InicializarBotonK3_Buscador(Optional ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = ObtenerHoja("BUSCADOR CLIENTE")
    If ws Is Nothing Then Exit Sub
    Dim bEv As Boolean: bEv = Application.EnableEvents
    Application.EnableEvents = False
    With ws.Range(CEL_LIMPIAR)
        .Value               = ChrW(&H2715) & " LIMPIAR"
        .Interior.Color      = RGB(192, 0, 0)
        .Font.Color          = RGB(255, 255, 255)
        .Font.Bold           = True
        .Font.Size           = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With
    Application.EnableEvents = bEv
End Sub

'================================================================
'  EnviarWADesdeBuscador
'================================================================
Public Sub EnviarWADesdeBuscador(ByVal filaRes As Long)
    Dim wsBusc As Worksheet: Set wsBusc = ObtenerHoja("BUSCADOR CLIENTE")
    Dim wsOp   As Worksheet: Set wsOp   = ObtenerHoja("OPERACIONES")
    If wsBusc Is Nothing Or wsOp Is Nothing Then Exit Sub

    Dim sCli  As String: sCli  = Trim(CStr(wsBusc.Cells(filaRes, RES_CLI).Value))
    Dim sRFC  As String: sRFC  = Trim(CStr(wsBusc.Cells(filaRes, RES_RFC).Value))
    Dim sConc As String: sConc = Trim(CStr(wsBusc.Cells(filaRes, RES_CONC).Value))
    If sCli = "" Then Exit Sub

    Dim filaOp As Long: filaOp = BuscarFilaEnOp(wsOp, sCli, sRFC, sConc)
    If filaOp > 0 Then Call EnviarMensajeInteligente(filaOp)
End Sub

'================================================================
'  GenerarPDFDesdeBuscador
'================================================================
Public Sub GenerarPDFDesdeBuscador(ByVal filaRes As Long)
    Dim wsBusc As Worksheet: Set wsBusc = ObtenerHoja("BUSCADOR CLIENTE")
    Dim wsOp   As Worksheet: Set wsOp   = ObtenerHoja("OPERACIONES")
    If wsBusc Is Nothing Or wsOp Is Nothing Then Exit Sub

    Dim sCli  As String: sCli  = Trim(CStr(wsBusc.Cells(filaRes, RES_CLI).Value))
    Dim sRFC  As String: sRFC  = Trim(CStr(wsBusc.Cells(filaRes, RES_RFC).Value))
    Dim sConc As String: sConc = Trim(CStr(wsBusc.Cells(filaRes, RES_CONC).Value))
    If sCli = "" Then Exit Sub

    Dim filaOp As Long: filaOp = BuscarFilaEnOp(wsOp, sCli, sRFC, sConc)
    If filaOp > 0 Then Call GenerarEstadoCuentaPDF(filaOp)
End Sub

'================================================================
'  BuscarFilaEnOp -- helper
'================================================================
Private Function BuscarFilaEnOp(wsOp As Worksheet, _
                                  sCli As String, sRFC As String, _
                                  sConc As String) As Long
    BuscarFilaEnOp = 0
    Dim uF As Long: uF = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row
    Dim i As Long
    For i = 2 To uF
        Dim cC As String: cC = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_CLIENTE).Value)))
        Dim cR As String: cR = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_RFC).Value)))
        Dim cN As String: cN = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_CONCEPTO).Value)))
        If cC = UCase(sCli) And cR = UCase(sRFC) And cN = UCase(sConc) Then
            BuscarFilaEnOp = i: Exit For
        End If
    Next i
    If BuscarFilaEnOp = 0 Then
        For i = 2 To uF
            Dim cC2 As String: cC2 = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_CLIENTE).Value)))
            Dim cR2 As String: cR2 = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_RFC).Value)))
            If cC2 = UCase(sCli) And cR2 = UCase(sRFC) Then
                BuscarFilaEnOp = i: Exit For
            End If
        Next i
    End If
End Function

'================================================================
'  LimpiarResultados -- borra area fila 7+ cols A-K
'================================================================
Private Sub LimpiarResultados(wsBusc As Worksheet)
    Dim uFilaRes As Long
    uFilaRes = wsBusc.Cells(wsBusc.Rows.Count, RES_CLI).End(xlUp).Row
    If uFilaRes >= BUSC_FILA_DATOS Then
        With wsBusc.Range(wsBusc.Cells(BUSC_FILA_DATOS, 1), _
                          wsBusc.Cells(uFilaRes, RES_PDF))
            .ClearContents
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex     = xlAutomatic
            .Font.Bold           = False
        End With
    End If
End Sub

'================================================================
'  OrdenarResultados -- burbuja simple
'================================================================
Private Sub OrdenarResultados(arrCli() As String, arrRsp() As String, _
                               arrRFC() As String, arrReg() As String, _
                               arrConc() As String, arrM() As Double, _
                               arrV() As Variant, arrE() As String, _
                               arrD() As Long, arrF() As Long, _
                               n As Integer, sOrden As String, sDir As String)
    Dim asc As Boolean: asc = (InStr(UCase(sDir), "MENOR") > 0)
    Dim i As Integer, j As Integer

    For i = 0 To n - 2
        For j = 0 To n - i - 2
            Dim swap As Boolean: swap = False
            Select Case sOrden
                Case "MONTO":
                    swap = IIf(asc, arrM(j) > arrM(j + 1), arrM(j) < arrM(j + 1))
                Case "CLIENTE":
                    swap = IIf(asc, arrCli(j) > arrCli(j + 1), arrCli(j) < arrCli(j + 1))
                Case Else ' Vencimiento / Dias
                    swap = IIf(asc, arrD(j) < arrD(j + 1), arrD(j) > arrD(j + 1))
            End Select
            If swap Then
                Dim tC As String:  tC  = arrCli(j):  arrCli(j)  = arrCli(j+1):  arrCli(j+1)  = tC
                Dim tR As String:  tR  = arrRsp(j):  arrRsp(j)  = arrRsp(j+1):  arrRsp(j+1)  = tR
                Dim tRF As String: tRF = arrRFC(j):  arrRFC(j)  = arrRFC(j+1):  arrRFC(j+1)  = tRF
                Dim tRG As String: tRG = arrReg(j):  arrReg(j)  = arrReg(j+1):  arrReg(j+1)  = tRG
                Dim tCN As String: tCN = arrConc(j): arrConc(j) = arrConc(j+1): arrConc(j+1) = tCN
                Dim tM As Double:  tM  = arrM(j):    arrM(j)    = arrM(j+1):    arrM(j+1)    = tM
                Dim tV As Variant: tV  = arrV(j):    arrV(j)    = arrV(j+1):    arrV(j+1)    = tV
                Dim tE As String:  tE  = arrE(j):    arrE(j)    = arrE(j+1):    arrE(j+1)    = tE
                Dim tD As Long:    tD  = arrD(j):    arrD(j)    = arrD(j+1):    arrD(j+1)    = tD
                Dim tF As Long:    tF  = arrF(j):    arrF(j)    = arrF(j+1):    arrF(j+1)    = tF
            End If
        Next j
    Next i
End Sub
