Attribute VB_Name = "Mod_MasivoPDF"
'================================================================
' MODULO: Mod_MasivoPDF -- BajaTax v4 FINAL
' GenerarPDFMasivo: un PDF por RFC con todos sus adeudos
'   Crea carpeta SALIDA_PDF/DDMMYYYY/ automaticamente
'   Sin pausas (proceso local)
'================================================================
Option Explicit

Public Sub GenerarPDFMasivo()
    If Not HojasOK() Then Exit Sub

    Dim wsOp   As Worksheet: Set wsOp   = ObtenerHoja("OPERACIONES")
    Dim wsConf As Worksheet: Set wsConf = ObtenerHoja("CONFIGURACION")

    Dim resp As Integer
    resp = MsgBox("GENERACION MASIVA DE PDF" & Chr(13) & Chr(13) & _
                  "Se generara un PDF por cliente con todos sus adeudos pendientes." & Chr(13) & _
                  "Los archivos se guardaran en:" & Chr(13) & _
                  "  SALIDA_PDF/" & Format(Date, "ddmmyyyy") & "/" & Chr(13) & Chr(13) & _
                  "Se omitiran clientes EXCLUIDOS o SUSPENDIDOS." & Chr(13) & _
                  ChrW(&H26A0) & " Continuar?", _
                  vbYesNo + vbQuestion, "BajaTax - PDF Masivo")
    If resp = vbNo Then Exit Sub

    ' Crear carpeta de salida
    Dim sBase As String
    sBase = ThisWorkbook.Path & Application.PathSeparator & _
            "SALIDA_PDF" & Application.PathSeparator & _
            Format(Date, "ddmmyyyy") & Application.PathSeparator

    On Error Resume Next
    MkDir ThisWorkbook.Path & Application.PathSeparator & "SALIDA_PDF"
    MkDir Left(sBase, Len(sBase) - 1)
    On Error GoTo 0

    ' Recopilar RFCs unicos con adeudos pendientes
    Dim uFila As Long
    uFila = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row

    Dim arrRFC(500)     As String
    Dim arrCliente(500) As String
    Dim nRFCs As Integer: nRFCs = 0

    Dim i As Long
    For i = 2 To uFila
        Dim sCliente As String: sCliente = Trim(CStr(wsOp.Cells(i, COL_OP_CLIENTE).Value))
        If sCliente = "" Then GoTo SigScan

        Dim sPago    As String: sPago    = Trim(CStr(wsOp.Cells(i, COL_OP_REG_PAGO).Value))
        Dim sExcluir As String: sExcluir = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_EXCLUIR).Value)))
        Dim sEst     As String: sEst     = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_ESTATUS).Value)))
        Dim sRFC     As String: sRFC     = Trim(CStr(wsOp.Cells(i, COL_OP_RFC).Value))

        If sPago <> "" Then GoTo SigScan
        If sExcluir = "SI" Or sExcluir = "S" & ChrW(237) Or sExcluir = "X" Then GoTo SigScan
        If sEst = "PAGADO" Then GoTo SigScan
        If RFCSuspendido(sRFC) Then GoTo SigScan

        ' Solo pendientes, vencidos, hoy vence
        If sEst <> "PENDIENTE" And sEst <> "VENCIDO" And sEst <> "HOY VENCE" Then GoTo SigScan

        ' Agregar RFC si no esta ya en la lista
        Dim yaExiste As Boolean: yaExiste = False
        Dim k As Integer
        For k = 0 To nRFCs - 1
            If arrRFC(k) = sRFC And arrCliente(k) = sCliente Then
                yaExiste = True: Exit For
            End If
        Next k

        If Not yaExiste And nRFCs < 500 Then
            arrRFC(nRFCs)     = sRFC
            arrCliente(nRFCs) = sCliente
            nRFCs             = nRFCs + 1
        End If
SigScan:
    Next i

    If nRFCs = 0 Then
        MsgBox "No hay clientes con adeudos pendientes.", vbInformation, "BajaTax"
        Exit Sub
    End If

    ' Confirmar cantidad
    Dim confCant As Integer
    confCant = MsgBox("Se generaran " & nRFCs & " PDF(s)." & Chr(13) & _
                      "Carpeta: " & sBase & Chr(13) & Chr(13) & "Continuar?", _
                      vbYesNo + vbQuestion, "BajaTax")
    If confCant = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents   = False

    Dim contOK  As Integer: contOK  = 0
    Dim contErr As Integer: contErr = 0

    ' Leer config
    Dim sDespacho     As String: sDespacho     = LeerConfig("B5")
    Dim sBeneficiario As String: sBeneficiario = LeerConfig("B6")
    Dim sBanco        As String: sBanco        = LeerConfig("B7")
    Dim sCLABE        As String: sCLABE        = LeerConfig("B8")
    Dim sTelDespacho  As String: sTelDespacho  = LeerConfig("B9")
    Dim sCorreo       As String: sCorreo       = LeerConfig("B10")
    Dim sDepartamento As String: sDepartamento = LeerConfig("B12")
    Dim sRutaLogo     As String: sRutaLogo     = LeerConfig("B25")

    Dim colorSec1  As Long: colorSec1  = RGB(31, 78, 121)
    Dim colorSec2  As Long: colorSec2  = RGB(56, 86, 35)
    Dim colorZebra As Long: colorZebra = RGB(242, 242, 242)

    ' -- Generar un PDF por cliente --------------------------
    Dim g As Integer
    For g = 0 To nRFCs - 1
        Dim sClienteG As String: sClienteG = arrCliente(g)
        Dim sRFCG     As String: sRFCG     = arrRFC(g)

        On Error GoTo ErrPDF

        ' Limpiar hoja temporal
        Dim wsTmp As Worksheet
        Set wsTmp = ObtenerHoja("TEMP_BAJATAX")
        If Not wsTmp Is Nothing Then
            Application.DisplayAlerts = False
            wsTmp.Delete
            Application.DisplayAlerts = True
        End If
        Set wsTmp = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTmp.Name    = "TEMP_BAJATAX"
        wsTmp.Visible = xlSheetVisible

        ' Configuracion de pagina
        With wsTmp.PageSetup
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .Zoom           = False
            .TopMargin    = Application.CentimetersToPoints(1.5)
            .BottomMargin = Application.CentimetersToPoints(1.5)
            .LeftMargin   = Application.CentimetersToPoints(1.5)
            .RightMargin  = Application.CentimetersToPoints(1.5)
            .LeftFooter   = "&""Calibri""&8 " & sDespacho & " | " & sTelDespacho
            .RightFooter  = "&""Calibri""&8 P" & ChrW(225) & "g. &P de &N"
            .CenterFooter = "&""Calibri""&8 CLABE: " & sCLABE
        End With

        wsTmp.Columns("A").ColumnWidth = 5
        wsTmp.Columns("B").ColumnWidth = 36
        wsTmp.Columns("C").ColumnWidth = 13
        wsTmp.Columns("D").ColumnWidth = 13
        wsTmp.Columns("E").ColumnWidth = 14
        wsTmp.Columns("F").ColumnWidth = 12
        wsTmp.Columns("G").ColumnWidth = 9

        ' Logo
        If sRutaLogo <> "" Then
            On Error Resume Next
            Dim oPic As Object
            Set oPic = wsTmp.Pictures.Insert(sRutaLogo)
            If Not oPic Is Nothing Then
                With oPic
                    .Left = wsTmp.Cells(1, 1).Left
                    .Top  = wsTmp.Cells(1, 1).Top
                    .Height = 55
                    .Width  = .Width * (55 / .Height)
                    .Placement = xlFreeFloating
                End With
            End If
            On Error GoTo ErrPDF
        End If

        ' Membrete
        Dim fila As Long: fila = 1
        With wsTmp.Cells(fila, 7)
            .Value = sDespacho: .Font.Bold = True: .Font.Size = 14
            .Font.Color = colorSec1: .HorizontalAlignment = xlRight
        End With
        fila = 2
        With wsTmp.Cells(fila, 7)
            .Value = sDepartamento: .Font.Size = 9
            .Font.Color = RGB(80, 80, 80): .HorizontalAlignment = xlRight
        End With
        fila = 3
        With wsTmp.Cells(fila, 7)
            .Value = sTelDespacho & "  |  " & sCorreo: .Font.Size = 8
            .Font.Color = RGB(80, 80, 80): .HorizontalAlignment = xlRight
        End With
        fila = 4
        With wsTmp.Range("A" & fila & ":G" & fila).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Color = colorSec1: .Weight = xlMedium
        End With

        fila = 5
        With wsTmp.Cells(fila, 1)
            .Value = "ESTADO DE CUENTA": .Font.Bold = True: .Font.Size = 16: .Font.Color = colorSec1
        End With
        With wsTmp.Cells(fila, 6)
            .Value = "Generado el " & Format(Now, "dd-mmm-yyyy")
            .Font.Size = 9: .HorizontalAlignment = xlRight
        End With
        fila = 6
        wsTmp.Cells(fila, 1).Value = "CLIENTE:": wsTmp.Cells(fila, 1).Font.Bold = True: wsTmp.Cells(fila, 1).Font.Size = 11
        wsTmp.Cells(fila, 2).Value = UCase(sClienteG): wsTmp.Cells(fila, 2).Font.Bold = True: wsTmp.Cells(fila, 2).Font.Size = 11
        fila = 7
        wsTmp.Cells(fila, 1).Value = "RFC:": wsTmp.Cells(fila, 1).Font.Bold = True
        wsTmp.Cells(fila, 2).Value = UCase(sRFCG)

        ' Seccion 1
        fila = 9
        With wsTmp.Range("A" & fila & ":G" & fila)
            .Merge: .Value = "  SECCI" & ChrW(211) & "N 1: CONCEPTOS PENDIENTES"
            .Interior.Color = colorSec1: .Font.Bold = True: .Font.Size = 11: .Font.Color = RGB(255, 255, 255)
        End With
        fila = 10
        Dim enc As Variant: enc = Array("No.", "Concepto", "F. Cobro", "Vencimiento", "Monto", "Estatus", "D" & ChrW(237) & "as")
        Dim col As Integer
        For col = 0 To 6
            With wsTmp.Cells(fila, col + 1)
                .Value = enc(col): .Font.Bold = True: .Font.Size = 10
                .Font.Color = RGB(255, 255, 255): .Interior.Color = colorSec1
                .HorizontalAlignment = IIf(col >= 4, xlRight, xlLeft)
            End With
        Next col

        fila = 11
        Dim contador  As Integer: contador  = 0
        Dim montoPend As Double:  montoPend = 0
        Dim zebra     As Integer: zebra     = 0

        For i = 2 To uFila
            Dim sCR As String: sCR = Trim(CStr(wsOp.Cells(i, COL_OP_CLIENTE).Value))
            If UCase(sCR) <> UCase(sClienteG) Then GoTo SigP
            sPago = Trim(CStr(wsOp.Cells(i, COL_OP_REG_PAGO).Value))
            sEst  = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_ESTATUS).Value)))
            If sPago <> "" Or sEst = "PAGADO" Then GoTo SigP

            Dim dM   As Double: dM   = 0
            Dim sC   As String: sC   = Trim(CStr(wsOp.Cells(i, COL_OP_CONCEPTO).Value))
            Dim fCob As Variant: fCob = wsOp.Cells(i, COL_OP_FECHA_COB).Value
            Dim fV   As Variant: fV   = wsOp.Cells(i, COL_OP_VENCIMIENTO).Value
            On Error Resume Next: dM = CDbl(wsOp.Cells(i, COL_OP_MONTO).Value): On Error GoTo ErrPDF

            Dim dV As Long: dV = 0
            If IsDate(fV) Then dV = DateDiff("d", CDate(fV), Date)

            contador = contador + 1: montoPend = montoPend + dM: zebra = zebra + 1
            Dim bg As Long: bg = IIf(zebra Mod 2 = 0, colorZebra, RGB(255, 255, 255))

            wsTmp.Cells(fila, 1).Value = contador
            wsTmp.Cells(fila, 2).Value = sC
            wsTmp.Cells(fila, 3).Value = IIf(IsDate(fCob), Format(CDate(fCob), "dd-mmm-yyyy"), "")
            wsTmp.Cells(fila, 4).Value = IIf(IsDate(fV), Format(CDate(fV), "dd-mmm-yyyy"), "")
            wsTmp.Cells(fila, 5).Value = dM: wsTmp.Cells(fila, 5).NumberFormat = "$#,##0.00"
            wsTmp.Cells(fila, 5).HorizontalAlignment = xlRight
            wsTmp.Cells(fila, 7).Value = IIf(dV > 0, dV, ""): wsTmp.Cells(fila, 7).HorizontalAlignment = xlRight

            Dim cF As Long: Dim cT As Long
            Select Case sEst
                Case "VENCIDO": cF = RGB(255, 199, 206): cT = RGB(156, 0, 6)
                Case "HOY VENCE", "HOY_VENCE": cF = RGB(255, 235, 156): cT = RGB(156, 101, 0)
                Case Else: cF = RGB(221, 235, 247): cT = colorSec1
            End Select
            wsTmp.Cells(fila, 6).Value = sEst: wsTmp.Cells(fila, 6).Interior.Color = cF
            wsTmp.Cells(fila, 6).Font.Color = cT: wsTmp.Cells(fila, 6).Font.Bold = True
            wsTmp.Cells(fila, 6).HorizontalAlignment = xlCenter

            Dim cz As Integer
            For cz = 1 To 5: wsTmp.Cells(fila, cz).Interior.Color = bg: Next cz
            wsTmp.Cells(fila, 7).Interior.Color = bg
            wsTmp.Range("A" & fila & ":G" & fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
            fila = fila + 1
SigP:
        Next i

        fila = fila + 1
        With wsTmp.Cells(fila, 4): .Value = "TOTAL PENDIENTE:": .Font.Bold = True: .HorizontalAlignment = xlRight: End With
        With wsTmp.Cells(fila, 5): .Value = montoPend: .NumberFormat = "$#,##0.00": .Font.Bold = True: .Font.Color = RGB(156, 0, 6): .HorizontalAlignment = xlRight: End With

        ' Datos bancarios
        fila = fila + 3
        With wsTmp.Range("A" & fila & ":G" & fila)
            .Merge: .Value = "  DATOS PARA TRANSFERENCIA": .Interior.Color = colorSec1
            .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        End With
        fila = fila + 1: wsTmp.Cells(fila, 1).Value = "Beneficiario:": wsTmp.Cells(fila, 1).Font.Bold = True: wsTmp.Cells(fila, 2).Value = sBeneficiario
        fila = fila + 1: wsTmp.Cells(fila, 1).Value = "Banco:": wsTmp.Cells(fila, 1).Font.Bold = True: wsTmp.Cells(fila, 2).Value = sBanco
        fila = fila + 1: wsTmp.Cells(fila, 1).Value = "CLABE:": wsTmp.Cells(fila, 1).Font.Bold = True: wsTmp.Cells(fila, 2).Value = sCLABE

        wsTmp.Columns("B:G").AutoFit

        ' Guardar PDF
        Dim sClienteLimpio As String
        sClienteLimpio = Replace(Replace(Replace(sClienteG, " ", "_"), "/", "-"), "\", "-")
        Dim sRutaPDF As String
        sRutaPDF = sBase & "EdoCuenta_" & sClienteLimpio & "_" & Format(Date, "ddmmyyyy") & ".pdf"

        wsTmp.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=sRutaPDF, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False

        ' Limpiar hoja temporal
        Application.DisplayAlerts = False
        wsTmp.Delete
        Application.DisplayAlerts = True

        contOK = contOK + 1
        Application.StatusBar = "BajaTax: PDF " & contOK & " de " & nRFCs & " generado..."
        GoTo SigPDF

ErrPDF:
        contErr = contErr + 1
        On Error Resume Next
        Application.DisplayAlerts = False
        Set wsTmp = ObtenerHoja("TEMP_BAJATAX")
        If Not wsTmp Is Nothing Then wsTmp.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
        Err.Clear
SigPDF:
    Next g

    Application.StatusBar     = False
    Application.ScreenUpdating = True
    Application.EnableEvents   = True

    MsgBox "Generacion masiva completada." & Chr(13) & Chr(13) & _
           "  PDFs generados: " & contOK & Chr(13) & _
           "  Errores:        " & contErr & Chr(13) & Chr(13) & _
           "Carpeta: " & sBase, vbInformation, "BajaTax - PDF Masivo"
End Sub
