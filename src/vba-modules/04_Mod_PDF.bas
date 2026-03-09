Attribute VB_Name = "PDF"
'================================================================
' MODULO: PDF -- BajaTax v4 FINAL
' GenerarEstadoCuentaPDF: diseno mejorado con zebra striping,
'   colores correctos (#1F4E78 / #385623), logo desde B25
'================================================================
Option Explicit

Public Sub GenerarEstadoCuentaPDF(ByVal NL As Long)
    If Not HojasOK() Then Exit Sub

    Dim wsOp   As Worksheet: Set wsOp   = ObtenerHoja("OPERACIONES")
    Dim wsConf As Worksheet: Set wsConf = ObtenerHoja("CONFIGURACION")

    Dim sCliente As String: sCliente = Trim(CStr(wsOp.Cells(NL, COL_OP_CLIENTE).Value))
    Dim sRFC     As String: sRFC     = Trim(CStr(wsOp.Cells(NL, COL_OP_RFC).Value))

    If sCliente = "" Then
        MsgBox "Fila " & NL & " sin cliente.", vbExclamation, "BajaTax": Exit Sub
    End If

    ' Leer configuracion
    Dim sDespacho     As String: sDespacho     = LeerConfig("B5")
    Dim sBeneficiario As String: sBeneficiario = LeerConfig("B6")
    Dim sBanco        As String: sBanco        = LeerConfig("B7")
    Dim sCLABE        As String: sCLABE        = LeerConfig("B8")
    Dim sTelDespacho  As String: sTelDespacho  = LeerConfig("B9")
    Dim sCorreo       As String: sCorreo       = LeerConfig("B10")
    Dim sWeb          As String: sWeb          = LeerConfig("B11")
    Dim sDepartamento As String: sDepartamento = LeerConfig("B12")
    Dim sRutaLogo     As String: sRutaLogo     = LeerConfig("B25")

    ' Colores de seccion (manual v4)
    Dim colorSec1 As Long: colorSec1 = RGB(31, 78, 121)     ' #1F4E78 Azul Oscuro
    Dim colorSec2 As Long: colorSec2 = RGB(56, 86, 35)      ' #385623 Verde Bosque
    Dim colorHdr1 As Long: colorHdr1 = RGB(31, 78, 121)     ' header col S1
    Dim colorHdr2 As Long: colorHdr2 = RGB(0, 97, 0)        ' header col S2
    Dim colorZebra As Long: colorZebra = RGB(242, 242, 242) ' gris zebra

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    On Error GoTo ErrorHandler

    ' Eliminar hoja temporal previa
    Dim wsTmp As Worksheet
    Set wsTmp = ObtenerHoja("TEMP_BAJATAX")
    If Not wsTmp Is Nothing Then
        Application.DisplayAlerts = False
        wsTmp.Delete
        Application.DisplayAlerts = True
    End If

    ' Crear hoja temporal
    Set wsTmp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTmp.Name    = "TEMP_BAJATAX"
    wsTmp.Visible = xlSheetVisible

    ' Configuracion de pagina
    With wsTmp.PageSetup
        .TopMargin    = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin   = Application.CentimetersToPoints(1.5)
        .RightMargin  = Application.CentimetersToPoints(1.5)
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom           = False
        .LeftFooter  = "&""Calibri""&8 " & sDespacho & " | " & sTelDespacho & " | " & sCorreo
        .RightFooter = "&""Calibri""&8 P" & ChrW(225) & "gina &P de &N"
        .CenterFooter = "&""Calibri""&8 CLABE: " & sCLABE & " | Beneficiario: " & sBeneficiario
    End With

    ' Anchos de columna
    wsTmp.Columns("A").ColumnWidth = 5
    wsTmp.Columns("B").ColumnWidth = 36
    wsTmp.Columns("C").ColumnWidth = 13
    wsTmp.Columns("D").ColumnWidth = 13
    wsTmp.Columns("E").ColumnWidth = 14
    wsTmp.Columns("F").ColumnWidth = 12
    wsTmp.Columns("G").ColumnWidth = 9

    ' -- LOGO ----------------------------------------------------------
    If sRutaLogo <> "" Then
        On Error Resume Next
        Dim oPic As Object
        Set oPic = wsTmp.Pictures.Insert(sRutaLogo)
        If Not oPic Is Nothing Then
            With oPic
                .Left  = wsTmp.Cells(1, 1).Left
                .Top   = wsTmp.Cells(1, 1).Top
                .Height = 55
                .Width  = .Width * (55 / .Height)
                .Placement = xlFreeFloating
            End With
        End If
        On Error GoTo ErrorHandler
    End If

    ' -- MEMBRETE ------------------------------------------------------
    Dim fila As Long: fila = 1

    With wsTmp.Cells(fila, 7)
        .Value              = sDespacho
        .Font.Bold          = True
        .Font.Size          = 14
        .Font.Color         = colorSec1
        .HorizontalAlignment = xlRight
    End With

    fila = 2
    With wsTmp.Cells(fila, 7)
        .Value              = sDepartamento
        .Font.Size          = 9
        .Font.Color         = RGB(80, 80, 80)
        .HorizontalAlignment = xlRight
    End With

    fila = 3
    With wsTmp.Cells(fila, 7)
        .Value              = sTelDespacho & "  |  " & sCorreo
        .Font.Size          = 8
        .Font.Color         = RGB(80, 80, 80)
        .HorizontalAlignment = xlRight
    End With

    fila = 4
    With wsTmp.Range("A" & fila & ":G" & fila).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color     = colorSec1
        .Weight    = xlMedium
    End With

    ' -- TITULO Y DATOS DEL CLIENTE ------------------------------------
    fila = 5
    With wsTmp.Cells(fila, 1)
        .Value    = "ESTADO DE CUENTA"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = colorSec1
    End With
    With wsTmp.Cells(fila, 6)
        .Value              = "Generado el " & Format(Now, "dd-mmm-yyyy")
        .Font.Size          = 9
        .Font.Color         = RGB(100, 100, 100)
        .HorizontalAlignment = xlRight
    End With

    fila = 6
    wsTmp.Cells(fila, 1).Value    = "CLIENTE:"
    wsTmp.Cells(fila, 1).Font.Bold = True
    wsTmp.Cells(fila, 1).Font.Size = 11
    wsTmp.Cells(fila, 2).Value    = UCase(sCliente)
    wsTmp.Cells(fila, 2).Font.Bold = True
    wsTmp.Cells(fila, 2).Font.Size = 11

    fila = 7
    wsTmp.Cells(fila, 1).Value    = "RFC:"
    wsTmp.Cells(fila, 1).Font.Bold = True
    wsTmp.Cells(fila, 2).Value    = UCase(sRFC)

    ' -- SECCION 1: CONCEPTOS PENDIENTES -------------------------------
    fila = 9
    With wsTmp.Range("A" & fila & ":G" & fila)
        .Merge
        .Value              = "  SECCI" & ChrW(211) & "N 1: CONCEPTOS PENDIENTES"
        .Interior.Color     = colorSec1
        .Font.Bold          = True
        .Font.Size          = 11
        .Font.Color         = RGB(255, 255, 255)
    End With

    fila = 10
    Dim encCols As Variant
    encCols = Array("No.", "Concepto", "F. Cobro", "Vencimiento", "Monto", "Estatus", "D" & ChrW(237) & "as Venc.")
    Dim col As Integer
    For col = 0 To 6
        With wsTmp.Cells(fila, col + 1)
            .Value              = encCols(col)
            .Font.Bold          = True
            .Font.Size          = 10
            .Font.Color         = RGB(255, 255, 255)
            .Interior.Color     = colorHdr1
            .HorizontalAlignment = IIf(col >= 4, xlRight, xlLeft)
        End With
    Next col

    ' Datos pendientes
    Dim uFilaOp As Long
    uFilaOp = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row

    fila = 11
    Dim contador    As Integer: contador    = 0
    Dim montoPend   As Double:  montoPend   = 0
    Dim zebraCont   As Integer: zebraCont   = 0

    Dim i As Long
    For i = 2 To uFilaOp
        Dim sClienteReg As String: sClienteReg = Trim(CStr(wsOp.Cells(i, COL_OP_CLIENTE).Value))
        If UCase(sClienteReg) <> UCase(sCliente) Then GoTo SigPend

        Dim sPago       As String: sPago    = Trim(CStr(wsOp.Cells(i, COL_OP_REG_PAGO).Value))
        Dim sEst        As String: sEst     = UCase(Trim(CStr(wsOp.Cells(i, COL_OP_ESTATUS).Value)))
        If sPago <> "" Or sEst = "PAGADO" Then GoTo SigPend

        Dim dMonto      As Double: dMonto   = 0
        Dim sConcReg    As String: sConcReg = Trim(CStr(wsOp.Cells(i, COL_OP_CONCEPTO).Value))
        Dim fCobro      As Variant: fCobro  = wsOp.Cells(i, COL_OP_FECHA_COB).Value
        Dim fVencReg    As Variant: fVencReg = wsOp.Cells(i, COL_OP_VENCIMIENTO).Value
        On Error Resume Next
        dMonto = CDbl(wsOp.Cells(i, COL_OP_MONTO).Value)
        On Error GoTo ErrorHandler

        Dim diasV As Long: diasV = 0
        If IsDate(fVencReg) Then diasV = DateDiff("d", CDate(fVencReg), Date)

        contador  = contador + 1
        montoPend = montoPend + dMonto
        zebraCont = zebraCont + 1

        ' Zebra striping
        Dim bgColor As Long
        bgColor = IIf(zebraCont Mod 2 = 0, colorZebra, RGB(255, 255, 255))

        wsTmp.Cells(fila, 1).Value = contador
        wsTmp.Cells(fila, 2).Value = sConcReg
        wsTmp.Cells(fila, 3).Value = IIf(IsDate(fCobro), Format(CDate(fCobro), "dd-mmm-yyyy"), "")
        wsTmp.Cells(fila, 4).Value = IIf(IsDate(fVencReg), Format(CDate(fVencReg), "dd-mmm-yyyy"), "")
        wsTmp.Cells(fila, 5).Value = dMonto
        wsTmp.Cells(fila, 5).NumberFormat = "$#,##0.00"
        wsTmp.Cells(fila, 5).HorizontalAlignment = xlRight
        wsTmp.Cells(fila, 7).Value = IIf(diasV > 0, diasV, "")
        wsTmp.Cells(fila, 7).HorizontalAlignment = xlRight

        ' Color estatus
        Dim colorFondo As Long: Dim colorTxt As Long
        Select Case sEst
            Case "VENCIDO"
                colorFondo = RGB(255, 199, 206): colorTxt = RGB(156, 0, 6)
            Case "HOY VENCE", "HOY_VENCE"
                colorFondo = RGB(255, 235, 156): colorTxt = RGB(156, 101, 0)
            Case Else
                colorFondo = RGB(221, 235, 247): colorTxt = colorSec1
        End Select

        wsTmp.Cells(fila, 6).Value              = sEst
        wsTmp.Cells(fila, 6).Interior.Color     = colorFondo
        wsTmp.Cells(fila, 6).Font.Color         = colorTxt
        wsTmp.Cells(fila, 6).Font.Bold          = True
        wsTmp.Cells(fila, 6).HorizontalAlignment = xlCenter

        ' Aplicar color de fondo zebra a columnas A-E y G
        Dim cz As Integer
        For cz = 1 To 5
            If wsTmp.Cells(fila, cz).Interior.ColorIndex = xlNone Or _
               wsTmp.Cells(fila, cz).Interior.Color = RGB(255, 255, 255) Then
                wsTmp.Cells(fila, cz).Interior.Color = bgColor
            End If
        Next cz
        wsTmp.Cells(fila, 7).Interior.Color = bgColor

        ' Borde inferior
        wsTmp.Range("A" & fila & ":G" & fila).Borders(xlEdgeBottom).LineStyle = xlContinuous

        fila = fila + 1
SigPend:
    Next i

    ' Total pendiente
    fila = fila + 1
    With wsTmp.Cells(fila, 4)
        .Value              = "TOTAL PENDIENTE:"
        .Font.Bold          = True
        .HorizontalAlignment = xlRight
    End With
    With wsTmp.Cells(fila, 5)
        .Value        = montoPend
        .NumberFormat = "$#,##0.00"
        .Font.Bold    = True
        .Font.Color   = RGB(156, 0, 6)
        .HorizontalAlignment = xlRight
    End With

    ' -- SECCION 2: HISTORIAL DE PAGOS ---------------------------------
    fila = fila + 2
    With wsTmp.Range("A" & fila & ":G" & fila)
        .Merge
        .Value          = "  SECCI" & ChrW(211) & "N 2: HISTORIAL DE CONCEPTOS LIQUIDADOS"
        .Interior.Color = colorSec2
        .Font.Bold      = True
        .Font.Size      = 11
        .Font.Color     = RGB(255, 255, 255)
    End With

    fila = fila + 1
    Dim encPagos As Variant
    encPagos = Array("No.", "Concepto", "F. Cobro", "Fecha Pago", "Monto", "M" & ChrW(233) & "todo", "")
    For col = 0 To 6
        With wsTmp.Cells(fila, col + 1)
            .Value          = encPagos(col)
            .Font.Bold      = True
            .Font.Color     = RGB(255, 255, 255)
            .Interior.Color = colorHdr2
            .Font.Size      = 10
        End With
    Next col

    fila = fila + 1
    Dim contPag      As Integer: contPag      = 0
    Dim montoLiqui   As Double:  montoLiqui   = 0
    Dim zebraContP   As Integer: zebraContP   = 0

    For i = 2 To uFilaOp
        sClienteReg = Trim(CStr(wsOp.Cells(i, COL_OP_CLIENTE).Value))
        If UCase(sClienteReg) <> UCase(sCliente) Then GoTo SigPag
        sPago = Trim(CStr(wsOp.Cells(i, COL_OP_REG_PAGO).Value))
        If sPago = "" Then GoTo SigPag

        contPag    = contPag + 1
        zebraContP = zebraContP + 1
        dMonto     = 0
        On Error Resume Next
        dMonto = CDbl(wsOp.Cells(i, COL_OP_MONTO).Value)
        On Error GoTo ErrorHandler
        montoLiqui = montoLiqui + dMonto
        sConcReg   = Trim(CStr(wsOp.Cells(i, COL_OP_CONCEPTO).Value))
        fCobro     = wsOp.Cells(i, COL_OP_FECHA_COB).Value

        Dim bgP As Long: bgP = IIf(zebraContP Mod 2 = 0, colorZebra, RGB(255, 255, 255))

        wsTmp.Cells(fila, 1).Value = contPag
        wsTmp.Cells(fila, 2).Value = sConcReg
        wsTmp.Cells(fila, 3).Value = IIf(IsDate(fCobro), Format(CDate(fCobro), "dd-mmm-yyyy"), "")
        wsTmp.Cells(fila, 4).Value = sPago
        wsTmp.Cells(fila, 5).Value = dMonto
        wsTmp.Cells(fila, 5).NumberFormat = "$#,##0.00"
        wsTmp.Cells(fila, 5).HorizontalAlignment = xlRight
        wsTmp.Cells(fila, 6).Value              = "Registrado"
        wsTmp.Cells(fila, 6).Interior.Color     = RGB(198, 239, 206)
        wsTmp.Cells(fila, 6).Font.Color         = RGB(0, 97, 0)
        wsTmp.Cells(fila, 6).HorizontalAlignment = xlCenter

        For cz = 1 To 5
            wsTmp.Cells(fila, cz).Interior.Color = bgP
        Next cz
        wsTmp.Cells(fila, 7).Interior.Color = bgP
        wsTmp.Range("A" & fila & ":G" & fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
        fila = fila + 1
SigPag:
    Next i

    fila = fila + 1
    With wsTmp.Cells(fila, 4)
        .Value              = "TOTAL LIQUIDADO:"
        .Font.Bold          = True
        .HorizontalAlignment = xlRight
    End With
    With wsTmp.Cells(fila, 5)
        .Value        = montoLiqui
        .NumberFormat = "$#,##0.00"
        .Font.Bold    = True
        .Font.Color   = RGB(0, 97, 0)
        .HorizontalAlignment = xlRight
    End With

    ' -- DATOS BANCARIOS -----------------------------------------------
    fila = fila + 3
    With wsTmp.Range("A" & fila & ":G" & fila)
        .Merge
        .Value          = "  DATOS PARA TRANSFERENCIA"
        .Interior.Color = colorSec1
        .Font.Bold      = True
        .Font.Color     = RGB(255, 255, 255)
    End With

    fila = fila + 1
    wsTmp.Cells(fila, 1).Value    = "Beneficiario:"
    wsTmp.Cells(fila, 1).Font.Bold = True
    wsTmp.Cells(fila, 2).Value    = sBeneficiario
    fila = fila + 1
    wsTmp.Cells(fila, 1).Value    = "Banco:"
    wsTmp.Cells(fila, 1).Font.Bold = True
    wsTmp.Cells(fila, 2).Value    = sBanco
    fila = fila + 1
    wsTmp.Cells(fila, 1).Value    = "CLABE:"
    wsTmp.Cells(fila, 1).Font.Bold = True
    wsTmp.Cells(fila, 2).Value    = sCLABE

    fila = fila + 2
    wsTmp.Cells(fila, 1).Value   = "Cualquier duda estamos a sus " & ChrW(243) & "rdenes."
    wsTmp.Cells(fila, 1).Font.Italic = True
    wsTmp.Cells(fila, 1).Font.Color  = RGB(80, 80, 80)

    fila = fila + 1
    wsTmp.Cells(fila, 1).Value   = sDepartamento & "  |  " & sTelDespacho & "  |  " & sCorreo
    wsTmp.Cells(fila, 1).Font.Size = 8
    wsTmp.Cells(fila, 1).Font.Color = RGB(100, 100, 100)

    wsTmp.Columns("B:G").AutoFit

    ' -- GUARDAR PDF ---------------------------------------------------
    Dim sFolderBase As String
    sFolderBase = ThisWorkbook.Path
    If Right(sFolderBase, 1) <> Application.PathSeparator Then
        sFolderBase = sFolderBase & Application.PathSeparator
    End If

    Dim sClienteLimpio As String
    sClienteLimpio = Replace(Replace(Replace(sCliente, " ", "_"), "/", "-"), "\", "-")
    Dim sNombreArch As String
    sNombreArch = "EdoCuenta_" & sClienteLimpio & "_" & Format(Now, "ddmmyyyy")

    Dim sRutaElegida As Variant
    If EsMac() Then
        sRutaElegida = Application.GetSaveAsFilename(InitialFileName:=sFolderBase & sNombreArch)
    Else
        sRutaElegida = Application.GetSaveAsFilename( _
            InitialFileName:=sFolderBase & sNombreArch, _
            FileFilter:="PDF Files (*.pdf), *.pdf")
    End If

    If sRutaElegida = False Or CStr(sRutaElegida) = "" Then GoTo Limpieza

    Dim sRutaPDF As String: sRutaPDF = CStr(sRutaElegida)
    If LCase(Right(sRutaPDF, 4)) <> ".pdf" Then sRutaPDF = sRutaPDF & ".pdf"

    wsTmp.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=sRutaPDF, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "PDF generado:" & Chr(13) & sRutaPDF, vbInformation, "BajaTax"
    GoTo Limpieza

ErrorHandler:
    MsgBox "Error al generar PDF: " & Err.Description & " (#" & Err.Number & ")", _
           vbCritical, "BajaTax"

Limpieza:
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsTmp = ObtenerHoja("TEMP_BAJATAX")
    If Not wsTmp Is Nothing Then wsTmp.Delete
    On Error GoTo 0
    Application.DisplayAlerts  = True
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
End Sub
