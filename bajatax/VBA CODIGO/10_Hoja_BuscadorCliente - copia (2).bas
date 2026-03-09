Attribute VB_Name = "Hoja_BUSCADOR"
'================================================================
' CODIGO DE HOJA: BUSCADOR CLIENTE -- BajaTax v4 FINAL
'
' BeforeDoubleClick : H3 -> EjecutarBusqueda
'                    K3 -> LimpiarBuscador
' SelectionChange   : Col J/K fila 7+ -> WA / PDF
'                    Salir de E3/F3/G3 -> EjecutarBusqueda
' Worksheet_Change  : Restaura H3, J7+, K7+, K3 si se borran
' Worksheet_Activate: Refresca dropdown de Regimen (C3)
'================================================================
Option Explicit

Private bBusy         As Boolean
Private mCeldaFiltro  As String   ' Rastrea celda de filtro texto activa

'================================================================
'  InicializarHojaBuscador -- configura dropdowns y botones
'  (llamado por Bootstrap al instalar)
'================================================================
Public Sub InicializarHojaBuscador()
    Dim bEv As Boolean: bEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo FinInit

    ' --- Dropdown B3: Responsable --------------------------------
    With Me.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="TODOS,JOSSELYN,DENISSE,FERNANDA,OMAR,JAVIER"
        .ShowError = False
    End With
    If Trim(CStr(Me.Range("B3").Value)) = "" Then Me.Range("B3").Value = "TODOS"

    ' --- Dropdown D3: Estatus ------------------------------------
    With Me.Range("D3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="TODOS,PENDIENTE,VENCIDO,HOY VENCE,PAGADO"
        .ShowError = False
    End With
    If Trim(CStr(Me.Range("D3").Value)) = "" Then Me.Range("D3").Value = "TODOS"

    ' --- Dropdown I3: Ordenar ------------------------------------
    With Me.Range("I3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="Vencimiento,Monto,Cliente,D" & ChrW(&HED) & "as"
        .ShowError = False
    End With
    If Trim(CStr(Me.Range("I3").Value)) = "" Then Me.Range("I3").Value = "Vencimiento"

    ' --- Dropdown J3: Direccion ----------------------------------
    With Me.Range("J3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="Mayor a menor,Menor a mayor"
        .ShowError = False
    End With
    If Trim(CStr(Me.Range("J3").Value)) = "" Then Me.Range("J3").Value = "Mayor a menor"

    ' --- Activador H3: BUSCAR -----------------------------------
    Call RestaurarActivadorH3

    ' --- Encabezados fila 6 -------------------------------------
    Call EscribirEncabezados

    ' --- Boton K3: LIMPIAR ----------------------------------------
    Call InicializarBotonK3_Buscador(Me)

    ' --- Dropdown C3: Regimen (unicos de OPERACIONES) -----------
    Call PopularDropdownsRegimen

FinInit:
    Application.EnableEvents = bEv
End Sub

'================================================================
'  BeforeDoubleClick -- H3 (buscar) y Z1 (limpiar)
'================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.CountLarge > 1 Then Exit Sub
    If bBusy Then Exit Sub

    ' --- Boton K3: LIMPIAR ----------------------------------------
    If Target.Address = "$K$3" Then
        Cancel = True
        bBusy = True
        Application.EnableEvents = False
        On Error GoTo FinDC
        Call LimpiarBuscador
        GoTo FinDC
    End If

    ' --- Activador H3: BUSCAR ----------------------------------
    If Target.Address = "$H$3" Then
        Cancel = True
        bBusy = True
        Application.EnableEvents = False
        On Error GoTo FinDC
        Call EjecutarBusqueda
    End If

FinDC:
    bBusy = False
    Application.EnableEvents = True
End Sub

'================================================================
'  SelectionChange -- WA (col J), PDF (col K), salir E3/F3/G3
'================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub
    If bBusy Then Exit Sub

    ' --- Detectar si se salio de celda de filtro texto ---------
    Dim eraCelda As String: eraCelda = mCeldaFiltro
    mCeldaFiltro = ""

    If Target.Address = "$E$3" Or Target.Address = "$F$3" Or _
       Target.Address = "$G$3" Then
        mCeldaFiltro = Target.Address
    End If

    ' Si se salio de una celda de filtro -> ejecutar busqueda
    If eraCelda <> "" And Target.Address <> eraCelda Then
        bBusy = True
        Application.EnableEvents = False
        On Error GoTo FinSC
        Call EjecutarBusqueda
        GoTo FinSC
    End If

    ' --- Boton WA (col J = 10) fila 7+ -----------------------
    If Target.Column = 10 And Target.Row >= BUSC_FILA_DATOS Then
        If Trim(CStr(Target.Value)) = ChrW(&H25B6) & " WA" Then
            If Trim(CStr(Me.Cells(Target.Row, 2).Value)) <> "" Then
                bBusy = True
                Application.EnableEvents = False
                On Error GoTo FinSC
                Call EnviarWADesdeBuscador(Target.Row)
                GoTo FinSC
            End If
        End If
    End If

    ' --- Boton PDF (col K = 11) fila 7+ ----------------------
    If Target.Column = 11 And Target.Row >= BUSC_FILA_DATOS Then
        If Trim(CStr(Target.Value)) = ChrW(&H25A0) & " PDF" Then
            If Trim(CStr(Me.Cells(Target.Row, 2).Value)) <> "" Then
                bBusy = True
                Application.EnableEvents = False
                On Error GoTo FinSC
                Call GenerarPDFDesdeBuscador(Target.Row)
                GoTo FinSC
            End If
        End If
    End If

FinSC:
    bBusy = False
    Application.EnableEvents = True
End Sub

'================================================================
'  Worksheet_Change -- auto-regenerar celdas borradas
'================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    If bBusy Then Exit Sub
    If Target.CountLarge > 50 Then Exit Sub

    bBusy = True
    Application.EnableEvents = False
    On Error GoTo FinCh

    Dim NL As Long: NL = Target.Row
    Dim NC As Long: NC = Target.Column

    ' --- K3 vaciada -> restaurar boton LIMPIAR ----------------
    If Target.Address = "$K$3" Then
        If Trim(CStr(Target.Value)) = "" Then
            Call InicializarBotonK3_Buscador(Me)
        End If
        GoTo FinCh
    End If

    ' --- H3 vaciada -> restaurar activador BUSCAR -------------
    If Target.Address = "$H$3" Then
        If Trim(CStr(Target.Value)) = "" Then
            Call RestaurarActivadorH3
        End If
        GoTo FinCh
    End If

    ' --- Col J (WA) vaciada fila 7+ -> restaurar boton -------
    If NC = 10 And NL >= BUSC_FILA_DATOS Then
        If Trim(CStr(Target.Value)) = "" Then
            If Trim(CStr(Me.Cells(NL, 2).Value)) <> "" Then
                With Me.Cells(NL, 10)
                    .Value               = ChrW(&H25B6) & " WA"
                    .Interior.Color      = RGB(198, 224, 180)
                    .Font.Bold           = True
                    .HorizontalAlignment = xlCenter
                End With
            End If
        End If
        GoTo FinCh
    End If

    ' --- Col K (PDF) vaciada fila 7+ -> restaurar boton ------
    If NC = 11 And NL >= BUSC_FILA_DATOS Then
        If Trim(CStr(Target.Value)) = "" Then
            If Trim(CStr(Me.Cells(NL, 2).Value)) <> "" Then
                With Me.Cells(NL, 11)
                    .Value               = ChrW(&H25A0) & " PDF"
                    .Interior.Color      = RGB(189, 215, 238)
                    .Font.Bold           = True
                    .HorizontalAlignment = xlCenter
                End With
            End If
        End If
        GoTo FinCh
    End If

FinCh:
    bBusy = False
    Application.EnableEvents = True
End Sub

'================================================================
'  Worksheet_Activate -- refresca dropdown de Regimen (C3)
'================================================================
Private Sub Worksheet_Activate()
    If Not bBusy Then
        On Error Resume Next
        Call PopularDropdownsRegimen
        On Error GoTo 0
    End If
End Sub

'================================================================
'  RestaurarActivadorH3 -- escribe/restaura el cell-btn H3
'================================================================
Private Sub RestaurarActivadorH3()
    With Me.Range("H3")
        .Value               = "BUSCAR " & ChrW(&H25B6)
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .Font.Bold           = True
        .Font.Size           = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment   = xlCenter
    End With
End Sub

'================================================================
'  EscribirEncabezados -- fila 6 (fijos, no se borran)
'================================================================
Private Sub EscribirEncabezados()
    Dim bEv As Boolean: bEv = Application.EnableEvents
    Application.EnableEvents = False
    Dim hdr(1 To 11) As String
    hdr(1) = "No":          hdr(2)  = "Cliente":      hdr(3)  = "Responsable"
    hdr(4) = "RFC":         hdr(5)  = "R" & ChrW(&HE9) & "gimen"
    hdr(6) = "Concepto":    hdr(7)  = "Monto":        hdr(8)  = "Vencimiento"
    hdr(9) = "Estatus":     hdr(10) = ChrW(&H25B6) & "WA"
    hdr(11) = ChrW(&H25A0) & "PDF"
    Dim c As Integer
    For c = 1 To 11
        With Me.Cells(BUSC_FILA_HEADERS, c)
            .Value               = hdr(c)
            .Interior.Color      = RGB(68, 114, 196)
            .Font.Color          = RGB(255, 255, 255)
            .Font.Bold           = True
            .HorizontalAlignment = xlCenter
        End With
    Next c
    Application.EnableEvents = bEv
End Sub
