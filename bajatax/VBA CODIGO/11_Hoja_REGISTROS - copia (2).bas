Attribute VB_Name = "Hoja_REGISTROS"
'================================================================
' CODIGO DE HOJA: REGISTROS -- BajaTax v4 FINAL  (MODULO NUEVO)
'
' La hoja REGISTROS es de solo lectura por defecto.
' Edicion requiere confirmacion por doble-clic.
'
' Columnas A-N (espejo de OPERACIONES):
'   A=Responsable  B=ID_Factura  C=Regimen    D=Cliente    E=RFC
'   F=Fecha_Cob    G=Concepto    H=Monto      I=Estatus    J=Vencimiento
'   K=Dias_Venc    L=Reg_Pago    M=Telefono   N=Correo
'   Z=Botones accion (Z1-Z5)
'
' Clave de sincronizacion: RFC (col E) + ID_Factura (col B)
'================================================================
Option Explicit

' --- Variables de control de edicion ---------------------------
Private mEditando   As Boolean   ' True = edicion autorizada en curso
Private mEditRow    As Long       ' Fila bajo edicion
Private mEditCol    As Integer    ' Columna bajo edicion
Private mEditOldVal As String     ' Valor original antes de editar

' --- Constantes internas de columnas REGISTROS ----------------
Private Const REG_RESP  As Integer = 1   ' A Responsable
Private Const REG_ID    As Integer = 2   ' B ID_Factura  <- clave
Private Const REG_REG   As Integer = 3   ' C Regimen
Private Const REG_CLI   As Integer = 4   ' D Cliente
Private Const REG_RFC   As Integer = 5   ' E RFC         <- clave
Private Const REG_FECHA As Integer = 6   ' F Fecha_Cob
Private Const REG_CONC  As Integer = 7   ' G Concepto
Private Const REG_MON   As Integer = 8   ' H Monto
Private Const REG_EST   As Integer = 9   ' I Estatus
Private Const REG_VENC  As Integer = 10  ' J Vencimiento
Private Const REG_DIAS  As Integer = 11  ' K Dias_Venc
Private Const REG_PAGO  As Integer = 12  ' L Reg_Pago
Private Const REG_TEL   As Integer = 13  ' M Telefono
Private Const REG_COR   As Integer = 14  ' N Correo

'================================================================
'  InicializarBotonesZ_REGISTROS -- escribe/restaura Z1-Z5
'  (llamado por Bootstrap y por Worksheet_Change)
'================================================================
Public Sub InicializarBotonesZ_REGISTROS()
    Dim bEv As Boolean: bEv = Application.EnableEvents
    Application.EnableEvents = False

    Dim textos(1 To 5) As String
    textos(1) = ChrW(&H25B6) & " IMPORTAR"
    textos(2) = ChrW(&H25B6) & " PROCESAR TODO"
    textos(3) = ChrW(&H25B6) & " ENVIO MASIVO WA"
    textos(4) = ChrW(&H25A0) & " PDF MASIVO"
    textos(5) = ChrW(&H21BA) & " REGENERAR"

    Dim i As Integer
    For i = 1 To 5
        With Me.Cells(i, 26)
            .Value               = textos(i)
            .Interior.Color      = RGB(68, 114, 196)
            .Font.Color          = RGB(255, 255, 255)
            .Font.Bold           = True
            .Font.Size           = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment   = xlCenter
            .WrapText            = False
        End With
    Next i

    ' Encabezado N1 = ESTATUS
    With Me.Cells(1, 14)
        .Value          = "ESTATUS"
        .Font.Bold      = True
        .Font.Color     = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With

    ' Aplicar formato visual a los valores de ESTATUS (columna N)
    Call AplicarColoresEstatusRegistros

    Application.EnableEvents = bEv
End Sub

'================================================================
'  AplicarColoresEstatusRegistros
'  - Resalta filas cuyo ESTATUS (col N) sea "OMITIDO"
'    con un fondo naranja tenue y texto ligeramente mas oscuro.
'  - No modifica formulas ni logica, solo formato visual.
'================================================================
Public Sub AplicarColoresEstatusRegistros()
    Dim uFila As Long
    Dim i As Long
    Dim sVal As String

    uFila = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row

    For i = 2 To uFila
        sVal = UCase$(Trim$(CStr(Me.Cells(i, REG_COR).Value))) ' REG_COR = col 14 (N)

        With Me.Cells(i, REG_COR)
            If sVal = "OMITIDO" Then
                .Interior.Color = RGB(255, 229, 204)   ' naranja tenue
                .Font.Color     = RGB(191, 97, 0)      ' texto naranja mas oscuro
                .Font.Bold      = True
            ElseIf sVal <> "" Then
                ' Otros estatus: formato neutro, pero respetando posibles formatos previos
                ' (no se limpia explicitamente para no interferir con reglas existentes)
            Else
                ' Celda vacia: quitar formato especifico de "OMITIDO"
                .Interior.ColorIndex = xlColorIndexNone
                .Font.ColorIndex     = xlColorIndexAutomatic
                .Font.Bold           = False
            End If
        End With
    Next i
End Sub

'================================================================
'  BeforeDoubleClick -- Z1-Z5 (macros) y datos A-N (edicion ctrl.)
'================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.CountLarge > 1 Then Exit Sub

    Dim NL As Long: NL = Target.Row
    Dim NC As Long: NC = Target.Column

    '--- Botones Z (col 26, filas 1-5) -------------------------
    If NC = 26 And NL >= 1 And NL <= 5 Then
        Cancel = True
        On Error GoTo FinDC
        Select Case NL
            Case 1: Call ImportarArchivosExternos
            Case 2: Call ProcesarTodoBajaTax
            Case 3: Call EnvioMasivoAutomatico
            Case 4: Call GenerarPDFMasivo
            Case 5: Call RegenerarFaltantes
        End Select
        Exit Sub
    End If

    '--- Edicion controlada: cols A-N, fila 2+ ----------------
    If NL < 2 Then Exit Sub
    If NC < 1 Or NC > 14 Then Exit Sub
    If Trim(CStr(Target.Value)) = "" Then Exit Sub   ' celda vacia -> no bloquear

    Cancel = True   ' bloquear por defecto hasta confirmar

    Dim resp As Integer
    resp = MsgBox("Deseas actualizar este dato del cliente?" & Chr(13) & Chr(13) & _
                  "  Campo:   " & NombreColumnaReg(NC) & Chr(13) & _
                  "  Valor:   " & CStr(Target.Value), _
                  vbYesNo + vbQuestion, "BajaTax - Editar Registro")

    If resp = vbYes Then
        mEditRow    = NL
        mEditCol    = NC
        mEditOldVal = CStr(Target.Value)
        mEditando   = True
        Cancel      = False   ' permitir edicion
    End If

    Exit Sub

FinDC:
    Err.Clear
End Sub

'================================================================
'  Worksheet_Change -- post-edicion: sincronizar o restaurar
'================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim NL As Long: NL = Target.Row
    Dim NC As Long: NC = Target.Column

    '--- Botones Z vaciados -> restaurar -----------------------
    If NC = 26 And NL >= 1 And NL <= 5 Then
        If Trim(CStr(Target.Value)) = "" Then
            Application.EnableEvents = False
            Call InicializarBotonesZ_REGISTROS
            Application.EnableEvents = True
        End If
        Exit Sub
    End If

    '--- Solo procesar si hay edicion autorizada en curso -----
    If Not mEditando Then Exit Sub
    If Target.Row <> mEditRow Or Target.Column <> mEditCol Then
        mEditando = False
        Exit Sub
    End If

    mEditando = False
    Dim nuevoVal As String: nuevoVal = Trim(CStr(Target.Value))

    Application.EnableEvents = False
    On Error GoTo FinCh

    If nuevoVal = "" Then
        '--- Celda vaciada: ofrecer restaurar desde OPERACIONES ----
        Dim resp As Integer
        Application.EnableEvents = True
        resp = MsgBox("La celda quedo vacia." & Chr(13) & Chr(13) & _
                      "Restaurar el valor original desde OPERACIONES?" & Chr(13) & _
                      "  (" & NombreColumnaReg(mEditCol) & " = """ & mEditOldVal & """)", _
                      vbYesNo + vbQuestion, "BajaTax - Celda Vaciada")
        Application.EnableEvents = False

        If resp = vbYes Then
            Target.Value = mEditOldVal
        End If
        ' Si No: dejar vacio -- inconsistencia elegida por el usuario

    Else
        '--- Valor nuevo: sincronizar a OPERACIONES + DIRECTORIO ---
        Call SincronizarEdicionRegistros(mEditRow, mEditCol, nuevoVal, mEditOldVal)
    End If

FinCh:
    Application.EnableEvents = True
End Sub

'================================================================
'  NombreColumnaReg -- devuelve nombre legible de columna
'================================================================
Private Function NombreColumnaReg(NC As Integer) As String
    Select Case NC
        Case 1:  NombreColumnaReg = "Responsable"
        Case 2:  NombreColumnaReg = "ID_Factura"
        Case 3:  NombreColumnaReg = "Regimen"
        Case 4:  NombreColumnaReg = "Cliente"
        Case 5:  NombreColumnaReg = "RFC"
        Case 6:  NombreColumnaReg = "Fecha_Cob"
        Case 7:  NombreColumnaReg = "Concepto"
        Case 8:  NombreColumnaReg = "Monto"
        Case 9:  NombreColumnaReg = "Estatus"
        Case 10: NombreColumnaReg = "Vencimiento"
        Case 11: NombreColumnaReg = "Dias Venc"
        Case 12: NombreColumnaReg = "Reg_Pago"
        Case 13: NombreColumnaReg = "Telefono"
        Case 14: NombreColumnaReg = "Correo"
        Case Else: NombreColumnaReg = "Col " & NC
    End Select
End Function
