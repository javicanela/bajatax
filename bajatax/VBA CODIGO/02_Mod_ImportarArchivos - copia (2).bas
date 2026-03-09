Attribute VB_Name = "Mod_ImportarArchivos"
'================================================================
' MODULO: Mod_ImportarArchivos
' DONDE INSTALARLO:
'   Editor VBA > Insertar > Modulo > Renombrar a "Mod_ImportarArchivos"
'
' QUE HACE:
'   Permite importar datos desde multiples archivos Excel externos
'   (.xlsx, .xls, .xlsm) hacia la hoja REGISTROS.
'
' FLUJO:
'   1. Abre un FileDialog para seleccionar uno o mas archivos
'   2. Para cada archivo:
'      a. Abre el workbook en modo solo-lectura
'      b. Auto-detecta las columnas por nombre de encabezado
'      c. Lee todas las filas de datos
'      d. Las inserta en REGISTROS respetando el formato
'   3. Cierra los archivos importados
'   4. Muestra resumen de importacion
'
' MAPEO DE COLUMNAS:
'   El sistema busca estas palabras clave en los encabezados:
'     "NOMBRE" o "CONTRIBUYENTE" -> col B (Nombre)
'     "RFC"                       -> col C (RFC)
'     "EMAIL" o "CORREO"         -> col D (Email)
'     "TELEFONO" o "TEL"         -> col E (Telefono)
'     "FECHA"                     -> col F (Fecha)
'     "CONCEPTO" o "SERVICIO"    -> col G (Concepto)
'     "MONTO" o "IMPORTE"        -> col H (Monto)
'     "FACTURA"                   -> col I (Factura)
'     "REGIMEN" o "REGIMEN"      -> col J (Regimen)
'     "VENCIMIENTO" o "VENCE"    -> col K (Vencimiento)
'     "RESPONSABLE"               -> col A (Responsable)
'================================================================

Option Explicit

' ---------------------------------------------------------------
'  Estructura para almacenar el mapeo de columnas
' ---------------------------------------------------------------
Private Type ColMap
    colNombre As Integer
    colRFC As Integer
    colEmail As Integer
    colTelefono As Integer
    colFecha As Integer
    colConcepto As Integer
    colMonto As Integer
    colFactura As Integer
    colRegimen As Integer
    colVencimiento As Integer
    colResponsable As Integer
    hayMapeo As Boolean
End Type

' ---------------------------------------------------------------
'  PUNTO DE ENTRADA: ImportarArchivosExternos
'  Asignarlo al boton "IMPORTAR ARCHIVOS" en REGISTROS
' ---------------------------------------------------------------
Public Sub ImportarArchivosExternos()

    If Not HojasOK() Then Exit Sub

    Dim wsReg As Worksheet
    Set wsReg = ObtenerHoja("REGISTROS")
    If wsReg Is Nothing Then Exit Sub

    ' --- Abrir dialogo para seleccionar archivos ---
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "BajaTax -- Seleccionar archivos Excel a importar"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xlsx; *.xls; *.xlsm; *.csv"
        .Filters.Add "Todos los archivos", "*.*"

        ' Intentar abrir en carpeta IMPORTAR si existe
        Dim sCarpetaImportar As String
        sCarpetaImportar = ThisWorkbook.Path & Application.PathSeparator & "IMPORTAR"
        On Error Resume Next
        If Dir(sCarpetaImportar, vbDirectory) <> "" Then
            .InitialFileName = sCarpetaImportar & Application.PathSeparator
        End If
        On Error GoTo 0

        If .Show = 0 Then
            ' Usuario cancelo
            Exit Sub
        End If

        ' Confirmar importacion
        Dim respConf As Integer
        respConf = MsgBox("Se importaran " & .SelectedItems.Count & " archivo(s)." & Chr(13) & Chr(13) & _
                          "Los datos se agregaran al final de REGISTROS." & Chr(13) & _
                          "El sistema detectara automaticamente las columnas." & Chr(13) & Chr(13) & _
                          "Deseas continuar?", _
                          vbYesNo + vbQuestion, "BajaTax -- Confirmar Importacion")
        If respConf = vbNo Then Exit Sub

        Application.ScreenUpdating = False
        Application.EnableEvents = False

        On Error GoTo ErrorImport

        Dim totalImportados As Long
        Dim totalArchivos As Integer
        Dim archivosError As String
        totalImportados = 0
        totalArchivos = 0
        archivosError = ""

        Dim k As Integer
        For k = 1 To .SelectedItems.Count
            Dim sRuta As String
            sRuta = .SelectedItems(k)

            Dim resultado As Long
            resultado = ImportarUnArchivo(sRuta, wsReg)

            If resultado >= 0 Then
                totalArchivos = totalArchivos + 1
                totalImportados = totalImportados + resultado
            Else
                archivosError = archivosError & "  - " & ExtraerNombreArchivo(sRuta) & Chr(13)
            End If
        Next k

        Application.ScreenUpdating = True
        Application.EnableEvents = True

        ' Resumen
        Dim sMensaje As String
        sMensaje = "Importacion completada." & Chr(13) & Chr(13) & _
                   "  Archivos procesados: " & totalArchivos & Chr(13) & _
                   "  Registros importados: " & totalImportados

        If archivosError <> "" Then
            sMensaje = sMensaje & Chr(13) & Chr(13) & _
                       "Archivos con error:" & Chr(13) & archivosError
        End If

        sMensaje = sMensaje & Chr(13) & Chr(13) & _
                   "Ahora presiona 'PROCESAR REGISTROS' para enviar " & _
                   "los datos a OPERACIONES y DIRECTORIO."

        MsgBox sMensaje, vbInformation, "BajaTax -- Importacion"
    End With

    Exit Sub

ErrorImport:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error durante la importacion: " & Err.Description & Chr(13) & _
           "Error #" & Err.Number, vbCritical, "BajaTax -- Error"
End Sub

' ---------------------------------------------------------------
'  Importa un archivo Excel individual
'  Retorna: numero de filas importadas, o -1 si hubo error
' ---------------------------------------------------------------
Private Function ImportarUnArchivo(sRuta As String, wsReg As Worksheet) As Long

    On Error GoTo ErrorArchivo

    ' Abrir archivo en modo solo lectura
    Dim wbExt As Workbook
    Set wbExt = Workbooks.Open(Filename:=sRuta, ReadOnly:=True, UpdateLinks:=0)

    ' Buscar la primera hoja con datos
    Dim wsExt As Worksheet
    Dim mejorHoja As Worksheet
    Dim maxFilas As Long
    maxFilas = 0

    Dim sh As Worksheet
    For Each sh In wbExt.Sheets
        Dim uF As Long
        uF = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
        If uF > maxFilas Then
            maxFilas = uF
            Set mejorHoja = sh
        End If
    Next sh

    If mejorHoja Is Nothing Or maxFilas < 2 Then
        wbExt.Close SaveChanges:=False
        ImportarUnArchivo = 0
        Exit Function
    End If

    Set wsExt = mejorHoja

    ' Auto-detectar columnas por encabezados
    Dim mapa As ColMap
    mapa = DetectarColumnas(wsExt)

    If Not mapa.hayMapeo Then
        ' Intentar con fila 2 como encabezados
        mapa = DetectarColumnasEnFila(wsExt, 2)
    End If

    ' Si no hay mapeo, intentar mapeo posicional por defecto
    If Not mapa.hayMapeo Then
        ' Usar mapeo generico: A=Nombre, B=RFC, etc.
        mapa.colNombre = 1
        mapa.colRFC = 2
        mapa.colEmail = 3
        mapa.colTelefono = 4
        mapa.colFecha = 5
        mapa.colConcepto = 6
        mapa.colMonto = 7
        mapa.hayMapeo = True
    End If

    ' Determinar fila de inicio de datos (despues de encabezados)
    Dim filaInicio As Long
    filaInicio = 2  ' Por defecto, fila 2

    ' Verificar si fila 1 parece encabezado
    Dim primerVal As String
    primerVal = UCase(Trim(CStr(wsExt.Cells(1, 1).Value)))
    If InStr(primerVal, "NOMBRE") > 0 Or InStr(primerVal, "RFC") > 0 Or _
       InStr(primerVal, "RESPONSABLE") > 0 Or InStr(primerVal, "CLIENTE") > 0 Or _
       InStr(primerVal, "CONTRIBUYENTE") > 0 Or IsNumeric(wsExt.Cells(2, 1).Value) Then
        filaInicio = 2
    ElseIf Len(primerVal) > 0 And Not IsNumeric(primerVal) Then
        filaInicio = 2
    End If

    ' Leer datos y agregar a REGISTROS
    Dim uFilaExt As Long
    uFilaExt = wsExt.Cells(wsExt.Rows.Count, 1).End(xlUp).Row
    ' Tambien verificar otras columnas
    Dim uFC As Long
    If mapa.colNombre > 0 Then
        uFC = wsExt.Cells(wsExt.Rows.Count, mapa.colNombre).End(xlUp).Row
        If uFC > uFilaExt Then uFilaExt = uFC
    End If

    Dim contImportados As Long
    contImportados = 0

    Dim i As Long
    For i = filaInicio To uFilaExt
        ' Verificar que la fila tenga al menos un dato relevante
        Dim sNombreExt As String
        sNombreExt = ""
        If mapa.colNombre > 0 Then
            sNombreExt = Trim(CStr(wsExt.Cells(i, mapa.colNombre).Value))
        End If

        Dim sRFCExt As String
        sRFCExt = ""
        If mapa.colRFC > 0 Then
            sRFCExt = Trim(CStr(wsExt.Cells(i, mapa.colRFC).Value))
        End If

        ' Saltar filas vacias
        If sNombreExt = "" And sRFCExt = "" Then GoTo SiguienteFilaExt

        ' Encontrar siguiente fila vacia en REGISTROS
        Dim uFilaReg As Long
        uFilaReg = wsReg.Cells(wsReg.Rows.Count, COL_REG_NOMBRE).End(xlUp).Row + 1

        ' Escribir datos en REGISTROS
        If mapa.colResponsable > 0 Then
            wsReg.Cells(uFilaReg, COL_REG_RESPONSABLE).Value = _
                Trim(CStr(wsExt.Cells(i, mapa.colResponsable).Value))
        End If

        wsReg.Cells(uFilaReg, COL_REG_NOMBRE).Value = sNombreExt
        wsReg.Cells(uFilaReg, COL_REG_RFC).Value = sRFCExt

        If mapa.colEmail > 0 Then
            wsReg.Cells(uFilaReg, COL_REG_EMAIL).Value = _
                Trim(CStr(wsExt.Cells(i, mapa.colEmail).Value))
        End If

        If mapa.colTelefono > 0 Then
            Dim vTel As Variant
            vTel = wsExt.Cells(i, mapa.colTelefono).Value
            If IsNumeric(vTel) Then
                wsReg.Cells(uFilaReg, COL_REG_TELEFONO).Value = CLng(vTel)
            Else
                wsReg.Cells(uFilaReg, COL_REG_TELEFONO).Value = Trim(CStr(vTel))
            End If
        End If

        If mapa.colFecha > 0 Then
            Dim vFecha As Variant
            vFecha = wsExt.Cells(i, mapa.colFecha).Value
            If IsDate(vFecha) Then
                wsReg.Cells(uFilaReg, COL_REG_FECHA).Value = CDate(vFecha)
            Else
                wsReg.Cells(uFilaReg, COL_REG_FECHA).Value = vFecha
            End If
        End If

        If mapa.colConcepto > 0 Then
            wsReg.Cells(uFilaReg, COL_REG_CONCEPTO).Value = _
                Trim(CStr(wsExt.Cells(i, mapa.colConcepto).Value))
        End If

        If mapa.colMonto > 0 Then
            Dim vMonto As Variant
            vMonto = wsExt.Cells(i, mapa.colMonto).Value
            On Error Resume Next
            wsReg.Cells(uFilaReg, COL_REG_MONTO).Value = CDbl(vMonto)
            On Error GoTo 0
        End If

        If mapa.colFactura > 0 Then
            wsReg.Cells(uFilaReg, COL_REG_FACTURA).Value = _
                Trim(CStr(wsExt.Cells(i, mapa.colFactura).Value))
        End If

        If mapa.colRegimen > 0 Then
            wsReg.Cells(uFilaReg, COL_REG_REGIMEN).Value = _
                Trim(CStr(wsExt.Cells(i, mapa.colRegimen).Value))
        End If

        If mapa.colVencimiento > 0 Then
            Dim vVenc As Variant
            vVenc = wsExt.Cells(i, mapa.colVencimiento).Value
            If IsDate(vVenc) Then
                wsReg.Cells(uFilaReg, COL_REG_VENCIMIENTO).Value = CDate(vVenc)
            Else
                wsReg.Cells(uFilaReg, COL_REG_VENCIMIENTO).Value = vVenc
            End If
        End If

        ' Marcar origen del registro
        wsReg.Cells(uFilaReg, COL_REG_PROCESADO).Value = "IMPORTADO: " & ExtraerNombreArchivo(sRuta)
        wsReg.Cells(uFilaReg, COL_REG_PROCESADO).Font.Color = RGB(0, 70, 127)
        wsReg.Cells(uFilaReg, COL_REG_PROCESADO).Font.Size = 8

        contImportados = contImportados + 1

SiguienteFilaExt:
    Next i

    wbExt.Close SaveChanges:=False
    ImportarUnArchivo = contImportados
    Exit Function

ErrorArchivo:
    On Error Resume Next
    If Not wbExt Is Nothing Then wbExt.Close SaveChanges:=False
    On Error GoTo 0
    ImportarUnArchivo = -1
End Function

' ---------------------------------------------------------------
'  Auto-detecta columnas por encabezados (fila 1)
' ---------------------------------------------------------------
Private Function DetectarColumnas(ws As Worksheet) As ColMap
    DetectarColumnas = DetectarColumnasEnFila(ws, 1)
End Function

' ---------------------------------------------------------------
'  Auto-detecta columnas por encabezados en una fila especifica
' ---------------------------------------------------------------
Private Function DetectarColumnasEnFila(ws As Worksheet, ByVal filaEnc As Long) As ColMap
    Dim mapa As ColMap
    mapa.hayMapeo = False

    Dim uCol As Integer
    uCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column

    Dim j As Integer
    For j = 1 To uCol
        Dim sEnc As String
        sEnc = UCase(Trim(CStr(ws.Cells(filaEnc, j).Value)))

        ' Limpiar caracteres especiales
        sEnc = Replace(sEnc, ChrW(&HC9), "E")
        sEnc = Replace(sEnc, ChrW(&HD3), "O")
        sEnc = Replace(sEnc, ChrW(&HCD), "I")
        sEnc = Replace(sEnc, ChrW(&HC1), "A")
        sEnc = Replace(sEnc, ChrW(&HDA), "U")
        sEnc = Replace(sEnc, ChrW(&HD1), "N")

        If InStr(sEnc, "CONTRIBUYENTE") > 0 Or _
           (InStr(sEnc, "NOMBRE") > 0 And InStr(sEnc, "ARCHIVO") = 0) Or _
           InStr(sEnc, "CLIENTE") > 0 Or _
           InStr(sEnc, "RAZON SOCIAL") > 0 Then
            mapa.colNombre = j
            mapa.hayMapeo = True

        ElseIf sEnc = "RFC" Or InStr(sEnc, "R.F.C") > 0 Then
            mapa.colRFC = j
            mapa.hayMapeo = True

        ElseIf InStr(sEnc, "EMAIL") > 0 Or InStr(sEnc, "CORREO") > 0 Or _
               InStr(sEnc, "E-MAIL") > 0 Then
            mapa.colEmail = j
            mapa.hayMapeo = True

        ElseIf InStr(sEnc, "TELEFONO") > 0 Or InStr(sEnc, "TEL") > 0 Or _
               InStr(sEnc, "CELULAR") > 0 Or InStr(sEnc, "WHATSAPP") > 0 Then
            mapa.colTelefono = j
            mapa.hayMapeo = True

        ElseIf InStr(sEnc, "CONCEPTO") > 0 Or InStr(sEnc, "SERVICIO") > 0 Or _
               InStr(sEnc, "DESCRIPCION") > 0 Then
            mapa.colConcepto = j
            mapa.hayMapeo = True

        ElseIf InStr(sEnc, "MONTO") > 0 Or InStr(sEnc, "IMPORTE") > 0 Or _
               InStr(sEnc, "TOTAL") > 0 Or InStr(sEnc, "SALDO") > 0 Then
            mapa.colMonto = j
            mapa.hayMapeo = True

        ElseIf InStr(sEnc, "FACTURA") > 0 Or InStr(sEnc, "FOLIO") > 0 Or _
               InStr(sEnc, "CFDI") > 0 Then
            mapa.colFactura = j
            mapa.hayMapeo = True

        ElseIf InStr(sEnc, "REGIMEN") > 0 Or sEnc = "PF" Or sEnc = "PM" Then
            mapa.colRegimen = j
            mapa.hayMapeo = True

        ElseIf InStr(sEnc, "VENCIMIENTO") > 0 Or InStr(sEnc, "VENCE") > 0 Or _
               InStr(sEnc, "LIMITE") > 0 Then
            mapa.colVencimiento = j
            mapa.hayMapeo = True

        ElseIf InStr(sEnc, "FECHA") > 0 And mapa.colFecha = 0 Then
            ' Solo el primer campo "FECHA" generico
            mapa.colFecha = j
            mapa.hayMapeo = True

        ElseIf InStr(sEnc, "RESPONSABLE") > 0 Or InStr(sEnc, "ASIGNADO") > 0 Or _
               InStr(sEnc, "SOPORTE") > 0 Then
            mapa.colResponsable = j
            mapa.hayMapeo = True
        End If
    Next j

    DetectarColumnasEnFila = mapa
End Function

' ---------------------------------------------------------------
'  Extrae solo el nombre del archivo de una ruta completa
' ---------------------------------------------------------------
Private Function ExtraerNombreArchivo(sRuta As String) As String
    Dim partes As Variant
    If InStr(sRuta, "\") > 0 Then
        partes = Split(sRuta, "\")
    ElseIf InStr(sRuta, "/") > 0 Then
        partes = Split(sRuta, "/")
    Else
        ExtraerNombreArchivo = sRuta
        Exit Function
    End If
    ExtraerNombreArchivo = partes(UBound(partes))
End Function

'================================================================
' SincronizarEdicionRegistros
' Propaga una edicion hecha en REGISTROS hacia OPERACIONES y DIRECTORIO
'
' Parametros:
'   editRow  - fila de REGISTROS que fue editada
'   editCol  - columna de REGISTROS que fue editada (1-14)
'   nuevoVal - nuevo valor ingresado por el usuario
'   oldVal   - valor original antes de la edicion (para referencia)
'
' Clave de busqueda: RFC (col E=5) + ID_Factura (col B=2)
'
' Mapeo REGISTROS -> DIRECTORIO:
'   Col 4 (D=Cliente)   -> COL_DIR_CLIENTE (2)
'   Col 5 (E=RFC)       -> COL_DIR_RFC     (1)
'   Col 13 (M=Telefono) -> COL_DIR_NUMERO  (4)
'   Col 14 (N=Correo)   -> COL_DIR_CORREO  (3)
'   Col 3  (C=Regimen)  -> COL_DIR_REGIMEN (5)
'================================================================
Public Sub SincronizarEdicionRegistros(editRow As Long, editCol As Integer, _
                                        nuevoVal As String, oldVal As String)
    Dim wsReg As Worksheet: Set wsReg = ObtenerHoja("REGISTROS")
    Dim wsOp  As Worksheet: Set wsOp  = ObtenerHoja("OPERACIONES")
    Dim wsDir As Worksheet: Set wsDir = ObtenerHoja("DIRECTORIO")

    If wsReg Is Nothing Or wsOp Is Nothing Then Exit Sub

    ' Clave compuesta: RFC (col 5) + ID_Factura (col 2)
    Dim sRFC As String: sRFC = Trim(CStr(wsReg.Cells(editRow, 5).Value))
    Dim sID  As String: sID  = Trim(CStr(wsReg.Cells(editRow, 2).Value))

    If sRFC = "" And sID = "" Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo FinSync

    ' -- 1. Actualizar OPERACIONES -------------------------------------
    Dim uOp As Long
    uOp = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row

    Dim i As Long
    For i = 2 To uOp
        Dim oRFC As String: oRFC = Trim(CStr(wsOp.Cells(i, COL_OP_RFC).Value))
        Dim oID  As String: oID  = Trim(CStr(wsOp.Cells(i, COL_OP_ID_FACTURA).Value))

        Dim coincide As Boolean
        If sRFC <> "" And sID <> "" Then
            coincide = (UCase(oRFC) = UCase(sRFC) And UCase(oID) = UCase(sID))
        ElseIf sRFC <> "" Then
            coincide = (UCase(oRFC) = UCase(sRFC))
        Else
            coincide = (UCase(oID) = UCase(sID))
        End If

        If coincide Then
            ' Las columnas de REGISTROS A-N (1-14) coinciden 1:1 con OPERACIONES A-N
            wsOp.Cells(i, editCol).Value = nuevoVal
            Exit For
        End If
    Next i

    ' -- 2. Actualizar DIRECTORIO si la columna aplica ------------------
    If wsDir Is Nothing Then GoTo FinSync
    If sRFC = "" Then GoTo FinSync   ' sin RFC no podemos identificar en DIRECTORIO

    Dim dirCol As Integer: dirCol = 0
    Select Case editCol
        Case 4:  dirCol = COL_DIR_CLIENTE   ' D (Cliente/Nombre) -> B
        Case 5:  dirCol = COL_DIR_RFC        ' E (RFC)            -> A
        Case 13: dirCol = COL_DIR_NUMERO    ' M (Telefono)       -> D
        Case 14: dirCol = COL_DIR_CORREO    ' N (Correo)         -> C
        Case 3:  dirCol = COL_DIR_REGIMEN   ' C (Regimen)        -> E
    End Select

    If dirCol > 0 Then
        Dim uDir As Long
        uDir = wsDir.Cells(wsDir.Rows.Count, COL_DIR_RFC).End(xlUp).Row
        Dim j As Long
        For j = 2 To uDir
            If UCase(Trim(CStr(wsDir.Cells(j, COL_DIR_RFC).Value))) = UCase(sRFC) Then
                wsDir.Cells(j, dirCol).Value = nuevoVal
                Exit For
            End If
        Next j
    End If

FinSync:
    Application.EnableEvents = True
End Sub

'================================================================
' SincronizarEdicionDirectorio
' Propaga una edicion hecha en DIRECTORIO hacia OPERACIONES y REGISTROS
'
' Parametros:
'   editRow  - fila de DIRECTORIO que fue editada
'   editCol  - columna de DIRECTORIO que fue editada (1-8)
'   nuevoVal - nuevo valor ingresado por el usuario
'   oldVal   - valor original antes de la edicion
'
' Clave de busqueda: RFC (col A=1 del DIRECTORIO)
' Actualiza TODAS las filas que coincidan con ese RFC.
'
' Mapeo DIRECTORIO -> OPERACIONES:
'   Col 1 (A=RFC)          -> col 5 (E=RFC)
'   Col 2 (B=CLIENTE)      -> col 4 (D=CLIENTE)
'   Col 3 (C=CORREO)       -> col 14 (N=CORREO)
'   Col 4 (D=NUMERO)       -> col 13 (M=TELEFONO)
'   Col 5 (E=REGIMEN)      -> col 3 (C=REGIMEN)
'   Col 6 (F=RESPONSABLE)  -> col 1 (A=RESPONSABLE)
'
' Mapeo DIRECTORIO -> REGISTROS:
'   Col 2 (B=CLIENTE)      -> col 2 (B=NOMBRE)
'   Col 3 (C=CORREO)       -> col 4 (D=EMAIL)
'   Col 4 (D=NUMERO)       -> col 5 (E=TELEFONO)
'   Col 5 (E=REGIMEN)      -> col 10 (J=REGIMEN)
'   Col 6 (F=RESPONSABLE)  -> col 1 (A=RESPONSABLE)
'================================================================
Public Sub SincronizarEdicionDirectorio(editRow As Long, editCol As Integer, _
                                         nuevoVal As String, oldVal As String)
    Dim wsDir As Worksheet: Set wsDir = ObtenerHoja("DIRECTORIO")
    Dim wsOp  As Worksheet: Set wsOp  = ObtenerHoja("OPERACIONES")
    Dim wsReg As Worksheet: Set wsReg = ObtenerHoja("REGISTROS")

    If wsDir Is Nothing Or wsOp Is Nothing Then Exit Sub

    ' Clave: RFC (col 1 de DIRECTORIO)
    Dim sRFC As String
    If editCol = 1 Then
        ' Si se edito el RFC, usar el valor ANTERIOR para buscar
        sRFC = UCase(Trim(oldVal))
    Else
        sRFC = UCase(Trim(CStr(wsDir.Cells(editRow, COL_DIR_RFC).Value)))
    End If

    If sRFC = "" Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo FinSyncD

    ' --- Mapeo DIRECTORIO col -> OPERACIONES col ---
    Dim opCol As Integer: opCol = 0
    Select Case editCol
        Case 1: opCol = COL_OP_RFC          ' A -> E
        Case 2: opCol = COL_OP_CLIENTE      ' B -> D
        Case 3: opCol = COL_OP_CORREO       ' C -> N
        Case 4: opCol = COL_OP_TELEFONO     ' D -> M
        Case 5: opCol = COL_OP_REGIMEN      ' E -> C
        Case 6: opCol = COL_OP_RESPONSABLE  ' F -> A
    End Select

    ' --- 1. Actualizar OPERACIONES (TODAS las filas con ese RFC) ---
    If opCol > 0 Then
        Dim uOp As Long
        uOp = wsOp.Cells(wsOp.Rows.Count, COL_OP_CLIENTE).End(xlUp).Row
        Dim i As Long
        For i = 2 To uOp
            If UCase(Trim(CStr(wsOp.Cells(i, COL_OP_RFC).Value))) = sRFC Then
                wsOp.Cells(i, opCol).Value = nuevoVal
            End If
        Next i
    End If

    ' --- Mapeo DIRECTORIO col -> REGISTROS col ---
    Dim regCol As Integer: regCol = 0
    Select Case editCol
        Case 2: regCol = COL_REG_NOMBRE     ' B -> B
        Case 3: regCol = COL_REG_EMAIL      ' C -> D
        Case 4: regCol = COL_REG_TELEFONO   ' D -> E
        Case 5: regCol = COL_REG_REGIMEN    ' E -> J
        Case 6: regCol = COL_REG_RESPONSABLE ' F -> A
    End Select

    ' --- 2. Actualizar REGISTROS (TODAS las filas con ese RFC) ---
    If regCol > 0 And Not wsReg Is Nothing Then
        Dim uReg As Long
        uReg = wsReg.Cells(wsReg.Rows.Count, COL_REG_NOMBRE).End(xlUp).Row
        Dim j As Long
        For j = 2 To uReg
            If UCase(Trim(CStr(wsReg.Cells(j, COL_REG_RFC).Value))) = sRFC Then
                wsReg.Cells(j, regCol).Value = nuevoVal
            End If
        Next j
    End If

FinSyncD:
    Application.EnableEvents = True
End Sub
