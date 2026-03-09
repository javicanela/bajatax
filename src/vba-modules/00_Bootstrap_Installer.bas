Attribute VB_Name = "Bootstrap_Installer"
Option Explicit
'===========================================================================
' BOOTSTRAP INSTALLER - BajaTax v4 FINAL
'
' INSTRUCCIONES:
'   1. Abre tu archivo .xlsm en Excel
'   2. Activa: Excel > Preferencias > Seguridad > "Confiar en el acceso
'              al modelo de objetos del proyecto de VBA"
'   3. Ve al Editor VBA: Herramientas > Macro > Editor de Visual Basic
'   4. Archivo > Importar Archivo > selecciona ESTE archivo
'   5. Ejecutar > Ejecutar Sub > selecciona InstalarSistema_BajaTax
'   6. Espera el mensaje de exito
'   7. Guarda como AUTOMATIZACION_v4_FINAL.xlsm
'
' Este modulo se AUTO-ELIMINA al terminar la instalacion.
'===========================================================================

Sub InstalarSistema_BajaTax()
    Dim vbp As Object
    Dim ruta As String
    Dim wsTemp As Worksheet

    ' --- 1. Verificar acceso al modelo VBA ---
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "PASO NECESARIO ANTES DE CONTINUAR:" & vbCrLf & vbCrLf & _
               "1. Ve a Excel > Preferencias" & vbCrLf & _
               "2. Haz clic en Seguridad" & vbCrLf & _
               "3. Marca la casilla:" & vbCrLf & _
               "   Confiar en el acceso al modelo de objetos" & vbCrLf & _
               "   del proyecto de VBA" & vbCrLf & vbCrLf & _
               "4. Cierra este cuadro y vuelve a ejecutar el macro.", _
               vbCritical + vbOKOnly, "Permiso requerido - BajaTax"
        Exit Sub
    End If
    On Error GoTo 0

    If vbp Is Nothing Then
        MsgBox "No se pudo acceder al proyecto VBA.", vbCritical, "Error"
        Exit Sub
    End If

    ' --- 2. Ruta base de los modulos ---
    ruta = "/Users/javieravila/Documents/AUTOMATIZACION/VBA_CODIGO/"

    If Dir(ruta, vbDirectory) = "" Then
        MsgBox "No se encontro la carpeta:" & vbCrLf & ruta & vbCrLf & vbCrLf & _
               "Verifica que los archivos .bas esten en esa ubicacion.", _
               vbCritical, "Carpeta no encontrada"
        Exit Sub
    End If

    ' --- 3. Modulos estandar (importar/reemplazar) ---
    Application.StatusBar = "BajaTax: Instalando Mod_Sistema..."
    ImportarModulo vbp, ruta, "01_Mod_Sistema.bas", "Mod_Sistema"

    Application.StatusBar = "BajaTax: Instalando Mod_ImportarArchivos..."
    ImportarModulo vbp, ruta, "02_Mod_ImportarArchivos.bas", "Mod_ImportarArchivos"

    Application.StatusBar = "BajaTax: Instalando WhatsApp..."
    ImportarModulo vbp, ruta, "03_Mod_WhatsApp.bas", "WhatsApp"

    Application.StatusBar = "BajaTax: Instalando PDF..."
    ImportarModulo vbp, ruta, "04_Mod_PDF.bas", "PDF"

    Application.StatusBar = "BajaTax: Instalando Mod_MasivoPDF..."
    ImportarModulo vbp, ruta, "07_Mod_MasivoPDF.bas", "Mod_MasivoPDF"

    Application.StatusBar = "BajaTax: Instalando Mod_BuscadorCliente..."
    ImportarModulo vbp, ruta, "08_Mod_BuscadorCliente.bas", "Mod_BuscadorCliente"

    ' Nota: Mod_ReportesCXC eliminado en v4 FINAL

    ' --- 4. Modulos de hoja (eventos Worksheet) ---
    Application.StatusBar = "BajaTax: Actualizando hoja OPERACIONES..."
    ActualizarCodigoHoja vbp, "OPERACIONES", ruta & "05_Hoja_OPERACIONES.bas"

    Application.StatusBar = "BajaTax: Actualizando hoja DIRECTORIO..."
    ActualizarCodigoHoja vbp, "DIRECTORIO", ruta & "06_Hoja_DIRECTORIO.bas"

    Application.StatusBar = "BajaTax: Actualizando hoja REGISTROS..."
    ActualizarCodigoHoja vbp, "REGISTROS", ruta & "11_Hoja_REGISTROS.bas"

    Application.StatusBar = "BajaTax: Actualizando hoja BUSCADOR CLIENTE..."
    ActualizarCodigoHoja vbp, "BUSCADOR CLIENTE", ruta & "10_Hoja_BuscadorCliente.bas"

    ' --- 5. Inicializar botones Z en todas las hojas ---
    Application.StatusBar = "BajaTax: Inicializando botones..."
    On Error Resume Next

    Set wsTemp = Nothing
    Set wsTemp = ThisWorkbook.Worksheets("OPERACIONES")
    If Not wsTemp Is Nothing Then
        Application.Run "'" & ThisWorkbook.Name & "'!InicializarBotonesZ"
    End If

    Set wsTemp = Nothing
    Set wsTemp = ThisWorkbook.Worksheets("DIRECTORIO")
    If Not wsTemp Is Nothing Then
        Application.Run "'" & ThisWorkbook.Name & "'!InicializarBotonesZ_DIRECTORIO"
    End If

    Set wsTemp = Nothing
    Set wsTemp = ThisWorkbook.Worksheets("REGISTROS")
    If Not wsTemp Is Nothing Then
        Application.Run "'" & ThisWorkbook.Name & "'!InicializarBotonesZ_REGISTROS"
    End If

    Set wsTemp = Nothing
    Set wsTemp = ThisWorkbook.Worksheets("BUSCADOR CLIENTE")
    If Not wsTemp Is Nothing Then
        Application.Run "'" & ThisWorkbook.Name & "'!InicializarHojaBuscador"
    End If

    Err.Clear
    On Error GoTo 0

    ' --- 6. Auto-eliminar este modulo Bootstrap ---
    Application.StatusBar = "BajaTax: Finalizando instalacion..."
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("Bootstrap_Installer")
    On Error GoTo 0

    Application.StatusBar = False

    MsgBox "INSTALACION COMPLETADA" & vbCrLf & vbCrLf & _
           "BajaTax v4 FINAL instalado correctamente." & vbCrLf & vbCrLf & _
           "SIGUIENTE PASO:" & vbCrLf & _
           "Archivo > Guardar Como > AUTOMATIZACION_v4_FINAL.xlsm" & vbCrLf & _
           "(formato: Libro de Excel habilitado para macros)", _
           vbInformation + vbOKOnly, "BajaTax v4 Instalado"
End Sub


'--------------------------------------------------------------------------
' ImportarModulo: Importa un archivo .bas como modulo estandar
'--------------------------------------------------------------------------
Private Sub ImportarModulo(vbp As Object, ruta As String, _
                           archivo As String, nombre As String)
    Dim comp As Object

    ' Eliminar modulo existente con el mismo nombre
    On Error Resume Next
    Set comp = vbp.VBComponents(nombre)
    If Not comp Is Nothing Then
        vbp.VBComponents.Remove comp
        Application.Wait Now + TimeValue("00:00:01")
    End If
    Set comp = Nothing
    Err.Clear
    On Error GoTo 0

    ' Verificar que el archivo existe
    If Dir(ruta & archivo) = "" Then
        MsgBox "Archivo no encontrado: " & archivo & vbCrLf & _
               "Verifica la carpeta VBA_CODIGO/", vbExclamation, "Archivo faltante"
        Exit Sub
    End If

    ' Importar modulo
    On Error Resume Next
    vbp.VBComponents.Import ruta & archivo
    If Err.Number <> 0 Then
        MsgBox "Error importando " & archivo & ":" & vbCrLf & Err.Description, _
               vbExclamation, "Error de importacion"
        Err.Clear
    End If
    On Error GoTo 0
End Sub


'--------------------------------------------------------------------------
' ActualizarCodigoHoja: Inyecta codigo .bas en el modulo de una hoja
'--------------------------------------------------------------------------
Private Sub ActualizarCodigoHoja(vbp As Object, nombreHoja As String, _
                                  rutaArchivo As String)
    Dim ws As Worksheet
    Dim wsFound As Worksheet
    Dim nNum As Integer
    Dim contenido As String
    Dim linea As String
    Dim saltarAtributos As Boolean

    ' Buscar la hoja por nombre visible
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nombreHoja Then
            Set wsFound = ws
            Exit For
        End If
    Next ws

    If wsFound Is Nothing Then
        MsgBox "La hoja '" & nombreHoja & "' no existe en este libro." & vbCrLf & _
               "Saltando codigo de hoja.", vbExclamation, "Hoja no encontrada"
        Exit Sub
    End If

    ' Verificar que el archivo .bas existe
    If Dir(rutaArchivo) = "" Then
        MsgBox "Archivo no encontrado: " & rutaArchivo, vbExclamation, "Archivo faltante"
        Exit Sub
    End If

    ' Leer contenido del archivo, saltando lineas Attribute del encabezado
    nNum = FreeFile
    contenido = ""
    saltarAtributos = True

    Open rutaArchivo For Input As #nNum
    Do While Not EOF(nNum)
        Line Input #nNum, linea
        If saltarAtributos Then
            If Left(Trim(linea), 9) <> "Attribute" Then
                saltarAtributos = False
            End If
        End If
        If Not saltarAtributos Then
            contenido = contenido & linea & vbLf
        End If
    Loop
    Close #nNum

    ' Reemplazar todo el codigo en el modulo de la hoja
    On Error Resume Next
    With vbp.VBComponents(wsFound.CodeName).CodeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
        If Len(Trim(contenido)) > 0 Then
            .AddFromString contenido
        End If
    End With
    If Err.Number <> 0 Then
        MsgBox "Error actualizando hoja " & nombreHoja & ":" & vbCrLf & _
               Err.Description, vbExclamation, "Error"
        Err.Clear
    End If
    On Error GoTo 0
End Sub
