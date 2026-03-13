Sub InstalarTodo()
    Dim vbp As Object
    Set vbp = ThisWorkbook.VBProject
    Dim baseDir As String: baseDir = "C:/Users/LENOVO/Desktop/vscode/bajatax/src/vba-modules/"

    ' Instalar Mod_Sistema
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("Mod_Sistema")
    On Error GoTo 0
    vbp.VBComponents.Import "C:/Users/LENOVO/Desktop/vscode/bajatax/src/vba-modules/01_Mod_Sistema.bas"

    ' Instalar Mod_ImportarArchivos
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("Mod_ImportarArchivos")
    On Error GoTo 0
    vbp.VBComponents.Import "C:/Users/LENOVO/Desktop/vscode/bajatax/src/vba-modules/02_Mod_ImportarArchivos.bas"

    ' Instalar WhatsApp
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("WhatsApp")
    On Error GoTo 0
    vbp.VBComponents.Import "C:/Users/LENOVO/Desktop/vscode/bajatax/src/vba-modules/03_Mod_WhatsApp.bas"

    ' Instalar PDF
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("PDF")
    On Error GoTo 0
    vbp.VBComponents.Import "C:/Users/LENOVO/Desktop/vscode/bajatax/src/vba-modules/04_Mod_PDF.bas"

    ' Instalar Mod_MasivoPDF
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("Mod_MasivoPDF")
    On Error GoTo 0
    vbp.VBComponents.Import "C:/Users/LENOVO/Desktop/vscode/bajatax/src/vba-modules/07_Mod_MasivoPDF.bas"

    ' Instalar Mod_BuscadorCliente
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("Mod_BuscadorCliente")
    On Error GoTo 0
    vbp.VBComponents.Import "C:/Users/LENOVO/Desktop/vscode/bajatax/src/vba-modules/08_Mod_BuscadorCliente.bas"

    ' Instalar Mod_FormatoGlobal
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("Mod_FormatoGlobal")
    On Error GoTo 0
    vbp.VBComponents.Import "C:/Users/LENOVO/Desktop/vscode/bajatax/src/vba-modules/09_Mod_FormatoGlobal.bas"

    MsgBox "MÃ³dulos instalados correctamente.", vbInformation, "BajaTax"
End Sub