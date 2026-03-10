---
name: cross-platform
description: >
  Usar esta skill cuando código VBA o Python necesite funcionar tanto en Mac como
  en Windows. Activar cuando el usuario mencione "funciona en Mac pero no en Windows",
  "error de ruta", "PathSeparator", "Shell en Mac", "el archivo no se encuentra",
  "PDF no genera en Mac", "TODAY() da error", o cuando se escriba cualquier código
  con rutas de archivo, comandos Shell, o funciones de fecha en VBA. También activar
  si algo funcionaba antes y dejó de funcionar al cambiar de computadora.
  Esta skill previene los 5 errores de compatibilidad más comunes de BajaTax.
---

# Cross-Platform Mac ↔ Windows — BajaTax

## Los 5 puntos de falla garantizados

### 1. Rutas de archivo
```vba
' ❌ Rompe en Mac:
ruta = "C:\Users\Javi\Desktop\vscode\BajataxV2\OUTPUT\"

' ❌ Rompe en Mac (barra al revés):
ruta = ThisWorkbook.Path & "\" & "OUTPUT"

' ✅ Funciona en ambos:
ruta = ThisWorkbook.Path & Application.PathSeparator & "OUTPUT"

' ✅ Rutas anidadas:
ruta = ThisWorkbook.Path & Application.PathSeparator & _
       "OUTPUT" & Application.PathSeparator & Format(Date, "DDMMYYYY")
```

### 2. Detectar OS en runtime
```vba
Function Util_IsMac() As Boolean
    Util_IsMac = (Application.OperatingSystem Like "*Mac*")
End Function
```

### 3. Shell commands — completamente distintos por OS
```vba
Sub Util_RunPython(scriptPath As String, args As String)
    On Error GoTo ErrorHandler
    Dim cmd As String
    If Util_IsMac() Then
        cmd = "python3 """ & scriptPath & """ " & args
        Shell "bash -c '" & cmd & "'"
    Else
        cmd = "python """ & scriptPath & """ " & args
        Shell "cmd /c " & cmd
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Error ejecutando Python: " & Err.Description, vbCritical
End Sub
```

### 4. Fechas — @TODAY() obligatorio en Mac
```vba
' ❌ En Mac falla o da resultado incorrecto:
ws.Range("K2").Formula = "=TODAY()-J2"

' ✅ Funciona en Mac y Windows:
ws.Range("K2").Formula = "=@TODAY()-J2"

' ✅ Fecha actual en código VBA (no fórmula) — igual en ambos OS:
Dim hoy As Date: hoy = Date
Dim ahora As Date: ahora = Now()
```

### 5. Nombres de archivo — sanitizar para ambos OS
```vba
Function Util_SanitizeFileName(nombre As String) As String
    Dim r As String: r = nombre
    r = Replace(r, "á","a"): r = Replace(r, "Á","A")
    r = Replace(r, "é","e"): r = Replace(r, "É","E")
    r = Replace(r, "í","i"): r = Replace(r, "Í","I")
    r = Replace(r, "ó","o"): r = Replace(r, "Ó","O")
    r = Replace(r, "ú","u"): r = Replace(r, "Ú","U")
    r = Replace(r, "ñ","n"): r = Replace(r, "Ñ","N")
    r = Replace(r, " ","_")
    r = Replace(r, "/","-"): r = Replace(r, "\","-")
    r = Replace(r, ":","_"): r = Replace(r, "*","_")
    Util_SanitizeFileName = r
End Function

' Uso para PDF:
nombrePDF = "EdoCuenta_" & Util_SanitizeFileName(cliente) & _
            "_" & Format(Date,"DDMMYYYY") & ".pdf"
```

---

## Verificar y crear carpetas

```vba
Sub Util_EnsureFolderExists(folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
End Sub

' Abrir archivo externo de forma segura:
Function Util_OpenWorkbook(filePath As String) As Workbook
    On Error GoTo ErrorHandler
    If Dir(filePath) = "" Then
        MsgBox "Archivo no encontrado:" & vbLf & filePath, vbCritical
        Set Util_OpenWorkbook = Nothing
        Exit Function
    End If
    Set Util_OpenWorkbook = Workbooks.Open(filePath, ReadOnly:=False)
    Exit Function
ErrorHandler:
    MsgBox "Error abriendo archivo: " & Err.Description, vbCritical
    Set Util_OpenWorkbook = Nothing
End Function
```

---

## Ejecutar pdf_server.py correctamente

```vba
Sub Util_GeneratePDF(jsonPath As String, expectedPDFPath As String)
    On Error GoTo ErrorHandler
    
    Dim scriptPath As String
    scriptPath = ThisWorkbook.Path & Application.PathSeparator & _
                 "src" & Application.PathSeparator & _
                 "python" & Application.PathSeparator & "pdf_server.py"
    
    If Dir(scriptPath) = "" Then
        MsgBox "pdf_server.py no encontrado en: " & scriptPath, vbCritical
        Exit Sub
    End If
    
    If Util_IsMac() Then
        Shell "bash -c 'python3 " & Chr(34) & scriptPath & Chr(34) & _
              " " & Chr(34) & jsonPath & Chr(34) & "'"
    Else
        Shell "cmd /c python " & Chr(34) & scriptPath & Chr(34) & _
              " " & Chr(34) & jsonPath & Chr(34)
    End If
    
    ' Esperar hasta 30 segundos a que aparezca el PDF
    Dim timeout As Date: timeout = Now() + TimeValue("00:00:30")
    Do While Dir(expectedPDFPath) = "" And Now() < timeout
        Application.Wait Now() + TimeValue("00:00:01")
        DoEvents
    Loop
    
    If Dir(expectedPDFPath) = "" Then
        MsgBox "Timeout: PDF no generado en 30 segundos", vbCritical
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error ejecutando pdf_server.py: " & Err.Description, vbCritical
End Sub
```

---

## Checklist antes de entregar código

- [ ] Rutas usan `Application.PathSeparator`
- [ ] Fórmulas de fecha usan `@TODAY()`
- [ ] Comandos Shell detectan OS con `Util_IsMac()`
- [ ] Nombres de archivo pasan por `Util_SanitizeFileName()`
- [ ] Existencia de archivo verificada con `Dir()` antes de abrir
- [ ] Carpetas de salida creadas con `Util_EnsureFolderExists()`
