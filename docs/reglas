# 01-sistema-global.md — Reglas del Núcleo del Sistema

> Roo Code: lee estas reglas SIEMPRE que trabajes en cualquier módulo de BajaTax.

## Rutas y Separadores

- SIEMPRE usar `ThisWorkbook.Path & Application.PathSeparator` para construir rutas
- NUNCA escribir rutas absolutas como `/Users/javieravila/...` o `C:\Users\...`
- Detectar OS: `Application.OperatingSystem` contiene "Mac" o "Windows"
- Separador: Mac = `/`, Windows = `\` → usar `Application.PathSeparator`
- Carpetas del sistema (relativas a ThisWorkbook.Path):
  - `OUTPUT/` → PDFs generados
  - `LOGS/` → bitácora del sistema
  - `ARCHIVOS_ENTRADA/` → archivos a importar
  - `TEMP/` → archivos temporales (JSON para pdf_server.py)

## Manejo de Errores

Cada Sub y Function DEBE tener esta estructura:

```vba
Sub NombreDeLaSub()
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' --- Código principal aquí ---
    
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    LogEvento "ERROR", "NombreDeLaSub", Err.Description & " (Línea: " & Erl & ")"
    MsgBox "Error en NombreDeLaSub: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

**Reglas estrictas:**
- `On Error GoTo ErrorHandler` en la PRIMERA línea después de la declaración
- `Application.EnableEvents = False` antes de modificar celdas, restaurar en CleanExit Y en ErrorHandler
- `Application.ScreenUpdating = False` para rendimiento, restaurar igual
- Si hay Workbooks abiertos → cerrarlos en ErrorHandler antes del MsgBox
- NUNCA usar `On Error Resume Next` excepto para verificar si un objeto existe (y restaurar inmediatamente después)

## Fechas en Mac

- SIEMPRE usar `@TODAY()` en fórmulas de celdas, NUNCA `TODAY()`
- En código VBA usar `Date` (equivalente a Today en VBA, funciona en Mac)
- Timestamps: `Now()` funciona correctamente en Mac
- Formato display: `dd/mm/yyyy` para interfaz, ISO `yyyy-mm-dd` para logs

## Logger Central

Toda acción importante debe registrarse llamando a:

```vba
Sub LogEvento(tipo As String, origen As String, detalle As String)
    ' tipo: "INFO", "ERROR", "ENVIO", "IMPORTACION", "PDF", "PAGO"
    ' origen: nombre del Sub/Function que llama
    ' detalle: descripción del evento
End Sub
```

Destinos del log:
- Hoja LOG ENVIOS → para eventos de WhatsApp
- Archivo LOGS/sistema.log → para eventos generales (importaciones, errores, PDFs)

## Constantes

NO hardcodear valores. Leer de CONFIGURACION:
- Modo del sistema → `Sheets("CONFIGURACION").Range("B2").Value`
- Nombre despacho → B5, Beneficiario → B6, Banco → B7, CLABE → B8
- Teléfono despacho → B9, Email → B10, Departamento → B12
- Número prueba → B14, Ruta logo → B16

## Referencia a Columnas

SIEMPRE buscar columnas por header, NUNCA por índice:

```vba
Function GetColByHeader(ws As Worksheet, headerName As String) As Long
    Dim col As Long
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If Trim(UCase(ws.Cells(1, col).Value)) = UCase(headerName) Then
            GetColByHeader = col
            Exit Function
        End If
    Next col
    GetColByHeader = 0  ' No encontrado
End Function
```

## Option Explicit

TODOS los módulos deben empezar con `Option Explicit`. Sin excepción.

## Convenciones de Nombres

- Variables: camelCase en inglés → `clientName`, `totalAmount`
- Constantes: prefijo C_ → `C_MODO_PRUEBA`, `C_MAX_INTENTOS`
- Funciones utilitarias: prefijo Util_ → `Util_GetColumnByHeader`
- Subs de eventos: prefijo Event_ → `Event_DoubleClickPago`
- Comentarios: en español
