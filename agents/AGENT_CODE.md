# AGENT_CODE.md — bajatax-code · Programador BajaTax

> Actívate cuando la tarea sea generar o editar un módulo `.bas` o script `.py`.
> Lee este archivo + `skills/vba-excel/SKILL.md` + `.roo/rules/05-hojas-excel.md` antes de escribir una sola línea.
> Las Reglas Inquebrantables están en CLAUDE.md — no se repiten aquí, pero se aplican sin excepción.

---

## Antes de escribir código

**Checklist de entrada — responde estas preguntas:**

1. ¿Qué módulo estoy generando? ¿Cuál es su única responsabilidad?
2. ¿Qué módulos necesita importar (dependencias)?
3. ¿Qué columnas del Excel toca? → verificar nombre exacto en `05-hojas-excel.md`
4. ¿Tiene activadores de doble clic? → usar `Worksheet_BeforeDoubleClick`, no botones
5. ¿Modifica celdas? → `EnableEvents = False` antes, `True` después
6. ¿Abre Workbooks externos? → cerrar en `ErrorHandler` siempre

Si no puedes responder alguna → usar modo `bajatax-architect` primero.

---

## Estructura estándar de cada módulo .bas

```vba
Attribute VB_Name = "Mod_NombreModulo"
Option Explicit

' ============================================================
' Mod_NombreModulo.bas — BajaTax v7
' Responsabilidad: [UNA sola frase aquí]
' Dependencias: Mod_Sistema
' Autor: BajaTax-Bot
' ============================================================

' [Constantes del módulo si aplica]
' Private Const C_NOMBRE As String = "valor"

' -------------------------------------------------------
' Sub principal pública
' -------------------------------------------------------
Public Sub NombrePrincipal()
    On Error GoTo ErrorHandler
    
    ' Variables
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    
    ' Validar modo antes de cualquier envío (si aplica)
    ' [lógica]
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' [lógica principal]
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    LogEvento "INFO", "Mod_NombreModulo.NombrePrincipal", "Completado OK"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    ' Cerrar cualquier Workbook externo abierto:
    ' If Not wbExterno Is Nothing Then wbExterno.Close False
    LogEvento "ERROR", "Mod_NombreModulo.NombrePrincipal", Err.Number & ": " & Err.Description
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "BajaTax"
End Sub
```

---

## Reglas de generación de código

### Hojas — siempre por nombre exacto

```vba
' ✅ Correcto
Dim wsOps As Worksheet
Set wsOps = ThisWorkbook.Sheets("OPERACIONES")

' ✅ Correcto con validación
On Error Resume Next
Set wsOps = ThisWorkbook.Sheets("OPERACIONES")
On Error GoTo 0
If wsOps Is Nothing Then
    MsgBox "Hoja OPERACIONES no encontrada", vbCritical
    Exit Sub
End If

' ❌ Nunca por índice
Set wsOps = ThisWorkbook.Sheets(3)
```

### Columnas — siempre por header, nunca por número

```vba
' ✅ Usar Util_GetColumnByHeader de Mod_Sistema:
Dim colRFC As Long
colRFC = Util_GetColumnByHeader(wsOps, "RFC")
If colRFC = -1 Then
    MsgBox "Columna RFC no encontrada en OPERACIONES", vbCritical
    Exit Sub
End If

' ❌ Nunca:
ws.Cells(i, 5).Value  ' ¿qué es la columna 5?
```

### Fórmulas en Excel — siempre @TODAY() en Mac

```vba
' Estatus (Col.I de OPERACIONES) — fórmula correcta para Mac:
ws.Cells(2, colEstatus).Formula = _
    "=IF(D2="""","""",IF(L2<>"""",""PAGADO"",IF(J2="""",""PENDIENTE""," & _
    "IF(@TODAY()=J2,""HOY VENCE"",IF(@TODAY()>J2,""VENCIDO"",""PENDIENTE"")))))"

' Días vencidos (Col.K):
ws.Cells(2, colDias).Formula = "=IF(J2="""","""",@TODAY()-J2)"
```

### Validación de RFC

```vba
Function Util_ValidarRFC(rfc As String) As Boolean
    ' RFC persona moral: 3 letras + 6 dígitos + 3 alfanuméricos = 12 chars
    ' RFC persona física: 4 letras + 6 dígitos + 3 alfanuméricos = 13 chars
    rfc = Trim(UCase(rfc))
    If Len(rfc) < 12 Or Len(rfc) > 13 Then
        Util_ValidarRFC = False
        Exit Function
    End If
    ' Patrón básico: letras al inicio, dígitos en medio
    Dim i As Integer
    For i = 1 To Len(rfc) - 9  ' primeras 3-4 son letras
        If Not (Mid(rfc, i, 1) Like "[A-Z]") Then
            Util_ValidarRFC = False
            Exit Function
        End If
    Next i
    Util_ValidarRFC = True
End Function
```

### Normalización de texto (para detección y comparación)

```vba
Function Util_Normalizar(texto As String) As String
    ' Trim + minúsculas + sin acentos + sin puntuación
    Dim s As String
    s = Trim(LCase(texto))
    ' Remover acentos
    s = Replace(s, "á", "a"): s = Replace(s, "é", "e")
    s = Replace(s, "í", "i"): s = Replace(s, "ó", "o")
    s = Replace(s, "ú", "u"): s = Replace(s, "ü", "u")
    s = Replace(s, "ñ", "n"): s = Replace(s, "ã", "a")
    ' Remover puntuación común
    s = Replace(s, ".", ""): s = Replace(s, ",", "")
    s = Replace(s, "_", " "): s = Replace(s, "-", " ")
    Util_Normalizar = s
End Function
```

### Detección de tipo por contenido (motor de importación)

```vba
' Patrones para detectar qué tipo de dato hay en una columna:

' RFC: 12-13 chars, patrón de letras+dígitos
Function EsColumnaRFC(col As Range) As Boolean
    Dim hits As Integer, total As Integer
    Dim cell As Range
    For Each cell In col.Cells
        If cell.Value <> "" Then
            total = total + 1
            Dim v As String: v = Trim(CStr(cell.Value))
            If Len(v) >= 12 And Len(v) <= 13 Then
                If v Like "[A-Za-z][A-Za-z][A-Za-z]*[0-9][0-9][0-9][0-9][0-9][0-9]*" Then
                    hits = hits + 1
                End If
            End If
        End If
        If total >= 10 Then Exit For  ' Muestra de 10 celdas
    Next cell
    EsColumnaRFC = (total > 0 And hits / total >= 0.5)
End Function

' Email: contiene @ y punto después del @
Function EsColumnaEmail(col As Range) As Boolean
    Dim hits As Integer, total As Integer
    Dim cell As Range
    For Each cell In col.Cells
        If cell.Value <> "" Then
            total = total + 1
            If InStr(cell.Value, "@") > 0 And InStr(cell.Value, ".") > InStr(cell.Value, "@") Then
                hits = hits + 1
            End If
        End If
        If total >= 10 Then Exit For
    Next cell
    EsColumnaEmail = (total > 0 And hits / total >= 0.5)
End Function

' Teléfono: 10 dígitos (con o sin prefijo)
Function EsColumnaTelefono(col As Range) As Boolean
    Dim hits As Integer, total As Integer
    Dim cell As Range
    For Each cell In col.Cells
        If cell.Value <> "" Then
            total = total + 1
            Dim cleaned As String
            cleaned = Replace(Replace(Replace(Replace(CStr(cell.Value), " ", ""), "-", ""), "(", ""), ")", "")
            cleaned = Replace(Replace(cleaned, "+52", ""), "52", "")
            If Len(cleaned) = 10 And IsNumeric(cleaned) Then hits = hits + 1
        End If
        If total >= 10 Then Exit For
    Next cell
    EsColumnaTelefono = (total > 0 And hits / total >= 0.5)
End Function
```

---

## Diccionario de aliases (mapeo de columnas)

Cuando un archivo tiene headers, normalizar y comparar contra este diccionario:

```vba
' En 02_Mod_ImportarArchivos.bas — función de mapeo
Function MapearHeader(header As String) As String
    Dim h As String
    h = Util_Normalizar(header)
    
    ' → RESPONSABLE
    If h = "responsable" Or h = "asesor" Or h = "vendedor" Or h = "agente" Then
        MapearHeader = "RESPONSABLE": Exit Function
    End If
    ' → RFC
    If h = "rfc" Or h = "r.f.c" Or h = "registro federal" Then
        MapearHeader = "RFC": Exit Function
    End If
    ' → CLIENTE / NOMBRE
    If h = "cliente" Or h = "nombre" Or h = "razon social" Or h = "razon_social" _
       Or h = "contribuyente" Or h = "nombre del contribuyente" Or h = "nombre contribuyente" Then
        MapearHeader = "NOMBRE DEL CONTRIBUYENTE": Exit Function
    End If
    ' → EMAIL
    If h = "email" Or h = "correo" Or h = "mail" Or h = "e-mail" Or h = "correo electronico" Then
        MapearHeader = "EMAIL": Exit Function
    End If
    ' → TELÉFONO
    If h = "telefono" Or h = "tel" Or h = "celular" Or h = "movil" Or h = "whatsapp" Then
        MapearHeader = "TELÉFONO": Exit Function
    End If
    ' → FECHA
    If h = "fecha" Or h = "fecha cobranza" Or h = "fecha emision" Or h = "fecha factura" Then
        MapearHeader = "FECHA": Exit Function
    End If
    ' → CONCEPTO
    If h = "concepto" Or h = "descripcion" Or h = "servicio" Or h = "detalle" Or h = "descripcion servicio" Then
        MapearHeader = "CONCEPTO": Exit Function
    End If
    ' → MONTO
    If h = "monto" Or h = "importe" Or h = "total" Or h = "cantidad" Or h = "valor" Or h = "precio" Then
        MapearHeader = "MONTO": Exit Function
    End If
    ' → FACTURA / ID
    If h = "factura" Or h = "folio" Or h = "uuid" Or h = "id factura" Or h = "no factura" Then
        MapearHeader = "FACTURA": Exit Function
    End If
    ' → RÉGIMEN
    If h = "regimen" Or h = "regimen fiscal" Or h = "tipo persona" Or h = "tipo" Then
        MapearHeader = "REGIMEN": Exit Function
    End If
    ' → VENCIMIENTO
    If h = "vencimiento" Or h = "fecha vencimiento" Or h = "fecha limite" Or h = "fecha pago" Then
        MapearHeader = "VENCIMIENTO": Exit Function
    End If
    
    MapearHeader = ""  ' No reconocido
End Function
```

---

## Deduplicación — clave compuesta

```vba
' Clave única: RFC + Concepto + FechaCobranza + Monto
' Usada tanto en REGISTROS→OPERACIONES como en importación
Function Util_ClaveDedup(rfc As String, concepto As String, fecha As String, monto As String) As String
    Util_ClaveDedup = UCase(Trim(rfc)) & "|" & _
                      LCase(Trim(concepto)) & "|" & _
                      Trim(fecha) & "|" & _
                      Format(CDbl(IIf(IsNumeric(monto), monto, 0)), "0.00")
End Function

' Uso: verificar antes de insertar en OPERACIONES
' Si la clave ya existe → preguntar al usuario antes de duplicar
```

---

## Nunca hacer esto

```vba
' ❌ Referencia a columna por número:
ws.Cells(i, 5).Value = rfc

' ❌ TODAY() en Mac:
ws.Range("K2").Formula = "=TODAY()-J2"

' ❌ Ruta hardcodeada:
path = "/Users/javieravila/Documents/BajaTax/OUTPUT"

' ❌ Sub sin ErrorHandler:
Public Sub HacerAlgo()
    ' ... código sin On Error GoTo ErrorHandler
End Sub

' ❌ Escribir a celda sin apagar eventos:
ws.Cells(2, 9).Value = "PAGADO"  ' sin EnableEvents = False antes

' ❌ Abrir Workbook sin cerrar en error:
Set wb = Workbooks.Open(path)
' ... si falla aquí, wb queda abierto y Excel se cuelga
```
